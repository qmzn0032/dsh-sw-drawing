"""SW 桥的 MCP 服务端（stdio，一行一帧 JSON-RPC，零第三方依赖）。

由 dsh --profile acp 按 session/new 的 mcpServers 配置拉起；stdout 只走协议，日志走 stderr。

工具：
  sw_status           SW 是否在运行、当前文档是什么
  sw_part_summary     当前零件的名称/类型/路径/特征树/材料（只读）
  sw_set_material     改当前配置的材料（只改内存，不保存；重建后返回改前改后）
  sw_mass_properties  体积/质量/表面积/重心/包围盒（只读，建模后的自校验用）
  sw_rebuild          重建一次并回读体积与包围盒（只读：不改数据、不保存）
  sw_new              用 SW 模板新建空零件/装配体（写，走审批）
  sw_build            草图 → 拉伸/切除（写，走审批；一步失败整步撤销）

用法（手动调试）：
    echo {"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18"}} | python py/mcp_sw_server.py
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import time
import traceback
import urllib.request

import sw_com

SERVER_INFO = {"name": "sw-bridge", "version": "0.1.0"}

TOOLS = [
    {
        "name": "sw_status",
        "description": (
            "查询 SolidWorks 当前状态：是否在运行、打开的是哪个文档（标题/路径/类型/配置）。"
            "不启动 SW、不修改任何东西。"
        ),
        "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
    },
    {
        "name": "sw_part_summary",
        "description": (
            "读取当前活动文档的概况：名称、类型、路径、配置、特征树（名称+类型）、当前材料。"
            "只读，不修改模型。零件未打开或 SW 未运行时会在结果里说明原因。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "max_features": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 500,
                    "default": 60,
                    "description": "最多返回多少个特征（默认 60）",
                }
            },
            "additionalProperties": False,
        },
    },
    {
        "name": "sw_set_material",
        "description": (
            "把当前文档「当前配置」的材料改成指定名称（例如 45 钢 / 普通碳钢 / Alloy Steel）。"
            "只改内存并重建，不保存文件；返回改前/改后材料名。写操作，调用前应让用户确认。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "material": {"type": "string", "minLength": 1, "description": "目标材料名，例如 45 钢"},
                "database": {
                    "type": "string",
                    "description": "材料库（可选；留空沿用当前库，例如 SolidWorks Materials.sldmat）",
                    "default": "",
                },
            },
            "required": ["material"],
            "additionalProperties": False,
        },
    },
    {
        "name": "sw_mass_properties",
        "description": (
            "读活动零件/装配体的质量属性：体积（mm3）、质量（kg）、表面积（mm2）、密度、"
            "重心（mm）和包围盒（mm）。只读，不改模型。建模后用它对账："
            "体积和包围盒跟你算的对不上，就是建错了。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc": {
                    "type": "string",
                    "description": "已打开文档的标题 / 全路径 / 文件名；留空用活动文档",
                    "default": "",
                }
            },
            "additionalProperties": False,
        },
    },
    {
        "name": "sw_rebuild",
        "description": (
            "重建一次当前文档，返回重建结果、特征树（名称+类型）以及体积/包围盒。"
            "不改数据、不保存。注意：逐特征的错误码本版还没接。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc": {
                    "type": "string",
                    "description": "已打开文档的标题 / 全路径 / 文件名；留空用活动文档",
                    "default": "",
                }
            },
            "additionalProperties": False,
        },
    },
    {
        "name": "sw_new",
        "description": (
            "用 SOLIDWORKS 模板新建一个空零件（或装配体）并激活它。只在内存里，"
            "不保存、不覆盖任何文件；保存与否由用户在 SW 里决定。"
            "新建后用 sw_build 建模。写操作，调用前应让用户确认。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "type": {
                    "type": "string",
                    "enum": ["part", "assembly"],
                    "default": "part",
                    "description": "默认 part",
                },
                "template": {
                    "type": "string",
                    "description": (
                        "模板全路径（可选）。留空则用 SW 选项里的默认模板，"
                        "再退到安装目录里的 gb_part.prtdot / gb_assembly.asmdot"
                    ),
                    "default": "",
                },
            },
            "additionalProperties": False,
        },
    },
    {
        "name": "sw_build",
        "description": (
            "在零件上建模：按顺序跑 sketch（在基准面上画矩形/圆/线/多段线/弧）"
            "和 extrude（拉伸加料）或 cut（切除）。坐标一律是**模型空间 mm**，"
            "草图平面可选 front/top/right 或基准面名字。"
            "默认 atomic：某一步失败就把这次已经建出来的特征全删掉，零件回到调用前的样子。"
            "返回每个新特征的名称、体积、质量、包围盒——拿它跟自己算的数对账。"
            "写操作，调用前应让用户确认。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc": {
                    "type": "string",
                    "description": "已打开零件的标题 / 全路径 / 文件名；留空用活动文档",
                    "default": "",
                },
                "steps": {
                    "type": "array",
                    "maxItems": 60,
                    "description": "按顺序执行的建模步骤，最多 60 步",
                    "items": {
                        "type": "object",
                        "properties": {
                            "op": {
                                "type": "string",
                                "enum": ["sketch", "extrude", "cut"],
                                "description": "sketch 画草图；extrude 加料；cut 切除",
                            },
                            "on": {
                                "type": "string",
                                "description": "sketch：front / top / right 或基准面名字。默认 front",
                                "default": "front",
                            },
                            "entities": {
                                "type": "array",
                                "description": "sketch：画什么。闭合环成轮廓，环套环成孔",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "type": {
                                            "type": "string",
                                            "enum": ["line", "centerline", "rectangle", "circle", "arc", "polyline"],
                                        },
                                        "from": {
                                            "type": "array",
                                            "items": {"type": "number"},
                                            "description": "line / centerline / arc：起点 [x, y] 或 [x, y, z]（mm）",
                                        },
                                        "to": {
                                            "type": "array",
                                            "items": {"type": "number"},
                                            "description": "line / centerline / arc：终点",
                                        },
                                        "through": {
                                            "type": "array",
                                            "items": {"type": "number"},
                                            "description": "arc：弧上任意一点",
                                        },
                                        "corner": {
                                            "type": "array",
                                            "items": {"type": "number"},
                                            "description": "rectangle：一个角",
                                        },
                                        "opposite": {
                                            "type": "array",
                                            "items": {"type": "number"},
                                            "description": "rectangle：对角",
                                        },
                                        "center": {
                                            "type": "array",
                                            "items": {"type": "number"},
                                            "description": "circle：圆心",
                                        },
                                        "diameter": {"type": "number", "description": "circle：直径 mm"},
                                        "radius": {"type": "number", "description": "circle：半径 mm"},
                                        "points": {
                                            "type": "array",
                                            "items": {"type": "array", "items": {"type": "number"}},
                                            "description": "polyline：依次的拐点",
                                        },
                                        "closed": {
                                            "type": "boolean",
                                            "description": "polyline：末点连回起点，默认 true",
                                            "default": True,
                                        },
                                    },
                                    "required": ["type"],
                                    "additionalProperties": False,
                                },
                            },
                            "sketch": {
                                "type": "string",
                                "description": "extrude / cut：用哪个草图，默认用最后一个草图",
                                "default": "",
                            },
                            "depth": {
                                "type": "number",
                                "description": "extrude / cut：深度 mm（blind 必填）。mid_plane 时是总厚度",
                                "default": 10.0,
                            },
                            "end": {
                                "type": "string",
                                "enum": ["blind", "through_all", "through_all_both", "mid_plane"],
                                "default": "blind",
                                "description": "终止条件。默认 blind（要 depth）",
                            },
                            "cut": {
                                "type": "boolean",
                                "description": "extrude 步骤里写 cut=true 等价于 op=cut",
                                "default": False,
                            },
                            "reverse": {
                                "type": "boolean",
                                "description": "反方向。默认沿草图法线（基准面外法线为正）",
                                "default": False,
                            },
                            "merge": {
                                "type": "boolean",
                                "description": "拉伸是否并入已有实体，默认 true",
                                "default": True,
                            },
                            "draft": {
                                "type": "number",
                                "description": "拔模角（度），默认 0。⚠ 本机未验证，可能被忽略",
                                "default": 0.0,
                            },
                        },
                        "required": ["op"],
                        "additionalProperties": False,
                    },
                },
                "atomic": {
                    "type": "boolean",
                    "description": "默认 true：某一步失败就撤掉本次调用已经建出来的特征",
                    "default": True,
                },
            },
            "required": ["steps"],
            "additionalProperties": False,
        },
    },
]


# ---------------------------------------------------------------- 并发骨架
# 主线程只做「读 stdin + 处理轻量请求（ping / tools/list）」，所有工具调用丢给
# 一条工作线程串行执行。这样确认框弹着的时候，服务端对 ping 依然有响应——
# 之前是单线程 for line in stdin，弹窗期间整个服务端是死的。
_OUT_LOCK = threading.Lock()
_JOB_Q: "queue.Queue[dict | None]" = queue.Queue()
_APPROVE_ALL = False  # 本进程内「全部允许」（重启即失效）


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def send(payload: dict) -> None:
    # 工作线程和主线程都会写 stdout，必须串行，否则两行 JSON 会粘成一行
    line = json.dumps(payload, ensure_ascii=False) + "\n"
    with _OUT_LOCK:
        sys.stdout.write(line)
        sys.stdout.flush()


def ok(msg_id, result: dict) -> None:
    send({"jsonrpc": "2.0", "id": msg_id, "result": result})


def fail(msg_id, code: int, message: str) -> None:
    send({"jsonrpc": "2.0", "id": msg_id, "error": {"code": code, "message": message}})


def tool_text(text: str, is_error: bool = False) -> dict:
    return {"content": [{"type": "text", "text": text}], "isError": is_error}


TOOLS.extend(
    [
        {
            "name": "sw_drawing_build",
            "description": (
                "从零件出一张 GB 工程图：GB 图框（a3-gb 等）→ 第一角三视图 → "
                "按 GB/T 14690 自动选比例 → 存 .SLDDRW + 导 PDF。返回图纸尺寸、比例、"
                "每个视图的位置/轮廓/尺寸、视图重叠与出界检查、文件字节数。"
                "只新建文件，不改零件。"
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "model": {"type": "string", "description": "零件全路径（.SLDPRT）"},
                    "outDir": {"type": "string", "description": "输出目录（不存在会创建）"},
                    "paper": {
                        "type": "string",
                        "enum": ["A4", "A3", "A2", "A1", "A0"],
                        "default": "A3",
                        "description": "GB/T 14689 图幅",
                    },
                    "basename": {"type": "string", "description": "输出文件名主干，默认取零件名"},
                    "font": {
                        "type": "string",
                        "default": "汉仪长仿宋体",
                        "description": (
                            "图纸字体（按文档属性写：注释/尺寸/注解三项）。默认汉仪长仿宋体；"
                            "传空字符串表示不动字体"
                        ),
                    },
                    "keepOpen": {
                        "type": "boolean",
                        "default": True,
                        "description": "画完是否留在 SW 里打开",
                    },
                },
                "required": ["model", "outDir"],
            },
        },
        {
            "name": "sw_export",
            "description": (
                "把当前文档（默认活动文档）导出到指定路径。格式由扩展名决定："
                ".pdf / .dwg / .dxf / .bmp。返回是否成功与文件字节数。"
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "目标文件全路径，含扩展名"},
                    "doc": {
                        "type": "string",
                        "description": "文档标题/全路径/文件名；留空用活动文档",
                    },
                },
                "required": ["path"],
            },
        },
    ]
)


TOOLS.extend(
    [
        {
            "name": "sw_face_sketch",
            "description": (
                "在**实体面上**画草图（不是基准面），可选直接切除或加料。"
                "面用「法向 + 一个点」定位：faceNormal 是面法向，facePoint 是该面上的任意一点。"
                "坐标一律模型空间 mm。切除后返回体积/质量属性用于对账。"
                "注意：草图轮廓不要与实体边正好相切（差 0.1mm 即可），否则 SW 会拒绝出特征。"
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "doc": {"type": "string", "description": "文档标题/路径；留空用活动文档"},
                    "entities": {
                        "type": "array",
                        "description": "草图实体，字段与 sw_build 的 sketch 一致",
                        "items": {"type": "object"},
                    },
                    "faceNormal": {
                        "type": "array", "items": {"type": "number"},
                        "description": "面法向 [x,y,z]，例如顶面 [0,0,1]",
                    },
                    "facePoint": {
                        "type": "array", "items": {"type": "number"},
                        "description": "面上的点 [x,y,z]（mm），用来在同法向的多个面里挑最近的那个",
                    },
                    "mode": {
                        "type": "string", "enum": ["cut", "boss", "none"], "default": "cut",
                        "description": "画完草图后做什么：切除 / 加料 / 什么都不做",
                    },
                    "depthMm": {"type": "number", "default": 10},
                    "end": {
                        "type": "string",
                        "enum": ["blind", "through_all", "through_all_both", "mid_plane"],
                        "default": "through_all",
                    },
                    "reverse": {"type": "boolean", "default": False},
                },
                "required": ["entities"],
            },
        },
        {
            "name": "sw_pattern",
            "description": (
                "线性阵列。seeds 是特征名数组，direction 是方向向量 [x,y,z]。"
                "底层用「辅助草图直线」当方向参考（本机无法直接选边）。"
                "返回阵列特征名、辅助草图名，并用体积对账。"
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "doc": {"type": "string"},
                    "seeds": {"type": "array", "items": {"type": "string"},
                              "description": "要阵列的特征名，例如 [\"切除-拉伸1\"]"},
                    "direction": {"type": "array", "items": {"type": "number"},
                                  "description": "方向向量，例如 [1,0,0]"},
                    "spacingMm": {"type": "number", "description": "间距 mm"},
                    "count": {"type": "integer", "minimum": 2, "description": "总数（含种子）"},
                    "flip": {"type": "boolean", "default": False},
                    "geometryPattern": {"type": "boolean", "default": True},
                    "direction2": {"type": "array", "items": {"type": "number"},
                                   "description": "第二方向向量（二维阵列才给，例如 [0,1,0]）"},
                    "spacing2Mm": {"type": "number", "description": "第二方向间距 mm"},
                    "count2": {"type": "integer", "minimum": 2,
                               "description": "第二方向总数（给了 direction2 才生效）"},
                    "flip2": {"type": "boolean", "default": False},
                },
                "required": ["seeds", "direction", "spacingMm", "count"],
            },
        },
        {
            "name": "sw_rename",
            "description": "给特征起语义名（例如把 `切除-拉伸1` 改成 `功能牌1第三安装孔`）。返回改后特征树。",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "doc": {"type": "string"},
                    "old": {"type": "string", "description": "现名"},
                    "new": {"type": "string", "description": "新名"},
                },
                "required": ["old", "new"],
            },
        },
    ]
)


TOOLS.extend(
    [
        {
            "name": "sw_fillet",
            "description": (
                "圆角或倒角。edges 用「射线」指定：每项 {point:[x,y,z], dir:[dx,dy,dz], radiusMm?}，"
                "point 要落在实体外面朝目标边打（半径默认 0.5mm）。"
                "kind=fillet 用半径 sizeMm；kind=chamfer 用等距宽度 sizeMm（角度默认 45°）。"
                "返回特征名 + 体积/质量属性用于对账。"
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "doc": {"type": "string"},
                    "kind": {"type": "string", "enum": ["fillet", "chamfer"], "default": "fillet"},
                    "sizeMm": {"type": "number", "description": "圆角半径或倒角宽度（mm）"},
                    "edges": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "point": {"type": "array", "items": {"type": "number"}},
                                "dir": {"type": "array", "items": {"type": "number"}},
                                "radiusMm": {"type": "number"},
                            },
                            "required": ["point"],
                        },
                    },
                    "angleDeg": {"type": "number", "default": 45},
                },
                "required": ["sizeMm", "edges"],
            },
        },
        {
            "name": "sw_shell",
            "description": (
                "抽壳。faces 列的是要开口去掉的面（每项 {point:[x,y,z], normal:[nx,ny,nz]}，"
                "point 落在该面内任意处），留空则整块挖成闭合空腔。"
                "thicknessMm 是壁厚 mm，outward=true 表示往外长（默认往里挖）。返回特征名 + 体积。"
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "doc": {"type": "string"},
                    "thicknessMm": {"type": "number", "description": "壁厚 mm"},
                    "faces": {
                        "type": "array",
                        "description": "开口面（不填=闭合空腔）",
                        "items": {
                            "type": "object",
                            "properties": {
                                "point": {"type": "array", "items": {"type": "number"}},
                                "normal": {"type": "array", "items": {"type": "number"}},
                            },
                            "required": ["point"],
                        },
                    },
                    "outward": {"type": "boolean", "default": False},
                },
                "required": ["thicknessMm"],
            },
        },
        {
            "name": "sw_circular_pattern",
            "description": (
                "圆周阵列。seeds 是要阵列的特征名，count 含种子本身。"
                "angleDeg 默认 360 表示整圈；equalSpacing=true 时它是**总角度**"
                "（整圈按 360/N 等分，非整圈按 总角/(N-1)），false 时它是每份间隔角。"
                "flip=false 为顺时针。axis 用 x/y/z 自动建参考轴，也能用 axisFeature 指定现成的基准轴。"
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "doc": {"type": "string"},
                    "seeds": {"type": "array", "items": {"type": "string"}},
                    "count": {"type": "integer", "minimum": 2},
                    "angleDeg": {"type": "number", "default": 360},
                    "equalSpacing": {"type": "boolean", "default": True},
                    "flip": {"type": "boolean", "default": False},
                    "axis": {"type": "string", "enum": ["x", "y", "z"], "default": "z"},
                    "axisFeature": {"type": "string"},
                    "geometryPattern": {"type": "boolean", "default": True},
                },
                "required": ["seeds", "count"],
            },
        },
        {
            "name": "sw_hole",
            "description": (
                "在面上打孔（几何等价实现，非异型孔向导）：通孔 / 盲孔 / 柱形沉头孔。"
                "points 是孔心模型坐标 mm 数组；face 给 {point,normal} 定位放置面。"
                "end=through_all（默认）或 blind（需 depthMm）。"
                "给 counterboreDiameterMm + counterboreDepthMm 就叠一层沉头。返回特征名 + 体积。"
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "doc": {"type": "string"},
                    "points": {
                        "type": "array",
                        "items": {"type": "array", "items": {"type": "number"}},
                        "description": "孔心 [[x,y,z], ...]（mm）",
                    },
                    "diameterMm": {"type": "number"},
                    "face": {
                        "type": "object",
                        "properties": {
                            "point": {"type": "array", "items": {"type": "number"}},
                            "normal": {"type": "array", "items": {"type": "number"}},
                        },
                    },
                    "end": {"type": "string", "enum": ["through_all", "blind"], "default": "through_all"},
                    "depthMm": {"type": "number"},
                    "counterboreDiameterMm": {"type": "number"},
                    "counterboreDepthMm": {"type": "number"},
                },
                "required": ["points", "diameterMm"],
            },
        },
        {
            "name": "sw_hole_wizard",
            "description": (
                "异型孔向导（SOLIDWORKS Hole Wizard 标准库，真机已验证）：柱形沉头孔 / 锥形沉头孔 / "
                "螺栓通孔 / 螺纹孔。尺寸与沉头直径/深度都取自 SW 的 ISO 标准表，不是手算的。"
                "kind=counterbore（内六角圆柱头 ISO 4762）/ countersink（沉头 ISO 10642）/ "
                "clearance（螺栓通孔 ISO 273）/ tap（螺纹孔，按 ISO 2306 底孔切除 + 装饰螺纹线）。"
                "size 用标称串 M5/M6/M8/M10/M12；points 是孔心模型坐标 mm；"
                "face 给 {point,normal} 定位放置面；end=through_all 或 blind（需 depthMm）。"
                "返回特征名、回读的标准/螺钉类型/实际孔径/沉头尺寸和体积，用于对账。"
                "比 sw_hole 好在：沉头/螺纹尺寸来自标准表，特征类型是真正的 HoleWzd（工程图能出孔标注）。"
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "doc": {"type": "string"},
                    "kind": {
                        "type": "string",
                        "enum": ["counterbore", "countersink", "clearance", "tap"],
                        "default": "counterbore",
                    },
                    "size": {"type": "string", "default": "M8",
                             "description": "标称尺寸 M5/M6/M8/M10/M12"},
                    "points": {
                        "type": "array",
                        "items": {"type": "array", "items": {"type": "number"}},
                        "description": "孔心 [[x,y,z], ...]（mm）",
                    },
                    "face": {
                        "type": "object",
                        "properties": {
                            "point": {"type": "array", "items": {"type": "number"}},
                            "normal": {"type": "array", "items": {"type": "number"}},
                        },
                    },
                    "end": {"type": "string", "enum": ["through_all", "blind"], "default": "through_all"},
                    "depthMm": {"type": "number"},
                    "tapDrillMm": {"type": "number", "description": "仅螺纹孔：覆盖底孔直径（默认查 ISO 2306）"},
                    "name": {"type": "string", "description": "给新特征起的名字（可选）"},
                },
                "required": ["points"],
            },
        },
    ]
)


def request_approval(tool: str, arguments: dict) -> bool:
    """写操作前要授权：浮窗模式走 HTTP；GUI 模式（SW_BRIDGE_GUI_APPROVAL=1）弹确认框。

    两者都没有时按「拒绝」处理，绝不静默放行。

    超时一律 45 s（硬上限 55 s）：DSH 的 `toolCallTimeoutMs` 默认 60 s，
    弹窗要是等得比它久，就会出现「DSH 记失败、SW 真改了」的错位。
    """
    global _APPROVE_ALL
    if _APPROVE_ALL:
        log(f"[sw-bridge] 本会话已放行全部写操作，跳过确认：{tool}")
        return True

    url = os.environ.get("SW_BRIDGE_APPROVAL_URL")
    token = os.environ.get("SW_BRIDGE_APPROVAL_TOKEN") or ""
    if not url:
        if os.environ.get("SW_BRIDGE_GUI_APPROVAL") == "1":
            allow, remember = _gui_approval(tool, arguments)
            if allow and remember:
                _APPROVE_ALL = True
                log("[sw-bridge] 用户勾选了「本次启动内全部允许」")
            return allow
        log("[sw-bridge] 未配置审批方式（浮窗 URL 或 GUI 弹窗），拒绝写操作")
        return False
    body = json.dumps({"tool": tool, "arguments": arguments}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url.rstrip("/") + "/approve",
        data=body,
        headers={"Content-Type": "application/json; charset=utf-8", "X-Bridge-Token": token},
    )
    try:
        with urllib.request.urlopen(req, timeout=50) as resp:
            payload = json.loads(resp.read().decode("utf-8") or b"{}")
        allow = bool(payload.get("allow"))
        if allow and payload.get("remember"):
            _APPROVE_ALL = True
        return allow
    except Exception as exc:  # noqa: BLE001
        log(f"[sw-bridge] 审批请求失败：{type(exc).__name__}: {exc}")
        return False


def _gui_approval(tool: str, arguments: dict) -> tuple[bool, bool]:
    """在**子进程**里弹确认框，返回 (允许, 本会话全部放行)。

    子进程而不是线程：Tk 在主线程之外跑主循环容易把整个进程挂住；
    子进程还能在超时后硬杀，不留幽灵窗口。子进程的 stdout 必须 DEVNULL，
    不然 Tk 的任何输出都会污染 JSON-RPC 通道。
    """
    import approval_dialog

    timeout = approval_dialog.resolve_timeout()
    payload = json.dumps(
        {"tool": tool, "arguments": arguments, "timeout": timeout}, ensure_ascii=False
    )
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        proc = subprocess.run(
            [sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), "approval_dialog.py")],
            input=payload.encode("utf-8"),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=timeout + 8,
            creationflags=flags,
        )
    except subprocess.TimeoutExpired:
        log(f"[sw-bridge] 确认框超时（{timeout}s），按拒绝处理：{tool}")
        return (False, False)
    except Exception as exc:  # noqa: BLE001
        log(f"[sw-bridge] 确认框起不来，按拒绝处理：{type(exc).__name__}: {exc}")
        return (False, False)
    log(f"[sw-bridge] 确认框返回码={proc.returncode}（0允许/2允许并记住/其他拒绝）")
    if proc.returncode == 0:
        return (True, False)
    if proc.returncode == 2:
        return (True, True)
    return (False, False)


def call_tool(name: str, args: dict) -> dict:
    if name == "sw_status":
        status = sw_com.get_status()
        payload = {"running": status.running, "document": status.document, "error": status.error}
        return tool_text(sw_com.to_json(payload), is_error=not status.running)

    if name == "sw_part_summary":
        max_features = int(args.get("max_features", 60))
        summary = sw_com.part_summary(max_features=max_features)
        return tool_text(sw_com.to_json(summary), is_error=not summary.get("ok", False))

    if name == "sw_set_material":
        material = str(args.get("material", "")).strip()
        database = str(args.get("database", "") or "")
        if not request_approval("sw_set_material", {"material": material, "database": database}):
            return tool_text("用户拒绝了这次写操作（材料未改动）。如果确实要改，请重新发起并在确认框里点「允许」。", is_error=True)
        try:
            sw = sw_com.connect()
            doc = sw.ActiveDoc
            if doc is None:
                return tool_text("SW 在运行，但当前没有打开的文档", is_error=True)
            result = sw_com.set_material(doc, material, database)
            rebuilt = sw_com.rebuild(doc)
            result["rebuilt"] = rebuilt
            return tool_text(sw_com.to_json(result))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)

    if name == "sw_mass_properties":
        try:
            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, str(args.get("doc", "") or ""))
            payload = {"ok": True, "doc": sw_com.doc_title(doc)}
            payload.update(sw_com.geometry_snapshot(doc))
            return tool_text(sw_com.to_json(payload))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)

    if name == "sw_rebuild":
        try:
            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, str(args.get("doc", "") or ""))
            return tool_text(sw_com.to_json(sw_com.rebuild_report(doc)))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)

    if name == "sw_new":
        kind = str(args.get("type", "part") or "part")
        template = str(args.get("template", "") or "")
        if not request_approval("sw_new", {"type": kind, "template": template or "(默认模板)"}):
            return tool_text("用户拒绝了新建文档（什么都没做）。", is_error=True)
        try:
            sw = sw_com.connect()
            return tool_text(sw_com.to_json(sw_com.new_document(sw, kind=kind, template=template)))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)

    if name == "sw_build":
        steps = args.get("steps")
        if not isinstance(steps, list) or not steps:
            return tool_text("BAD_REQUEST：steps 必须是非空数组", is_error=True)
        doc_name = str(args.get("doc", "") or "")
        atomic = args.get("atomic", True)
        summary = [
            {k: v for k, v in (s or {}).items() if k in ("op", "on", "depth", "end", "cut", "sketch")}
            for s in steps
            if isinstance(s, dict)
        ]
        if not request_approval("sw_build", {"doc": doc_name or "(活动文档)", "steps": summary}):
            return tool_text("用户拒绝了这次建模（零件未改动）。", is_error=True)
        try:
            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, doc_name)
            report = sw_com.build(doc, steps, atomic=bool(atomic))
            return tool_text(sw_com.to_json(report), is_error=not report.get("ok", False))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)

    if name == "sw_drawing_build":
        model = str(args.get("model", "") or "")
        out_dir = str(args.get("outDir", "") or "")
        if not model or not out_dir:
            return tool_text("BAD_REQUEST：model 和 outDir 必填", is_error=True)
        paper = str(args.get("paper", "A3") or "A3").upper()
        if not request_approval(
            "sw_drawing_build",
            {"model": model, "outDir": out_dir, "paper": paper},
        ):
            return tool_text("用户拒绝了这次出图（没有新建任何文件）。", is_error=True)
        try:
            import sw_drawing

            report = sw_drawing.build(
                model,
                out_dir,
                paper=paper,
                basename=(args.get("basename") or None),
                keep_open=bool(args.get("keepOpen", True)),
                font=(args.get("font", "汉仪长仿宋体") or None),
            )
            return tool_text(
                json.dumps(report, ensure_ascii=False, indent=2),
                is_error=not report.get("ok", False),
            )
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return tool_text(f"DRAWING_FAILED：{exc}", is_error=True)

    if name == "sw_export":
        path = str(args.get("path", "") or "")
        if not path:
            return tool_text("BAD_REQUEST：path 必填", is_error=True)
        fmt = os.path.splitext(path)[1].lstrip(".").lower() or "pdf"
        doc_name = str(args.get("doc", "") or "")
        if not request_approval(
            "sw_export", {"path": path, "format": fmt, "doc": doc_name or "(活动文档)"}
        ):
            return tool_text("用户拒绝了这次导出。", is_error=True)
        try:
            import sw_drawing

            sw = sw_drawing.attach_sw()
            doc = sw_com.focus_doc(sw, doc_name)
            if fmt == "pdf":
                res = sw_drawing.export_pdf(sw, doc, path)
            elif fmt == "bmp":
                os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
                okk = doc.SaveBmp(path, 1600, 1132)
                res = {"ok": bool(okk) and os.path.exists(path), "path": path}
            else:
                res = sw_drawing.save_as(doc, path)
            return tool_text(json.dumps(res, ensure_ascii=False), is_error=not res.get("ok", False))
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return tool_text(f"EXPORT_FAILED：{exc}", is_error=True)

    if name == "sw_face_sketch":
        ents = args.get("entities")
        if not isinstance(ents, list) or not ents:
            return tool_text("BAD_REQUEST：entities 必须是非空数组", is_error=True)
        mode = str(args.get("mode", "cut") or "cut").lower()
        doc_name = str(args.get("doc", "") or "")
        if not request_approval(
            "sw_face_sketch",
            {"doc": doc_name or "(活动文档)", "mode": mode, "entities": ents,
             "faceNormal": args.get("faceNormal"), "facePoint": args.get("facePoint")},
        ):
            return tool_text("用户拒绝了这次面草图（零件未改动）。", is_error=True)
        try:
            import sw_features

            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, doc_name)
            sk = sw_features.sketch_on_face(
                doc, ents,
                point_mm=args.get("facePoint"),
                normal=args.get("faceNormal"),
            )
            out = {"sketch": sk}
            if mode in ("cut", "boss"):
                feat = sw_com.extrude(
                    doc,
                    sketch_name=sk["name"],
                    depth_mm=float(args.get("depthMm", 10)),
                    end=str(args.get("end", "through_all")),
                    cut=(mode == "cut"),
                    reverse=bool(args.get("reverse", False)),
                )
                out["feature"] = feat
            out["mass"] = sw_com.mass_properties(doc)
            return tool_text(json.dumps(out, ensure_ascii=False, indent=2))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return tool_text(f"FACE_SKETCH_FAILED：{exc}", is_error=True)

    if name == "sw_pattern":
        seeds = args.get("seeds")
        direction = args.get("direction")
        if not isinstance(seeds, list) or not seeds:
            return tool_text("BAD_REQUEST：seeds 必须是非空数组", is_error=True)
        if not isinstance(direction, list) or len(direction) != 3:
            return tool_text("BAD_REQUEST：direction 必须是 3 个数", is_error=True)
        doc_name = str(args.get("doc", "") or "")
        if not request_approval(
            "sw_pattern",
            {"doc": doc_name or "(活动文档)", "seeds": seeds, "direction": direction,
             "spacingMm": args.get("spacingMm"), "count": args.get("count")},
        ):
            return tool_text("用户拒绝了这次阵列（零件未改动）。", is_error=True)
        try:
            import sw_features

            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, doc_name)
            rep = sw_features.linear_pattern(
                doc,
                [str(s) for s in seeds],
                [float(v) for v in direction],
                float(args.get("spacingMm", 10)),
                int(args.get("count", 2)),
                flip=bool(args.get("flip", False)),
                geometry_pattern=bool(args.get("geometryPattern", True)),
                direction2=([float(v) for v in args["direction2"]]
                            if isinstance(args.get("direction2"), list) and len(args["direction2"]) == 3
                            else None),
                spacing2_mm=float(args.get("spacing2Mm", 0) or 0),
                count2=int(args.get("count2", 1) or 1),
                flip2=bool(args.get("flip2", False)),
            )
            rep["mass"] = sw_com.mass_properties(doc)
            return tool_text(json.dumps(rep, ensure_ascii=False, indent=2))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return tool_text(f"PATTERN_FAILED：{exc}", is_error=True)

    if name == "sw_rename":
        old = str(args.get("old", "") or "")
        new = str(args.get("new", "") or "")
        if not old or not new:
            return tool_text("BAD_REQUEST：old 和 new 必填", is_error=True)
        doc_name = str(args.get("doc", "") or "")
        if not request_approval("sw_rename", {"doc": doc_name or "(活动文档)", "old": old, "new": new}):
            return tool_text("用户拒绝了这次改名（零件未改动）。", is_error=True)
        try:
            import sw_features

            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, doc_name)
            rep = sw_features.rename_feature(doc, old, new)
            rep["features"] = sw_com.feature_names(doc)
            return tool_text(json.dumps(rep, ensure_ascii=False, indent=2))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return tool_text(f"RENAME_FAILED：{exc}", is_error=True)

    if name == "sw_fillet":
        edges = args.get("edges")
        if not isinstance(edges, list) or not edges:
            return tool_text("BAD_REQUEST：edges 必须是非空数组", is_error=True)
        kind = str(args.get("kind", "fillet") or "fillet").lower()
        size = float(args.get("sizeMm", 0) or 0)
        if size <= 0:
            return tool_text("BAD_REQUEST：sizeMm 必须大于 0", is_error=True)
        doc_name = str(args.get("doc", "") or "")
        if not request_approval("sw_fillet", {"doc": doc_name or "(活动文档)",
                                              "kind": kind, "sizeMm": size}):
            return tool_text("用户拒绝了这次圆角/倒角（零件未改动）。", is_error=True)
        try:
            import sw_features

            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, doc_name)
            if kind == "chamfer":
                rep = sw_features.chamfer(doc, size, edges, float(args.get("angleDeg", 45)))
            else:
                rep = sw_features.fillet(doc, size, edges)
            rep["mass"] = sw_com.mass_properties(doc)
            return tool_text(json.dumps(rep, ensure_ascii=False, indent=2))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return tool_text(f"FILLET_FAILED：{exc}", is_error=True)

    if name == "sw_shell":
        thickness = float(args.get("thicknessMm", 0) or 0)
        if thickness <= 0:
            return tool_text("BAD_REQUEST：thicknessMm 必须大于 0", is_error=True)
        faces = args.get("faces") or []
        if not isinstance(faces, list):
            return tool_text("BAD_REQUEST：faces 必须是数组", is_error=True)
        outward = bool(args.get("outward", False))
        doc_name = str(args.get("doc", "") or "")
        if not request_approval("sw_shell", {"doc": doc_name or "(活动文档)",
                                             "thicknessMm": thickness, "outward": outward,
                                             "openFaces": len(faces)}):
            return tool_text("用户拒绝了这次抽壳（零件未改动）。", is_error=True)
        try:
            import sw_features

            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, doc_name)
            rep = sw_features.shell(doc, thickness, faces, outward)
            rep["mass"] = sw_com.mass_properties(doc)
            return tool_text(json.dumps(rep, ensure_ascii=False, indent=2))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return tool_text(f"SHELL_FAILED：{exc}", is_error=True)

    if name == "sw_circular_pattern":
        seeds = args.get("seeds")
        if not isinstance(seeds, list) or not seeds:
            return tool_text("BAD_REQUEST：seeds 必须是非空数组", is_error=True)
        count = int(args.get("count", 0) or 0)
        if count < 2:
            return tool_text("BAD_REQUEST：count 至少 2", is_error=True)
        angle = float(args.get("angleDeg", 360) or 0)
        if angle <= 0:
            return tool_text("BAD_REQUEST：angleDeg 必须大于 0", is_error=True)
        doc_name = str(args.get("doc", "") or "")
        if not request_approval("sw_circular_pattern", {"doc": doc_name or "(活动文档)",
                                                        "seeds": seeds, "count": count,
                                                        "angleDeg": angle}):
            return tool_text("用户拒绝了这次圆周阵列（零件未改动）。", is_error=True)
        try:
            import sw_features

            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, doc_name)
            rep = sw_features.circular_pattern(
                doc,
                [str(s) for s in seeds],
                count,
                angle,
                bool(args.get("equalSpacing", True)),
                bool(args.get("flip", False)),
                str(args.get("axis", "z") or "z"),
                str(args.get("axisFeature", "") or ""),
                bool(args.get("geometryPattern", True)),
            )
            rep["mass"] = sw_com.mass_properties(doc)
            return tool_text(json.dumps(rep, ensure_ascii=False, indent=2))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return tool_text(f"CIRCULAR_PATTERN_FAILED：{exc}", is_error=True)

    if name == "sw_hole":
        points = args.get("points")
        if not isinstance(points, list) or not points:
            return tool_text("BAD_REQUEST：points 必须是非空数组", is_error=True)
        dia = float(args.get("diameterMm", 0) or 0)
        if dia <= 0:
            return tool_text("BAD_REQUEST：diameterMm 必须大于 0", is_error=True)
        doc_name = str(args.get("doc", "") or "")
        if not request_approval("sw_hole", {"doc": doc_name or "(活动文档)",
                                            "points": len(points), "diameterMm": dia,
                                            "end": args.get("end", "through_all")}):
            return tool_text("用户拒绝了这次打孔（零件未改动）。", is_error=True)
        try:
            import sw_features

            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, doc_name)
            rep = sw_features.hole(
                doc,
                points,
                dia,
                args.get("face") or None,
                str(args.get("end", "through_all") or "through_all"),
                args.get("depthMm"),
                args.get("counterboreDiameterMm"),
                args.get("counterboreDepthMm"),
            )
            rep["mass"] = sw_com.mass_properties(doc)
            return tool_text(json.dumps(rep, ensure_ascii=False, indent=2))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return tool_text(f"HOLE_FAILED：{exc}", is_error=True)

    if name == "sw_hole_wizard":
        points = args.get("points")
        if not isinstance(points, list) or not points:
            return tool_text("BAD_REQUEST：points 必须是非空数组", is_error=True)
        doc_name = str(args.get("doc", "") or "")
        if not request_approval("sw_hole_wizard", {"doc": doc_name or "(活动文档)",
                                                   "kind": args.get("kind", "counterbore"),
                                                   "size": args.get("size", "M8"),
                                                   "points": len(points),
                                                   "end": args.get("end", "through_all")}):
            return tool_text("用户拒绝了这次异型孔（零件未改动）。", is_error=True)
        try:
            import sw_features

            sw = sw_com.connect()
            doc = sw_com.focus_doc(sw, doc_name)
            rep = sw_features.hole_wizard(
                doc,
                points,
                str(args.get("kind", "counterbore") or "counterbore"),
                str(args.get("size", "M8") or "M8"),
                args.get("face") or None,
                str(args.get("end", "through_all") or "through_all"),
                args.get("depthMm"),
                args.get("tapDrillMm"),
                str(args.get("name", "") or ""),
            )
            rep["mass"] = sw_com.mass_properties(doc)
            return tool_text(json.dumps(rep, ensure_ascii=False, indent=2))
        except sw_com.SwError as exc:
            return tool_text(str(exc), is_error=True)
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return tool_text(f"HOLE_WIZARD_FAILED：{exc}", is_error=True)

    raise KeyError(name)


def handle(msg: dict) -> None:
    method = msg.get("method")
    msg_id = msg.get("id")

    if method == "initialize":
        requested = (msg.get("params") or {}).get("protocolVersion") or "2025-06-18"
        ok(
            msg_id,
            {
                "protocolVersion": requested,
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": SERVER_INFO,
            },
        )
        return

    if method in ("notifications/initialized", "notifications/cancelled"):
        return

    if method == "ping":
        ok(msg_id, {})
        return

    if method == "tools/list":
        ok(msg_id, {"tools": TOOLS})
        return

    if method == "tools/call":
        # 丢给工作线程执行：确认框可能弹十几秒，主线程必须还能回 ping
        _JOB_Q.put(msg)
        return

    if method in ("resources/list", "prompts/list"):
        key = "resources" if method == "resources/list" else "prompts"
        ok(msg_id, {key: []})
        return

    if method == "logging/setLevel":
        ok(msg_id, {})
        return

    if msg_id is not None:
        fail(msg_id, -32601, f"不支持的方法：{method}")


def _worker_loop() -> None:
    """唯一的工具执行线程。串行执行，避免两个写操作同时碰同一个文档。"""
    try:
        import pythoncom

        pythoncom.CoInitialize()  # 工具线程自己初始化 COM 套间（STA）
    except Exception as exc:  # noqa: BLE001
        log(f"[sw-bridge] CoInitialize 失败（继续，pywin32 会兜底）：{exc}")

    while True:
        msg = _JOB_Q.get()
        if msg is None:
            return
        params = msg.get("params") or {}
        name = params.get("name")
        args = params.get("arguments") or {}
        msg_id = msg.get("id")
        started = time.monotonic()
        try:
            result = call_tool(name, args)
        except KeyError:
            fail(msg_id, -32602, f"未知工具：{name}")
            continue
        except Exception:  # noqa: BLE001
            log(traceback.format_exc())
            fail(msg_id, -32603, "工具执行异常（详见 stderr）")
            continue
        ok(msg_id, result)
        log(f"[sw-bridge] {name} 用时 {time.monotonic() - started:.2f}s")


def main() -> int:
    # 协议必须是 UTF-8：Windows 下 Python 的 stdio 默认跟随控制台代码页（GBK），
    # 不显式改会让中文 JSON 在客户端那边解码失败。
    for stream, encoding in ((sys.stdin, "utf-8"), (sys.stdout, "utf-8"), (sys.stderr, "utf-8")):
        try:
            stream.reconfigure(encoding=encoding, errors="replace")
        except Exception:  # noqa: BLE001
            pass

    log(f"[sw-bridge] MCP server started (pid={__import__('os').getpid()})")
    threading.Thread(target=_worker_loop, name="sw-tool-worker", daemon=True).start()
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            log(f"[sw-bridge] 非 JSON 输入：{line[:200]}")
            continue
        if "id" not in msg and "method" in msg:
            if msg.get("method") not in ("notifications/initialized", "notifications/cancelled"):
                log(f"[sw-bridge] 通知：{msg.get('method')}")
        try:
            handle(msg)
        except Exception:  # noqa: BLE001
            log(traceback.format_exc())
    log("[sw-bridge] stdin 关闭，退出")
    return 0


if __name__ == "__main__":
    sys.exit(main())
