"""SolidWorks COM 桥（按 SW2024 真机校准过的形态）。

2026-11-16 在 SW2024 SP0.1（`功能牌1 - 副本.SLDPRT`）上实测：
- `GetTitle` / `GetPathName` / `GetType` / `GetFeatureCount` / `FirstFeature` /
  `Feature.Name` / `Feature.GetTypeName2` / `Feature.GetNextFeature` **都是属性式访问**，
  加括号会报“找不到成员”；
- 材料读取用 `IModelDoc2.MaterialUserName` / `MaterialIdName`（属性）；
- 材料写入用 `IPartDoc.SetMaterialPropertyName2(ConfigName, Database, Name)`（方法）；

2026-11-18 增加建模（新建零件 / 草图 / 拉伸 / 切除），同样在真机上逐条校准过，
那批结论写在文件末尾「建模」那一节的注释里。

设计约束：
- 只连接**已经运行**的 SolidWorks，绝不自己启动（`sw_new` 用的是 SW 自己的 NewDocument）；
- 不保存文档；改材料、建模都只改内存，是否保存由用户在 SW 里决定。
"""

from __future__ import annotations

import json
import math
import os
import time
from dataclasses import dataclass, field
from typing import Any

try:  # pywin32 缺失时留给 _require_win32com 报可读的错
    import pythoncom
    import win32com.client
    from win32com.client import VARIANT
except Exception:  # noqa: BLE001
    pythoncom = None  # type: ignore[assignment]
    VARIANT = None  # type: ignore[assignment]

SW_DOC_PART = 1
SW_DOC_ASSEMBLY = 2
SW_DOC_DRAWING = 3

DOC_TYPE_NAMES = {SW_DOC_PART: "零件", SW_DOC_ASSEMBLY: "装配体", SW_DOC_DRAWING: "工程图"}

PROGID = "SldWorks.Application"
DEFAULT_MATERIAL_DB = "solidworks materials.sldmat"


class SwError(RuntimeError):
    """可预期的 SW 侧错误（未运行、无文档、API 不适用）。"""


@dataclass
class SwStatus:
    running: bool
    document: dict[str, Any] | None = None
    error: str | None = None
    extra: dict[str, Any] = field(default_factory=dict)


def _require_win32com():
    try:
        import win32com.client  # noqa: F401
    except Exception as exc:  # noqa: BLE001
        raise SwError(f"缺少 pywin32（{exc}）；请用装了 pywin32 的 Python 运行") from exc
    return win32com.client


def _value(obj, name: str, *args):
    """属性式/方法式兼容访问。

    本机 SW 的零参方法多被暴露成属性；而属性返回的 COM 对象（CDispatch）本身
    又可调用，所以不能用 callable() 判断——按类型名区分：`method` 才是方法。
    """
    attr = getattr(obj, name)
    is_method = type(attr).__name__ in ("method", "builtin_function_or_method")
    if is_method:
        return attr(*args)
    if args:
        raise SwError(f"{name} 是属性，不接受参数")
    return attr


SW_EXE = os.environ.get("SW_EXE") or r"D:\SW\SOLIDWORKS\sldworks.exe"


def _start_sw_process() -> None:
    """起一个**真实的** SOLIDWORKS 进程。

    不能用 `win32com.client.Dispatch(PROGID)`：本机实测它起的实例不登记到 ROT，
    于是同一进程里再调 `GetActiveObject` 也看不到，MCP 子进程更看不到 ——
    表现出来就是「SW 明明在跑，却一直报 SW_NOT_RUNNING」，而且这种幽灵实例
    退出时会带走许可证、看起来像崩溃。
    """
    import subprocess

    if not os.path.isfile(SW_EXE):
        raise SwError(f"找不到 SOLIDWORKS 可执行文件：{SW_EXE}（可用 SW_EXE 环境变量指定）")
    subprocess.Popen([SW_EXE], close_fds=True)


def connect(launch: bool = False):
    win32com = _require_win32com()
    if launch:
        try:
            return win32com.GetActiveObject(PROGID)
        except Exception:  # noqa: BLE001
            _start_sw_process()
        deadline = time.time() + 120
        last = None
        while time.time() < deadline:
            try:
                return win32com.GetActiveObject(PROGID)
            except Exception as exc:  # noqa: BLE001
                last = exc
                time.sleep(2.0)
        raise SwError(f"启动 SolidWorks 超时（120s）：{last}")
    try:
        return win32com.GetActiveObject(PROGID)
    except Exception as exc:  # noqa: BLE001
        raise SwError(
            "SW_NOT_RUNNING：SolidWorks 没有运行（或与本次会话不在同一用户/完整性级别）。"
            "请先手动打开 SOLIDWORKS 2024 并打开一个零件，再重试。"
            f"（底层错误：{type(exc).__name__}: {exc}）"
        ) from exc


def active_doc(sw):
    doc = sw.ActiveDoc
    if doc is None:
        raise SwError("SW 在运行，但当前没有打开的文档")
    return doc


def doc_config(doc) -> str:
    try:
        return str(doc.ConfigurationManager.ActiveConfiguration.Name)
    except Exception:  # noqa: BLE001
        return ""


def get_status() -> SwStatus:
    try:
        sw = connect()
        doc = sw.ActiveDoc
        if doc is None:
            return SwStatus(running=True, error="SW 在运行，但当前没有打开的文档")
        doc_type = int(_value(doc, "GetType"))
        info = {
            "title": str(_value(doc, "GetTitle")),
            "path": str(_value(doc, "GetPathName")),
            "type": DOC_TYPE_NAMES.get(doc_type, f"未知({doc_type})"),
            "typeId": doc_type,
            "configuration": doc_config(doc),
        }
        return SwStatus(running=True, document=info)
    except SwError as exc:
        return SwStatus(running=False, error=str(exc))
    except Exception as exc:  # noqa: BLE001
        return SwStatus(running=True, error=f"读取活动文档失败：{type(exc).__name__}: {exc}")


def list_features(doc, limit: int = 200) -> list[dict[str, str]]:
    features: list[dict[str, str]] = []
    feat = _value(doc, "FirstFeature")
    guard = 0
    while feat is not None and guard < limit:
        guard += 1
        try:
            name = str(_value(feat, "Name"))
        except Exception:  # noqa: BLE001
            name = "?"
        try:
            type_name = str(_value(feat, "GetTypeName2"))
        except Exception:  # noqa: BLE001
            type_name = "?"
        features.append({"name": name, "type": type_name})
        try:
            feat = _value(feat, "GetNextFeature")
        except Exception:  # noqa: BLE001
            break
    return features


def read_material(doc) -> dict[str, Any]:
    """材料名与材料库标识（空字符串表示未指定材料）。"""
    material = ""
    material_id = ""
    try:
        material = str(_value(doc, "MaterialUserName"))
    except Exception:  # noqa: BLE001
        pass
    try:
        material_id = str(_value(doc, "MaterialIdName"))
    except Exception:  # noqa: BLE001
        pass
    # 写材料后 MaterialUserName 可能不刷新，但 MaterialIdName 会立刻变成
    # "库名|材料名|序号"，以它为准补出材料名。
    if not material and material_id:
        parts = material_id.split("|")
        if len(parts) >= 2 and parts[1].strip():
            material = parts[1].strip()
    return {
        "material": material,
        "materialId": material_id,
        "config": doc_config(doc),
        "assigned": bool(material or material_id),
    }


def set_material(doc, material: str, database: str = "") -> dict[str, Any]:
    if not material or not material.strip():
        raise SwError("material 不能为空")
    before = read_material(doc)
    config = before["config"] or "默认"
    db = database or DEFAULT_MATERIAL_DB
    try:
        _value(doc, "SetMaterialPropertyName2", config, db, material)
    except Exception as exc:  # noqa: BLE001
        raise SwError(
            f"写入材料失败：{type(exc).__name__}: {exc}；"
            f"参数=({config!r}, {db!r}, {material!r})"
        ) from exc
    after = read_material(doc)
    return {
        "before": before["material"],
        "after": after["material"],
        "requested": material,
        "database": db,
        "config": config,
        "changed": after["material"] != before["material"],
    }


def rebuild(doc) -> bool | None:
    try:
        return bool(_value(doc, "EditRebuild3"))
    except Exception:  # noqa: BLE001
        return None


def part_summary(max_features: int = 60) -> dict[str, Any]:
    try:
        sw = connect()
        doc = active_doc(sw)
    except SwError as exc:
        return {"ok": False, "error": str(exc)}

    try:
        doc_type = int(_value(doc, "GetType"))
    except Exception:  # noqa: BLE001
        doc_type = 0
    summary: dict[str, Any] = {
        "ok": True,
        "title": str(_value(doc, "GetTitle")),
        "path": str(_value(doc, "GetPathName")),
        "type": DOC_TYPE_NAMES.get(doc_type, f"未知({doc_type})"),
        "configuration": doc_config(doc),
    }
    try:
        summary["featureCount"] = int(_value(doc, "GetFeatureCount"))
    except Exception:  # noqa: BLE001
        summary["featureCount"] = None
    try:
        features = list_features(doc)
        summary["featureCountFromTree"] = len(features)
        summary["features"] = features[:max_features]
        summary["featuresTruncated"] = len(features) > max_features
    except Exception as exc:  # noqa: BLE001
        summary["featureError"] = f"{type(exc).__name__}: {exc}"
    try:
        summary["material"] = read_material(doc)
    except Exception as exc:  # noqa: BLE001
        summary["materialError"] = str(exc)
    return summary


def to_json(data: Any) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2)


# ====================================================================== #
# 建模：新建零件 / 草图 / 拉伸 / 切除 / 自校验
#
# 工具契约的形状参考 dshworks/dsh-solidworks（MIT）的 protocol/tools.json，
# 代码是按本机 SW2024 SP0.1 的实测重写的。**下面每一条都是真跑出来的，不是推断**：
#
# 1. `IModelDocExtension.SelectByID2` 在这台机器上必然失败：
#    9 参（Callout 传 None 或 pythoncom.Empty）报 DISP_E_TYPEMISMATCH，
#    7/8 参报「非选择性的参数」。
#    → 选基准面一律用 `IFeature.Select2(False, 0)`；选面用 `Extension.SelectByRay`。
# 2. 切除必须 `Sd=False` 且 T1/T2 都给端条件。`Sd=True` 的 `FeatureCut4`
#    返回 None、一个特征都不生成（盲切、贯穿都如此）。加材料则相反，`Sd=True` 正常。
# 3. `CreateMassProperty` / `CreateMassProperty2` 在晚期绑定下报「找不到成员」，
#    体积要从 `IModelDoc2.GetMassProperties2(status)` 取：12 个 double，SI 单位
#    （0-2 重心 m，3 体积 m3，4 表面积 m2，5 质量 kg）。
# 4. `IPartDoc.GetPartBox(False)` 返回**文档单位**（本机是 mm）、顺序
#    (minX, minY, minZ, maxX, maxY, maxZ)；`True` 才是米。
# 5. 草图坐标与拉伸深度走 COM 时是**米**，所以对外一律收 mm、进来乘 1e-3。
# 6. 中文模板下基准面叫「前视/上视/右视基准面」，树里它们的前三个 `RefPlane`
#    就是这个顺序；草图叫「草图1」，拉伸叫「凸台-拉伸1」，切除叫「切除-拉伸2」
#    （序号共用一条计数）。
# ====================================================================== #

MM = 0.001

# swEndConditions_e
END_BLIND = 0
END_THROUGH_ALL = 1
END_MID_PLANE = 6
END_THROUGH_ALL_BOTH = 9

END_CONDITIONS = {
    "blind": END_BLIND,
    "through_all": END_THROUGH_ALL,
    "through_all_both": END_THROUGH_ALL_BOTH,
    "mid_plane": END_MID_PLANE,
}

# swUserPreferenceStringValue_e
PREF_DEFAULT_TEMPLATE = {"part": 8, "assembly": 9, "drawing": 10}
TEMPLATE_NAMES = {
    "part": ("gb_part.prtdot", "part.prtdot"),
    "assembly": ("gb_assembly.asmdot", "assembly.asmdot"),
    "drawing": ("gb.drwdot", "drawing.drwdot"),
}

# 树里不属于建模的节点（文件夹、传感器、注释…），统计特征时跳过
TREE_SKIP_TYPES = {
    "FavoriteFolder", "HistoryFolder", "SelectionSetFolder", "SensorFolder",
    "DocsFolder", "DetailCabinet", "InkMarkupFolder", "EnvFolder",
    "SolidBodyFolder", "SurfaceBodyFolder", "MaterialFolder", "CommentsFolder",
    "CommentFolder", "ScenesFolder", "LightsFolder", "CamerasFolder",
    "EqnFolder", "ConfigTableFolder", "Tables",
}


def _byref_int(value: int = 0):
    """给 COM 的 byref 出参用（传普通 int 会 DISP_E_TYPEMISMATCH）。"""
    _require_win32com()
    return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, int(value))


def _doc_type(doc) -> int:
    try:
        return int(_value(doc, "GetType"))
    except Exception:  # noqa: BLE001
        return 0


def _require_part(doc, what: str = "这个操作") -> None:
    doc_type = _doc_type(doc)
    if doc_type != SW_DOC_PART:
        raise SwError(
            f"WRONG_DOC_TYPE：{what}只适用于零件，当前文档是"
            f"{DOC_TYPE_NAMES.get(doc_type, f'未知({doc_type})')}"
        )


def open_documents(sw) -> list:
    """当前打开的所有文档（本机 GetDocuments 是属性，空时返回 None）。"""
    try:
        raw = _value(sw, "GetDocuments")
    except Exception:  # noqa: BLE001
        return []
    if raw is None:
        return []
    if isinstance(raw, (list, tuple)):
        return [d for d in raw if d is not None]
    return [raw]


def doc_title(doc) -> str:
    try:
        return str(_value(doc, "GetTitle"))
    except Exception:  # noqa: BLE001
        return ""


def doc_path(doc) -> str:
    try:
        return str(_value(doc, "GetPathName"))
    except Exception:  # noqa: BLE001
        return ""


def focus_doc(sw, name: str = ""):
    """按标题 / 全路径 / 文件名尾匹配定位文档，并把它激活（草图必须活动文档）。"""
    if not name:
        return active_doc(sw)
    target = name.strip()
    lowered = target.lower()
    for doc in open_documents(sw):
        title = doc_title(doc)
        path = doc_path(doc)
        if (
            lowered in (title.lower(), path.lower())
            or path.lower().endswith(os.sep + lowered)
            or os.path.basename(path).lower() == lowered
        ):
            try:
                sw.ActivateDoc3(title, False, 0, _byref_int(0))
            except Exception:  # noqa: BLE001
                sw.ActivateDoc2(title, False, _byref_int(0))
            return sw.ActiveDoc
    raise SwError(f"DOC_NOT_OPEN：没有打开名为 {name!r} 的文档")


def tree_items(doc) -> list[tuple[str, str, Any]]:
    """特征树里（去掉文件夹）的 (名字, 类型, 特征对象)，保持树里的顺序。"""
    items: list[tuple[str, str, Any]] = []
    try:
        feat = _value(doc, "FirstFeature")
    except Exception:  # noqa: BLE001
        return items
    guard = 0
    while feat is not None and guard < 2000:
        guard += 1
        try:
            name = str(_value(feat, "Name"))
        except Exception:  # noqa: BLE001
            name = "?"
        try:
            kind = str(_value(feat, "GetTypeName2"))
        except Exception:  # noqa: BLE001
            kind = "?"
        if kind not in TREE_SKIP_TYPES:
            items.append((name, kind, feat))
        try:
            feat = _value(feat, "GetNextFeature")
        except Exception:  # noqa: BLE001
            break
    return items


def feature_names(doc) -> list[str]:
    return [name for name, _kind, _feat in tree_items(doc)]


def find_feature(doc, name: str):
    lowered = name.strip().lower()
    for item_name, _kind, feat in tree_items(doc):
        if item_name.lower() == lowered:
            return feat
    return None


def mass_properties(doc) -> dict[str, Any]:
    """体积/质量/表面积/重心。单位统一成 mm / kg / mm2。"""
    raw = None
    try:
        raw = _value(doc, "GetMassProperties2", _byref_int(0))
    except Exception:  # noqa: BLE001
        raw = None
    if raw is None:
        try:
            raw = _value(doc, "GetMassProperties")
        except Exception as exc:  # noqa: BLE001
            raise SwError(f"COM_NULL：读不到质量属性（{type(exc).__name__}: {exc}）") from exc
    if raw is None:
        raise SwError("COM_NULL：读不到质量属性")
    vals = [float(v) for v in raw]
    if len(vals) < 6:
        raise SwError(f"COM_ERROR：质量属性返回 {len(vals)} 个数，少于 6 个")
    volume_m3 = vals[3]
    density = (vals[5] / volume_m3) if volume_m3 else 0.0
    return {
        "volume_mm3": volume_m3 * 1e9,
        "surface_area_mm2": vals[4] * 1e6,
        "mass_kg": vals[5],
        "density_kg_m3": density,
        "center_mm": [vals[0] / MM, vals[1] / MM, vals[2] / MM],
    }


def part_box_mm(doc) -> dict[str, Any] | None:
    """包围盒（mm）。GetPartBox(False) 给文档单位，本机是 mm。"""
    try:
        raw = _value(doc, "GetPartBox", False)
    except Exception:  # noqa: BLE001
        return None
    if raw is None:
        return None
    try:
        vals = [float(v) for v in raw]
    except Exception:  # noqa: BLE001
        return None
    if len(vals) < 6:
        return None
    lo = vals[0:3]
    hi = vals[3:6]
    return {
        "min_mm": lo,
        "max_mm": hi,
        "size_mm": [hi[i] - lo[i] for i in range(3)],
    }


def geometry_snapshot(doc) -> dict[str, Any]:
    """自校验用的那一组数：体积、质量、表面积、包围盒。"""
    out: dict[str, Any] = {}
    try:
        out.update(mass_properties(doc))
    except SwError as exc:
        out["massError"] = str(exc)
    box = part_box_mm(doc)
    if box is not None:
        out["bbox"] = box
    return out


def _template_guess(kind: str) -> str:
    import glob

    names = TEMPLATE_NAMES.get(kind, ())
    roots = [
        r"C:\ProgramData\SOLIDWORKS",
        r"C:\Users\Public\Documents\SOLIDWORKS",
    ]
    hits: list[str] = []
    for root in roots:
        if not os.path.isdir(root):
            continue
        for name in names:
            hits.extend(sorted(glob.glob(os.path.join(root, "*", "**", name), recursive=True)))
    return hits[0] if hits else ""


def default_template(sw, kind: str) -> str:
    """先问 SW 自己的默认模板设置，再按安装目录里的模板文件名兜底。"""
    pref = PREF_DEFAULT_TEMPLATE.get(kind)
    if pref is not None:
        try:
            path = str(_value(sw, "GetUserPreferenceStringValue", pref) or "")
        except Exception:  # noqa: BLE001
            path = ""
        if path and os.path.isfile(path):
            return path
    return _template_guess(kind)


def new_document(sw, kind: str = "part", template: str = "") -> dict[str, Any]:
    """用 SW 模板新建一个空文档并激活它。不保存、不覆盖任何文件。"""
    kind = (kind or "part").strip().lower()
    if kind not in PREF_DEFAULT_TEMPLATE:
        raise SwError(f"BAD_REQUEST：type 只能是 part / assembly / drawing，收到 {kind!r}")
    given = (template or "").strip()
    if given:
        if not os.path.isfile(given):
            raise SwError(f"FILE_NOT_FOUND：模板不存在 {given}")
        tpl = given
    else:
        tpl = default_template(sw, kind)
        if not tpl:
            raise SwError(
                f"FILE_NOT_FOUND：找不到 {kind} 模板。请在 SW 里设置默认模板，"
                "或直接传 template（.prtdot/.asmdot 的全路径）"
            )
    doc = _value(sw, "NewDocument", tpl, 0, 0, 0)
    if doc is None:
        raise SwError(f"OPEN_FAILED：SW 没能用模板新建文档（{tpl}）")
    return {
        "ok": True,
        "kind": kind,
        "template": tpl,
        "title": doc_title(doc),
        "path": doc_path(doc),
        "type": DOC_TYPE_NAMES.get(_doc_type(doc), "未知"),
        "untitled": not doc_path(doc),
    }


def _clear_selection(doc) -> None:
    try:
        _value(doc, "ClearSelection2", True)
    except Exception:  # noqa: BLE001
        pass


def _select(feat) -> bool:
    try:
        return bool(feat.Select2(False, 0))
    except Exception:  # noqa: BLE001
        return False


def base_planes(doc) -> list[tuple[str, Any]]:
    return [(name, feat) for name, kind, feat in tree_items(doc) if kind == "RefPlane"]


def _pick_plane(doc, on: str):
    """on: front/top/right（任何语言的模板都认）或基准面自己的名字。"""
    planes = base_planes(doc)
    if not planes:
        raise SwError("COM_NULL：这个零件里没有基准面")
    key = (on or "front").strip().lower()
    by_key = {"front": 0, "top": 1, "right": 2, "前视": 0, "上视": 1, "右视": 2}
    if key in by_key:
        idx = by_key[key]
        if idx >= len(planes):
            raise SwError(f"COM_NULL：找不到第 {idx + 1} 个基准面")
        return planes[idx]
    for name, feat in planes:
        if name.lower() == key:
            return name, feat
    raise SwError(f"BAD_REQUEST：找不到基准面 {on!r}（可用 front/top/right 或基准面名字）")


def _point(p) -> tuple[float, float, float]:
    if not isinstance(p, (list, tuple)) or len(p) < 2 or len(p) > 3:
        raise SwError(f"BAD_REQUEST：点要写成 [x, y] 或 [x, y, z]（mm），收到 {p!r}")
    vals = [float(v) for v in p] + [0.0] * (3 - len(p))
    return vals[0] * MM, vals[1] * MM, vals[2] * MM


def _draw_entities(sm, entities: list[dict[str, Any]]) -> dict[str, int]:
    counts: dict[str, int] = {}
    if not entities:
        raise SwError("BAD_REQUEST：sketch 步骤至少要有一个 entities")
    for raw in entities:
        if not isinstance(raw, dict):
            raise SwError(f"BAD_REQUEST：entities 里每一项都必须是对象，收到 {raw!r}")
        kind = str(raw.get("type", "")).strip().lower()
        if kind == "line":
            a, b = _point(raw.get("from")), _point(raw.get("to"))
            sm.CreateLine(*a, *b)
        elif kind == "centerline":
            a, b = _point(raw.get("from")), _point(raw.get("to"))
            sm.CreateCenterLine(*a, *b)
        elif kind == "rectangle":
            a, b = _point(raw.get("corner")), _point(raw.get("opposite"))
            sm.CreateCornerRectangle(*a, *b)
        elif kind == "circle":
            cx, cy, cz = _point(raw.get("center"))
            if raw.get("diameter") is not None:
                radius_mm = float(raw["diameter"]) / 2.0
            elif raw.get("radius") is not None:
                radius_mm = float(raw["radius"])
            else:
                raise SwError("BAD_REQUEST：circle 需要 diameter 或 radius")
            if radius_mm <= 0:
                raise SwError("BAD_REQUEST：圆的直径/半径必须大于 0")
            sm.CreateCircle(cx, cy, cz, cx + radius_mm * MM, cy, cz)
        elif kind == "arc":
            a, b = _point(raw.get("from")), _point(raw.get("to"))
            through = _point(raw.get("through"))
            sm.Create3PointArc(*a, *b, *through)
        elif kind == "polyline":
            points = raw.get("points") or []
            if len(points) < 2:
                raise SwError("BAD_REQUEST：polyline 至少要两个点")
            pts = [_point(p) for p in points]
            if raw.get("closed", True):
                pts.append(pts[0])
            for i in range(len(pts) - 1):
                sm.CreateLine(*pts[i], *pts[i + 1])
        elif kind == "point":
            # 异型孔靠草图点定位（不是圆），所以这里必须支持单独的点
            p = raw.get("point") or raw.get("from") or raw.get("center")
            sm.CreatePoint(*_point(p))
        else:
            raise SwError(
                f"BAD_REQUEST：不支持的实体类型 {kind!r}"
                "（可用 line/centerline/rectangle/circle/arc/polyline/point）"
            )
        counts[kind] = counts.get(kind, 0) + 1
    return counts


def sketch(doc, on: str = "front", entities: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    """在基准面上画草图并退出草图（退出后草图是选中态，直接可以拉伸）。"""
    _require_part(doc, "草图")
    plane_name, plane_feat = _pick_plane(doc, on)
    _clear_selection(doc)
    if not _select(plane_feat):
        raise SwError(f"COM_ERROR：选不中基准面 {plane_name!r}")
    before = set(feature_names(doc))
    sm = doc.SketchManager
    try:
        _value(sm, "InsertSketch", True)
    except Exception as exc:  # noqa: BLE001
        raise SwError(f"COM_ERROR：进不了草图（{type(exc).__name__}: {exc}）") from exc
    active = None
    try:
        active = _value(sm, "ActiveSketch")
    except Exception:  # noqa: BLE001
        active = None
    if active is None:
        raise SwError(
            f"COM_ERROR：草图没开始（基准面 {plane_name!r} 可能没选中）。"
            "SW 里如果有未完成的草图或正在等对话框，也会这样"
        )
    try:
        counts = _draw_entities(sm, entities or [])
    except SwError:
        try:
            _value(sm, "InsertSketch", True)  # 退出，别把 SW 留在草图里
        except Exception:  # noqa: BLE001
            pass
        leftover = [n for n in feature_names(doc) if n not in before]
        if leftover:  # 半成品草图不要留在树里
            rollback(doc, leftover)
        raise
    _value(sm, "InsertSketch", True)
    created = [n for n in feature_names(doc) if n not in before]
    new_sketch = created[0] if created else ""
    for candidate in created:
        feat = find_feature(doc, candidate)
        if feat is None:
            continue
        try:
            if str(_value(feat, "GetTypeName2")) == "ProfileFeature":
                new_sketch = candidate
                break
        except Exception:  # noqa: BLE001
            continue
    if not new_sketch:
        raise SwError("COM_ERROR：草图退出后特征树里没看到新草图")
    return {
        "op": "sketch",
        "on": plane_name,
        "name": new_sketch,
        "entities": counts,
        "segments": sum(counts.values()),
    }


def _extrusion_args(*, single: bool, end1: int, end2: int, depth: float,
                    reverse: bool, merge: bool, draft_deg: float) -> tuple:
    return (
        single, reverse, False,
        end1, end2,
        depth, depth if not single else 0.0,
        False, False, False, False,
        draft_deg, 0.0,
        False, False, False, False,
        merge, True, True,
        0, 0.0, False,
    )


def _cut_args(*, single: bool, end1: int, end2: int, depth: float, reverse: bool,
              draft_deg: float) -> tuple:
    return (
        single, reverse, False,
        end1, end2,
        depth, depth if not single else 0.0,
        False, False, False, False,
        draft_deg, 0.0,
        False, False, False, False,
        False,  # NormalCut
        True, True,
        False, False, False,
        0, 0.0, False,
        True,
    )


def extrude(
    doc,
    sketch_name: str = "",
    depth_mm: float = 10.0,
    end: str = "blind",
    cut: bool = False,
    reverse: bool = False,
    merge: bool = True,
    draft_deg: float = 0.0,
) -> dict[str, Any]:
    """拉伸/切除。cut=True 时用的是切除配方（Sd=False，两侧同端条件）。"""
    _require_part(doc, "拉伸")
    end = (end or "blind").strip().lower()
    if end not in END_CONDITIONS:
        raise SwError(
            f"BAD_REQUEST：end 只能是 {sorted(END_CONDITIONS)}，收到 {end!r}"
        )
    cond = END_CONDITIONS[end]
    if end == "blind" and float(depth_mm) <= 0:
        raise SwError("BAD_REQUEST：blind 拉伸/切除需要正的 depth（mm）")

    profile = find_feature(doc, sketch_name) if sketch_name else None
    if sketch_name and profile is None:
        raise SwError(f"NO_ACTIVE_DOC：找不到草图 {sketch_name!r}")
    if profile is None:
        sketches = [feat for _name, kind, feat in tree_items(doc) if kind == "ProfileFeature"]
        if not sketches:
            raise SwError("NO_ACTIVE_DOC：这个零件里没有草图可拉伸")
        profile = sketches[-1]
    try:
        profile_name = str(_value(profile, "Name"))
    except Exception:  # noqa: BLE001
        profile_name = sketch_name
    _clear_selection(doc)
    if not _select(profile):
        raise SwError("COM_ERROR：选不中草图")
    depth = float(depth_mm) * MM
    before = set(feature_names(doc))
    fm = doc.FeatureManager
    feature = None
    attempts: list[str] = []
    if cut:
        end1 = END_THROUGH_ALL if cond == END_THROUGH_ALL_BOTH else cond
        for single in (False, True):
            end2 = end1 if not single else END_BLIND
            attempts.append(f"FeatureCut4(single={single}, end={end}, d={float(depth_mm):g}mm)")
            feature = _value(
                fm, "FeatureCut4",
                *_cut_args(single=single, end1=end1, end2=end2,
                           depth=depth, reverse=reverse, draft_deg=draft_deg),
            )
            if feature is not None:
                break
    else:
        if cond == END_THROUGH_ALL_BOTH:
            end1, end2, single = END_THROUGH_ALL, END_THROUGH_ALL, False
        else:
            end1, end2, single = cond, END_BLIND, True
        attempts.append(f"FeatureExtrusion3(end={end}, d={float(depth_mm):g}mm)")
        feature = _value(
            fm, "FeatureExtrusion3",
            *_extrusion_args(single=single, end1=end1, end2=end2, depth=depth,
                             reverse=reverse, merge=merge, draft_deg=draft_deg),
        )
    if feature is None:
        raise SwError(
            "COM_NULL：SW 没有生成特征（" + "；".join(attempts) + "）。"
            "常见原因：草图没有封闭轮廓、草图被别的特征吃掉、或选中的不是草图"
        )
    created = [n for n in feature_names(doc) if n not in before]
    name = created[0] if created else ""
    feat = find_feature(doc, name) if name else feature
    try:
        feat_type = str(_value(feat, "GetTypeName2")) if feat is not None else ""
    except Exception:  # noqa: BLE001
        feat_type = ""
    return {
        "op": "cut" if cut else "extrude",
        "name": name,
        "type": feat_type,
        "sketch": profile_name,
        "end": end,
        "depth_mm": float(depth_mm),
        "reverse": bool(reverse),
        "merge": bool(merge) and not cut,
        "method": attempts[-1],
    }


def rollback(doc, names: list[str]) -> list[str]:
    """按名字倒序删掉本次建的特征（一步失败整步撤销的后半截）。"""
    removed: list[str] = []
    for name in reversed(names):
        feat = find_feature(doc, name)
        if feat is None:
            continue
        try:
            _clear_selection(doc)
            if not _select(feat):
                continue
            _value(doc, "EditDelete")
            removed.append(name)
        except Exception:  # noqa: BLE001
            continue
    return removed


def build(doc, steps: list[dict[str, Any]], atomic: bool = True) -> dict[str, Any]:
    """按顺序跑 sketch / extrude（含 cut）步骤，最后回读体积与包围盒。

    atomic=True 时某一步失败会删掉本次调用已经建出来的特征，零件回到调用前。
    """
    _require_part(doc, "建模")
    if not steps:
        raise SwError("BAD_REQUEST：steps 不能为空")
    if len(steps) > 60:
        raise SwError("BAD_REQUEST：一次最多 60 步")
    before = set(feature_names(doc))
    created: list[dict[str, Any]] = []
    started = time.perf_counter()
    try:
        for index, step in enumerate(steps, start=1):
            if not isinstance(step, dict):
                raise SwError(f"BAD_REQUEST：第 {index} 步不是对象")
            op = str(step.get("op", "")).strip().lower()
            if op == "sketch":
                created.append(sketch(doc, on=str(step.get("on", "front")),
                                      entities=step.get("entities") or []))
            elif op in ("extrude", "cut"):
                created.append(extrude(
                    doc,
                    sketch_name=str(step.get("sketch", "") or ""),
                    depth_mm=float(step.get("depth", 10.0)),
                    end=str(step.get("end", "blind") or "blind"),
                    cut=(op == "cut") or bool(step.get("cut", False)),
                    reverse=bool(step.get("reverse", False)),
                    merge=bool(step.get("merge", True)),
                    draft_deg=float(step.get("draft", 0.0) or 0.0),
                ))
            else:
                raise SwError(
                    f"BAD_REQUEST：第 {index} 步的 op={op!r} 还不支持"
                    "（本版只有 sketch / extrude / cut）"
                )
            created[-1]["step"] = index
    except SwError as exc:
        report: dict[str, Any] = {
            "ok": False,
            "error": str(exc),
            "created": created,
            "doc": doc_title(doc),
        }
        if atomic:
            names = [c.get("name") for c in created if c.get("name")]
            report["rolled_back"] = rollback(doc, names)
        try:
            report.update(geometry_snapshot(doc))
        except Exception:  # noqa: BLE001
            pass
        report["duration_ms"] = int((time.perf_counter() - started) * 1000)
        return report

    after = set(feature_names(doc))
    report = {
        "ok": True,
        "doc": doc_title(doc),
        "created": created,
        "newFeatures": [n for n in feature_names(doc) if n not in before and n in after],
        "duration_ms": int((time.perf_counter() - started) * 1000),
    }
    report.update(geometry_snapshot(doc))
    return report


def rebuild_report(doc) -> dict[str, Any]:
    """重建一次并回读体积/包围盒。逐特征的错误码还没接（本机未验证）。"""
    flag = None
    try:
        flag = bool(_value(doc, "EditRebuild3"))
    except Exception as exc:  # noqa: BLE001
        raise SwError(f"COM_ERROR：重建失败（{type(exc).__name__}: {exc}）") from exc
    report: dict[str, Any] = {
        "ok": True,
        "rebuilt": flag,
        "doc": doc_title(doc),
        "features": [{"name": n, "type": k} for n, k, _f in tree_items(doc)],
    }
    report.update(geometry_snapshot(doc))
    return report
