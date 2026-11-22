# -*- coding: utf-8 -*-
"""GB 工程图：建图 / 设图幅 / 插三视图 / 自动选比例 / 导出。

设计约定（与本项目建模层一致）：
  * 对外一律 mm；进 COM 乘 MM。
  * 只新建文件，绝不改用户已有文档。
  * 每个写操作都能回读——出图最大的失败模式是「返回 True 但什么都没建」。
  * GB 规则全部来自 gb_rules.json，不写死数值。

真机校准（2026-11-20，SW2024 SP0.1 / interop 32.0.1.4）：
  1. `sw.NewDocument(gb_a3.drwdot,0,0,0)` 得到的是 **GB 图框**（a3 - gb.slddrt），
     不是用户现有图用的 ISO 图框。
  2. `Create1stAngleViews(path)` 0.06 s 建出 3 个视图，但**按 1:1 摆，长零件会重叠**。
     所以必须两趟：先 1:1 建、量出轮廓、再定比例。
  3. `Extension.SaveAs(..., None, ...)` 存 .slddrw 报「类型不匹配」；
     必须传 `VARIANT(VT_DISPATCH, None)`。导 PDF 时第四参要真 `IExportPdfData`。
  4. `GetViews()` 返回 tuple 套一层，元素是 PyIDispatch，要再 Dispatch 一次。
  5. `sheet.GetProperties()` 7 个值：下标 2/3 是比例分子分母，4 是第一角标志（GB=1）。
"""

from __future__ import annotations

import json
import os
from typing import Any

try:  # 与 sw_com 保持一致的软依赖
    import pythoncom
    from win32com.client import VARIANT
    import win32com.client as wc
except Exception:  # pragma: no cover
    pythoncom = None
    VARIANT = None
    wc = None

import sw_com as C

MM = 0.001
RULES_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gb_rules.json")

SW_DOC_DRAWING = 3

# 图纸默认字体。族名必须和 Windows 字体里的族名一致：
#   汉仪长仿宋体.ttf（%LOCALAPPDATA%\Microsoft\Windows\Fonts，2017 版，4 665 776 字节）
#   hyswlongfangsong.ttf（C:\Windows\Fonts，2023 版，4 682 840 字节）
# 两个文件的 name 表一模一样：family=HYSWLongFangSong / 汉仪长仿宋体，
# postscript=HYSWLongFangSong（PDF 里看到的就是这个名字）。SW 只认族名。
#
# 2026-11-21 实测（GDI+ 系统族 vs 只加载单份文件，逐像素比）：
#   系统渲染 == 2023 版（0 差异），== 2017 版差 14 922 像素
#   → 本机 Windows 解析到的是 **2023 版**，SW 用的就是它。
# 两份文件不是完全一样：字符集都 8 633，但 426 个字形轮廓不同（ASCII 数字/字母为主），
# 同一串数字 2017 版墨宽 795px、2023 版 771px（宽 ~3%）。
DEFAULT_FONT = "汉仪长仿宋体"

# swUserPreferenceTextFormat_e：决定「文件属性 > 注解 > 字体」里那几项
_PREF_TEXT_FORMATS = (
    (0, "note"),        # swDetailingNoteTextFormat    —— 注释/标题栏文字
    (1, "dimension"),   # swDetailingDimensionTextFormat —— 尺寸、孔标注
    (19, "annotation"),  # swDetailingAnnotationTextFormat
)
_NO_OPTION = 0  # swUserPreferenceOption_e.swDetailingNoOptionSpecified


# ---------------------------------------------------------------- 规则表
_rules_cache: dict[str, Any] | None = None


def rules() -> dict[str, Any]:
    global _rules_cache
    if _rules_cache is None:
        with open(RULES_PATH, "r", encoding="utf-8") as f:
            _rules_cache = json.load(f)
    return _rules_cache


def paper_size(paper: str) -> tuple[float, float]:
    p = rules()["paper"]
    key = str(paper or "A3").upper().replace(" ", "")
    if key not in p:
        raise ValueError(f"不支持的图幅 {paper!r}，只认 {[k for k in p if k.startswith('A')]}")
    return float(p[key]["width"]), float(p[key]["height"])


def template_for(paper: str) -> str:
    """找 SW 安装目录里的 GB 图纸模板。"""
    key = str(paper or "A3").upper().replace(" ", "")
    name = "gb_a4p.drwdot" if key == "A4P" else f"gb_{key.lower()}.drwdot"
    roots = [
        os.environ.get("SW_TEMPLATE_DIR", ""),
        r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2024\templates",
        r"D:\SW\SOLIDWORKS\data\templates",
    ]
    for root in roots:
        if root and os.path.isdir(root):
            p = os.path.join(root, name)
            if os.path.isfile(p):
                return p
    raise FileNotFoundError(
        f"找不到 GB 图纸模板 {name}；请把 SW_TEMPLATE_DIR 指向含 gb_a*.drwdot 的目录"
    )


# ---------------------------------------------------------------- 小工具
def _val(obj, name, *args, default=None):
    return C._value(obj, name, *args) if args else C._value(obj, name)


def _dispatch_all(raw) -> list:
    """GetViews 会多包一层，元素还是 PyIDispatch。"""
    items = list(raw or [])
    if len(items) == 1 and isinstance(items[0], (tuple, list)):
        items = list(items[0])
    out = []
    for it in items:
        if "PyIDispatch" in type(it).__name__:
            try:
                it = wc.Dispatch(it)
            except Exception:
                continue
        out.append(it)
    return out


def _rect_mm(outline) -> tuple[float, float, float, float]:
    """view.GetOutline() 给的是图纸坐标系、单位米。"""
    x1, y1, x2, y2 = [float(v) * 1000.0 for v in outline]
    if x1 > x2:
        x1, x2 = x2, x1
    if y1 > y2:
        y1, y2 = y2, y1
    return x1, y1, x2, y2


def _num_pair(v) -> tuple[float, float]:
    try:
        if isinstance(v, (tuple, list)) and len(v) >= 2:
            return float(v[0]), float(v[1])
    except Exception:
        pass
    return 1.0, 1.0


# ---------------------------------------------------------------- 图纸页
def attach_sw():
    if wc is None:
        raise RuntimeError("pywin32 不可用")
    return wc.GetActiveObject("SldWorks.Application")


def sheet_state(drw) -> dict[str, Any]:
    sh = _val(drw, "GetCurrentSheet")
    props = _val(sh, "GetProperties") or ()
    num = den = 1.0
    if len(props) >= 4:
        num, den = float(props[2]), float(props[3])
    return {
        "name": _val(sh, "GetName"),
        "sheet_format": _val(sh, "GetSheetFormatName"),
        "template": _val(sh, "GetTemplateName"),
        "size_mm": [round(float(props[5]) * 1000, 2), round(float(props[6]) * 1000, 2)] if len(props) >= 7 else None,
        "scale": [num, den],
        "first_angle": bool(props[4]) if len(props) >= 5 else None,
    }


def set_sheet_scale(drw, num: float, den: float) -> bool:
    """改图纸比例。成功与否一律用 GetProperties 回读判定。"""
    sh = _val(drw, "GetCurrentSheet")
    attempts = [
        lambda: sh.SetScale(num, den, False, False),
        lambda: sh.SetScale(num, den, False, True),
        lambda: sh.SetScale(num, den),
    ]
    for fn in attempts:
        try:
            fn()
        except Exception:
            continue
        st = sheet_state(drw)
        if abs(st["scale"][0] - num) < 1e-6 and abs(st["scale"][1] - den) < 1e-6:
            return True
    return False


# ---------------------------------------------------------------- 字体
def set_document_font(drw, font: str = DEFAULT_FONT, prefs=None) -> dict[str, Any]:
    """按**文件属性**改字体：等价于 工具 > 选项 > 文档属性 > 注解 > 字体。

    `IModelDocExtension::SetUserPreferenceTextFormat(pref, swDetailingNoOptionSpecified, TextFormat)`
    —— 只换 `TypeFaceName`，其余（字高/粗体/斜体）沿用原对象，所以标题栏字号不会被带跑。

    这一点很关键：图纸格式里的那 80 多个标题栏文字通常 `GetUseDocTextFormat()=True`，
    也就是**跟着文档属性走**，改完文档属性它们自动就变了，不用逐条改（也改不到：
    这些对象从 SelectAll 拿回来的是 INote，本机 pywin32 上 `GetTextFormat` 报
    「找不到成员」）。真正带本地格式的只有视图里的尺寸，见 `apply_font_to_dimensions`。

    返回每项的前后字体，`ok` 以**回读**为准。
    """
    ext = _val(drw, "Extension")
    out: dict[str, Any] = {}
    for pref, name in (prefs or _PREF_TEXT_FORMATS):
        cur = _val(ext, "GetUserPreferenceTextFormat", pref, _NO_OPTION)
        if cur is None:
            out[name] = {"ok": False, "reason": "拿不到 TextFormat 对象"}
            continue
        before = _val(cur, "TypeFaceName")
        if before != font:
            try:
                cur.TypeFaceName = font
            except Exception as exc:  # noqa: BLE001
                out[name] = {"ok": False, "reason": f"赋 TypeFaceName 失败：{exc}"}
                continue
        ok = _val(ext, "SetUserPreferenceTextFormat", pref, _NO_OPTION, cur)
        after = _val(_val(ext, "GetUserPreferenceTextFormat", pref, _NO_OPTION), "TypeFaceName")
        out[name] = {
            "ok": bool(ok) and after == font,
            "before": before,
            "after": after,
            "height_mm": round(float(_val(cur, "CharHeight", default=0) or 0) / MM, 3),
        }
    return out


def apply_font_to_dimensions(drw, font: str = DEFAULT_FONT) -> list[dict[str, Any]]:
    """把视图里带**本地格式**的尺寸/孔标注逐个换成 `font`（不动字高）。

    `IAnnotation.GetTextFormat(i)` → `ITextFormat.TypeFaceName` → `SetTextFormat(i, False, tf)`；
    第二参 `UseDoc=False` 表示保持本地格式（传 True 会丢字号，标题栏里各格字高不同）。
    """
    changed: list[dict[str, Any]] = []
    for item in display_dimensions(drw):
        ann = item["_ann"]
        count = int(_val(ann, "GetTextFormatCount", default=0) or 0)
        for i in range(count):
            tf = _val(ann, "GetTextFormat", i)
            if tf is None:
                continue
            before = _val(tf, "TypeFaceName")
            if before == font:
                continue
            tf.TypeFaceName = font
            ok = _val(ann, "SetTextFormat", i, False, tf)
            changed.append({
                "dim": item["name"], "idx": i, "from": before,
                "to": _val(tf, "TypeFaceName"), "ok": bool(ok),
            })
    return changed


def apply_font(drw, font: str = DEFAULT_FONT) -> dict[str, Any]:
    """整套换字体：先改文档属性（标题栏等跟随的文字），再改尺寸的本地格式。"""
    prefs = set_document_font(drw, font)
    dims = apply_font_to_dimensions(drw, font)
    _val(drw, "GraphicsRedraw2")
    return {"font": font, "document_properties": prefs, "dimensions_changed": dims}


# ---------------------------------------------------------------- 视图
def measure_views(drw) -> list[dict[str, Any]]:
    views = _dispatch_all(_val(drw, "GetViews"))
    out = []
    for v in views:
        name = _val(v, "GetName2", default="?")
        vtype = _val(v, "Type", default=-1)
        try:
            rect = _rect_mm(_val(v, "GetOutline"))
        except Exception:
            rect = None
        pos = _num_pair(_val(v, "Position"))
        out.append(
            {
                "name": name,
                "type": vtype,
                "is_sheet": vtype == 1,
                "model": _val(v, "GetReferencedModelName", default="") or "",
                "position_mm": [round(pos[0] * 1000, 2), round(pos[1] * 1000, 2)],
                "outline_mm": [round(x, 2) for x in rect] if rect else None,
                "size_mm": [round(rect[2] - rect[0], 2), round(rect[3] - rect[1], 2)] if rect else None,
                "scale": list(_num_pair(_val(v, "ScaleRatio"))),
                "uses_sheet_scale": bool(_val(v, "UseSheetScale", default=True)),
            }
        )
    return out


def find_overlaps(views: list[dict[str, Any]], tol_mm: float = 0.5) -> list[dict[str, Any]]:
    rects = [v for v in views if not v["is_sheet"] and v.get("outline_mm")]
    bad = []
    for i in range(len(rects)):
        for j in range(i + 1, len(rects)):
            a, b = rects[i]["outline_mm"], rects[j]["outline_mm"]
            w = min(a[2], b[2]) - max(a[0], b[0])
            h = min(a[3], b[3]) - max(a[1], b[1])
            if w > tol_mm and h > tol_mm:
                bad.append(
                    {
                        "a": rects[i]["name"],
                        "b": rects[j]["name"],
                        "overlap_mm": [round(w, 2), round(h, 2)],
                    }
                )
    return bad


def out_of_sheet(views, paper: str) -> list[dict[str, Any]]:
    w, h = paper_size(paper)
    bad = []
    for v in views:
        if v["is_sheet"] or not v.get("outline_mm"):
            continue
        x1, y1, x2, y2 = v["outline_mm"]
        if x1 < 0 or y1 < 0 or x2 > w or y2 > h:
            bad.append({"name": v["name"], "outline_mm": v["outline_mm"], "sheet_mm": [w, h]})
    return bad


def insert_standard_views(drw, model_path: str) -> bool:
    """第一角三视图。GB 只给第一角，不给第三角。"""
    for meth in ("Create1stAngleViews",):
        try:
            return bool(getattr(drw, meth)(model_path))
        except Exception:
            continue
    raise RuntimeError("Create1stAngleViews 调用失败")


# ---------------------------------------------------------------- 选比例
def pick_scale(block_w_mm: float, block_h_mm: float, paper: str) -> tuple[int, int]:
    """在 GB 优先序列里挑第一个放得下的比例。"""
    w, h = paper_size(paper)
    r = rules()
    m = r["paper"]["margin"]
    shr = r["layout"]["usable_shrink"]
    avail_w = w - m["left"] - m["other"] - shr.get("right", 0)
    avail_h = h - m["other"] - shr.get("bottom", 0) - shr.get("top", 0)
    gap = float(r["layout"]["gap"])
    for num, den in r["scale"]["auto_order"]:
        k = num / den
        if block_w_mm * k + gap <= avail_w and block_h_mm * k + gap <= avail_h:
            return int(num), int(den)
    return 1, 100


def _usable_rect(paper: str) -> tuple[float, float, float, float]:
    w, h = paper_size(paper)
    r = rules()
    m = r["paper"]["margin"]
    shr = r["layout"]["usable_shrink"]
    x0 = m["left"] + m["other"]
    y0 = m["other"] + shr.get("bottom", 0) + 45.0  # 给标题栏/分区号留位
    x1 = w - m["other"] - shr.get("right", 0)
    y1 = h - m["other"] - shr.get("top", 0)
    return x0, y0, x1, y1


def relayout_views(drw, paper: str) -> dict[str, Any]:
    """按第一角规矩重排：前视左上、俯视正下、侧视正右。"""
    views = [v for v in measure_views(drw) if not v["is_sheet"] and v.get("outline_mm")]
    if len(views) < 2:
        return {"ok": False, "reason": "视图不足"}
    base = max(views, key=lambda v: v["size_mm"][0] * v["size_mm"][1])
    others = [v for v in views if v is not base]
    below = [v for v in others if v["position_mm"][1] < base["position_mm"][1]]
    right = [v for v in others if v["position_mm"][1] >= base["position_mm"][1]]
    gap = float(rules()["layout"]["gap"])
    x0, y0, x1, y1 = _usable_rect(paper)
    bw = base["size_mm"][0] + (right[0]["size_mm"][0] + gap if right else 0)
    bh = base["size_mm"][1] + (below[0]["size_mm"][1] + gap if below else 0)
    ox = x0 + max(0.0, (x1 - x0 - bw) / 2.0)
    oy = y0 + max(0.0, (y1 - y0 - bh) / 2.0) + bh

    moved = []
    targets = [(base, ox + base["size_mm"][0] / 2, oy - base["size_mm"][1] / 2)]
    for v in below:
        vw = v["size_mm"][0]
        targets.append(
            (v, ox + base["size_mm"][0] / 2, oy - base["size_mm"][1] - gap - v["size_mm"][1] / 2)
        )
    for v in right:
        targets.append(
            (v, ox + base["size_mm"][0] + gap + v["size_mm"][0] / 2, oy - base["size_mm"][1] / 2)
        )
    for v, cx, cy in targets:
        for raw in _raw_view(drw, v["name"]):
            try:
                raw.Position = (cx * MM, cy * MM)
                moved.append({"view": v["name"], "to_mm": [round(cx, 2), round(cy, 2)]})
            except Exception as ex:
                moved.append({"view": v["name"], "error": str(ex)})
    return {"ok": True, "moved": moved}


def _raw_view(drw, name: str):
    for v in _dispatch_all(_val(drw, "GetViews")):
        if _val(v, "GetName2", default=None) == name:
            yield v


# ---------------------------------------------------------------- 导出
def save_as(drw, path: str) -> dict[str, Any]:
    if os.path.exists(path):
        os.remove(path)
    ext = drw.Extension
    errs = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warns = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    null = VARIANT(pythoncom.VT_DISPATCH, None)
    try:
        ok = ext.SaveAs(path, 0, 1, null, errs, warns)
    except Exception as ex:
        return {"ok": False, "error": str(ex), "path": path}
    return {
        "ok": bool(ok) and os.path.exists(path),
        "path": path,
        "bytes": os.path.getsize(path) if os.path.exists(path) else 0,
        "errors": int(errs.value) if hasattr(errs, "value") else None,
        "warnings": int(warns.value) if hasattr(warns, "value") else None,
    }


def export_pdf(sw, drw, path: str) -> dict[str, Any]:
    if os.path.exists(path):
        os.remove(path)
    data = None
    try:
        data = sw.GetExportFileData(1)  # swExportDataPdf
    except Exception:
        data = None
    ext = drw.Extension
    errs = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warns = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    try:
        ok = ext.SaveAs(path, 0, 1, data, errs, warns)
    except Exception as ex:
        return {"ok": False, "error": str(ex), "path": path}
    return {
        "ok": bool(ok) and os.path.exists(path),
        "path": path,
        "bytes": os.path.getsize(path) if os.path.exists(path) else 0,
    }


def _find_open(sw, path: str):
    """按路径找已经打开的文档。绝不因为「看起来是我们开的」就关用户的东西。"""
    target = os.path.normcase(os.path.abspath(path))
    for d in _dispatch_all(_val(sw, "GetDocuments", default=None)):
        p = _val(d, "GetPathName", default="") or ""
        if p and os.path.normcase(os.path.abspath(p)) == target:
            return d
    return None


def _close_quietly(sw, doc) -> None:
    """关掉自己新建的文档，不弹「是否保存」。"""
    try:
        doc.SetSaveFlag()
    except Exception:
        pass
    try:
        sw.CloseDoc(_val(doc, "GetTitle", default=""))
    except Exception:
        pass


# ---------------------------------------------------------------- 编排
def build(
    model_path: str,
    out_dir: str,
    paper: str = "A3",
    basename: str | None = None,
    relayout: bool = False,
    keep_open: bool = True,
    font: str | None = DEFAULT_FONT,
) -> dict[str, Any]:
    """从零件出一张 GB 图纸，返回可对账的报告。

    `font` 默认 `DEFAULT_FONT`（汉仪长仿宋体），通过**文档属性**写入图纸：
    标题栏那些 `GetUseDocTextFormat()=True` 的文字跟着一起变，不需要逐条改。
    传 `font=None` 则不动字体。

    注意 `relayout` 默认 False：本机 `IView.Position` 赋值行为异常
    （赋 (0.15,0.20) 回读成 (0.0,0.15)，且会把三个视图压到同一点），
    所以手动排布这条路暂时关掉，只靠「先定比例再插视图」让 SW 自己摆。
    """
    model_path = os.path.abspath(model_path)
    if not os.path.isfile(model_path):
        raise FileNotFoundError(model_path)
    os.makedirs(out_dir, exist_ok=True)
    base = basename or os.path.splitext(os.path.basename(model_path))[0]

    sw = attach_sw()

    # 1) 零件开成只读，只为拿路径；已打开就直接用，绝不动它
    opened_here = False
    part = _find_open(sw, model_path)
    if part is None:
        errs = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warns = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        part = sw.OpenDoc6(model_path, 1, 1, "", errs, warns)
        opened_here = part is not None

    try:
        # 2) 找 GB 模板
        tpl = template_for(paper)

        # 3) 第一趟：临时图纸，1:1 建视图，只为量出真实尺寸
        probe = sw.NewDocument(tpl, 0, 0, 0)
        if probe is None:
            raise RuntimeError(f"NewDocument 失败：{tpl}")
        try:
            insert_standard_views(probe, model_path)
            first = measure_views(probe)
            real = [v for v in first if not v["is_sheet"] and v.get("size_mm")]
            if not real:
                raise RuntimeError("没建出视图")
            base_v = max(real, key=lambda v: v["size_mm"][0] * v["size_mm"][1])
            others = [v for v in real if v is not base_v]
            gap = float(rules()["layout"]["gap"])
            block_w = base_v["size_mm"][0] + max([v["size_mm"][0] for v in others] or [0]) + gap
            block_h = base_v["size_mm"][1] + max([v["size_mm"][1] for v in others] or [0]) + gap
        finally:
            _close_quietly(sw, probe)

        num, den = pick_scale(block_w, block_h, paper)

        # 4) 第二趟：正式图纸。先在 1:1 之外把比例定好，再插视图——
        #    这样 SW 自己的排布就是按最终尺寸算的，不用手动挪（手动挪会跟投影对齐打架）
        drw = sw.NewDocument(tpl, 0, 0, 0)
        if drw is None:
            raise RuntimeError(f"NewDocument 失败：{tpl}")
        scale_ok = set_sheet_scale(drw, num, den) if (num, den) != (1, 1) else True
        insert_standard_views(drw, model_path)
        # 字体按文件属性设，视图插完就设好，后面的尺寸才会跟着走
        font_report = apply_font(drw, font) if font else None
        try:
            drw.ViewZoomtofit2()
        except Exception:
            pass

        views = measure_views(drw)
        overlaps = find_overlaps(views)
        laid = None
        if overlaps and relayout:  # 兜底：万一还是挤在一起，手动排
            laid = relayout_views(drw, paper)
            views = measure_views(drw)
            overlaps = find_overlaps(views)

        report = {
            "ok": True,
            "model": model_path,
            "template": tpl,
            "paper": paper,
            "sheet": sheet_state(drw),
            "scale": [num, den],
            "scale_applied": scale_ok,
            "font": font_report,
            "block_mm_1to1": [round(block_w, 2), round(block_h, 2)],
            "views": views,
            "overlaps": find_overlaps(views),
            "out_of_sheet": out_of_sheet(views, paper),
            "relayout": laid,
        }

        # 5) 落盘
        slddrw = os.path.join(out_dir, f"{base}_{paper}.SLDDRW")
        pdf = os.path.join(out_dir, f"{base}_{paper}.pdf")
        report["slddrw"] = save_as(drw, slddrw)
        report["pdf"] = export_pdf(sw, drw, pdf)
        if not keep_open:
            try:
                drw.SetSaveFlag()
                sw.CloseDoc(_val(drw, "GetTitle"))
            except Exception:
                pass
        return report
    finally:
        if opened_here:
            try:
                part.SetSaveFlag()
                sw.CloseDoc(_val(part, "GetTitle", default=os.path.basename(model_path)))
            except Exception:
                pass


def view_objects(drw) -> list:
    """图纸里的**模型视图**（跳过图纸本身）。"""
    out = []
    for v in _dispatch_all(_val(drw, "GetViews")) or []:
        try:
            if int(_val(v, "Type", default=-1)) == 1:  # 1 = 图纸
                continue
        except Exception:  # noqa: BLE001
            continue
        out.append(v)
    return out


def display_dimensions(drw) -> list[dict[str, Any]]:
    """列出所有视图里的尺寸/标注：名字、数值、是否孔标注、文本位置(mm)。

    `GetDimension2(0)` 拿到的 IDimension.Value **单位是文档单位(mm)**，不是米 ——
    按米去比对会全部对不上。文本位置从 `IDisplayDimension.GetAnnotation().GetPosition()`
    取，单位米。刚 `OpenDoc6` 完立刻读可能读到空，等 1 秒再读。
    """
    out: list[dict[str, Any]] = []
    for vi, v in enumerate(view_objects(drw)):
        for dd in _dispatch_all(_val(v, "GetDisplayDimensions") or []) or []:
            dim = _val(dd, "GetDimension2", 0)
            if dim is None:
                continue
            ann = _val(dd, "GetAnnotation")
            pos = _val(ann, "GetPosition") if ann is not None else None
            out.append({
                "view": vi,
                "view_name": _val(v, "GetName2"),
                "name": str(_val(dim, "FullName")),
                "short": str(_val(dim, "FullName")).split("@")[0],
                "value": _val(dim, "Value"),
                "is_hole_callout": bool(_val(dd, "IsHoleCallout")),
                "position_mm": [round(float(c) / MM, 3) for c in pos] if pos else None,
                "_dd": dd,
                "_dim": dim,
                "_ann": ann,
            })
    return out


def find_dimension(drw, value_mm: float | None = None, name_prefix: str | None = None,
                   tol: float = 1e-4, hole_callout: bool | None = None):
    """按数值(mm)或名字前缀找第一条尺寸，返回 display_dimensions() 的那条 dict。"""
    for item in display_dimensions(drw):
        if hole_callout is not None and item["is_hole_callout"] != hole_callout:
            continue
        if name_prefix and not item["name"].startswith(name_prefix):
            continue
        if value_mm is not None:
            v = item["value"]
            if v is None or abs(float(v) - float(value_mm)) > tol:
                continue
        return item
    return None


def delete_annotation(drw, name: str) -> bool:
    """按名字删一条尺寸/标注（名字形如 `RD5@工程图视图1`）。

    坑：`IAnnotation.Select3(False, NULL)` 之后 `EditDelete` 对图纸尺寸不生效；
    必须先 `Extension.SelectByID2(name, "DIMENSION", ...)` 选中（这个能选中），
    再调 `EditDelete`。它返回值不可靠，删没删要以重新列一遍为准。
    """
    _val(drw, "ClearSelection2", True)
    ok = drw.Extension.SelectByID2(name, "DIMENSION", 0, 0, 0, False, 0,
                                   VARIANT(pythoncom.VT_DISPATCH, None), 0)
    if not ok:
        return False
    _val(drw, "EditDelete")
    _val(drw, "ClearSelection2", True)
    return True


def add_hole_callout(drw, edge_point_mm: list[float], text_point_mm: list[float]):
    """选孔的一条圆边，在图纸上生成**孔标注**（螺纹孔会带 M8-6H 那种标注）。

    - `edge_point_mm`：图纸坐标(mm)，落在孔的圆边上（如视图中心上方半径处）
    - `text_point_mm`：标注文字落点(mm)
    用 `IDrawingDoc.AddHoleCallout2`，前提是孔的圆边已选中；标注内容由模型里的
    异型孔特征决定，所以先得把孔建成真的 `HoleWzd`（见 sw_features.hole_wizard）。
    """
    _val(drw, "ClearSelection2", True)
    picked = drw.Extension.SelectByID2("", "EDGE", edge_point_mm[0] * MM,
                                       edge_point_mm[1] * MM, 0, False, 0,
                                       VARIANT(pythoncom.VT_DISPATCH, None), 0)
    if not picked:
        raise RuntimeError(f"HOLECALLOUT_NO_EDGE：({edge_point_mm}) 处没选中孔的圆边")
    info = _val(drw, "AddHoleCallout2", text_point_mm[0] * MM, text_point_mm[1] * MM, 0.0)
    _val(drw, "ClearSelection2", True)
    if info is None:
        raise RuntimeError("HOLECALLOUT_FAILED：AddHoleCallout2 返回空（孔边是不是真异型孔？）")
    return True


def set_dimension_tolerance(drw, name: str, plus_mm: float, minus_mm: float | None = None,
                            symmetric: bool = False) -> dict[str, Any]:
    """给尺寸加公差。`symmetric=True` 时 `plus_mm` 就是 ±值（忽略 minus_mm）。

    坑（2026-11-21 实测 SW2024）：
      1. `IDimensionTolerance.SetValues2(...)` **一律返回 False、值进不去**；
         能生效的是老接口 `SetValues(MinValue, MaxValue)`。
      2. 参数顺序是**先最小后最大**，单位**米**：+0.2/-0.5 → SetValues(-0.0005, 0.0002)。
      3. 对称公差用 `Type=4 (swTolSYMMETRIC)`，双向公差用 `Type=2 (swTolBILAT)`。
      4. 尺寸数值本身是 mm，公差值是 m，两个单位不一样，别混。
    """
    item = find_dimension(drw, name_prefix=name)
    if item is None:
        raise RuntimeError(f"TOLERANCE_NO_DIM：找不到尺寸 {name}")
    tol = _val(item["_dim"], "Tolerance")
    if tol is None:
        raise RuntimeError(f"TOLERANCE_NO_OBJECT：{name} 上没有公差对象")
    if symmetric:
        lo = hi = abs(float(plus_mm))
    else:
        hi = abs(float(plus_mm))
        lo = -abs(float(minus_mm if minus_mm is not None else plus_mm))
    tol.Type = 4 if symmetric else 2
    _value_call = _val(tol, "SetValues", lo * MM, hi * MM)
    read = {"type": _val(tol, "Type"), "max_mm": _val(tol, "GetMaxValue") / MM,
            "min_mm": _val(tol, "GetMinValue") / MM, "set_ok": _value_call}
    return read


if __name__ == "__main__":
    import sys

    sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(build(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "_tmp/drawing_out"),
                     ensure_ascii=False, indent=2))
