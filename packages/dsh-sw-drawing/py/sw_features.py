# -*- coding: utf-8 -*-
"""P0 三维能力：在实体面上画草图 / 特征改名 / 线性阵列。

与 sw_com 的分工：sw_com 管会话、单位、草图绘制原语、拉伸切除；
这里只做「选面」「改名」「阵列」三件 sw_com 还没有的事，复用它的原语。

真机校准（SW2024 SP0.1 / interop 32.0.1.4，2026-11-20）：
  1. 面对象没有 Select2 的把握，优先 `IEntity.Select4(Append, Callout)`；
     Callout 传 None 会 DISP_E_TYPEMISMATCH，必须传 `VARIANT(VT_DISPATCH, None)`。
  2. 草图里 CreateLine/CreateCircle 收的是**模型坐标**（SW 会投影到草图平面），
     所以「在某个面上画」只要把面选对，坐标仍按模型空间 mm 写。
  3. `FeatureLinearPattern4` 精确签名（反射得到，20 个参数）：
     (Num1, Spacing1, Num2, Spacing2, FlipDir1, FlipDir2, DName1, DName2,
      GeometryPattern, VaryInstance, HasOffset1, HasOffset2, CtrlByNum1, CtrlByNum2,
      FromCentroid1, FromCentroid2, RevOffset1, RevOffset2, Offset1, Offset2)
"""

from __future__ import annotations

import math
import re
from typing import Any

import sw_com as C

try:
    import win32com.client as wc
except Exception:  # pragma: no cover
    wc = None

MM = C.MM


def _wrap(obj):
    """COM 数组元素常常是原始 PyIDispatch，不包一层成员全找不到。"""
    if wc is not None and "PyIDispatch" in type(obj).__name__:
        try:
            return wc.Dispatch(obj)
        except Exception:  # noqa: BLE001
            return obj
    return obj


def _null():
    return C.VARIANT(C.pythoncom.VT_DISPATCH, None)


def _r(v, nd: int = 3) -> list[float]:
    return [round(float(x), 4) for x in list(v)[:nd]]


# ---------------------------------------------------------------- 面
def solid_faces(doc) -> list[dict[str, Any]]:
    """枚举实体面。坐标一律 mm。"""
    C._require_part(doc, "读面")
    try:
        bodies = C._value(doc, "GetBodies2", 0, True) or []
    except Exception as exc:  # noqa: BLE001
        raise C.SwError(f"COM_ERROR：读不到实体（{exc}）") from exc
    out: list[dict[str, Any]] = []
    for bi, body in enumerate(bodies):
        try:
            faces = C._value(body, "GetFaces") or []
        except Exception:  # noqa: BLE001
            continue
        for fi, face in enumerate(faces):
            rec: dict[str, Any] = {"body": bi, "index": fi}
            try:
                box = list(C._value(face, "GetBox"))
                if len(box) >= 6:
                    rec["box_mm"] = [round(float(v) / MM, 3) for v in box]
                    rec["center_mm"] = [
                        round((box[i] + box[i + 3]) / 2.0 / MM, 3) for i in range(3)
                    ]
            except Exception:  # noqa: BLE001
                pass
            try:
                n = list(C._value(face, "Normal"))
                if len(n) >= 3:
                    rec["normal"] = _r(n)
            except Exception:  # noqa: BLE001
                pass
            try:
                rec["area_mm2"] = round(float(C._value(face, "GetArea")) / (MM * MM), 2)
            except Exception:  # noqa: BLE001
                pass
            rec["_face"] = face
            out.append(rec)
    return out


def find_face(doc, point_mm: list[float] | None = None, normal: list[float] | None = None):
    """按「法向 + 离某点最近」挑一个面。返回 (face 对象, 面信息)。"""
    faces = solid_faces(doc)
    if not faces:
        raise C.SwError("COM_NULL：零件里没有实体面")
    cands = faces
    if normal:
        nv = [float(v) for v in normal]
        ln = math.sqrt(sum(v * v for v in nv)) or 1.0
        nv = [v / ln for v in nv]

        def dot(f):
            n = f.get("normal")
            if not n:
                return -2.0
            return sum(n[i] * nv[i] for i in range(3))

        cands = [f for f in cands if dot(f) > 0.98]
        if not cands:
            avail = sorted({str(f.get("normal")) for f in faces})
            raise C.SwError(f"COM_NULL：找不到法向≈{normal} 的面；现有法向 {avail[:6]}")
    if point_mm:
        p = [float(v) for v in point_mm]

        def dist(f):
            c = f.get("center_mm") or [1e9, 1e9, 1e9]
            return math.dist(p, c)

        cands = sorted(cands, key=dist)
    else:
        cands = sorted(cands, key=lambda f: -(f.get("area_mm2") or 0.0))
    return cands[0]["_face"], cands[0]


def _select_face(face) -> bool:
    try:
        if face.Select4(False, _null()):
            return True
    except Exception:  # noqa: BLE001
        pass
    try:
        return bool(face.Select2(False, 0))
    except Exception:  # noqa: BLE001
        return False


_POINT_KEYS = ("from", "to", "corner", "opposite", "center", "through")


def _to_sketch_space(entities, mat):
    """模型坐标 mm → 草图坐标 mm。

    真机实测：在实体面上画草图时，`CreateCircle` / `CreateLine` 收的是**草图坐标系**，
    不是模型坐标系。Ø20 的孔按模型坐标写就落到板外，`FeatureCut4` 一个特征都不出；
    Ø60 的大圆碰巧还压在板上才切成了。

    换算走 `ISketch.ModelToSketchTransform.ArrayData`（16 个 double，行主序、行向量）。
    本机 `IMathUtility` 取不到（IModelDoc2 和 ISldWorks 上都没有），所以矩阵自己在
    Python 里乘，不再依赖 IMathPoint/IMathTransform 的 COM 往返。
    返回仍是 mm，让 `_draw_entities` 自己的 MM 换算保持不变。
    """
    m = [float(v) for v in list(mat)]
    if len(m) != 16:
        raise C.SwError(f"COM_ERROR：ModelToSketchTransform 长度 {len(m)}，应当 16")

    def tp(p):
        vals = [float(v) for v in p] + [0.0] * (3 - len(list(p)))
        x, y, z = vals[0] * MM, vals[1] * MM, vals[2] * MM
        return [
            (x * m[0] + y * m[4] + z * m[8] + m[12]) / MM,
            (x * m[1] + y * m[5] + z * m[9] + m[13]) / MM,
            (x * m[2] + y * m[6] + z * m[10] + m[14]) / MM,
        ]

    out = []
    for e in entities:
        e2 = dict(e)
        for k in _POINT_KEYS:
            if isinstance(e.get(k), (list, tuple)):
                e2[k] = tp(e[k])
        if isinstance(e.get("points"), list):
            e2["points"] = [tp(p) for p in e["points"]]
        out.append(e2)
    return out


def pick_edges_by_ray(doc, edges: list[dict[str, Any]]) -> int:
    """用射线选边。edges 每项：{"point":[x,y,z], "dir":[dx,dy,dz], "radiusMm":0.5}

    为什么用射线：本机 `body.GetEdges()` 返回的边对象成员全取不到，
    但 `SelectByRay(..., TypeWanted=1 /*EDGE*/)` 能稳定选中（实测 4/4）。
    """
    if not edges:
        raise C.SwError("BAD_REQUEST：edges 不能为空")
    ext = doc.Extension
    C._clear_selection(doc)
    for i, e in enumerate(edges):
        p = e.get("point") or []
        d = e.get("dir") or [0, 1, 0]
        if len(p) != 3:
            raise C.SwError(f"BAD_REQUEST：edges[{i}].point 要 3 个数")
        rad = float(e.get("radiusMm", 0.5))
        ok = ext.SelectByRay(
            float(p[0]) * MM, float(p[1]) * MM, float(p[2]) * MM,
            float(d[0]), float(d[1]), float(d[2]),
            rad * MM, 1, bool(i), 0, 0,
        )
        if not ok:
            raise C.SwError(
                f"COM_ERROR：射线没打到第 {i + 1} 条边（point={p}, dir={d}）。"
                "点要落在实体外面朝边打，半径默认 0.5mm"
            )
    return len(edges)


def fillet(doc, radius_mm: float, edges: list[dict[str, Any]]) -> dict[str, Any]:
    """等半径圆角。`Options` 必须传 **2**，传 0/1 一律返回空（真机实测）。"""
    if float(radius_mm) <= 0:
        raise C.SwError("BAD_REQUEST：圆角半径必须大于 0")
    n = pick_edges_by_ray(doc, edges)
    fm = doc.FeatureManager
    try:
        feat = fm.FeatureFillet3(
            2, float(radius_mm) * MM, 0.0, 0.0, 0, 0, 0,
            None, None, None, None, None, None, None,
        )
    except Exception as exc:  # noqa: BLE001
        raise C.SwError(f"COM_ERROR：FeatureFillet3 失败（{type(exc).__name__}: {exc}）") from exc
    if feat is None:
        raise C.SwError("COM_NULL：圆角没建出来（Options 要 2；半径不要超过相邻面尺寸）")
    return {"op": "fillet", "name": str(C._value(feat, "Name")),
            "radius_mm": float(radius_mm), "edges": n}


def chamfer(doc, width_mm: float, edges: list[dict[str, Any]],
            angle_deg: float = 45.0) -> dict[str, Any]:
    """等距倒角。`ChamferType` 必须传 **1**；传 0 会长出特征但**不切任何材料**（很坑，真机实测）。"""
    if float(width_mm) <= 0:
        raise C.SwError("BAD_REQUEST：倒角宽度必须大于 0")
    n = pick_edges_by_ray(doc, edges)
    fm = doc.FeatureManager
    try:
        feat = fm.InsertFeatureChamfer(
            0, 1, float(width_mm) * MM, math.radians(float(angle_deg)),
            float(width_mm) * MM, 0.0, 0.0, 0.0,
        )
    except Exception as exc:  # noqa: BLE001
        raise C.SwError(f"COM_ERROR：InsertFeatureChamfer 失败（{type(exc).__name__}: {exc}）") from exc
    if feat is None:
        raise C.SwError("COM_NULL：倒角没建出来（ChamferType 要 1）")
    return {"op": "chamfer", "name": str(C._value(feat, "Name")),
            "width_mm": float(width_mm), "angle_deg": float(angle_deg), "edges": n}


def _in_plane_axes(normal):
    """面内两个轴：取与法向最不平行的两个坐标轴。"""
    n = [abs(float(v)) for v in normal] + [0.0] * (3 - len(normal))
    order = sorted(range(3), key=lambda i: n[i])
    return order[0], order[1]


def boundary_warnings(face_info: dict[str, Any], entities: list[dict[str, Any]]) -> list[str]:
    """圆孔是否与实体边相切。

    真机实测：孔边与实体边正好相切时 `FeatureCut4` 一律返回空，而且只报
    「草图没有封闭轮廓」，完全误导。差 0.1 mm 就正常。这里提前拦。
    """
    box = face_info.get("box_mm")
    n = face_info.get("normal")
    if not box or not n or len(box) < 6:
        return []
    a, b = _in_plane_axes(n)
    lo = {a: box[a], b: box[b]}
    hi = {a: box[a + 3], b: box[b + 3]}
    out: list[str] = []
    for e in entities or []:
        if not isinstance(e, dict) or str(e.get("type", "")).lower() != "circle":
            continue
        c = e.get("center") or []
        if len(c) < 2:
            continue
        if e.get("diameter") is not None:
            r = float(e["diameter"]) / 2.0
        elif e.get("radius") is not None:
            r = float(e["radius"])
        else:
            continue
        if r <= 0:
            continue
        for ax, name in ((a, "xyz"[a]), (b, "xyz"[b])):
            d = min(abs(float(c[ax]) - lo[ax]), abs(hi[ax] - float(c[ax])))
            if abs(d - r) < 0.05:
                out.append(
                    f"约束警告：圆 ({c[0]}, {c[1]}) R{r:g} 与实体边相切（{name} 向净距 {d:.3f} mm）。"
                    f"SW 会拒绝出特征，请把圆往里收至少 0.1 mm"
                )
    return out


def sketch_on_face(
    doc,
    entities: list[dict[str, Any]],
    point_mm: list[float] | None = None,
    normal: list[float] | None = None,
    use_transform: bool = False,
) -> dict[str, Any]:
    """在实体面上建草图并退出。点坐标仍按模型空间 mm 写。"""
    C._require_part(doc, "面草图")
    face, info = find_face(doc, point_mm=point_mm, normal=normal)
    C._clear_selection(doc)
    if not _select_face(face):
        raise C.SwError(
            "COM_ERROR：选不中面。模型可能被压缩、或 SW 里有未完成的操作"
        )
    before = set(C.feature_names(doc))
    sm = doc.SketchManager
    try:
        C._value(sm, "InsertSketch", True)
    except Exception as exc:  # noqa: BLE001
        raise C.SwError(f"COM_ERROR：进不了草图（{type(exc).__name__}: {exc}）") from exc
    try:
        active = C._value(sm, "ActiveSketch")
    except Exception:  # noqa: BLE001
        active = None
    if active is None:
        raise C.SwError("COM_ERROR：面选中了但草图没开始（SW 可能在等对话框）")
    # 模型坐标 → 草图坐标（不过这一步，孔会画到实体外面去）
    if use_transform:
        try:
            mat = list(C._value(active, "ModelToSketchTransform").ArrayData)
            entities = _to_sketch_space(entities or [], mat)
        except C.SwError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise C.SwError(
                f"COM_ERROR：取不到草图的 ModelToSketchTransform，"
                f"不换算会把几何画到实体外面（{type(exc).__name__}: {exc}）"
            ) from exc
    try:
        counts = C._draw_entities(sm, entities or [])
    except C.SwError:
        try:
            C._value(sm, "InsertSketch", True)
        except Exception:  # noqa: BLE001
            pass
        leftover = [n for n in C.feature_names(doc) if n not in before]
        if leftover:
            C.rollback(doc, leftover)
        raise
    C._value(sm, "InsertSketch", True)
    created = [n for n in C.feature_names(doc) if n not in before]
    name = ""
    for cand in created:
        feat = C.find_feature(doc, cand)
        try:
            if feat is not None and str(C._value(feat, "GetTypeName2")) == "ProfileFeature":
                name = cand
                break
        except Exception:  # noqa: BLE001
            continue
    if not name and created:
        name = created[0]
    if not name:
        raise C.SwError("COM_ERROR：草图退出后特征树里没看到新草图")
    # 回读对账：SW 会在实体外画完也不报错，这里把落点交出来好判真假
    pts: list[list[float]] = []
    try:
        sk_obj = C._value(C.find_feature(doc, name), "GetSpecificFeature2")
        for q in (C._value(sk_obj, "GetSketchPoints2") or [])[:20]:
            pts.append([round(float(getattr(q, a, 0.0)) / MM, 3) for a in ("X", "Y", "Z")])
    except Exception:  # noqa: BLE001
        pass
    return {
        "op": "sketch_on_face",
        "name": name,
        "face": {
            "normal": info.get("normal"),
            "center_mm": info.get("center_mm"),
            "area_mm2": info.get("area_mm2"),
        },
        "entities": counts,
        "segments": sum(counts.values()),
        "points_mm": pts,
        "used_transform": bool(use_transform),
        "warnings": boundary_warnings(info, entities or []),
    }


# ---------------------------------------------------------------- 改名
def rename_feature(doc, old: str, new: str) -> dict[str, Any]:
    """给特征起语义名（用户零件里就是 `功能牌1第三安装孔` 这种）。"""
    new = (new or "").strip()
    if not new:
        raise C.SwError("BAD_REQUEST：新名字不能为空")
    feat = C.find_feature(doc, old)
    if feat is None:
        raise C.SwError(f"COM_NULL：特征树里没有 {old!r}")
    try:
        feat.Name = new
    except Exception as exc:  # noqa: BLE001
        raise C.SwError(f"COM_ERROR：改名失败（{type(exc).__name__}: {exc}）") from exc
    names = C.feature_names(doc)
    return {"op": "rename", "old": old, "name": new, "ok": new in names}


# ---------------------------------------------------------------- 阵列
def body_edges(doc) -> list[dict[str, Any]]:
    """枚举实体边（模型坐标 mm），给阵列方向用。"""
    C._require_part(doc, "读边")
    out: list[dict[str, Any]] = []
    try:
        bodies = C._value(doc, "GetBodies2", 0, True) or []
    except Exception:  # noqa: BLE001
        return out
    for bi, body in enumerate(bodies):
        try:
            edges = C._value(body, "GetEdges") or []
        except Exception:  # noqa: BLE001
            continue
        for ei, e in enumerate(edges):
            e = _wrap(e)
            rec: dict[str, Any] = {"body": bi, "index": ei}
            cp = None
            try:
                # 本机 GetCurveParams3 是属性不是方法，直接调用会「找不到成员」
                cp = C._value(e, "GetCurveParams3")
            except Exception:  # noqa: BLE001
                cp = None
            try:
                if cp is not None and len(list(cp)) >= 6:
                    arr = [float(v) for v in list(cp)]
                    rec["start_mm"] = [round(v / MM, 3) for v in arr[0:3]]
                    rec["end_mm"] = [round(v / MM, 3) for v in arr[3:6]]
                else:  # 退路：用两个顶点
                    a = C._value(C._value(e, "GetStartVertex"), "GetPoint")
                    b = C._value(C._value(e, "GetEndVertex"), "GetPoint")
                    rec["start_mm"] = [round(float(v) / MM, 3) for v in list(a)[:3]]
                    rec["end_mm"] = [round(float(v) / MM, 3) for v in list(b)[:3]]
                d = [rec["end_mm"][i] - rec["start_mm"][i] for i in range(3)]
                ln = math.sqrt(sum(v * v for v in d)) or 1.0
                rec["dir"] = [round(v / ln, 4) for v in d]
                rec["length_mm"] = round(ln, 3)
            except Exception:  # noqa: BLE001
                continue
            rec["_edge"] = e
            out.append(rec)
    return out


_PLANE_NORMALS = (("front", (0.0, 0.0, 1.0)), ("top", (0.0, 1.0, 0.0)), ("right", (1.0, 0.0, 0.0)))


def _plane_for_direction(nv):
    """挑一个法向与方向最垂直的基准面，直线才画得进去。"""
    best = None
    for name, n in _PLANE_NORMALS:
        d = abs(sum(n[i] * nv[i] for i in range(3)))
        if best is None or d < best[1]:
            best = (name, d)
    return best[0]


def _plane_for_directions(*dirs):
    """挑一个法向同时与所有方向垂直的基准面，几条方向线才都画得进去。"""
    if len(dirs) == 1:
        return _plane_for_direction(dirs[0])
    a, b = dirs[0], dirs[1]
    cross = [
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    ]
    ln = math.sqrt(sum(v * v for v in cross))
    if ln < 1e-9:
        return _plane_for_direction(a)   # 两方向平行，退化成单方向
    cross = [v / ln for v in cross]
    best = None
    for name, n in _PLANE_NORMALS:
        d = abs(sum(n[i] * cross[i] for i in range(3)))
        if best is None or d > best[1]:
            best = (name, d)
    return best[0]


def linear_pattern(
    doc,
    seeds: list[str],
    direction: list[float],
    spacing_mm: float,
    count: int,
    flip: bool = False,
    geometry_pattern: bool = True,
    helper_name: str = "阵列方向",
    direction2: list[float] | None = None,
    spacing2_mm: float = 0.0,
    count2: int = 1,
    flip2: bool = False,
) -> dict[str, Any]:
    """沿一个（或两个）方向做线性阵列。

    真机校准（2026-11-20，SW2024）：
      1. 成败全在**选择标记**：要阵列的特征 `Select2(Append, 4)`；第一方向实体 `Select2(Append, 1)`；
         **第二方向实体 `Select2(Append, 2)`**（实测：3×20 + 2×15 → 6 个孔，孔位和体积全对）。
      2. `Select4(Append, Callout)` 没有 mark 参数，用它选方向 `FeatureLinearPattern4` 一律返回空（不报错）。
      3. 方向实体用「辅助草图里的一条直线」，因为本机 `body.GetEdges()` 返回的边对象成员全取不到。
         两方向就画两条线在一个草图里，分别按 mark 1 / mark 2 选。
      4. `FeatureLinearPattern4` 20 个参数：
         (Num1, Spacing1, Num2, Spacing2, FlipDir1, FlipDir2, DName1, DName2,
          GeometryPattern, VaryInstance, HasOffset1, HasOffset2, CtrlByNum1, CtrlByNum2,
          FromCentroid1, FromCentroid2, RevOffset1, RevOffset2, Offset1, Offset2)
    """
    if count < 2:
        raise C.SwError("BAD_REQUEST：count 至少 2")
    if spacing_mm <= 0:
        raise C.SwError("BAD_REQUEST：spacing 必须大于 0")
    two_dir = direction2 is not None and int(count2) >= 2
    if two_dir and float(spacing2_mm) <= 0:
        raise C.SwError("BAD_REQUEST：第二方向需要 spacing2Mm > 0")

    def _unit(v):
        ln = math.sqrt(sum(x * x for x in v)) or 1.0
        return [x / ln for x in v]

    nv = _unit([float(v) for v in direction])
    nv2 = _unit([float(v) for v in direction2]) if two_dir else None

    for n in seeds:
        if C.find_feature(doc, n) is None:
            raise C.SwError(f"COM_NULL：特征树里没有 {n!r}")

    # 1) 先画方向辅助草图（插草图会清选择，所以必须排在选种子之前）
    L = 50.0
    line1 = {"type": "line", "from": [0, 0, 0], "to": [nv[0] * L, nv[1] * L, nv[2] * L]}
    entities = [line1]
    if two_dir:
        entities.append({"type": "line", "from": [0, 0, 0],
                         "to": [nv2[0] * L, nv2[1] * L, nv2[2] * L]})
        plane = _plane_for_directions(nv, nv2)
    else:
        plane = _plane_for_directions(nv)
    helper = C.sketch(doc, on=plane, entities=entities)
    sk_obj = C._value(C.find_feature(doc, helper["name"]), "GetSpecificFeature2")
    segs = [_wrap(s) for s in (C._value(sk_obj, "GetSketchSegments") or [])]
    if not segs:
        raise C.SwError("COM_ERROR：辅助草图里没拿到线段")

    # 2) 选种子（mark 4）+ 选方向（mark 1 / mark 2）
    C._clear_selection(doc)
    for i, name in enumerate(seeds):
        f = C.find_feature(doc, name)
        if not f.Select2(bool(i), 4):
            raise C.SwError(f"COM_ERROR：选不中要阵列的特征 {name!r}")
    picked = False
    for call in (lambda: segs[0].Select2(True, 1), lambda: segs[0].Select2(True, 1)):
        try:
            if call():
                picked = True
                break
        except Exception:  # noqa: BLE001
            continue
    if not picked:
        raise C.SwError("COM_ERROR：选不中方向线段（mark 1）")
    if two_dir:
        if len(segs) < 2:
            raise C.SwError("COM_ERROR：辅助草图里只拿到 1 段，第二条方向线没画出来")
        if not segs[1].Select2(True, 2):
            raise C.SwError("COM_ERROR：选不中第二方向线段（mark 2）")

    fm = doc.FeatureManager
    feat = None
    try:
        feat = fm.FeatureLinearPattern4(
            int(count), float(spacing_mm) * MM,
            int(count2) if two_dir else 1, float(spacing2_mm) * MM if two_dir else 0.0,
            bool(flip), bool(flip2),
            "", "",
            bool(geometry_pattern), False,
            False, False, True, True, False, False, False, False, 0.0, 0.0,
        )
    except Exception as exc:  # noqa: BLE001
        raise C.SwError(f"COM_ERROR：FeatureLinearPattern4 失败（{type(exc).__name__}: {exc}）") from exc
    if feat is None:
        raise C.SwError(
            "COM_NULL：阵列没建出来。常见原因：种子/方向的选择标记不对（种子 4、方向1 用 1、方向2 用 2），"
            "或方向与种子在同一特征里"
        )
    # 3) 给方向草图一个人能看懂的名字
    try:
        C.find_feature(doc, helper["name"]).Name = helper_name
    except Exception:  # noqa: BLE001
        pass
    return {
        "op": "linear_pattern",
        "name": str(C._value(feat, "Name")),
        "seeds": seeds,
        "count": int(count),
        "spacing_mm": float(spacing_mm),
        "direction": [round(v, 4) for v in nv],
        "count2": int(count2) if two_dir else 1,
        "spacing2_mm": float(spacing2_mm) if two_dir else 0.0,
        "direction2": [round(v, 4) for v in nv2] if two_dir else None,
        "helper_sketch": helper_name,
        "helper_plane": plane,
    }


# ---------------------------------------------------------------- 抽壳
def shell(
    doc,
    thickness_mm: float,
    faces: list[dict[str, Any]] | None = None,
    outward: bool = False,
) -> dict[str, Any]:
    """抽壳。`faces` 里列的是要「开口去掉」的面（点 + 法向，交给 find_face 定位）；
    留空 = 整块挖成闭合空腔。

    真机校准（SW2024 SP0.1，2026-11-20）：
      1. 接口是 `IModelDoc2.InsertFeatureShell(Thickness, Outward)`，**不是** FeatureManager
         的方法（`IFeatureManager` 里只有个 `GetPlasticsShellType`，别被名字骗了）。
      2. 返回值不可靠——成功也返回 `None`，所以用特征树前后差集找新特征，
         并且按 SW 的类型名 `Shell` 认（跟界面语言无关）。
      3. 厚度是「往里挖」，`outward=True` 变成「往外长」。
      4. 100×60×10 板实测：闭合 27744 / 顶面开口 16992 / 上下开口 6240 / 向外 33184（mm³）。
      5. 多个开口面：第一个用 `Select4(False, None)`，其余用 `Select4(True, None)`（追加）。
    """
    if float(thickness_mm) <= 0:
        raise C.SwError("BAD_REQUEST：抽壳厚度必须大于 0")
    C._clear_selection(doc)
    opened: list[list[float]] = []
    for i, sel in enumerate(faces or []):
        point = sel.get("point")
        normal = sel.get("normal")
        face, info = find_face(doc, point_mm=point, normal=normal)
        ok = _select_face(face) if i == 0 else bool(face.Select4(True, _null()))
        if not ok:
            raise C.SwError(
                f"COM_ERROR：第 {i + 1} 个开口面选不中（point={point}, normal={normal}）"
            )
        opened.append(_r(info.get("center_mm") or []))

    before = {item[0] for item in C.tree_items(doc)}
    try:
        doc.InsertFeatureShell(float(thickness_mm) * MM, bool(outward))
    except Exception as exc:  # noqa: BLE001
        raise C.SwError(
            f"COM_ERROR：InsertFeatureShell 失败（{type(exc).__name__}: {exc}）"
        ) from exc
    new = [item for item in C.tree_items(doc) if item[0] not in before]
    feat = next((it for it in new if it[1] == "Shell"), None) or (new[0] if new else None)
    if feat is None:
        raise C.SwError(
            "COM_NULL：抽壳没建出来。厚度是不是超过零件最薄处？或者选面时把整个实体选中了？"
        )
    return {
        "op": "shell",
        "name": feat[0],
        "thickness_mm": float(thickness_mm),
        "outward": bool(outward),
        "open_faces": opened,
        "volume_mm3": C.mass_properties(doc).get("volume_mm3"),
    }


# ---------------------------------------------------------------- 参考轴 / 圆周阵列
_BASE_PLANE_KEYS = ("front", "top", "right")
# 两个基准面的交线 = 某根模型坐标轴（顺序按树里的前视/上视/右视）
_AXIS_PLANE_PAIRS = {
    "x": ("top", "front"),
    "y": ("right", "front"),
    "z": ("right", "top"),
}


def reference_axis(doc, kind: str = "z", name: str = "") -> dict[str, Any]:
    """用两个基准面的交线建一根参考轴。

    真机校准：先对两个基准面 `Select2(True, 0)`，再 `IModelDoc2.InsertAxis2(True)`；
    实测 z（右视 ∩ 上视）能建出来，而且能直接当圆周阵列的轴用。
    """
    key = (kind or "z").strip().lower()
    if key not in _AXIS_PLANE_PAIRS:
        raise C.SwError("BAD_REQUEST：axis 只能是 x / y / z")
    planes = dict(zip(_BASE_PLANE_KEYS, C.base_planes(doc)))
    pair = [(k, planes.get(k)) for k in _AXIS_PLANE_PAIRS[key]]
    missing = [k for k, v in pair if v is None]
    if missing:
        raise C.SwError(f"COM_NULL：零件里找不到基准面 {missing}")

    C._clear_selection(doc)
    for k, (plane_name, plane_feat) in pair:
        if not plane_feat.Select2(True, 0):
            raise C.SwError(f"COM_ERROR：选不中基准面 {plane_name}")
    before = {item[0] for item in C.tree_items(doc)}
    try:
        doc.InsertAxis2(True)
    except Exception as exc:  # noqa: BLE001
        raise C.SwError(f"COM_ERROR：InsertAxis2 失败（{type(exc).__name__}: {exc}）") from exc
    new = [item for item in C.tree_items(doc) if item[0] not in before]
    feat = next((it for it in new if it[1] in ("RefAxis", "Axis")), None) or (
        new[0] if new else None
    )
    if feat is None:
        raise C.SwError("COM_NULL：参考轴没建出来（两个基准面得相交）")
    axis_name = feat[0]
    if name and name != axis_name:
        rename_feature(doc, axis_name, name)
        axis_name = name
    return {
        "op": "reference_axis",
        "name": axis_name,
        "kind": key,
        "planes": [v[0] for _k, v in pair],
    }


def circular_pattern(
    doc,
    seeds: list[str],
    count: int,
    angle_deg: float | None = None,
    equal_spacing: bool = True,
    flip: bool = False,
    axis: str = "z",
    axis_feature: str = "",
    geometry_pattern: bool = True,
) -> dict[str, Any]:
    """绕一根轴做圆周阵列。

    真机校准（SW2024 SP0.1，2026-11-20，用凸台面心的角度逐个验证过）：
      1. 接口 `IFeatureManager.FeatureCircularPattern5(Number, Spacing(弧度), FlipDirection,
         DName, GeometryPattern, EqualSpacing, VaryInstance, SyncSubAssemblies, BDir2,
         BSymmetric, Number2, Spacing2, DName2, EqualSpacing2)`——一共 14 个参数。
      2. **Spacing 是弧度**，语义随 EqualSpacing 变（这条最容易踩）：
         - `equal_spacing=False` → 每份之间的间隔角；
         - `equal_spacing=True`  → **总角度**：整圈 360° 时相邻角 = 360/N，
           其它角度时相邻角 = 总角/(N-1)（首末都算实例，所以是 N-1）。
         实测：360°/4 → 0/90/180/270；120°/3 → 0/120/240；90°/3 → 0/-45/-90。
      3. `FlipDirection=False` 是**顺时针**（角度递减），True 逆时针。实测 60°/3/翻转 → 0/60/120。
      4. 选择标记跟线性阵列一致：种子 `Select2(Append, 4)`，轴 `Select2(Append, 1)`，
         同时把轴特征名当 DName 传进去。
      5. 轴别指望 `body.GetEdges()`，用参考轴（见 reference_axis）。
    """
    if count < 2:
        raise C.SwError("BAD_REQUEST：count 至少 2")
    ang = float(360.0 if angle_deg is None else angle_deg)
    if ang <= 0:
        raise C.SwError("BAD_REQUEST：angle_deg 必须大于 0")
    for n in seeds:
        if C.find_feature(doc, n) is None:
            raise C.SwError(f"COM_NULL：特征树里没有 {n!r}")

    if axis_feature:
        if C.find_feature(doc, axis_feature) is None:
            raise C.SwError(f"COM_NULL：特征树里没有轴 {axis_feature!r}")
        ax_nm = axis_feature
        axis_planes = []
    else:
        ax = reference_axis(doc, axis)
        ax_nm = ax["name"]
        axis_planes = ax["planes"]

    C._clear_selection(doc)
    for i, name in enumerate(seeds):
        if not C.find_feature(doc, name).Select2(bool(i), 4):
            raise C.SwError(f"COM_ERROR：选不中要阵列的特征 {name!r}（mark 4）")
    if not C.find_feature(doc, ax_nm).Select2(True, 1):
        raise C.SwError(f"COM_ERROR：选不中阵列轴 {ax_nm!r}（mark 1）")

    fm = doc.FeatureManager
    try:
        feat = fm.FeatureCircularPattern5(
            int(count),
            math.radians(ang),
            bool(flip),
            ax_nm,
            bool(geometry_pattern),
            bool(equal_spacing),
            False, False, False, False,
            1, 0.0, "", True,
        )
    except Exception as exc:  # noqa: BLE001
        raise C.SwError(
            f"COM_ERROR：FeatureCircularPattern5 失败（{type(exc).__name__}: {exc}）"
        ) from exc
    if feat is None:
        raise C.SwError(
            "COM_NULL：圆周阵列没建出来。检查选择标记（种子 4 / 轴 1）和轴是否真的存在"
        )

    if equal_spacing:
        step = ang / count if abs(ang - 360.0) < 1e-9 else ang / (count - 1)
    else:
        step = ang
    return {
        "op": "circular_pattern",
        "name": str(C._value(feat, "Name")),
        "seeds": seeds,
        "count": int(count),
        "angle_deg": ang,
        "equal_spacing": bool(equal_spacing),
        "step_deg": round(step, 4),
        "clockwise": not bool(flip),
        "axis": ax_nm,
        "axis_planes": axis_planes,
        "volume_mm3": C.mass_properties(doc).get("volume_mm3"),
    }


# ---------------------------------------------------------------- 孔
def hole(
    doc,
    points: list[list[float]],
    diameter_mm: float,
    face: dict[str, Any] | None = None,
    end: str = "through_all",
    depth_mm: float | None = None,
    counterbore_diameter_mm: float | None = None,
    counterbore_depth_mm: float | None = None,
) -> dict[str, Any]:
    """在面上打孔：通孔 / 盲孔 / 柱形沉头孔（几何等价实现，不是异型孔特征）。

    和 `hole_wizard()` 的区别：这里建的是**草图圆 + 切除**，没有标准件数据，
    特征树上就是一个普通「切除-拉伸」；要对得上 GB / ISO 标准（沉头、螺纹孔、
    底部螺纹孔）请走 `hole_wizard()`，那边是真正的异型孔向导特征。
    这里保留的是「几何兜底」：不想依赖标准库、或只想切个洞时用它，体积能严格对账。

    - `points`：孔心，模型坐标 mm（必须落在 `face` 上或投影后落在上面）
    - `face`：{point:[x,y,z], normal:[nx,ny,nz]}，交给 find_face 定位；不传就用最后一个草图所在的面
    - `counterbore_*`：给了就再叠一层柱形沉头（先打穿再扩沉头，几何上等价）
    """
    if not points:
        raise C.SwError("BAD_REQUEST：points 不能为空")
    for i, p in enumerate(points):
        if not isinstance(p, (list, tuple)) or len(p) != 3:
            raise C.SwError(f"BAD_REQUEST：points[{i}] 要 3 个数，收到 {p!r}")
    if float(diameter_mm) <= 0:
        raise C.SwError("BAD_REQUEST：diameter 必须大于 0")
    end_key = (end or "through_all").strip().lower()
    if end_key == "blind" and not depth_mm:
        raise C.SwError("BAD_REQUEST：end=blind 时必须给 depthMm")

    face = face or {}
    face_pt = face.get("point")
    face_n = face.get("normal")

    def cut_circles(dia: float, depth, end_cond):
        entities = [{"type": "circle", "center": [float(v) for v in p], "diameter": dia}
                    for p in points]
        sk = sketch_on_face(doc, entities, point_mm=face_pt, normal=face_n)
        cut = C.extrude(doc, sk["name"], float(depth or 10.0), end=end_cond, cut=True)
        return sk, cut

    sk1, cut1 = cut_circles(float(diameter_mm), depth_mm, end_key)
    result: dict[str, Any] = {
        "op": "hole",
        "name": cut1.get("name", ""),
        "sketch": sk1["name"],
        "points": [[round(float(v), 4) for v in p] for p in points],
        "diameter_mm": float(diameter_mm),
        "end": end_key,
        "depth_mm": float(depth_mm) if depth_mm else None,
        "shape": "hole",
        "volume_mm3_before_counterbore": C.mass_properties(doc).get("volume_mm3"),
    }

    if counterbore_diameter_mm:
        cb_d = float(counterbore_diameter_mm)
        cb_h = float(counterbore_depth_mm or 0)
        if cb_d <= float(diameter_mm):
            raise C.SwError("BAD_REQUEST：沉头直径必须大于孔径")
        if cb_h <= 0:
            raise C.SwError("BAD_REQUEST：沉头深度必须大于 0")
        sk2, cut2 = cut_circles(cb_d, cb_h, "blind")
        result["shape"] = "counterbore"
        result["counterbore"] = {
            "feature": cut2.get("name", ""),
            "sketch": sk2["name"],
            "diameter_mm": cb_d,
            "depth_mm": cb_h,
        }
    result["volume_mm3"] = C.mass_properties(doc).get("volume_mm3")
    return result


# ---------------------------------------------------------------- 异型孔向导
# 2026-11-21 真机打通。之前「10 个组合全 None」的两个根因：
#   1) 选的是**草图特征**，而异型孔要的是**草图点对象**（ISketchPoint.Select2，
#      且不要给选择标记）——选错对象 SW 一声不响返回 None；
#   2) FastenerTypeIndex 传了 0，而它必须是该标准下具体的 fastener 序号
#      （ISO：内六角圆柱头 139 / 沉头螺钉 140 / 螺栓通孔 144 / 内螺纹 147）。
# 参数语义（官方 Remarks）：单位米、角度弧度、要 -1 的参数 SW 会忽略。
ISO_STANDARD = 8
SPACE_GENERIC = {"counterbore": 0, "countersink": 1, "clearance": 2, "tap": 4}
SPACE_FASTENER = {"counterbore": 139, "countersink": 140, "clearance": 144, "tap": 147}
# 粗牙螺纹的底孔直径（ISO 2306）
ISO_TAP_DRILL_MM = {
    "M3": 2.5, "M4": 3.3, "M5": 4.2, "M6": 5.0, "M8": 6.8, "M10": 8.5,
    "M12": 10.2, "M14": 12.0, "M16": 14.0, "M20": 17.5,
}
# 螺纹孔的 Value1..Value12 有固定顺序 —— 见本机 API 帮助
# D:\SW\SOLIDWORKS\api\sldworksapi.chm → IFeatureManager::HoleWizard5 →
# 「Pipe/Straight Tap Holes」：
#   V1 螺纹深度 / V2 近端沉孔直径 / V3 近端沉孔角度 / V4 远端沉孔直径 /
#   V5 远端沉孔角度 / V6 钻尖角度 / V7 装饰螺纹线类型 / V8 螺纹终止条件 /
#   V9 钢丝螺套类型 / V10 偏移 / V11 V12 保留
# 坑（2026-11-21 实测）：螺纹孔这 12 个位置如果全传 0，SW 会理解成
# 「沉孔直径 0、角度 0」，然后退回**英寸默认值**（Ø25.4 孔 / Ø50.8 孔 / 1.25" 螺纹深），
# 建出来的 HoleWzd 是个 Ø50.8 的大坑。传 -1 让 SW 忽略，几何才对：
#   Ø6.8×15 盲孔 + 118° 钻尖 + 装饰螺纹线，切掉 569.4833 mm³
# （544.7522 圆柱 + 24.7311 锥尖 = πr²h + πr³/tan59°）。
# 另一个坑：Diameter 必须给**底孔直径**（M8 → 0.0068 m）；给 0 同样退回 Ø25.4。
SPACE_TAP_VALUES = {"cosmetic_thread_with_callout": 1.0, "thread_end_blind": 0.0}
_SIZE_RE = re.compile(r"^M\s*(\d+(?:\.\d+)?)")


def _tap_values(thread_depth_mm: float, thread_end: float = 0.0,
                cosmetic: float = 1.0) -> tuple[float, ...]:
    """螺纹孔的 Value1..Value12（顺序见上面官方表）。-1 = 让 SW 忽略该位置。"""
    return (
        float(thread_depth_mm) * C.MM,  # V1 螺纹深度
        -1.0, -1.0,                     # V2 V3 近端沉孔直径 / 角度
        -1.0, -1.0,                     # V4 V5 远端沉孔直径 / 角度
        -1.0,                           # V6 钻尖角度（-1 = 默认 118°）
        float(cosmetic),                # V7 装饰螺纹线：0 无 / 1 带标注 / 2 不带标注
        float(thread_end),              # V8 螺纹终止：0 盲孔 / 1 通孔 / 2 到下一边
        -1.0,                           # V9 钢丝螺套类型（非钢丝螺套时忽略）
        -1.0,                           # V10 偏移（仅「从面偏移」终止条件用）
        -1.0, -1.0,                     # V11 V12 保留
    )


def _nominal_diameter_mm(size: str) -> float:
    """从尺寸串里取标称直径（mm）："M8" → 8，"M8X1.25" → 8，"M10" → 10。取不到返回 0。"""
    m = _SIZE_RE.match(str(size or "").strip().upper())
    return float(m.group(1)) if m else 0.0


def _select_sketch_points(doc, sketch_name: str) -> int:
    """选中草图里的**点对象**（异型孔要的就是这个）。返回选中个数。"""
    feat = C.find_feature(doc, sketch_name)
    if feat is None:
        raise C.SwError(f"COM_ERROR：找不到草图 {sketch_name}")
    sketch = C._value(feat, "GetSpecificFeature2")
    points = list(C._value(sketch, "GetSketchPoints2") or [])
    C._clear_selection(doc)
    picked = 0
    for i, pt in enumerate(points):
        try:
            if pt.Select2(i > 0, 0):
                picked += 1
        except Exception:  # noqa: BLE001
            continue
    return picked


def hole_wizard(
    doc,
    points: list[list[float]],
    kind: str = "counterbore",
    size: str = "M8",
    face: dict[str, Any] | None = None,
    end: str = "through_all",
    depth_mm: float | None = None,
    tap_drill_mm: float | None = None,
    name: str = "",
) -> dict[str, Any]:
    """异型孔向导：柱形沉头 / 锥形沉头 / 螺栓通孔 / 螺纹孔，尺寸取自 SW 标准库。

    - `kind`：counterbore（柱形沉头，内六角圆柱头 ISO 4762）/ countersink（锥形沉头，
      ISO 10642）/ clearance（螺栓通孔 ISO 273）/ tap（螺纹孔）
    - `size`：标称尺寸串，如 "M5" "M6" "M8" "M10" "M12"（必须存在于该 fastener 的表里）
    - `points`：孔心，模型坐标 mm；`face`：{point, normal} 定位放置面
    - `end`：through_all / blind（blind 要给 depth_mm）
    - `tap_drill_mm`：仅螺纹孔，覆盖底孔直径（默认查 ISO_TAP_DRILL_MM）
    返回特征名、回读的标准/螺钉/尺寸、以及体积（用于对账）。
    """
    key = (kind or "counterbore").strip().lower()
    if key not in SPACE_GENERIC:
        raise C.SwError(f"BAD_REQUEST：kind 只能是 {sorted(SPACE_GENERIC)}，收到 {kind!r}")
    if not points:
        raise C.SwError("BAD_REQUEST：points 不能为空")
    end_key = (end or "through_all").strip().lower()
    if end_key == "blind" and not depth_mm:
        raise C.SwError("BAD_REQUEST：end=blind 时必须给 depthMm")
    end_type = 0 if end_key == "blind" else 1
    face = face or {}
    face_pt, face_n = face.get("point"), face.get("normal")
    size = str(size).strip().upper()
    before = C.mass_properties(doc).get("volume_mm3")

    if key == "tap":
        drill = float(tap_drill_mm or ISO_TAP_DRILL_MM.get(size, 0) or 0)
        if drill <= 0:
            raise C.SwError(
                f"BAD_REQUEST：不知道 {size} 的底孔直径，请用 tapDrillMm 显式给（常见："
                f"{', '.join(f'{k}={v}' for k, v in ISO_TAP_DRILL_MM.items())}）")
        # 标称直径（M8 → 8），用来算界面上的默认螺纹深度 2*DIA
        major = _nominal_diameter_mm(size) or drill / 0.85

    entities = [{"type": "point", "point": [float(v) for v in p]} for p in points]
    sk = sketch_on_face(doc, entities, point_mm=face_pt, normal=face_n)
    picked = _select_sketch_points(doc, sk["name"])
    if picked != len(points):
        raise C.SwError(f"COM_ERROR：草图里 {len(points)} 个点只选中 {picked} 个")
    diameter = 0.0
    length = 0.0
    depth = float(depth_mm or 0) * C.MM
    if key == "tap":
        # 底孔直径必须显式给：给 0 时 SW 退回英寸默认值 Ø25.4，孔就被打成 1 英寸
        diameter = drill * C.MM
        length = 2.0 * major * C.MM           # 界面上的「螺纹线 给定深度 (2 * DIA)」
        values = _tap_values(float(depth_mm or 0))
    else:
        values = (0.0,) * 12
    # 非螺纹孔 Value1..12 传 0：实测传 -1 时 clearance（swWzdHole）会静默返回 None，
    # 传 0 三类孔都能建（0 = 交给标准表取默认值）。螺纹孔见 _tap_values 的说明。
    try:
        feat = doc.FeatureManager.HoleWizard5(
            SPACE_GENERIC[key], ISO_STANDARD, SPACE_FASTENER[key], size, end_type,
            diameter, depth, length,
            *values,
            "", False, False, False, False, False, False)
    except Exception as exc:  # noqa: BLE001
        raise C.SwError(
            f"COM_ERROR：HoleWizard5 调用失败（{type(exc).__name__}: {exc}）") from exc
    if feat is None:
        raise C.SwError(
            f"COM_NULL：异型孔没建出来。检查 size={size!r} 是否在 ISO/{key} 的尺寸表里"
            "（可用 M5/M6/M8/M10/M12），以及草图点是落在实体的面上")
    C._clear_selection(doc)
    after = C.mass_properties(doc).get("volume_mm3")
    data = None
    try:
        data = C._value(feat, "GetDefinition")
    except Exception:  # noqa: BLE001
        data = None
    out: dict[str, Any] = {
        "op": "hole_wizard",
        "kind": key,
        "name": str(C._value(feat, "Name")),
        "sketch": sk["name"],
        "points": [[round(float(v), 4) for v in p] for p in points],
        "size": size,
        "end": end_key,
        "depth_mm": float(depth_mm) if depth_mm else None,
        "volume_mm3": after,
        "removed_mm3": (before - after) if (before is not None and after is not None) else None,
    }
    if key == "tap":
        out["tap_drill_mm"] = drill
        out["thread_depth_mm"] = float(depth_mm or 0)
        out["cosmetic_thread"] = "with_callout"
    for prop, field in (("Standard", "standard"), ("FastenerType", "fastener"),
                        ("FastenerSize", "size_readback"), ("EndCondition", "end_condition"),
                        ("ThruHoleDiameter", "hole_diameter_mm"),
                        ("CounterBoreDiameter", "counterbore_diameter_mm"),
                        ("CounterBoreDepth", "counterbore_depth_mm"),
                        ("CounterSinkDiameter", "countersink_diameter_mm"),
                        ("CounterSinkAngle", "countersink_angle_rad"),
                        ("ThreadDiameter", "thread_diameter_mm"),
                        ("ThreadDepth", "thread_depth_readback_mm"),
                        ("TapDrillDiameter", "tap_drill_readback_mm"),
                        ("TapDrillDepth", "tap_drill_depth_mm")):
        try:
            v = C._value(data, prop) if data is not None else None
        except Exception:  # noqa: BLE001
            v = None
        if v is None:
            continue
        if field.endswith("_mm") and isinstance(v, float):
            v = round(v / C.MM, 4)
        elif field.endswith("_rad") and isinstance(v, float):
            v = round(math.degrees(v), 4)
        out[field] = v
    if name:
        out["renamed"] = rename_feature(doc, out["name"], name)
        out["name"] = name
    return out


def _hole_cylinder_face(doc, center: list[float], diameter_mm: float):
    """找孔的那个圆柱面（装饰螺纹线要选它）。按包络盒中心 + 直径匹配。"""
    need = [float(center[0]), float(center[1])]
    for f in solid_faces(doc):
        box = f.get("box_mm") or []
        if len(box) < 6:
            continue
        cx, cy = (box[0] + box[3]) / 2.0, (box[1] + box[4]) / 2.0
        dx, dy = box[3] - box[0], box[4] - box[1]
        if (abs(cx - need[0]) < 0.02 and abs(cy - need[1]) < 0.02
                and abs(dx - diameter_mm) < 0.05 and abs(dy - diameter_mm) < 0.05):
            return f
    return None


def _add_cosmetic_thread(doc, points, drill_mm, size, depth_mm, end_kind="through_all"):
    """给底孔加装饰螺纹线（IFeatureManager.InsertCosmeticThread3）。失败返回 None。

    坑：`SelectionManager.CreateSelectData()` 在晚期绑定下报 DISP_E_MEMBERNOTFOUND，
    `SelectByID2` 的 Callout 参数传 None 报 DISP_E_TYPEMISMATCH —— 选面统一走
    `_select_face()`（`Select4(False, VARIANT(VT_DISPATCH, None))`，退回 `Select2`）。
    """
    added = []
    fm = doc.FeatureManager
    end_type = 0 if end_kind == "blind" else 1
    for p in points:
        info = _hole_cylinder_face(doc, p, drill_mm)
        if info is None or not info.get("_face"):
            continue
        C._clear_selection(doc)
        if not _select_face(info["_face"]):
            continue
        for std_type in ("螺纹孔", "ISO", ""):
            try:
                feat = fm.InsertCosmeticThread3(
                    ISO_STANDARD, std_type, size, drill_mm * C.MM, end_type,
                    float(depth_mm or 0) * C.MM, "")
            except Exception:  # noqa: BLE001
                continue
            if feat is not None:
                added.append(str(C._value(feat, "Name")))
                break
    C._clear_selection(doc)
    return added or None
