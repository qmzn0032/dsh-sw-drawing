# -*- coding: utf-8 -*-
"""一键把工程图里的字体换掉（走「文件属性」，不逐条改文字）。

为什么是文件属性：
    标题栏/图框里那几十上百个文字对象的 `GetUseDocTextFormat()` 是 True，
    也就是跟着「文件属性 > 注解 > 字体」走。改这三项，它们全跟着变；
    真正带本地格式的只有视图里的尺寸和孔标注，再单独扫一遍。

用法（三种都行）：
    1) 把 .SLDDRW 拖到同目录的「一键换字体.cmd」上
    2) 命令行：python font_switch.py 图纸1.SLDDRW 图纸2.SLDDRW --font 宋体
    3) 不带参数：处理 SolidWorks 里当前活动的那张工程图

常用开关：
    --check      只看现在是什么字体，不改
    --dry-run    走完整流程但不存盘
    --no-save    改了不存（= dry-run 的变体，留在 SW 里给你看）
    --no-backup  存盘前不备份
    --font 宋体  换别的字体（默认 汉仪长仿宋体）

真机校准（2026-11-21，SW2024 SP0.1）：
  * 文档属性字体：`IModelDocExtension.Get/SetUserPreferenceTextFormat(pref, 0, tf)`
    pref = 0 注释 / 1 尺寸 / 19 注解；只改 `TypeFaceName`，字高保持原样。
  * 带本地格式的尺寸：`IAnnotation.GetTextFormat(i)` -> 改 `TypeFaceName`
    -> `SetTextFormat(i, False, tf)`；第二参必须 False，传 True 会丢字高。
  * 存盘用 `Save3(1)`（swSaveAsOptions_Silent）；存完回读 mtime 和 GetSaveFlag 判定。
  * 图纸格式里的文字只读得到、改不到（pywin32 上 GetAnnotation 对它们报
    「找不到成员」），但它们本来就 UseDoc=True，改文档属性就够了。
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
import time
import winreg

HERE = os.path.dirname(os.path.abspath(__file__))

# 两种目录布局都要认：仓库里是..\_plugins\...，发布包里是..\packages\...
_CANDIDATES = (
    os.path.join(HERE, os.pardir, "packages", "dsh-sw-drawing", "py"),
    os.path.join(HERE, os.pardir, "_plugins", "dsh-sw-drawing", "py"),
)
PLUGIN_PY = next(
    (os.path.abspath(c) for c in _CANDIDATES
     if os.path.isfile(os.path.join(c, "sw_drawing.py"))),
    os.path.abspath(_CANDIDATES[0]),
)
sys.path.insert(0, PLUGIN_PY)

try:
    import pythoncom  # noqa: F401  仅确保 pywin32 可用
    import sw_com as C
    import sw_drawing as D
    from win32com.client import VARIANT
except Exception as exc:  # noqa: BLE001
    print(f"[×] 载入 SolidWorks 组件失败：{exc}")
    print("    需要 pywin32，且 _plugins\\dsh-sw-drawing\\py 要完整。")
    sys.exit(2)

SW_DOC_DRAWING = 3
SW_OPEN_SILENT = 1
SW_SAVE_SILENT = 1
DEFAULT_FONT = "汉仪长仿宋体"

# 汉仪长仿宋体的两份安装（族名相同，SW 只认族名，用哪份由 Windows 决定）
FONT_FILES = (
    ("2023 版（系统级）", r"C:\Windows\Fonts\hyswlongfangsong.ttf"),
    ("2017 版（仅当前用户）", os.path.join(
        os.environ.get("LOCALAPPDATA", ""), r"Microsoft\Windows\Fonts\汉仪长仿宋体.ttf")),
)


# ---------------------------------------------------------------- 控制台
def setup_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass


def hr(title: str = "", ch: str = "=") -> None:
    print("" if title else ch * 64)
    if title:
        print(f"\n{title}")


# ---------------------------------------------------------------- 字体盘点
def registered_fonts() -> list[tuple[str, str, str]]:
    """读注册表里的字体登记：[(作用域, 字体名, 文件)]。"""
    keys = (
        ("HKLM", winreg.HKEY_LOCAL_MACHINE),
        ("HKCU", winreg.HKEY_CURRENT_USER),
    )
    path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    rows: list[tuple[str, str, str]] = []
    for scope, hive in keys:
        try:
            with winreg.OpenKey(hive, path) as k:
                for i in range(winreg.QueryInfoKey(k)[1]):
                    name, val, _ = winreg.EnumValue(k, i)
                    rows.append((scope, name, str(val)))
        except OSError:
            continue
    return rows


def font_inventory(font: str) -> list[tuple[str, str, int | None, str | None]]:
    """这个字体在系统里到底有几份文件、多新。"""
    out = []
    for label, path in FONT_FILES:
        if font != DEFAULT_FONT:
            return out
        exp = os.path.expandvars(path)
        if os.path.isfile(exp):
            st = os.stat(exp)
            out.append((label, exp, st.st_size, time.strftime("%Y-%m-%d", time.localtime(st.st_mtime))))
        else:
            out.append((label, exp, None, None))
    return out


def print_font_env(font: str) -> None:
    hr("字体环境")
    hits = [r for r in registered_fonts()
            if ("汉仪" in r[1] or "HYSWLongFangSong" in r[1] or "长仿宋" in r[1])]
    if not hits:
        print(f"    注册表里没找到「汉仪长仿宋体」，请确认字体已安装。")
    for scope, name, val in hits:
        print(f"    {scope}  {name}  ->  {val}")
    files = font_inventory(font)
    if files:
        print()
        for label, path, size, mtime in files:
            if size:
                print(f"    {label:<22} {size:>9,} 字节  {mtime}  {path}")
            else:
                print(f"    {label:<22} 未安装")
        if all(f[2] for f in files):
            print()
            print("    [i] 两份同族字体同时存在。本机实测（GDI+ 逐像素比对）：Windows 解析到的是")
            print("        2023 版（系统级那份）。要永久锁定 2023 版，可注销 2017 版（「仅为我安装」）。")


# ---------------------------------------------------------------- 图纸读写
def open_path(sw, path: str):
    errs = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warns = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    return sw.OpenDoc6(os.path.abspath(path), SW_DOC_DRAWING, SW_OPEN_SILENT, "", errs, warns)


def read_state(drw) -> dict:
    ext = D._val(drw, "Extension")
    prefs = {}
    for pref, name in ((0, "注释"), (1, "尺寸"), (19, "注解")):
        tf = D._val(ext, "GetUserPreferenceTextFormat", pref, 0)
        prefs[name] = {
            "font": D._val(tf, "TypeFaceName") if tf else None,
            "height_mm": round(float(D._val(tf, "CharHeight", default=0) or 0) * 1000, 2) if tf else None,
        }
    dims = {}
    for item in D.display_dimensions(drw):
        ann = item["_ann"]
        tf = D._val(ann, "GetTextFormat", 0)
        dims[item["name"]] = D._val(tf, "TypeFaceName") if tf else None
    return {"document_properties": prefs, "dimensions": dims,
            "dirty": D._val(drw, "GetSaveFlag")}


def state_fonts(state: dict) -> set:
    s = {v["font"] for v in state["document_properties"].values() if v["font"]}
    s |= {v for v in state["dimensions"].values() if v}
    return s


def backup_file(path: str) -> str | None:
    folder = os.path.join(os.path.dirname(os.path.abspath(path)), "换字体备份")
    try:
        os.makedirs(folder, exist_ok=True)
        dst = os.path.join(folder, f"{os.path.splitext(os.path.basename(path))[0]}"
                                    f"_{time.strftime('%Y%m%d_%H%M%S')}.SLDDRW")
        shutil.copy2(path, dst)
        return dst
    except Exception as exc:  # noqa: BLE001
        print(f"    [!] 备份失败（继续）：{exc}")
        return None


def save_doc(drw, path: str) -> dict:
    """存盘并**回读**判定：mtime 变了、GetSaveFlag 落了，才算存上。

    校准（2026-11-21）：本机 `Save3(1)` 会抛「非选择性的参数」，
    `Save2(1)` 反而正常（返回值是 0，别拿它当成功标志）；所以按 Save2 → Save 顺序试。
    """
    t0 = os.path.getmtime(path) if os.path.exists(path) else 0
    size0 = os.path.getsize(path) if os.path.exists(path) else 0
    tried: list[str] = []
    for label, fn in (("Save2(1)", lambda: drw.Save2(SW_SAVE_SILENT)),
                      ("Save()", lambda: drw.Save())):
        try:
            fn()
        except Exception as exc:  # noqa: BLE001
            tried.append(f"{label} 异常: {exc}")
            continue
        time.sleep(0.5)
        mt = os.path.getmtime(path) if os.path.exists(path) else 0
        size = os.path.getsize(path) if os.path.exists(path) else 0
        if mt > t0 + 0.4 and size > 1024 and not D._val(drw, "GetSaveFlag"):
            return {"ok": True, "bytes": size, "how": label, "changed": size != size0}
        tried.append(f"{label} 没生效（mtime 未变 / 仍是 dirty）")
    return {"ok": False, "bytes": os.path.getsize(path) if os.path.exists(path) else 0,
            "how": "", "tried": tried}


def process(sw, path: str, font: str, *, dry_run: bool, do_save: bool,
            do_backup: bool) -> dict:
    path = os.path.abspath(path)
    if not path.lower().endswith(".slddrw"):
        return {"file": path, "ok": False, "why": "不是 .SLDDRW"}
    if not os.path.isfile(path):
        return {"file": path, "ok": False, "why": "文件不存在"}

    opened_here = False
    drw = D._find_open(sw, path)
    if drw is None:
        drw = open_path(sw, path)
        opened_here = drw is not None
    if drw is None:
        return {"file": path, "ok": False, "why": "SW 打不开（被占用/文件坏了？）"}

    time.sleep(0.6)
    if D._val(drw, "GetType") != SW_DOC_DRAWING:
        if opened_here:
            D._close_quietly(sw, drw)
        return {"file": path, "ok": False, "why": "不是工程图"}

    before = read_state(drw)
    dirty_before = bool(before["dirty"])
    stale = sorted(state_fonts(before) - {font})

    print(f"\n  文件: {path}")
    print(f"  当前字体: {', '.join(sorted(state_fonts(before))) or '（读不到）'}")
    if not stale:
        print(f"  -> 已经是「{font}」，不用动")
        result = {"file": path, "ok": True, "changed": False, "saved": False,
                  "opened_here": opened_here, "dirty_before": dirty_before}
        if opened_here:
            D._close_quietly(sw, drw)
        return result

    if dry_run:
        print(f"  -> 需要换成「{font}」：{', '.join(stale)}（dry-run，什么都没改）")
        return {"file": path, "ok": True, "changed": False, "dry": True, "would_replace": stale}

    print(f"  -> 换成「{font}」，原字体: {', '.join(stale)}")
    report = D.apply_font(drw, font)
    try:
        drw.EditRebuild3()
    except Exception:  # noqa: BLE001
        pass
    D._val(drw, "GraphicsRedraw2")
    time.sleep(0.5)

    after = read_state(drw)
    doc_ok = all(v["font"] == font for v in after["document_properties"].values())
    dims_ok = all(v == font for v in after["dimensions"].values()) if after["dimensions"] else True
    print(f"  文档属性: {doc_ok} | 尺寸/孔标注({len(after['dimensions'])} 个): {dims_ok}"
          f" | 这次改了 {len(report['dimensions_changed'])} 处本地格式")

    saved = {"ok": False}
    backup = None
    if do_save:
        if do_backup:
            backup = backup_file(path)
            if backup:
                print(f"  备份: {backup}")
        saved = save_doc(drw, path)
        print(f"  存盘: {'OK' if saved['ok'] else '失败'}  {saved.get('bytes', 0):,} 字节"
              f"  {saved.get('how') or ''}  dirty={D._val(drw, 'GetSaveFlag')}")
        for line in saved.get("tried") or []:
            print(f"    ! {line}")
    else:
        print("  没存盘（--no-save）：改动只在内存里，SW 里看一眼，满意了自己保存")

    if opened_here:
        if do_save:
            D._close_quietly(sw, drw)
        else:
            print("  [!] 这张图是我打开的，先不关：改动没存，关掉就白改了")

    return {"file": path, "ok": bool(doc_ok and dims_ok and (saved["ok"] or not do_save)),
            "changed": True, "saved": saved["ok"], "backup": backup,
            "doc_props_ok": doc_ok, "dims_ok": dims_ok,
            "dims_rewritten": len(report["dimensions_changed"]),
            "dirty_before": dirty_before, "opened_here": opened_here,
            "fonts_before": sorted(state_fonts(before))}


# ---------------------------------------------------------------- 入口
def pick_targets(sw, args) -> list[str]:
    if args.files:
        return list(args.files)
    doc = C.active_doc(sw)
    if doc is None:
        print("[×] SolidWorks 里没有打开的文档。请把 .SLDDRW 拖到「一键换字体.cmd」上。")
        return []
    kind = D._val(doc, "GetType")
    if kind != SW_DOC_DRAWING:
        print(f"[×] 当前活动文档是「{C.doc_title(doc)}」，不是工程图。"
              f"请先打开要换字体的 .SLDDRW，或把它拖到 .cmd 上。")
        return []
    print(f"目标（SW 里当前活动的那张图）: {C.doc_title(doc)}")
    return [C.doc_path(doc)] if C.doc_path(doc) else []


def main() -> int:
    setup_console()
    ap = argparse.ArgumentParser(
        description="把工程图里的字体换成指定字体（默认 汉仪长仿宋体）",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*", help="一个或多个 .SLDDRW；留空则用 SW 里当前活动的图")
    ap.add_argument("--font", default=DEFAULT_FONT, help=f"目标字体（默认 {DEFAULT_FONT}）")
    ap.add_argument("--check", action="store_true", help="只看现在是什么字体，不改")
    ap.add_argument("--dry-run", action="store_true", help="走流程但不存盘")
    ap.add_argument("--no-save", dest="no_save", action="store_true", help="改了不存（留在 SW 里给你看）")
    ap.add_argument("--no-backup", dest="no_backup", action="store_true", help="存盘前不备份")
    args = ap.parse_args()

    print("一键换字体  ·  目标字体:", args.font)
    setup_font = args.font

    try:
        sw = D.attach_sw()
    except Exception as exc:  # noqa: BLE001
        print(f"[×] 连不上 SolidWorks：{exc}\n    先启动 SW 再运行本工具。")
        return 3

    print_font_env(setup_font)

    if args.check:
        results = []
        for path in pick_targets(sw, args):
            drw = D._find_open(sw, path) or open_path(sw, path)
            if drw is None:
                print(f"[×] 打不开 {path}")
                continue
            before = read_state(drw)
            results.append({
                "file": path,
                "document_properties": before["document_properties"],
                "dimensions": before["dimensions"],
                "needs_change": sorted(state_fonts(before) - {setup_font}),
            })
            print(f"\n  {os.path.basename(path)}")
            for k, v in before["document_properties"].items():
                print(f"     {k}: {v['font']}  {v['height_mm']} mm")
            print(f"     尺寸/孔标注 {len(before['dimensions'])} 个: "
                  f"{', '.join(sorted(set(v for v in before['dimensions'].values() if v))) or '-'}")
            if results[-1]["needs_change"]:
                print(f"     需替换: {', '.join(results[-1]['needs_change'])}")
        print("\n" + ("=" * 64))
        return 0 if results else 1

    targets = pick_targets(sw, args)
    if not targets:
        return 1

    results = []
    for path in targets:
        try:
            results.append(process(sw, path, setup_font,
                                   dry_run=args.dry_run,
                                   do_save=not (args.dry_run or args.no_save),
                                   do_backup=not args.no_backup))
        except Exception:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            results.append({"file": path, "ok": False, "why": "异常，见上面的堆栈"})

    hr("汇总")
    ok_n = 0
    for r in results:
        mark = "OK " if r.get("ok") else "×  "
        if r.get("ok"):
            ok_n += 1
        detail = ("已是目标字体" if r.get("changed") is False and not r.get("dry")
                  else "dry-run，未改" if r.get("dry")
                  else f"已替换并保存 {r.get('dims_rewritten', 0)} 处尺寸格式" if r.get("saved")
                  else "已替换（未存盘）" if r.get("changed")
                  else f"失败：{r.get('why', '')}")
        print(f"  {mark}{os.path.basename(r['file'])}: {detail}")
    print(f"\n  成功 {ok_n}/{len(results)}")
    return 0 if ok_n == len(results) else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n中断")
        sys.exit(130)
