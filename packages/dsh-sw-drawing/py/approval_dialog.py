"""写操作确认框。

两种用法：
  1) 被 mcp_sw_server 以**子进程**方式拉起（推荐）：JSON 从 stdin 进来，退出码表示结果。
     这样 Tk 不在 MCP 服务进程里，服务端主循环能一直响应 ping / tools/list。
  2) 直接 import 调 `ask(tool, args)`（旧浮窗版还在用）。

退出码：0 = 允许，1 = 拒绝/超时，2 = 允许且本次启动后面全部放行。

两条硬约束，都是真机踩出来的：

* **超时必须短于 DSH 的 60 s**。DSH 的 `toolCallTimeoutMs` 默认 60 000 ms，
  弹窗要等 120 s 的话，DSH 先超时报失败、你后点「允许」→ SW 真改了、记录却是失败。
  所以默认 45 s，硬上限 55 s。

* **不用 ttk.Button 的默认点击语义**。2026-11-21 实测：弹窗总是出现在鼠标指针正下方，
  而 Tk 的按钮只要收到 `<ButtonRelease-1>` 就执行 command —— 指针底下悬着的
  指针底下还没抬起来的那一下（或窗口弹出瞬间的余波）就能把一次写操作「点同意」。
  所以这里用 Label 自绘按钮：窗口显示满 0.35 s 冷却期（吞掉双击余波）、
  按下和抬起都要落在按钮内、不绑任何键盘事件。
"""

from __future__ import annotations

import json
import os
import sys
import time

DEFAULT_TIMEOUT = 45
MAX_TIMEOUT = 55   # 必须小于 DSH 的 60 s 单次调用超时
CLICK_COOLDOWN = 0.35


def resolve_timeout(timeout: int | None = None) -> int:
    """超时优先级：入参 > 环境变量 > 45 s；一律夹到 [5, 55]。"""
    if not timeout:
        try:
            timeout = int(os.environ.get("SW_BRIDGE_APPROVAL_TIMEOUT", "") or DEFAULT_TIMEOUT)
        except ValueError:
            timeout = DEFAULT_TIMEOUT
    return max(5, min(int(timeout), MAX_TIMEOUT))


def click_decision(*, now: float, shown_at: float, press_inside: bool,
                   release_inside: bool, cooldown: float = CLICK_COOLDOWN) -> bool:
    """纯判定函数（便于单测）：这次点击算不算「用户真的点了允许」。"""
    return (now - shown_at) >= cooldown and press_inside and release_inside


def ask(tool: str, arguments: dict, timeout: int = 0) -> bool:
    """返回是否允许；弹窗失败/超时一律按拒绝处理（fail closed）。"""
    return ask_ex(tool, arguments, timeout)[0]


def ask_ex(tool: str, arguments: dict, timeout: int = 0) -> tuple[bool, bool]:
    """返回 (是否允许, 是否本次启动全部放行)。"""
    try:
        return _ask_tk(tool, arguments, resolve_timeout(timeout or None))
    except Exception:  # noqa: BLE001
        return (False, False)


def _place_away_from_pointer(root, winfo: dict) -> None:
    """把窗口放到不压指针的地方：指针在候选位置上就挪到左上角。"""
    try:
        px, py = root.winfo_pointerxy()
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w = max(winfo.get("w") or 1, 1)
        h = max(winfo.get("h") or 1, 1)
        x, y = (sw - w) // 2, (sh - h) // 3
        if x - 24 <= px <= x + w + 24 and y - 24 <= py <= y + h + 24:
            x, y = 40, 80
        root.geometry(f"+{x}+{y}")
    except Exception:  # noqa: BLE001
        pass


def _make_button(parent, text: str, on_click, *, bg: str, fg: str, px: int = 14) -> None:
    """自绘按钮：按下和抬起都要在同一个 Label 里才算一次点击。"""
    import tkinter as tk

    lbl = tk.Label(parent, text=text, bg=bg, fg=fg, font=("Segoe UI", 10))
    state = {"down": False}

    def on_press(_event) -> None:
        state["down"] = True
        lbl.configure(bg="#4a4a58")

    def on_release(event) -> None:
        was_down = state["down"]
        state["down"] = False
        lbl.configure(bg=bg)
        if not was_down:
            return
        inside = 0 <= event.x <= lbl.winfo_width() and 0 <= event.y <= lbl.winfo_height()
        on_click(inside)

    lbl.bind("<ButtonPress-1>", on_press)
    lbl.bind("<ButtonRelease-1>", on_release)
    lbl.pack(side="left", padx=(0, px), ipadx=10, ipady=4)



def _ask_tk(tool: str, arguments: dict, timeout: int) -> tuple[bool, bool]:
    import tkinter as tk

    shown_at = time.monotonic()
    result = {"allow": False, "remember": False}

    root = tk.Tk()
    root.title("SW 写操作确认")
    root.configure(bg="#27272f")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    args_text = json.dumps(arguments, ensure_ascii=False, indent=2)
    tk.Label(root, text=f"DSH 请求执行写操作：{tool}", bg="#27272f", fg="#f0b04c",
             font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=(12, 4))
    tk.Label(root, text=args_text, bg="#27272f", fg="#e6e6e6", justify="left",
             font=("Consolas", 9)).pack(anchor="w", padx=14)

    hint = tk.Label(root, text="", bg="#27272f", fg="#9a9aa5", font=("Segoe UI", 9))
    hint.pack(anchor="w", padx=14, pady=(8, 0))

    remember_var = tk.BooleanVar(value=False)
    tk.Checkbutton(root, text="本次启动内全部允许（直到 DSH 重启）", variable=remember_var,
                   bg="#27272f", fg="#e6e6e6", selectcolor="#3a3a45",
                   activebackground="#27272f", activeforeground="#e6e6e6",
                   font=("Segoe UI", 9)).pack(anchor="w", padx=14)

    row = tk.Frame(root, bg="#27272f")
    row.pack(fill="x", padx=14, pady=10)

    def finish(allow: bool) -> None:
        result["allow"] = allow
        result["remember"] = bool(allow and remember_var.get())
        try:
            root.destroy()
        except Exception:  # noqa: BLE001
            pass

    def allow_clicked(inside: bool) -> None:
        if not inside:
            return
        if not click_decision(now=time.monotonic(), shown_at=shown_at,
                              press_inside=True, release_inside=True):
            hint.configure(text="刚才那下像是弹窗时点在原处的余波，请再点一次「允许」", fg="#f0b04c")
            return
        finish(True)

    _make_button(row, "允许", allow_clicked, bg="#3a3a45", fg="#e6e6e6")
    _make_button(row, "拒绝", lambda inside: finish(False) if inside else None,
                 bg="#3a3a45", fg="#e6e6e6")
    tk.Label(row, text="只改内存，不保存文件", bg="#27272f", fg="#9a9aa5",
             font=("Segoe UI", 9)).pack(side="right")

    root.protocol("WM_DELETE_WINDOW", lambda: finish(False))
    root.update_idletasks()
    _place_away_from_pointer(root, {"w": root.winfo_reqwidth(), "h": root.winfo_reqheight()})

    remaining = {"left": timeout}

    def tick() -> None:
        remaining["left"] -= 1
        if remaining["left"] <= 0:
            finish(False)
            return
        hint.configure(text=f"{remaining['left']} 秒后自动拒绝", fg="#9a9aa5")
        root.after(1000, tick)

    hint.configure(text=f"{remaining['left']} 秒后自动拒绝", fg="#9a9aa5")
    root.after(1000, tick)
    root.mainloop()
    return (result["allow"], result["remember"])


# ---------------------------------------------------------------- 子进程入口
def main() -> int:
    """stdin 收一个 JSON：{"tool":..., "arguments":{...}, "timeout":45}。"""
    raw = ""
    try:
        raw = sys.stdin.read()
        payload = json.loads(raw or "{}")
    except Exception:  # noqa: BLE001
        print(f"[approval] 参数解析失败：{raw[:200]!r}", file=sys.stderr)
        return 1
    allow, remember = ask_ex(
        str(payload.get("tool") or "?"),
        payload.get("arguments") or {},
        int(payload.get("timeout") or 0),
    )
    if not allow:
        return 1
    return 2 if remember else 0


if __name__ == "__main__":
    sys.exit(main())
