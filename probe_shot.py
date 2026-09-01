r"""
화면 캡처 + 라벨 덤프 — 컨트롤이 화면에서 실제로 무슨 항목인지 눈으로 확인한다.
⚠️ Windows 전용.   py probe_shot.py [출력.bmp]

Pillow 없이 pywin32(PrintWindow)만으로 BMP 를 저장한다.
"""
from __future__ import annotations

import ctypes
import sys
import time

import ierp_progress as ip


def wm_gettext(h) -> str:
    try:
        buf = ctypes.create_unicode_buffer(512)
        ctypes.windll.user32.SendMessageW(h, 0x000D, 512, buf)   # WM_GETTEXT
        return buf.value.strip()
    except Exception:
        return ""


def capture(hwnd, path):
    import win32gui
    import win32ui
    import win32con
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r - l, b - t
    hdc = win32gui.GetWindowDC(hwnd)
    src = win32ui.CreateDCFromHandle(hdc)
    mem = src.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(src, w, h)
    mem.SelectObject(bmp)
    # PrintWindow(2 = PW_RENDERFULLCONTENT) — 가려져 있어도 그려진다
    ok = ctypes.windll.user32.PrintWindow(hwnd, mem.GetSafeHdc(), 2)
    if not ok:
        mem.BitBlt((0, 0), (w, h), src, (0, 0), win32con.SRCCOPY)
    bmp.SaveBitmapFile(mem, path)
    mem.DeleteDC()
    src.DeleteDC()
    win32gui.ReleaseDC(hwnd, hdc)
    return w, h


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else r"C:\ierp_day_prd\ierp_inspect\screen.bmp"
    import os
    os.makedirs(os.path.dirname(out), exist_ok=True)
    scfg = ip.load_screen()
    top = ip.ensure_screen(scfg, log=print)
    ip.focus_window(top)
    time.sleep(1.0)

    import win32gui
    print(f"■ 화면 hwnd={top} {win32gui.GetWindowText(top)!r}")

    # ── 조회조건 영역(상단)의 모든 자식: 클래스·텍스트·좌표 ──
    print("\n── 상단 조회조건 영역의 컨트롤 (y < 200) ─────────────")
    rows = []

    def _cb(h, _):
        try:
            l, t, r, b = win32gui.GetWindowRect(h)
            cls = win32gui.GetClassName(h)
            if t - win32gui.GetWindowRect(top)[1] < 260:
                aid = ""
                try:
                    aid = ip.elem_of(h).automation_id or ""
                except Exception:
                    pass
                rows.append((t, l, cls, aid, wm_gettext(h)))
        except Exception:
            pass
        return True
    win32gui.EnumChildWindows(top, _cb, None)
    rows.sort(key=lambda x: (x[0] // 10, x[1]))
    for t, l, cls, aid, txt in rows:
        short = cls.replace("WindowsForms10.", "").split(".app")[0]
        print(f"  y={t:<5} x={l:<5} {short:<14} {aid:<14} {txt!r}")

    w, h = capture(top, out)
    print(f"\n■ 캡처 저장: {out}  ({w}x{h})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
