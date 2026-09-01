r"""
조회일자(dtpFRDATE/dtpTODATE) 진단 — 왜 시작일이 비어 있는지 방법별로 확인한다.
⚠️ Windows 전용.   py probe_date.py
"""
from __future__ import annotations

import ctypes
import time

import ierp_progress as ip
from pywinauto.controls.uiawrapper import UIAWrapper

GWL_STYLE = -16
DTS_SHOWNONE = 0x0002           # 체크박스가 달린 DTP (값을 '없음' 으로 둘 수 있다)
DTM_GETSYSTEMTIME = 0x1001
GDT_VALID, GDT_NONE = 0, 1


class SYSTEMTIME(ctypes.Structure):
    _fields_ = [("wYear", ctypes.c_ushort), ("wMonth", ctypes.c_ushort),
                ("wDayOfWeek", ctypes.c_ushort), ("wDay", ctypes.c_ushort),
                ("wHour", ctypes.c_ushort), ("wMinute", ctypes.c_ushort),
                ("wSecond", ctypes.c_ushort), ("wMilliseconds", ctypes.c_ushort)]


def read_all(h, tag):
    print(f"  [{tag}] hwnd={h}")
    try:
        print(f"      UIA.name          = {ip.elem_of(h).name!r}")
    except Exception as e:
        print(f"      UIA.name          = !{e}")
    try:
        print(f"      UIA.ValuePattern  = {UIAWrapper(ip.elem_of(h)).iface_value.CurrentValue!r}")
    except Exception as e:
        print(f"      UIA.ValuePattern  = !{type(e).__name__}")
    try:
        buf = ctypes.create_unicode_buffer(256)
        ctypes.windll.user32.SendMessageW(h, 0x000D, 256, buf)
        print(f"      win32.WM_GETTEXT  = {buf.value!r}")
    except Exception as e:
        print(f"      win32.WM_GETTEXT  = !{e}")
    st = SYSTEMTIME()
    r = ctypes.windll.user32.SendMessageW(h, DTM_GETSYSTEMTIME, 0, ctypes.byref(st))
    print(f"      DTM_GETSYSTEMTIME = {'GDT_VALID' if r == GDT_VALID else 'GDT_NONE(비어있음)'}"
          f"  {st.wYear}-{st.wMonth:02d}-{st.wDay:02d}")
    import win32gui
    style = win32gui.GetWindowLong(h, GWL_STYLE)
    print(f"      style             = 0x{style & 0xFFFFFFFF:08X}  "
          f"DTS_SHOWNONE(체크박스)={'있음 ★' if style & DTS_SHOWNONE else '없음'}")
    print(f"      rect              = {win32gui.GetWindowRect(h)}")


def main():
    scfg = ip.load_screen()
    top = ip.ensure_screen(scfg, log=print)
    ip.focus_window(top)
    time.sleep(0.4)
    dtps = ip.find_children_by_class(top, scfg["dtp_class"])
    print(f"■ DTP {len(dtps)}개\n")
    for i, h in enumerate(dtps):
        aid = ""
        try:
            aid = ip.elem_of(h).automation_id
        except Exception:
            pass
        read_all(h, f"[{i}] {aid}")
        print()

    if not dtps:
        return 1
    h = dtps[0]
    print("── 시작일 칸에 체크박스가 있으면, 왼쪽 끝을 클릭해 켜 본다 ──")
    w = UIAWrapper(ip.elem_of(h))
    import win32gui
    l, t, r, b = win32gui.GetWindowRect(h)
    print(f"   왼쪽 끝 클릭 (컨트롤 내 좌표 (7,{(b-t)//2}))")
    w.click_input(coords=(7, (b - t) // 2))
    time.sleep(0.5)
    read_all(h, "왼쪽끝 클릭 후")
    print()
    print("── 그 다음 년 구역 클릭 + 20260831 타이핑 ──")
    from pywinauto.keyboard import send_keys
    w.click_input(coords=(20, (b - t) // 2))
    time.sleep(0.3)
    send_keys("20260831", pause=0.12)
    time.sleep(0.6)
    read_all(h, "타이핑 후")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
