r"""
작업반 입력칸(txtDPTNBR) 진단 — 왜 값이 안 들어가는지 **읽는 방법별로** 확인한다.
⚠️ Windows 전용. iERP 로그인 상태에서 실행.

  py probe_team.py [작업반코드]     기본 M2105
"""
from __future__ import annotations

import sys
import time

import ierp_progress as ip
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.keyboard import send_keys


def read_all(h, tag):
    """같은 컨트롤을 방법별로 읽어 본다 — 어떤 방법이 값을 싣는지 확인."""
    out = {}
    try:
        out["UIA.window_text"] = UIAWrapper(ip.elem_of(h)).window_text()
    except Exception as e:
        out["UIA.window_text"] = f"!{e}"
    try:
        out["UIA.ValuePattern"] = UIAWrapper(ip.elem_of(h)).iface_value.CurrentValue
    except Exception as e:
        out["UIA.ValuePattern"] = f"!{type(e).__name__}"
    try:
        out["UIA.name"] = ip.elem_of(h).name
    except Exception as e:
        out["UIA.name"] = f"!{e}"
    try:
        import win32gui
        out["win32.GetWindowText"] = win32gui.GetWindowText(h)
    except Exception as e:
        out["win32.GetWindowText"] = f"!{e}"
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(512)
        ctypes.windll.user32.SendMessageW(h, 0x000D, 512, buf)      # WM_GETTEXT
        out["win32.WM_GETTEXT"] = buf.value
    except Exception as e:
        out["win32.WM_GETTEXT"] = f"!{e}"
    print(f"  [{tag}]")
    for k, v in out.items():
        print(f"      {k:22} = {v!r}")
    return out


def modals():
    import win32gui
    found = []

    def _cb(h, _):
        try:
            if win32gui.IsWindowVisible(h):
                t = (win32gui.GetWindowText(h) or "").strip()
                cls = win32gui.GetClassName(h)
                if t and "WindowsForms" in cls:
                    found.append((h, cls, t))
        except Exception:
            pass
        return True
    win32gui.EnumWindows(_cb, None)
    return found


def main():
    team = sys.argv[1] if len(sys.argv) > 1 else "M2105"
    scfg = ip.load_screen()
    top = ip.ensure_screen(scfg, log=print)
    print(f"■ 화면 hwnd={top}")

    before = {h for h, _, _ in modals()}
    ip.focus_window(top)
    time.sleep(0.6)

    import win32gui
    print(f"  포그라운드 = {win32gui.GetForegroundWindow()} (화면 hwnd={top}) "
          f"{'일치' if win32gui.GetForegroundWindow() == top else '★불일치 — 키 입력이 딴 데로 갑니다'}")

    edits = ip.controls_by_id(top, "EDIT")
    h = edits.get(scfg.get("team_field", "txtDPTNBR"))
    if not h:
        print(f"※ {scfg.get('team_field')} 를 못 찾음. 있는 Edit: {list(edits)}")
        return 1
    hn = edits.get(scfg.get("team_name_field", "txtDPTNBRD"))
    print(f"  txtDPTNBR hwnd={h} / txtDPTNBRD hwnd={hn}")
    print(f"  클래스 = {win32gui.GetClassName(h)!r} / "
          f"enabled={win32gui.IsWindowEnabled(h)} visible={win32gui.IsWindowVisible(h)}")

    print("\n── ① 입력 전 ─────────────────────────────")
    read_all(h, "txtDPTNBR")

    print(f"\n── ② 클릭 + 타이핑 '{team}' ───────────────")
    w = UIAWrapper(ip.elem_of(h))
    w.click_input()
    time.sleep(0.3)
    print(f"  클릭 후 포커스 hwnd = {win32gui.GetFocus() if False else '(스레드밖이라 조회불가)'}")
    send_keys("^a{DELETE}")
    time.sleep(0.2)
    send_keys(team, with_spaces=True, pause=0.05)
    time.sleep(0.5)
    read_all(h, "타이핑 직후(TAB 전)")

    print("\n── ③ TAB 으로 코드 확정 ───────────────────")
    send_keys("{TAB}")
    time.sleep(1.5)
    read_all(h, "TAB 후 txtDPTNBR")
    if hn:
        read_all(hn, "TAB 후 txtDPTNBRD(작업반명)")

    now = modals()
    new = [m for m in now if m[0] not in before]
    print(f"\n── ④ 새로 뜬 창(코드도움 팝업 등) ──────────")
    if new:
        for hh, cls, t in new:
            print(f"  ★ hwnd={hh} class={cls} title={t!r}")
    else:
        print("  없음")

    print("\n── ⑤ set_edit_text(win32) 로도 시도 ───────")
    try:
        from pywinauto.controls.win32_controls import EditWrapper
        from pywinauto.application import Application
        app = Application(backend="win32").connect(handle=top)
        e32 = EditWrapper(app.window(handle=h).wrapper_object().element_info)
        e32.set_edit_text(team)
        time.sleep(0.8)
        read_all(h, "set_edit_text 후")
    except Exception as e:
        print(f"  win32 set_edit_text 실패: {type(e).__name__}: {e}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
