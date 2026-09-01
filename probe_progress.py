r"""
정찰 — 공정진척현황(PM60250Rv3) 의 창 제목·컨트롤 ID 를 **추측 없이 실측**한다.
⚠️ Windows 전용. iERP 로그인(iEMenu 떠 있음) 상태에서 실행.

  py probe_progress.py            # 화면을 열고 컨트롤을 전부 덤프 → screen.json 갱신
  py probe_progress.py --titles   # 지금 떠 있는 창 제목만 (창 제목이 다를 때)

⚠️ **조회를 누르기 전에** 돌린다. 그리드에 행이 차면 UIA 탐색이 63초 뒤 실패한다.

결과: ierp_inspect/probe_progress_출력.txt  +  screen.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import ierp_progress as ip

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "ierp_inspect"
OUT_TXT = OUT_DIR / "probe_progress_출력.txt"

_lines: list[str] = []


def say(msg=""):
    print(msg, flush=True)
    _lines.append(str(msg))


def dump_titles():
    import win32gui
    say("── 지금 떠 있는 창 ─────────────────────────────")
    rows = []

    def _cb(h, _):
        try:
            if win32gui.IsWindowVisible(h):
                t = (win32gui.GetWindowText(h) or "").strip()
                if t:
                    rows.append((h, win32gui.GetClassName(h), t))
        except Exception:
            pass
        return True
    win32gui.EnumWindows(_cb, None)
    for h, cls, t in rows:
        say(f"  hwnd={h:<10} class={cls:<28} {t}")
    say(f"  총 {len(rows)}개")


def rect_of(h):
    try:
        import win32gui
        l, t, r, b = win32gui.GetWindowRect(h)
        return f"({l},{t})-({r},{b}) {r-l}x{b-t}"
    except Exception:
        return "?"


def main():
    if "--titles" in sys.argv:
        dump_titles()
        OUT_DIR.mkdir(exist_ok=True)
        OUT_TXT.write_text("\n".join(_lines), encoding="utf-8")
        return 0

    import win32gui
    scfg = ip.load_screen()
    say("=" * 62)
    say(" 공정진척현황(PM60250Rv3) 정찰")
    say("=" * 62)
    say("⚠️ 조회를 누르기 전 상태에서 돌리세요 (그리드가 차면 UIA 가 느려집니다)")
    say("")

    top = ip.ensure_screen(scfg, log=say)
    title = win32gui.GetWindowText(top)
    say(f"■ 화면 hwnd={top}")
    say(f"  제목  : {title!r}")
    say(f"  클래스: {win32gui.GetClassName(top)}")
    say(f"  크기  : {rect_of(top)}")
    say("")

    found = {}

    # ── 날짜 (DateTimePicker) ──────────────────────────────
    say("── 조회일자 (SysDateTimePick32) ─────────────────")
    dtps = ip.find_children_by_class(top, scfg["dtp_class"])
    for i, h in enumerate(dtps):
        try:
            aid = ip.elem_of(h).automation_id
        except Exception:
            aid = "?"
        say(f"  [{i}] hwnd={h:<10} auto_id={aid!r:<18} 값={ip.read_date(h)!r:<12} {rect_of(h)}")
    say(f"  → {len(dtps)}칸. "
        + ("2칸이면 왼쪽=시작 오른쪽=종료로 같은 날짜를 넣습니다."
           if len(dtps) >= 2 else "1칸이면 그 칸에만 넣습니다."))
    if not dtps:
        say("  ※ 날짜 칸을 못 찾았습니다 — 이 화면은 날짜 입력이 Edit 일 수 있습니다(아래 목록 확인)")
    say("")

    # ── Edit (작업반·품목 등) ──────────────────────────────
    say("── 입력칸 (Edit) — 여기서 작업반 칸을 고릅니다 ──")
    edits = ip.controls_by_id(top, "EDIT")
    ordered = sorted(edits.items(), key=lambda kv: (win32gui.GetWindowRect(kv[1])[1] // 12,
                                                    win32gui.GetWindowRect(kv[1])[0]))
    for aid, h in ordered:
        mark = "  ←작업반?" if "DPT" in aid.upper() else ""
        say(f"  auto_id={aid!r:<22} 값={ip._edit_text(h)!r:<14} {rect_of(h)}{mark}")
    say(f"  총 {len(edits)}개")
    code_h, name_h, _ = ip.resolve_team_fields(top, scfg, log=say)
    if scfg.get("team_field"):
        found["team_field"] = scfg["team_field"]
        say(f"  → 작업반 코드칸 = {scfg['team_field']!r}")
    if scfg.get("team_name_field"):
        found["team_name_field"] = scfg["team_name_field"]
        say(f"  → 작업반명 칸  = {scfg['team_name_field']!r}")
    if not scfg.get("team_field"):
        say("  ※ 작업반 칸을 자동으로 못 골랐습니다 — 위 목록에서 골라 "
            "screen.json 의 team_field 에 적으세요")
    say("")

    # ── ComboBox (오더상태 등) ─────────────────────────────
    say("── 콤보 (ComboBox) ─────────────────────────────")
    combos = ip.controls_by_id(top, "COMBO")
    for aid, h in combos.items():
        items = []
        try:
            from pywinauto.controls.uiawrapper import UIAWrapper
            w = UIAWrapper(ip.elem_of(h))
            items = [i.window_text() for i in w.descendants(control_type="ListItem")]
        except Exception:
            pass
        say(f"  auto_id={aid!r:<22} 값={ip._edit_text(h)!r:<14} {rect_of(h)}")
        if items:
            say(f"      항목({len(items)}): {items}")
    say(f"  총 {len(combos)}개")
    for aid in combos:
        if any(k in aid.upper() for k in ("STAT", "SOR")):
            found["status_combo"] = aid
            say(f"  → 오더상태 콤보로 추정 = {aid!r}")
            break
    say("")

    # ── CheckBox (조회조건) ────────────────────────────────
    say("── 조회조건 (CheckBox) — 기본 상태를 그대로 적어 둡니다 ──")
    boxes = ip.controls_by_id(top, "BUTTON")
    for aid, h in boxes.items():
        state = None
        try:
            from pywinauto.controls.uiawrapper import UIAWrapper
            state = bool(UIAWrapper(ip.elem_of(h)).get_toggle_state())
        except Exception:
            pass
        txt = (win32gui.GetWindowText(h) or "").strip()
        if state is None:
            continue
        say(f"  auto_id={aid!r:<22} {txt!r:<20} = {'켜짐' if state else '꺼짐'}")
    say("  ※ 형제 화면(일일생산실적)에서는 '작업마감제외' 가 켜진 채로 떠서 데이터의 99%가")
    say("    빠졌던 적이 있습니다. 위 목록에 그런 조건이 있으면 알려주세요.")
    say("")

    # ── 툴바 ───────────────────────────────────────────────
    say("── 툴바 버튼 ───────────────────────────────────")
    try:
        from pywinauto.uia_element_info import UIAElementInfo
        root = UIAElementInfo(top)
        for k in root.children():
            names = [(b.name or "").strip() for b in k.children()]
            names = [n for n in names if n]
            if names:
                say(f"  auto_id={k.automation_id!r:<18} 버튼: {names}")
                if any(n == scfg["query_btn"] for n in names):
                    found["toolbar_id"] = k.automation_id
                    say(f"  → 툴바 = {k.automation_id!r} ('조회' 있음)")
    except Exception as e:
        say(f"  ※ 툴바 덤프 실패: {e}")
    say("")

    if title:
        import re
        m = re.match(r"^(.*?)\s*\(PM60250Rv3\)", title)
        if m:
            found["title_re"] = "^" + re.escape(m.group(1).strip()) + r".*\(PM60250Rv3\).*"
            say(f"  → 창 제목 정규식 = {found['title_re']!r}")

    if found:
        ip.save_screen(found)
        say(f"\n■ screen.json 갱신: {json.dumps(found, ensure_ascii=False, indent=2)}")

    OUT_DIR.mkdir(exist_ok=True)
    OUT_TXT.write_text("\n".join(_lines), encoding="utf-8")
    say(f"\n결과 저장: {OUT_TXT}")
    say("\n다음 — 작업반 1개로 시험:  py run_daily.py --date <YYYY-MM-DD>")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
