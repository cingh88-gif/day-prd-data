r"""
iERP 공정진척현황(작업보고) **PM60250Rv3** 화면 조작.  ⚠️ Windows 전용.

★ 이 화면을 다루는 핵심 원칙 — **UIA 깊은 탐색을 절대 하지 않는다**
  (ierp-prod-report / outsourcing-cost 실측을 그대로 이식)
  조회 결과가 그리드에 수천 행 차면 컨트롤 하나 찾는 데:
      win32 EnumChildWindows          0.00초
      UIA 깊은 탐색(child_window)     63초 뒤 **실패**
  → 컨트롤은 ① win32 로 찾고 ② 핸들 바인딩 UIA(O(1))로 읽고 누른다.
  '조회가 끝났는지' 도 그리드가 아니라 **앱이 다시 응답하는지**로 판단한다.

★ 컨트롤 ID
  PM60250Rv3 의 ID 는 아직 실측 확정 전이다. 그래서 두 경로를 둔다.
    1) screen.json 이 있으면 그 값을 쓴다  ← probe_progress.py 가 만든다(권장)
    2) 없으면 형제 화면(PM40510Uv3)의 규칙으로 **자동 탐지**한다
  자동 탐지가 실패하면 **조용히 넘어가지 않고** 정찰을 안내하며 멈춘다.
"""
from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path

try:
    from pywinauto.keyboard import send_keys
    from pywinauto.uia_element_info import UIAElementInfo
    from pywinauto.controls.uiawrapper import UIAWrapper
except ImportError:                       # WSL 에서 import 만 할 때
    send_keys = None
    UIAElementInfo = UIAWrapper = None

# 출력이 파이프로 넘어가면 cp949 로 인코딩된다 — 로그 한 줄에 수집이 죽지 않도록.
try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
SCREEN_JSON = HERE / "screen.json"

SCREEN = dict(
    # 창 제목 뒤에 날짜가 붙으므로 앞부분만 고정한다.
    #   정찰 실측(2026-09-01): '공정진척현황(작업보고) (PM60250Rv3) 2025-11-03 오전 9:40:28'
    title_re=r"^공정진척현황.*\(PM60250Rv3\).*",
    # ★ 버전 접미사(v3)까지 붙여야 열린다. 'PM60250R' 로는 검색창에 값만 들어가고 안 열렸다.
    program_id="PM60250Rv3",
    dtp_class="SysDateTimePick32",
    toolbar_id="toolStrip1",
    query_btn="조회",
    export_btn="엑셀출력",
    # ★ 정찰 실측 확정(2026-09-01) — 형제 화면(PM40510Uv3)과 같았다.
    team_field="txtDPTNBR",           # 작업반 코드 (옆에 코드도움 버튼 cmdHelp0)
    team_name_field="txtDPTNBRD",     # 작업반명 — 코드 확정(TAB) 후 iERP 가 채운다(검증용)
    # 화면이 바뀌었을 때를 위한 후보 목록(위 확정값이 없을 때만 쓰인다)
    team_field_candidates=["txtDPTNBR", "txtDPTCD", "txtWCCD", "txtDEPT"],
    team_name_candidates=["txtDPTNBRD", "txtDPTNBRNM", "txtDPTCDD", "txtWCCDD"],
    # 조회기간 = **개시예정일**(dtpFRDATE ~ dtpTODATE). 작업일자·보고일자가 아니다.
    date_from_field="dtpFRDATE",
    date_to_field="dtpTODATE",
    # 오더상태 콤보(있으면). 기본은 건드리지 않고 ALL 로 받아 집계에서 30을 뺀다.
    status_combo_candidates=["cmbSORSTAT", "cmbSTAT", "cmbORDSTAT"],
    status_value="ALL",
    set_status=False,          # 이 화면엔 오더상태 조회조건이 없다(정찰 실측) → 결과 컬럼으로만 나온다
    # ★★ 조회조건 체크박스 — 이 화면의 진짜 함정 (2026-09-01 정찰 실측)
    #   chkSELOPT1('작업완료 포함')이 **꺼진 채로 뜬다.** 그대로 조회하면 80작업완료 건이
    #   통째로 빠져 **진행률이 매일 0%** 로 나온다. 우리 진행률의 분자가 바로 80작업완료다.
    #   형제 화면(일일생산실적)의 '작업마감제외' 와 같은 종류 — 총계 대조로는 못 잡는다.
    #   그래서 수집 시작 전에 **명시적으로 켜고, 읽어서 확인**한다.
    checkboxes={
        "chkSELOPT1": True,    # 작업완료 포함 — 켜야 80작업완료가 들어온다
    },
)

QUERY_WAIT = 1.0
QUERY_TIMEOUT = 1800
TEAM_SETTLE = 1.2          # 작업반 코드 확정(TAB) 후 iERP 가 작업반명을 조회하는 시간


def load_screen() -> dict:
    """screen.json(정찰 결과)이 있으면 덮어쓴다."""
    cfg = dict(SCREEN)
    if SCREEN_JSON.exists():
        try:
            cfg.update(json.loads(SCREEN_JSON.read_text(encoding="utf-8")))
        except Exception:
            pass
    return cfg


def save_screen(found: dict) -> None:
    cur = {}
    if SCREEN_JSON.exists():
        try:
            cur = json.loads(SCREEN_JSON.read_text(encoding="utf-8"))
        except Exception:
            cur = {}
    cur.update(found)
    SCREEN_JSON.write_text(json.dumps(cur, ensure_ascii=False, indent=2), encoding="utf-8")


# ====================================================================
# win32 계층 — 그리드 크기와 무관하게 항상 빠르다
# ====================================================================
def find_hwnd(title_re: str):
    """제목 정규식으로 최상위 창 핸들. 앱이 바빠도 동작한다(UIA 와 달리).
    Windows 가 붙이는 ' (응답 없음)' 꼬리표는 떼고 비교한다."""
    try:
        import win32gui
    except ImportError:
        return None
    pat, hits = re.compile(title_re), []

    def _cb(h, _):
        try:
            if win32gui.IsWindowVisible(h):
                t = re.sub(r"\s*\(응답 없음\)\s*$", "", win32gui.GetWindowText(h) or "")
                if pat.match(t):
                    hits.append(h)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        return None
    return hits[0] if hits else None


def find_hwnds(title_re: str) -> list[int]:
    """제목 정규식에 맞는 **모든** 최상위 창 핸들. 같은 창이 여러 개 쌓였을 때 쓴다."""
    try:
        import win32gui
    except ImportError:
        return []
    pat, hits = re.compile(title_re), []

    def _cb(h, _):
        try:
            if win32gui.IsWindowVisible(h):
                t = re.sub(r"\s*\(응답 없음\)\s*$", "", win32gui.GetWindowText(h) or "")
                if pat.match(t):
                    hits.append(h)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        return []
    return hits


def find_children_by_class(top, class_sub: str) -> list[int]:
    """클래스명에 class_sub 가 들어가는 모든 자식 핸들. 화면 좌표순(위→아래, 왼→오른쪽).

    ★ 조회기간이 시작/종료 두 칸이면 왼쪽이 시작, 오른쪽이 종료가 되게 정렬한다."""
    try:
        import win32gui
    except ImportError:
        return []
    found = []

    def _cb(h, _):
        try:
            if class_sub in win32gui.GetClassName(h):
                found.append(h)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumChildWindows(top, _cb, None)
    except Exception:
        return []

    def key(h):
        try:
            l, t, r, b = win32gui.GetWindowRect(h)
            return (t // 12, l)
        except Exception:
            return (0, 0)
    return sorted(found, key=key)


def is_responsive(hwnd, ms: int = 300) -> bool:
    """창이 메시지를 처리하고 있는가(= 조회가 끝났는가). UIA 를 안 쓰므로 앱을 방해하지 않는다."""
    if not hwnd:
        return True
    try:
        import ctypes
        res = ctypes.c_ulong()
        ok = ctypes.windll.user32.SendMessageTimeoutW(
            hwnd, 0x0000, 0, 0, 0x0002, ms, ctypes.byref(res))    # WM_NULL, SMTO_ABORTIFHUNG
        return bool(ok)
    except Exception:
        return True


def wait_responsive(hwnd, timeout: float, quiet: float = 5.0, poll: float = 2.0,
                    should_stop=None) -> bool:
    """창이 다시 응답하고 quiet 초 동안 계속 응답할 때까지 대기.

    ★ 조회가 끝나기 전에 '엑셀출력' 을 누르면 **행 수는 맞는데 전부 빈 칸인 파일**이 나온다.
      그래서 이 대기는 생략 불가."""
    end = time.time() + timeout
    ok_since = None
    while time.time() < end:
        if should_stop and should_stop():
            return False
        if is_responsive(hwnd):
            ok_since = ok_since or time.time()
            if time.time() - ok_since >= quiet:
                return True
        else:
            ok_since = None
        time.sleep(poll)
    return False


# ====================================================================
# UIA 얕은 계층 — 핸들 바인딩 / 직계 자식만
# ====================================================================
def elem_of(hwnd):
    """HWND → UIA 엘리먼트. 트리 탐색이 아니라 O(1)."""
    return UIAElementInfo(hwnd)


def controls_by_id(top, class_sub: str = "") -> dict:
    """{automation_id: hwnd} — win32 로 자식을 훑고(즉시), 핸들마다 UIA 를 O(1) 로 바인딩해
    automation_id 만 읽는다. 깊은 탐색이 아니라 그리드가 차 있어도 빠르다."""
    try:
        import win32gui
    except ImportError:
        return {}
    hs = []

    def _cb(h, _):
        try:
            if not class_sub or class_sub.upper() in win32gui.GetClassName(h).upper():
                hs.append(h)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumChildWindows(top, _cb, None)
    except Exception:
        return {}
    out = {}
    for h in hs:
        try:
            aid = elem_of(h).automation_id
        except Exception:
            continue
        if aid and aid not in out:
            out[aid] = h
    return out


def toolbar_button(top_hwnd, name: str, toolbar_id: str = "toolStrip1"):
    """툴바 버튼 wrapper. 최상위 창의 직계 자식 → 툴바의 직계 자식만 본다(0.1초).

    ⚠️ win.child_window(title='조회') 같은 깊은 탐색은 그리드가 찬 뒤 63초 만에 실패한다."""
    root = UIAElementInfo(top_hwnd)
    tb = next((k for k in root.children() if k.automation_id == toolbar_id), None)
    if tb is None:
        for k in root.children():           # 툴바 id 가 다른 화면일 수 있다
            el = next((b for b in k.children() if (b.name or "").strip() == name), None)
            if el is not None:
                return UIAWrapper(el)
        raise RuntimeError(f"툴바({toolbar_id})를 찾지 못했습니다 — probe_progress.py 로 확인하세요")
    el = next((b for b in tb.children() if (b.name or "").strip() == name), None)
    if el is None:
        raise RuntimeError(f"툴바 버튼 '{name}' 을 찾지 못했습니다")
    return UIAWrapper(el)


def click_button(top_hwnd, name: str, cfg: dict):
    toolbar_button(top_hwnd, name, cfg.get("toolbar_id", "toolStrip1")).click_input()


# ====================================================================
# 사전점검 — 권한(무결성 레벨)이 맞는가
# ====================================================================
INTEGRITY_RID = {"0": ("Untrusted", 0), "4096": ("Low", 1), "8192": ("Medium", 2),
                 "8448": ("Medium+", 3), "12288": ("High(관리자)", 4),
                 "16384": ("System", 5)}


def process_integrity(pid: int):
    """(레벨이름, 순위, 관리자상승여부). 못 읽으면 (None, None, None)."""
    try:
        import win32api
        import win32con
        import win32security
        h = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        tok = win32security.OpenProcessToken(h, win32con.TOKEN_QUERY)
        sid = win32security.GetTokenInformation(tok, win32security.TokenIntegrityLevel)[0]
        rid = win32security.ConvertSidToStringSid(sid).rsplit("-", 1)[-1]
        name, rank = INTEGRITY_RID.get(rid, (f"RID {rid}", None))
        elev = bool(win32security.GetTokenInformation(tok, win32security.TokenElevation))
        return name, rank, elev
    except Exception:
        return None, None, None


def window_pid(hwnd) -> int | None:
    try:
        import win32process
        return int(win32process.GetWindowThreadProcessId(hwnd)[1])
    except Exception:
        return None


def preflight(log=print) -> None:
    """수집 시작 전 1회 점검. 문제가 있으면 **원인을 짚어** RuntimeError 를 던진다.

    ★★ 형제 프로젝트(ierp-manhour) 2026-07-27 실측 사고:
      iERP 를 **관리자 권한으로** 띄운 상태에서 보통 권한 스크립트를 돌렸더니 UIPI 에 막혀
      UIA 가 iEMenu 를 **0개**로 보았고, 매달 TimeoutError 로 죽어 **13개월을 통째로
      건너뛰었다.** 창은 win32 로 멀쩡히 보이므로 증상만 봐서는 원인을 알 수 없다.
    → 시작 전에 **무결성 레벨을 직접 비교**하고, UIA 핸들 접근까지 확인한다.
      (2026-09-02 실측 정상값: 이 도구도 iERP 도 둘 다 Medium / 관리자 아님)"""
    menu = find_hwnd(r"^iEMenu$") or find_hwnd(r".*iEMenu.*")
    if not menu:
        raise RuntimeError(
            "iEMenu 창을 찾을 수 없습니다.\n"
            "→ iERP 에 로그인해서 메뉴 창(iEMenu)을 띄운 뒤 다시 실행하세요.")

    import os
    my_name, my_rank, my_elev = process_integrity(os.getpid())
    pid = window_pid(menu)
    er_name, er_rank, er_elev = process_integrity(pid) if pid else (None, None, None)

    if my_rank is not None and er_rank is not None:
        if my_rank < er_rank:
            raise RuntimeError(
                f"권한이 맞지 않아 화면 자동화가 차단됩니다(UIPI).\n"
                f"   이 도구 : {my_name}\n"
                f"   iERP    : {er_name}   ← 더 높습니다\n"
                f"→ 해결: ① iERP 를 관리자 권한 **없이** 다시 실행하거나(권장)\n"
                f"        ② 이 프로그램을 우클릭 → '관리자 권한으로 실행'\n"
                f"   (Excel 도 iERP 가 띄우므로 셋의 권한이 같아야 합니다.)")
        if my_rank > er_rank:
            log(f"   ※ 권한이 다릅니다 — 이 도구 {my_name} / iERP {er_name}. "
                f"입력은 되지만 Excel COM 이 막힐 수 있습니다. "
                f"둘 다 관리자 없이 실행하는 것을 권합니다.")

    # UIPI 백스톱 — 레벨을 못 읽었더라도 핸들 접근이 막히면 자동화가 안 된다.
    try:
        _ = elem_of(menu).name
    except Exception as e:
        raise RuntimeError(
            f"iEMenu 창은 떠 있는데(hwnd={menu}) UIA 핸들 접근이 막혀 있습니다({e}).\n"
            f"→ iERP 와 이 도구의 실행 권한이 다를 때 나타납니다. 둘 다 관리자 권한 없이 "
            f"실행하세요.")

    log(f"   사전점검 OK — iEMenu hwnd={menu} / 권한 {my_name or '?'}"
        + (f" = iERP {er_name}" if er_name else ""))


# ====================================================================
# 화면 열기 / 재접속
# ====================================================================
def force_foreground(hwnd, tries: int = 5, log=print) -> bool:
    """창을 확실히 앞으로 가져오고 **정말 앞에 왔는지 확인**한다.

    ★★ 2026-09-01~02 실측 — 두 번 당했다.
      `SetForegroundWindow` 는 호출 프로세스가 포그라운드가 아니면 Windows 가 **조용히
      거부**한다(반환값만 실패, 예외 없음). 그걸 무시하고 `send_keys` 를 하면 키가
      **그때 포커스를 가진 아무 창에나 들어간다.** 무인 스케줄 실행에서 이러면
      'PM60250Rv3' 과 Enter 가 사용자의 메모장이나 브라우저에 찍힌다.
    → ① 최소화돼 있으면 복원 ② 포그라운드 스레드에 입력 큐를 붙여(AttachThreadInput)
      권한을 빌린 뒤 SetForegroundWindow ③ **확인**하고 아니면 재시도.
      끝까지 실패하면 **키를 보내지 않고** False 를 돌려준다."""
    try:
        import win32gui
        import win32con
        import ctypes
    except ImportError:
        return False
    user32 = ctypes.windll.user32
    for i in range(tries):
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.3)
            fg = user32.GetForegroundWindow()
            if fg == hwnd:
                return True
            # ⚠️ GetCurrentThreadId 는 win32process 가 아니라 kernel32 에 있다.
            cur_tid = ctypes.windll.kernel32.GetCurrentThreadId()
            fg_tid = user32.GetWindowThreadProcessId(fg, None) if fg else 0
            attached = False
            if fg_tid and fg_tid != cur_tid:
                attached = bool(user32.AttachThreadInput(fg_tid, cur_tid, True))
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            finally:
                if attached:
                    user32.AttachThreadInput(fg_tid, cur_tid, False)
            time.sleep(0.4)
            if user32.GetForegroundWindow() == hwnd:
                return True

            # ④ 그래도 안 되면 한 단계 더 — 2026-09-02 실측으로 통한 조합.
            #   막고 있던 것은 사용자가 열어 둔 **다른 iERP 화면**이었다.
            #   · ALT 를 눌렀다 떼서 '마지막 입력을 받은 프로세스' 조건을 만족시키고
            #   · TOPMOST 로 올렸다가 바로 되돌려 Z-순서를 끌어올린다
            #     (TOPMOST 로 두면 사용자 창을 계속 덮으므로 반드시 되돌린다)
            user32.keybd_event(0x12, 0, 0, 0)          # VK_MENU down
            time.sleep(0.05)
            user32.keybd_event(0x12, 0, 2, 0)          # KEYEVENTF_KEYUP
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            time.sleep(0.4)
            if user32.GetForegroundWindow() == hwnd:
                return True
        except Exception as e:
            log(f"      · 창 활성화 시도 {i + 1} 실패({e})")
        time.sleep(0.5)
    return False


SEARCH_DLG_RE = r"^프로그램 실행$"


def _find_search_dialog(timeout: float = 5.0):
    """F9 로 뜨는 '프로그램 실행' 검색창을 기다린다(없으면 None)."""
    end = time.time() + timeout
    while time.time() < end:
        h = find_hwnd(SEARCH_DLG_RE)
        if h:
            return h
        time.sleep(0.25)
    return None


def _close_search_dialog(log=print) -> int:
    r"""남아 있는 검색창을 **전부** 닫는다. 반환: 닫은 개수.

    ★★ 2026-09-02 실측 — 이걸 안 해서 화면이 아예 안 열렸다.
      실패할 때마다 F9 로 새 검색창을 띄우면서 닫지 않았더니 `프로그램 실행` 창이
      **4개 쌓였고**, 그 모달 창들이 iEMenu 를 막아 어떤 시도도 화면을 열지 못했다.
      증상은 '열기를 눌러도 화면이 안 뜸' 이라 원인이 전혀 안 보였다.
    → 시도 전마다 **모두** 닫는다(하나만 닫으면 소용없다)."""
    closed = 0
    for _ in range(10):                       # 쌓여 있을 수 있으니 반복
        hs = find_hwnds(SEARCH_DLG_RE)
        if not hs:
            break
        for h in hs:
            done = False
            try:
                btn = next((b for b in UIAWrapper(UIAElementInfo(h)).descendants(
                    control_type="Button") if (b.window_text() or "").strip() == "닫기"), None)
                if btn is not None:
                    btn.click_input()
                    done = True
            except Exception:
                pass
            if not done:                      # UIA 로 안 되면 창에 직접 닫기 메시지
                try:
                    import win32con
                    import win32gui
                    win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                except Exception:
                    continue
            closed += 1
            time.sleep(0.4)
    if closed:
        log(f"      · 남아 있던 검색창 {closed}개를 닫았습니다")
    return closed


def esc_keys(text: str) -> str:
    r"""send_keys 특수문자 이스케이프.

    ★ 2026-09-02 실측: `공정진척현황(작업보고)` 을 그냥 보냈더니 **괄호가 통째로 사라져**
      `공정진척현황작업보고` 가 입력됐다. `^ + % ~ ( ) { } [ ]` 는 send_keys 문법 문자다."""
    special = "^+%~(){}[]"
    return "".join("{%s}" % c if c in special else c for c in text)


def _dialog_combo(dlg):
    """검색창의 입력 콤보 핸들.

    ⚠️ 이건 Edit 이 아니라 **편집형 ComboBox**(WindowsForms10.COMBOBOX)다.
      안에 Edit(auto_id 1001)이 들어 있고, 목록 항목은 0개(이력용)다."""
    try:
        import win32gui
    except ImportError:
        return None
    found = []

    def _cb(h, _):
        try:
            if "COMBOBOX" in win32gui.GetClassName(h).upper():
                found.append(h)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumChildWindows(dlg, _cb, None)
    except Exception:
        return None
    return found[0] if found else None


def open_via_search(program_id: str, title_re: str = "", log=print, tries: int = 3):
    r"""iEMenu 에서 F9(메뉴검색) → 프로그램ID 입력 → **Enter**.
    같은 화면이 이미 떠 있으면 iERP 가 **기존 창을 재사용**한다(중복으로 안 열림).

    ⚠️ iEMenu 를 확실히 앞으로 못 가져오면 **키를 아예 보내지 않는다.**
      엉뚱한 창에 프로그램ID 를 타이핑하는 것보다 실패하는 게 낫다.

    ★★ 2026-09-02 실측 — 여기서 반나절을 헤맸다. 확정된 사실:
      · 검색창에 **'열기' 라는 이름의 Button 이 있지만 그건 콤보박스의 드롭다운 화살표다**
        (15x18px). 그걸 누르면 목록만 펼쳐지고 **프로그램은 안 열린다.**
        UIA 가 콤보 드롭다운 버튼에 붙이는 표준 한국어 이름이 하필 '열기' 다.
      · 이 창에 확인/실행 버튼은 없다. **제출 수단은 Enter 뿐이다.**
      · 실패한 시도의 검색창을 안 닫으면 **창이 쌓여**(4개까지 봤다) iEMenu 가 막혀
        그 뒤 어떤 시도도 화면을 열지 못한다. 증상이 원인과 전혀 안 닮아서 오진하기 쉽다.
      · 입력은 Edit 이 아니라 편집형 ComboBox 다. 좌표로 텍스트 영역을 클릭해 포커스를 준다.
    → 순서: 남은 창 전부 닫기 → F9 → 창 확인 → 콤보 클릭 → 입력 → **값 확인** → Enter
            → 화면 확인. 실패하면 창을 닫고 다시.
    """
    menu = find_hwnd(r"^iEMenu$") or find_hwnd(r".*iEMenu.*")
    if not menu:
        raise RuntimeError("iEMenu 창을 찾지 못했습니다 — iERP 에 로그인돼 있는지 확인하세요.")

    for attempt in range(1, tries + 1):
        _close_search_dialog(log=log)          # ★ 안 닫으면 창이 쌓여 전부 막힌다
        if not force_foreground(menu, log=log):
            raise RuntimeError(
                "iEMenu 를 앞으로 가져오지 못했습니다 — 키 입력을 보내지 않고 멈춥니다.\n"
                "  (다른 창이 포커스를 붙잡고 있습니다. 전체화면 앱이나 대화상자를 닫고 "
                "다시 실행하세요.)")
        time.sleep(0.4)

        # ① 검색창을 띄우고 **떴는지 확인**
        send_keys("{F9}")
        dlg = _find_search_dialog(5.0)
        if not dlg:
            log(f"      · F9 로 검색창이 안 떴습니다({attempt}/{tries})")
            continue

        # ② 콤보에 포커스를 주고 입력한 뒤 **들어갔는지 확인**
        force_foreground(dlg, tries=2, log=log)
        time.sleep(0.4)
        combo = _dialog_combo(dlg)
        try:
            if combo:
                UIAWrapper(elem_of(combo)).click_input(coords=(40, 10))   # 텍스트 영역
                time.sleep(0.3)
            send_keys("^a{DELETE}")
            time.sleep(0.2)
            send_keys(esc_keys(program_id), with_spaces=True, pause=0.05)
            time.sleep(0.7)
        except Exception as e:
            log(f"      · 프로그램ID 입력 예외({e})")
            _close_search_dialog(log=log)
            continue
        got = _edit_text(combo) if combo else program_id
        if combo and got.strip().upper() != program_id.strip().upper():
            log(f"      · 프로그램ID 가 안 들어갔습니다(화면값 {got!r})({attempt}/{tries})")
            _close_search_dialog(log=log)
            continue

        # ③ **Enter 로 제출** — '열기' 버튼은 콤보 드롭다운이라 누르면 안 된다
        send_keys("{ENTER}")

        # ④ 화면이 **실제로 떴는지** 확인
        for i in range(20):
            time.sleep(1.0)
            h = find_hwnd(title_re) if title_re else None
            if h:
                log(f"      · 화면 열림({i + 1}초, hwnd={h})")
                return h
        log(f"      · Enter 를 눌렀지만 화면이 뜨지 않았습니다({attempt}/{tries})")
        _close_search_dialog(log=log)

    raise RuntimeError(
        f"iEMenu 검색창으로 {program_id} 를 열지 못했습니다({tries}회 시도).\n"
        f"  · iERP 로그인 상태를 확인하세요\n"
        f"  · 프로그램ID 는 버전 접미사까지 필요합니다(예: PM60250Rv3)\n"
        f"  · iEMenu 에서 직접 F9 로 열리는지 확인해 보세요")


def ensure_screen(cfg: dict, log=print, busy_wait: float = QUERY_TIMEOUT):
    """화면 핸들을 돌려준다. 없으면 F9 로 연다. 바쁘면 한가해질 때까지 기다린다.

    ⚠️ '바쁘다' 를 '없다' 로 오해해 F9 를 누르면 안 된다 — 그래서 핸들 확인이 먼저다."""
    h = find_hwnd(cfg["title_re"])
    if h:
        if not is_responsive(h):
            log("   화면이 아직 작업 중입니다 — 끝날 때까지 기다립니다…")
            wait_responsive(h, timeout=busy_wait, quiet=3.0)
        return h
    log(f"   화면 자동 열기(F9 검색): {cfg['program_id']}")
    h = open_via_search(cfg["program_id"], cfg["title_re"], log=log)
    if h:
        wait_responsive(h, timeout=60, quiet=2.0)
        return h
    for _ in range(10):
        h = find_hwnd(cfg["title_re"])
        if h:
            wait_responsive(h, timeout=60, quiet=2.0)
            return h
        time.sleep(1.0)
    raise RuntimeError(
        f"화면({cfg['program_id']})을 열지 못했습니다 — iERP 로그인 상태와 "
        f"probe_progress.py --titles 로 창 제목을 확인하세요.")


def focus_window(hwnd, log=print) -> bool:
    """화면을 앞으로. 실패해도 진행은 하지만(클릭이 포커스를 가져오므로) 결과를 돌려준다."""
    return force_foreground(hwnd, tries=3, log=log)


# ====================================================================
# 작업반 입력
# ====================================================================
def _edit_text(hwnd) -> str:
    """Edit 값 읽기.

    ★★ 2026-09-01 실측 — 여기서 크게 틀렸다 (probe_team.py 로 방법별 대조).
      같은 컨트롤(txtDPTNBR, 값이 실제로 'M2105' 인 상태)을 방법별로 읽으면:
          UIA.window_text()      → ''        ← 처음에 쓰던 것
          win32.GetWindowText    → ''        ← 폴백도 마찬가지
          UIA ValuePattern       → 'M2105'   ✅
          win32 WM_GETTEXT       → 'M2105'   ✅
      `GetWindowText` 는 **다른 프로세스**의 자식 컨트롤에 대해 빈 값을 돌려준다
      (캐시된 제목만 준다 — 값을 받으려면 WM_GETTEXT 를 직접 보내야 한다).
      UIA 의 window_text() 도 이 WinForms TextBox 에서 Name 을 읽어 비어 있다.
    → **ValuePattern 을 먼저, 그다음 WM_GETTEXT.** 이걸 몰라서 값이 정상으로 들어갔는데도
      '작업반이 화면에 안 들어갔습니다' 로 오판하고 전 작업반을 건너뛰었다.
    """
    # ① UIA ValuePattern — WinForms TextBox 의 .Text 가 여기 실린다
    try:
        v = UIAWrapper(elem_of(hwnd)).iface_value.CurrentValue
        if v:
            return str(v).strip()
    except Exception:
        pass
    # ② WM_GETTEXT — 프로세스 경계를 넘어 실제 값을 가져온다
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(512)
        ctypes.windll.user32.SendMessageW(hwnd, 0x000D, 512, buf)      # WM_GETTEXT
        if buf.value:
            return buf.value.strip()
    except Exception:
        pass
    # ③ 마지막 폴백 (대개 빈 값이지만 컨트롤에 따라 이게 되는 경우도 있다)
    try:
        return (UIAWrapper(elem_of(hwnd)).window_text() or "").strip()
    except Exception:
        return ""


def resolve_team_fields(top, cfg: dict, log=print) -> tuple[int | None, int | None, dict]:
    """작업반 코드/명 입력칸 핸들을 찾는다.

    ① screen.json 에 확정값이 있으면 그것
    ② 없으면 automation_id 후보 목록
    ③ 그래도 없으면 'DPT' 가 들어간 Edit 을 좌표순으로 (코드, 명) 으로 추정
    """
    edits = controls_by_id(top, "EDIT")
    code_h = name_h = None

    fixed = cfg.get("team_field")
    if fixed and fixed in edits:
        code_h = edits[fixed]
    else:
        for aid in cfg["team_field_candidates"]:
            if aid in edits:
                code_h = edits[aid]
                cfg["team_field"] = aid
                break

    fixed_n = cfg.get("team_name_field")
    if fixed_n and fixed_n in edits:
        name_h = edits[fixed_n]
    else:
        for aid in cfg["team_name_candidates"]:
            if aid in edits:
                name_h = edits[aid]
                cfg["team_name_field"] = aid
                break

    if code_h is None:
        import win32gui
        cands = [(aid, h) for aid, h in edits.items() if "DPT" in aid.upper()]
        cands.sort(key=lambda kv: win32gui.GetWindowRect(kv[1])[0])
        if cands:
            code_h = cands[0][1]
            cfg["team_field"] = cands[0][0]
            log(f"   작업반 입력칸 자동 추정: {cands[0][0]}")
            if name_h is None and len(cands) > 1:
                name_h = cands[1][1]
                cfg["team_name_field"] = cands[1][0]
    return code_h, name_h, edits


def set_team(top, cfg, code_h, name_h, team: str, tries: int = 3,
             log=print) -> tuple[bool, str]:
    """작업반 코드 입력 → TAB 으로 확정 → 작업반명이 채워지는지 확인.

    ⚠️ 값을 넣고 **읽어서 확인**한다. 엉뚱한 작업반 데이터를 그 작업반 파일로 저장하는 게
      이 도구의 최악의 사고다. 이 화면의 엑셀에는 **작업반 열이 없어서**(작업장 열만 있다)
      저장 후 대조로는 못 잡는다 — 여기서 막는 게 유일한 방어선이다.

    ★★ 2026-09-01 실측 — 여기서 2개 작업반을 놓쳤다.
      앞 작업반을 마친 뒤 `cleanup_excel()` 이 Excel 인스턴스를 닫는데, 그때 **포커스가
      잠깐 Excel 로 갔다가 돌아온다.** 그 사이에 클릭+타이핑을 하면 키가 사라져 입력칸이
      **이전 작업반 값 그대로** 남고, 그대로 조회하면 앞 작업반 데이터를 다음 작업반
      파일로 저장하게 된다(실제로는 대조에 걸려 건너뛰었다).
    → ① 입력 직전에 창을 다시 앞으로 가져오고 ② 실패하면 재시도한다."""
    if code_h is None:
        return False, ""
    for attempt in range(1, tries + 1):
        try:
            focus_window(top)              # ★ Excel 정리 직후 포커스가 흔들린 상태를 복구
            time.sleep(0.3)
            w = UIAWrapper(elem_of(code_h))
            w.click_input()
            time.sleep(0.25)
            send_keys("^a{DELETE}")
            time.sleep(0.15)
            send_keys(team, with_spaces=True, pause=0.04)
            time.sleep(0.2)
            send_keys("{TAB}")             # 코드 확정 → iERP 가 작업반명을 조회한다
            time.sleep(TEAM_SETTLE)
        except Exception as e:
            log(f"      ※ 작업반 입력 예외({e})")
            time.sleep(0.5)
            continue

        got = _edit_text(code_h)
        name = _edit_text(name_h) if name_h else ""
        if got.strip().upper() == team.strip().upper():
            log(f"   작업반 {team} → 화면값={got!r} 작업반명={name!r} OK")
            return True, name
        log(f"   작업반 {team} → 화면값={got!r} ※불일치({attempt}/{tries}) — 재시도")
        time.sleep(0.6)
    return False, ""


# ====================================================================
# 조회일자 (DateTimePicker)
# ====================================================================
def read_date(dtp_hwnd) -> str:
    """DTP 표시값에서 숫자만. '2026-08-31' → '20260831'.
    ⚠️ win32 GetWindowText 는 DTP 에서 빈 문자열이 나온다 → UIA 핸들 바인딩으로 읽는다."""
    try:
        return re.sub(r"\D", "", elem_of(dtp_hwnd).name or "")
    except Exception:
        return ""


def materialize_dtp(dtp_hwnd, log=print) -> bool:
    r"""**빈 DateTimePicker 를 살린다.** ▼(달력) 을 클릭하고 Enter → 오늘 날짜가 들어간다.

    ★★ 2026-09-01 실측 (probe_date.py / probe_shot.py):
      이 화면의 `dtpFRDATE`(개시예정일 시작)는 열릴 때 **비어 있다**. WinForms 가
      `CustomFormat=" "` 로 빈 칸처럼 보이게 한 상태라 **편집할 날짜 구역 자체가 없어서
      클릭+타이핑이 통째로 무시된다**(3회 재시도가 전부 ''로 실패했다).
      값을 넣는 다른 길도 다 막혀 있었다:
          DTM_SETSYSTEMTIME  → 구조체 포인터라 **프로세스 경계를 못 넘는다**(ret=0)
          Legacy IAccessible → SetValue 없음(AttributeError)
          UIA ValuePattern   → NoPatternInterfaceError
      되는 것은 하나였다: **▼ 클릭 → Enter.** 오늘 날짜가 들어가면서 포맷이 살아나고,
      그 뒤에는 평소대로 8자리 타이핑이 먹는다.
    """
    if send_keys is None:
        return False
    try:
        import win32gui
        l, t, r, b = win32gui.GetWindowRect(dtp_hwnd)
        UIAWrapper(elem_of(dtp_hwnd)).click_input(coords=(r - l - 10, (b - t) // 2))
        time.sleep(1.0)
        send_keys("{ENTER}")                 # 달력에서 오늘 선택 → 값이 생긴다
        time.sleep(0.6)
    except Exception as e:
        log(f"      ※ 빈 날짜칸 활성화 예외({e})")
        return False
    if read_date(dtp_hwnd):
        return True
    try:
        send_keys("{ESC}")                   # 달력이 떠 있으면 닫는다
    except Exception:
        pass
    return False


def set_date(dtp_hwnd, ymd: str, tries: int = 3, log=print) -> bool:
    """날짜를 넣고 **읽어서 확인**한다. ymd = 'YYYYMMDD'.

    ⚠️ DTM_SETSYSTEMTIME 으로 바꾸면 표시만 바뀌고 WinForms 의 .Value 는 그대로라
      이전 날짜로 조회된다 → 반드시 실제 키 입력.
    ⚠️ {LEFT} 로 첫 칸을 맞추면 안 된다 — DTP 의 좌우 이동은 **순환**이라 엉뚱한 칸에 간다.
    → 년 구역을 클릭하고 화살표 없이 8자리를 친다. 칸이 차면 자동으로 다음 칸으로 넘어간다."""
    if send_keys is None:
        return False
    digits = re.sub(r"\D", "", ymd)
    # ★ 칸이 비어 있으면 편집할 구역이 없어 타이핑이 무시된다 → 달력으로 먼저 살린다.
    if not read_date(dtp_hwnd):
        log("      · 날짜 칸이 비어 있습니다 — 달력(▼)으로 값을 살립니다")
        if not materialize_dtp(dtp_hwnd, log=log):
            log("      ※ 빈 날짜 칸을 살리지 못했습니다")
            return False
    for attempt in range(1, tries + 1):
        try:
            UIAWrapper(elem_of(dtp_hwnd)).click_input(coords=(14, 10))    # 년 구역
            time.sleep(0.25)
            send_keys(digits, pause=0.12)
            time.sleep(0.4)
        except Exception as e:
            log(f"      ※ 일자 입력 예외({e})")
            return False
        if read_date(dtp_hwnd) == digits:
            return True
        log(f"      ※ 일자 확인 실패({attempt}/{tries}): 화면값 {read_date(dtp_hwnd)!r} "
            f"≠ {digits!r} — 재시도")
        time.sleep(0.5)
    return False


def set_period(dtps: list[int], d_from: str, d_to: str | None = None,
               log=print) -> bool:
    """개시예정일 기간을 시작/종료 DTP 에 넣는다.

    하루치면 d_from == d_to 로 같은 날짜가 들어간다(일마감).
    월마감이면 그 달 1일~말일이 들어간다.
    ⚠️ 날짜 칸이 1개뿐인 화면이면 시작일만 넣는다 — 그 경우 기간 조회가 안 되므로
      월 단위로 받을 수 없다(호출자가 판단하도록 False 가 아니라 True 를 돌려주되 로그를 남긴다)."""
    if not dtps:
        return False
    d_to = d_to or d_from
    if not set_date(dtps[0], d_from, log=log):
        return False
    if len(dtps) >= 2:
        return set_date(dtps[1], d_to, log=log)
    if d_from != d_to:
        log("      ※ 날짜 칸이 1개뿐이라 기간 조회를 할 수 없습니다 — 시작일만 넣었습니다")
    return True


def read_period(dtps: list[int]) -> str:
    return " ~ ".join(read_date(d) for d in dtps[:2])


# ====================================================================
# 조회조건 체크박스
# ====================================================================
def read_check(hwnd) -> bool | None:
    """체크 상태를 읽는다. **핸들 바인딩 UIA** 로 읽는다 — O(1) 이라 그리드가 차 있어도 빠르다.

    ⚠️ win32 BM_GETCHECK(0x00F0) 로 읽으면 안 된다. 메시지는 성공으로 돌아오는데 **값이 항상 0**
      이다(WinForms 체크박스는 이 메시지에 상태를 싣지 않는다). 형제 프로젝트에서 이걸로
      24개월 수집이 통째로 잘못된 조건으로 돌았다."""
    try:
        return bool(UIAWrapper(elem_of(hwnd)).get_toggle_state())
    except Exception:
        return None


def set_checkboxes(top, cfg, log=print) -> bool:
    """조회조건 체크박스를 cfg['checkboxes'] 대로 맞추고 **읽어서 확인**한다.

    ⚠️ 값은 **실제 마우스 클릭**으로 바꾼다. BM_SETCHECK 으로 바꾸면 표시만 바뀌고 WinForms 의
      .Checked 가 그대로일 수 있다(DTP 의 DTM_SETSYSTEMTIME 과 같은 함정)."""
    want = cfg.get("checkboxes") or {}
    if not want:
        return True
    boxes = controls_by_id(top, "BUTTON")
    if not boxes:
        log("   ※ 체크박스를 하나도 찾지 못했습니다 — 화면 구성을 확인하세요(probe_progress.py)")
        return False

    ok = True
    for aid, target in want.items():
        h = boxes.get(aid)
        if h is None:
            log(f"   ※ 체크박스 {aid} 없음 — 화면 구성이 바뀌었을 수 있습니다")
            ok = False
            continue
        try:
            import win32gui
            name = (win32gui.GetWindowText(h) or "").strip() or aid
        except Exception:
            name = aid
        cur = read_check(h)
        if cur is None:
            log(f"   ※ {aid}({name}) 상태를 읽지 못했습니다")
            ok = False
            continue
        if cur == target:
            log(f"   조회조건 {name}({aid}) = {'켜짐' if cur else '꺼짐'} (그대로)")
            continue
        try:
            UIAWrapper(elem_of(h)).click_input()      # 실제 클릭
            time.sleep(0.3)
            now = read_check(h)
        except Exception as e:
            log(f"   ※ {aid}({name}) 클릭 실패({e})")
            ok = False
            continue
        if now == target:
            log(f"   조회조건 {name}({aid}) = {'켜짐' if now else '꺼짐'} 으로 바꿈")
        else:
            log(f"   ※ {aid}({name}) 를 {'켜짐' if target else '꺼짐'} 으로 못 바꿨습니다 (현재 {now})")
            ok = False
    return ok


# ====================================================================
# 오더상태 콤보 (기본은 건드리지 않는다)
# ====================================================================
def set_status(top, cfg, value: str, log=print) -> str:
    """상태 콤보 선택. ①win32 select 로 강조 → ②'열기' 를 실제 클릭 → ③{ENTER}.
    ★ ①만으론 SelectionChangeCommitted 가 안 나 화면이 갱신되지 않는다(형제 프로젝트 실측)."""
    combos = controls_by_id(top, "COMBO")
    aid = cfg.get("status_combo") or next(
        (a for a in cfg["status_combo_candidates"] if a in combos), None)
    if not aid or aid not in combos:
        log("   ※ 오더상태 콤보를 찾지 못했습니다 — 화면 기본값으로 조회합니다")
        return ""
    h = combos[aid]
    try:
        w = UIAWrapper(elem_of(h))
        items = [i.window_text() for i in w.descendants(control_type="ListItem")] or []
        idx = next((i for i, t in enumerate(items) if (t or "").strip() == value), None)
        w.click_input()
        time.sleep(0.5)
        if idx is not None:
            send_keys("{HOME}" + "{DOWN}" * idx)
        else:
            send_keys(value, with_spaces=True, pause=0.03)
        time.sleep(0.3)
        send_keys("{ENTER}")
        time.sleep(0.8)
        shown = _edit_text(h)
        log(f"   오더상태 → {shown!r}")
        return shown
    except Exception as e:
        log(f"   ※ 오더상태 설정 실패({e}) — 화면 기본값으로 진행")
        return ""
