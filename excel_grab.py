r"""
iERP '엑셀출력' 이 연 워크북을 COM 으로 잡아 저장하고, 남는 빈 껍데기를 정리한다.

⚠️ Windows 전용. WSL 에서 import 만 하는 건 안전하다(호출하면 False 를 돌려준다).

★ 이 파일은 ierp-prod-report / outsourcing-cost 에서 **실측으로 확정된 로직을 그대로 이식**한
  것이다. 아래 함정들은 전부 실제로 당해서 알아낸 것이라 임의로 줄이면 안 된다.
  · 조회가 덜 끝난 채 출력 → 행 수는 맞는데 전부 빈 칸 → UsedRange 크기만 보지 말고 CountA 까지
  · iERP 가 셀을 채우기 전에 Close → HRESULT 0x800A01A8 → settle 대기 + 3회 연속 동일 확인
  · ROT 를 전부 Dispatch → iERP WCF 호스트에 COM 이 걸려 무한 정지 → 보이는 창 바인딩 우선
  · 사용자가 열어 둔 Excel 을 죽이면 안 된다 → guard_excel() 로 보호 목록 등록
"""
from __future__ import annotations

import time
from pathlib import Path

SETTLE = 1.5              # 엑셀출력 후 워크북이 뜨기까지 초기 대기
SAVE_TIMEOUT = 300

# ★★ Excel 이 바쁘면 COM 호출을 거부한다 — 0x80010001 RPC_E_CALL_REJECTED
#   ('피호출자가 호출을 거부했습니다'). 2026-09-01 실측: 사용자가 열어 둔 Excel 창 하나가
#   이 상태였는데, 예외를 조용히 삼키고 '실행 중인 Excel 없음' 으로 처리해 **5분 18초 동안
#   말없이 기다리다 저장 실패**했다. 원인이 로그에 한 줄도 안 남는 게 가장 나빴다.
#   → 거부는 대개 일시적이다. 짧게 재시도하고, 끝까지 거부하면 **이유를 로그에 남긴다.**
RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
_BUSY_HRESULTS = (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER)
_last_busy: list[str] = []          # 마지막으로 거부한 Excel 창 설명(진단용)


def _is_busy_error(e) -> bool:
    code = getattr(e, "hresult", None)
    if code is None:
        args = getattr(e, "args", ())
        code = args[0] if args and isinstance(args[0], int) else None
    return code in _BUSY_HRESULTS


def _com_retry(fn, tries: int = 3, wait: float = 0.4):
    """Excel 이 바쁠 때(호출 거부)만 짧게 재시도한다. 다른 예외는 그대로 올린다."""
    last = None
    for i in range(tries):
        try:
            return fn()
        except Exception as e:
            last = e
            if not _is_busy_error(e):
                raise
            time.sleep(wait * (i + 1))
    raise last

# 시작 시점에 이미 떠 있던 EXCEL.EXE PID(사용자 소유) — 절대 종료하지 않는다.
PROTECTED_EXCEL_PIDS: set[int] = set()


def _excel_apps_by_window():
    """보이는 Excel 창(XLMAIN→XLDESK→EXCEL7)에 COM 을 직접 바인딩해 Application 을 얻는다.

    ★ 이게 가장 확실하다. iERP 엑셀출력은 숨은 자동화 인스턴스와 별개로, 실제 데이터는
      '통합 문서1 - Excel' 같은 **보이는 창**에 넣는다. ROT/GetActiveObject 는 숨은 빈
      인스턴스(Workbooks=0)만 가리킨다."""
    try:
        import ctypes
        import comtypes
        import comtypes.client
        import win32gui
        from comtypes.automation import IDispatch
    except ImportError:
        return []
    OBJID_NATIVEOM = 0xFFFFFFF0
    apps, seen, tops = [], set(), []

    def _cb(h, _):
        try:
            if win32gui.GetClassName(h) == "XLMAIN":
                tops.append(h)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        return []

    _last_busy.clear()
    for top in tops:
        try:
            desk = win32gui.FindWindowEx(top, 0, "XLDESK", None)
            if not desk:
                continue
            x7 = win32gui.FindWindowEx(desk, 0, "EXCEL7", None)   # 워크북 있는 창만 존재
            if not x7:
                continue

            def _bind(_x7=x7, _top=top):
                ptr = ctypes.POINTER(IDispatch)()
                ctypes.oledll.oleacc.AccessibleObjectFromWindow(
                    _x7, OBJID_NATIVEOM, ctypes.byref(IDispatch._iid_), ctypes.byref(ptr))
                window = comtypes.client.GetBestInterface(ptr)
                return window.Application          # ← 여기서 거부가 난다

            app = _com_retry(_bind)
            key = int(_com_retry(lambda: getattr(app, "Hwnd", top)))
            if key not in seen:
                seen.add(key)
                apps.append(app)
        except Exception as e:
            if _is_busy_error(e):
                try:
                    title = win32gui.GetWindowText(top) or "(제목없음)"
                except Exception:
                    title = "?"
                _last_busy.append(title)
            continue
    return apps


def busy_hint() -> str:
    """왜 Excel 을 못 잡았는지 사람이 읽을 설명(없으면 빈 문자열)."""
    if not _last_busy:
        return ""
    return ("열려 있는 Excel 이 COM 호출을 거부하고 있습니다(바쁨): "
            + ", ".join(repr(t) for t in _last_busy)
            + "\n      → 그 Excel 에서 셀 편집 중이거나 대화상자가 떠 있는지 확인하고 닫으세요.")


def _excel_apps(win32, pythoncom):
    """실행 중인 Excel.Application 인스턴스.

    ⚠️ ROT 의 '모든' 항목을 Dispatch 하면 iERP WCF 호스트 같은 바쁜 STA COM 서버에 걸려
      무한정 멈춘다(실측). 보이는 창 바인딩을 먼저 쓰고, ROT 는 Excel CLSID 만 건드린다."""
    apps, seen = [], set()

    def _add(app):
        try:
            h = int(app.Hwnd)
        except Exception:
            return
        if h not in seen:
            seen.add(h)
            apps.append(app)

    for app in _excel_apps_by_window():
        _add(app)
    if apps:
        return apps
    try:
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        for mon in rot:
            try:
                nm = (mon.GetDisplayName(ctx, None) or "").lower()
            except Exception:
                continue
            if nm.startswith("!{") and "0002450" not in nm:
                continue                     # 남의 클래스 모니커 → 건드리지 않음(정지 방지)
            try:
                disp = win32.Dispatch(rot.GetObject(mon))
                _add(getattr(disp, "Application", disp))
            except Exception:
                continue
    except Exception:
        pass
    if not apps:
        try:
            _add(win32.GetActiveObject("Excel.Application"))
        except Exception:
            pass
    return apps


def _find_export_wb(xl):
    """저장 안 됐고(미저장) 데이터가 있는 워크북 1개(없으면 None)."""
    for i in range(xl.Workbooks.Count, 0, -1):
        try:
            wb = xl.Workbooks(i)
            if wb.Path:
                continue
            used = wb.Worksheets(1).UsedRange
            if used.Rows.Count < 2 and used.Columns.Count < 2:
                continue
            return wb
        except Exception:
            continue
    return None


def save_active_excel(out_path: Path, timeout=SAVE_TIMEOUT, settle=SETTLE,
                      stable_needed=3, poll=1.0, log=print) -> bool:
    """'엑셀출력' 이 연 워크북을 찾아 out_path 로 저장하고 닫는다."""
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError:
        log("※ pywin32 미설치 — 'py -m pip install -r requirements.txt' 후 자동저장됩니다.")
        return False
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    time.sleep(settle)
    end = time.time() + timeout
    prev_sig, stable, last_err = None, 0, None
    while time.time() < end:
        for xl in _excel_apps(win32, pythoncom):
            try:
                wb = _find_export_wb(xl)
                if wb is None:
                    continue
                used = wb.Worksheets(1).UsedRange
                # ⚠️ 크기만 보면 안 된다 — iERP 는 시트 크기를 먼저 잡고 값을 나중에 채운다.
                try:
                    filled = int(xl.WorksheetFunction.CountA(used))
                except Exception:
                    filled = -1
                sig = (wb.Name, used.Rows.Count, used.Columns.Count, filled)
                if sig == prev_sig:
                    stable += 1
                else:
                    prev_sig, stable = sig, 0
                if stable < stable_needed:
                    break
                if filled >= 0 and filled < used.Rows.Count:
                    prev_sig, stable = None, 0      # 아직 채우는 중
                    break
                xl.DisplayAlerts = False
                # ⚠️ comtypes 동적 디스패치는 키워드 인자를 못 받는다 → 위치 인자.
                wb.SaveAs(str(out_path), 51)        # 51 = .xlsx
                wb.Close(False)
                try:
                    if xl.Workbooks.Count == 0:
                        xl.Quit()
                except Exception:
                    pass
                return True
            except Exception as e:
                last_err = e
        time.sleep(poll)
    if last_err is not None:
        log(f"      (저장 예외: {last_err!r})")
    hint = busy_hint()
    if hint:
        log(f"      ※ {hint}")
    elif last_err is None:
        log("      ※ 엑셀출력이 연 워크북을 찾지 못했습니다 "
            "(엑셀출력 버튼이 안 눌렸거나 Excel 이 뜨지 않았습니다)")
    return False


def _excel_pids() -> set[int]:
    try:
        import win32com.client as win32
        wmi = win32.GetObject("winmgmts:")
        return {int(p.ProcessId) for p in wmi.InstancesOf("Win32_Process")
                if (p.Name or "").upper() == "EXCEL.EXE"}
    except Exception:
        return set()


def _visible_window_pids() -> set[int]:
    pids = set()
    try:
        import win32gui
        import win32process

        def _cb(hwnd, _):
            try:
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    pids.add(int(pid))
            except Exception:
                pass
            return True
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return pids


def guard_excel(log=print):
    """수집 시작 전 1회: 지금 떠 있는 Excel(사용자 것)을 보호 목록에 등록."""
    global PROTECTED_EXCEL_PIDS
    PROTECTED_EXCEL_PIDS = _excel_pids()
    if PROTECTED_EXCEL_PIDS:
        log(f"기존 Excel {len(PROTECTED_EXCEL_PIDS)}개는 보호(종료 안 함): "
            f"{sorted(PROTECTED_EXCEL_PIDS)}")


def _kill_orphan_excel() -> int:
    """iERP 가 남긴 '보이지 않는 좀비 EXCEL.EXE' 종료(백스톱).
    보호목록에 없고 + 보이는 창도 없는 것만 종료한다."""
    import os
    import signal
    cur = _excel_pids()
    if not cur:
        return 0
    visible = _visible_window_pids()
    killed = 0
    for pid in cur - PROTECTED_EXCEL_PIDS:
        if pid in visible:
            continue
        try:
            os.kill(pid, signal.SIGTERM)
            killed += 1
        except Exception:
            pass
    return killed


def cleanup_excel() -> tuple[int, int]:
    """엑셀출력이 매번 남기는, 화면에 안 보이는 빈 Excel 껍데기 정리.
    반환: (닫은 워크북 수, 종료한 인스턴스 수)."""
    try:
        import win32com.client as win32
        import pythoncom
    except Exception:
        return (0, 0)
    closed = quit_ = 0
    try:
        apps = _excel_apps(win32, pythoncom)
    except Exception:
        apps = []
    for xl in apps:
        try:
            for i in range(xl.Workbooks.Count, 0, -1):      # 뒤에서부터(인덱스 밀림 방지)
                try:
                    wb = xl.Workbooks(i)
                    if wb.Path:                             # 이미 저장된 파일 → 보존
                        continue
                    used = wb.Worksheets(1).UsedRange
                    if used.Rows.Count <= 1 and used.Columns.Count <= 1:
                        xl.DisplayAlerts = False
                        wb.Close(False)
                        closed += 1
                except Exception:
                    continue
            try:
                if xl.Workbooks.Count == 0:
                    xl.Quit()
                    quit_ += 1
            except Exception:
                pass
        except Exception:
            continue
    quit_ += _kill_orphan_excel()
    return (closed, quit_)


def remove_file(path: Path, tries: int = 12, wait: float = 1.0, log=print) -> bool:
    """저장 직후의 파일을 지운다 — **재시도한다.**

    ★ Close 직후에도 Excel 이 파일을 잠깐 더 쥐고 있어 unlink 가 PermissionError 를 낸다.
      빈 파일이 남으면 재개할 때 '이미 있음' 으로 건너뛰어 영영 빈 채로 남는다."""
    path = Path(path)
    for _ in range(tries):
        try:
            path.unlink()
            return True
        except FileNotFoundError:
            return True
        except OSError:
            time.sleep(wait)
    log(f"      ※ 빈 파일을 지우지 못했습니다: {path.name} — 재실행 전에 직접 지우세요")
    return False


def close_workbook_at(path, log=print) -> bool:
    """그 경로의 워크북이 Excel 에 열려 있으면 저장 없이 닫는다(파일 잠금 해제).

    ★ 보고서를 한 번 열어 보면 다음 실행 때 같은 이름으로 저장을 못 해 PermissionError 가
      난다. 이름을 바꿔 저장하는 폴백도 있지만, 파일이 두 개로 갈라져 헷갈린다.
      → 저장 전에 **그 파일만** 닫는다. 다른 워크북과 Excel 자체는 건드리지 않는다."""
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError:
        return False
    target = str(path).lower()
    closed = False
    try:
        apps = _excel_apps(win32, pythoncom)
    except Exception:
        return False
    for xl in apps:
        try:
            for i in range(xl.Workbooks.Count, 0, -1):
                wb = xl.Workbooks(i)
                try:
                    full = (wb.FullName or "").lower()
                except Exception:
                    continue
                if full == target:
                    xl.DisplayAlerts = False
                    wb.Close(False)
                    closed = True
                    log(f"      · 열려 있던 보고서를 닫았습니다: {Path(path).name}")
        except Exception:
            continue
    return closed
