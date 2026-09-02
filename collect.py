r"""
공정진척현황(PM60250Rv3) 작업반별 **하루치** 수집.  ⚠️ 수집은 Windows 전용.

흐름(작업반 1개):
  작업반 코드 입력 → TAB(코드 확정·작업반명 확인) → 대상일자 입력 → 조회
  → **앱이 다시 응답할 때까지 대기** → 엑셀출력 → 자동저장 → 저장파일 검증 → Excel 껍데기 정리

중단해도 같은 폴더로 다시 실행하면 **이어받는다**(이미 받은 작업반은 건너뜀).
"""
from __future__ import annotations

import csv
import datetime
import time
from pathlib import Path

import config
import ierp_progress as ip
import period as P


class Cancelled(RuntimeError):
    """사용자가 GUI 에서 중단을 눌렀을 때."""


def _now() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _hms(sec: float) -> str:
    sec = int(round(sec))
    h, m, s = sec // 3600, (sec % 3600) // 60, sec % 60
    if h:
        return f"{h}시간 {m}분 {s}초"
    if m:
        return f"{m}분 {s}초"
    return f"{s}초"


def target_date(offset_days: int = 1, today: datetime.date | None = None) -> str:
    """대상일자 'YYYYMMDD'. offset_days=1 이면 전날. (기존 호출부 호환용)"""
    d = (today or datetime.date.today()) - datetime.timedelta(days=int(offset_days))
    return d.strftime("%Y%m%d")


def new_run_label(per, now: datetime.datetime | None = None) -> str:
    """이번 실행의 라벨 = 기간 + **실행 시분초**. 예: '20260901_092048' / '202608_143022'.

    ★ 왜 시분초를 붙이나 (사용자 요청 2026-09-02)
      같은 기간을 **여러 번 돌릴 수 있다.** 시분초가 없으면 뒤 실행이 앞 결과를 덮어쓰거나,
      이미 있는 파일을 '이미 받음' 으로 건너뛰어 **새 데이터를 못 받는다.**
      실행마다 폴더가 갈리면 결과를 나란히 두고 비교할 수 있다."""
    now = now or datetime.datetime.now()
    return f"{per.label}_{now:%H%M%S}"


def find_latest_run_dir(export_dir, per) -> Path | None:
    """그 기간의 **가장 최근** 수집 폴더(없으면 None). 이어받기·재집계가 쓴다.

    시분초 없는 옛 폴더(day_20260831)도 같이 본다 — 기존 수집을 못 찾으면 안 된다."""
    base = config.to_local_path(str(export_dir))
    if not base.exists():
        return None
    cands = [d for d in base.glob(f"{per.mode}_{per.label}*") if d.is_dir()]
    if not cands:
        return None
    return sorted(cands, key=lambda d: (d.name, d.stat().st_mtime))[-1]


def run_dir_for(export_dir, per, run_label: str | None = None) -> Path:
    """실행 폴더. 이름이 '<모드>_<기간>_<시분초>' 라 실행마다 갈린다.

    ⚠️ 폴더 이름의 모드 접두어(day_/month_/range_)는 그대로 둔다 —
      '8월 하루치' 와 '8월 한 달치' 가 섞이면 안 된다."""
    if isinstance(per, str):                    # 'YYYYMMDD' 를 주면 하루로 본다(호환)
        per = P.day(per)
    label = run_label or new_run_label(per)
    return config.to_local_path(str(export_dir)) / f"{per.mode}_{label}"


def _log_row(run_dir: Path, team, status, detail=""):
    p = run_dir / "수집로그.csv"
    new = not p.exists()
    with open(p, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["시각", "작업반", "결과", "내용"])
        w.writerow([_now(), team, status, detail])


def verify_saved(xlsx_path: Path, team: str) -> tuple[bool, str, int]:
    r"""저장된 파일에 **실제 값이 들어 있는지** + **요청한 작업반이 맞는지** 확인.

    ★ 조회가 덜 끝난 채 저장되면 '행 수는 맞는데 전부 빈 칸' 인 파일이 나온다. 그런 파일을
      남겨 두면 재개 시 '이미 있음' 으로 건너뛰어 영영 빈 채로 남는다.
    ★ 작업반 대조가 핵심이다 — 작업반 입력이 안 먹은 채 조회되면 **다른 작업반 데이터가
      그 작업반 파일로** 저장된다. 총계 대조로는 절대 못 잡는 종류의 사고다.

    반환: (통과, 설명, 데이터행수)   행수 0 = 그 날 데이터 없음(정상)
    """
    import aggregate
    try:
        header, rows = aggregate.read_rows(xlsx_path)
    except Exception as e:
        return False, f"파일을 읽지 못했습니다({e})", -1
    if not header:
        return False, "시트가 비어 있습니다", -1
    if not rows:
        return True, "데이터 0건", 0

    i_status = aggregate.find_col(header, "status")
    if i_status < 0:
        return False, (f"'오더상태' 열이 없습니다 — 실제 헤더: {[h for h in header if h][:12]}"), len(rows)

    nonblank = sum(1 for r in rows
                   if any(v is not None and str(v).strip() for v in r))
    if nonblank == 0:
        return False, f"{len(rows):,}행이 전부 빈 칸 — 조회가 덜 끝난 파일", len(rows)
    if nonblank / len(rows) < 0.5:
        return False, f"{nonblank:,}/{len(rows):,}행에만 값 — 빈 파일", len(rows)

    i_team = aggregate.find_col(header, "team")
    if i_team >= 0:
        vals = {str(r[i_team]).strip().upper() for r in rows
                if i_team < len(r) and r[i_team] is not None and str(r[i_team]).strip()}
        if vals and vals != {team.strip().upper()}:
            return False, (f"작업반 불일치 — 요청 {team}, 파일 {sorted(vals)[:5]}\n"
                           f"      (파일 헤더: {[h for h in header if h][:14]})"), len(rows)
    else:
        # '작업반' 열이 없는 화면이면 대조를 건너뛴다.
        # ⚠️ 작업장(C40001 …)을 작업반으로 오인해 오탐을 냈던 자리다(2026-09-01).
        return True, (f"데이터 {len(rows):,}행 (작업반 열 없음 — 대조 생략, "
                      f"헤더: {[h for h in header if h][:8]})"), len(rows)
    return True, f"데이터 {len(rows):,}행", len(rows)


def collect_one(top, scfg, dtps, code_h, name_h, team: str, per,
                out_path: Path, log=print, should_stop=None) -> tuple[bool | None, int, str]:
    """작업반 1개 수집. 반환: (True 저장 / None 0건 / False 실패, 행수, 작업반명)"""
    import excel_grab

    if should_stop and should_stop():
        raise Cancelled()
    if not ip.wait_responsive(top, timeout=600, quiet=2.0, should_stop=should_stop):
        log("      ※ 앞 작업이 안 끝났습니다(앱 응답 없음) — 이 작업반은 건너뜁니다")
        return False, 0, ""

    # 1) 작업반
    ok_team, team_nm = ip.set_team(top, scfg, code_h, name_h, team, log=log)
    if not ok_team:
        log(f"      ※ 작업반 {team} 이 화면에 안 들어갔습니다 — 건너뜁니다"
            f" (없는 코드이거나 입력칸을 잘못 잡았을 수 있습니다)")
        return False, 0, team_nm

    # 2) 개시예정일 기간
    if not ip.set_period(dtps, per.d_from, per.d_to, log=log):
        log(f"      ※ 기간 {per.d_from}~{per.d_to} 입력 실패")
        return False, 0, team_nm

    # ⚠️ 조회 직전에 한 번 더 대조 — 엉뚱한 기간을 그 기간 파일로 저장하는 게 최악의 사고다.
    #   월마감은 한 번에 한 달치가 들어오므로 틀리면 피해가 더 크다.
    shown_from = ip.read_date(dtps[0])
    if shown_from and shown_from != per.d_from:
        log(f"      ※ 시작일 불일치(화면 {shown_from} ≠ 요청 {per.d_from}) — 건너뜁니다")
        return False, 0, team_nm
    if len(dtps) >= 2:
        shown_to = ip.read_date(dtps[1])
        if shown_to and shown_to != per.d_to:
            log(f"      ※ 종료일 불일치(화면 {shown_to} ≠ 요청 {per.d_to}) — 건너뜁니다")
            return False, 0, team_nm

    # 3) 조회 → 앱이 다시 응답할 때까지 대기 (= 조회 종료)
    t1 = time.monotonic()
    try:
        ip.click_button(top, scfg["query_btn"], scfg)
    except Exception as e:
        log(f"      ※ '조회' 클릭 실패({e})")
        return False, 0, team_nm
    time.sleep(ip.QUERY_WAIT)
    if not ip.wait_responsive(top, timeout=ip.QUERY_TIMEOUT, quiet=5.0, should_stop=should_stop):
        if should_stop and should_stop():
            raise Cancelled()
        log(f"      ※ 조회가 {ip.QUERY_TIMEOUT}초 안에 안 끝났습니다 — 건너뜁니다")
        return False, 0, team_nm
    log(f"      조회 {_hms(time.monotonic() - t1)} → 엑셀출력 → 저장 대기…")

    # 4) 엑셀출력 → 저장 → 검증
    try:
        ip.click_button(top, scfg["export_btn"], scfg)
    except Exception as e:
        log(f"      ※ '엑셀출력' 클릭 실패({e})")
        return False, 0, team_nm
    if not excel_grab.save_active_excel(out_path, log=log):
        log(f"      ※ 자동저장 실패 — 열린 엑셀을 '{out_path}' 로 직접 저장하세요")
        return False, 0, team_nm

    ok, detail, rows = verify_saved(out_path, team)
    if rows == 0:
        # ★ 지우기 전에 Excel 을 먼저 정리한다 — 방금 SaveAs 한 파일을 Excel 이 쥐고 있다.
        excel_grab.cleanup_excel()
        excel_grab.remove_file(out_path, log=log)
        return None, 0, team_nm
    if not ok:
        excel_grab.cleanup_excel()
        excel_grab.remove_file(out_path, log=log)   # 나쁜 파일을 남기면 재개 때 건너뛴다
        log(f"      ※ 저장파일 검증 실패({detail}) — 삭제하고 실패 처리")
        return False, rows, team_nm

    log(f"      저장됨: {out_path.name}  {detail}")
    return True, rows, team_nm


def run_collect(per=None, cfg: dict | None = None, log=print,
                should_stop=None, pump=None, resume: bool = False) -> dict:
    """기간 하나를 작업반별로 수집한다. per = period.Period (없으면 설정대로 전날).

    resume=True 면 **그 기간의 가장 최근 폴더를 이어서** 채운다(이미 받은 작업반은 건너뜀).
    기본(False)은 실행 시분초가 붙은 **새 폴더**를 만든다 — 여러 번 돌려도 안 겹친다.

    반환: {'run_dir':Path, 'period':Period, 'run_label':str, 'saved':[...],
           'empty':[...], 'failed':[...], 'names':{}}
    """
    import excel_grab

    cfg = cfg or config.load()
    if per is None:
        per = P.previous_day(cfg.get("offset_days", 1))
    elif isinstance(per, str):
        per = P.parse(per)
    teams = config.active_teams(cfg)

    run_dir = None
    if resume:
        run_dir = find_latest_run_dir(cfg["export_dir"], per)
        if run_dir:
            log(f"   이어받기 — 기존 폴더 사용: {run_dir.name}")
        else:
            log("   이어받기를 요청했지만 기존 폴더가 없습니다 — 새로 만듭니다")
    if run_dir is None:
        run_dir = run_dir_for(cfg["export_dir"], per)
    # 폴더 이름에서 이번 실행 라벨을 되찾는다(이어받기면 그 폴더의 라벨을 그대로 쓴다)
    run_label = run_dir.name.split("_", 1)[1]
    run_dir.mkdir(parents=True, exist_ok=True)

    scfg = ip.load_screen()
    kind = {"day": "일마감", "month": "월마감", "range": "기간"}[per.mode]
    log(f"■ {kind} {per.title} / 작업반 {len(teams)}개 → {run_dir}")

    # ★ 권한이 안 맞으면 아무것도 안 된다 — 창을 열기 전에 먼저 잡는다.
    ip.preflight(log=log)
    # ★ 좀비 정리를 guard_excel 보다 **먼저** — 보호 목록이 비었을 때 걷어내야
    #   사용자 Excel 만 보호 목록에 남는다.
    excel_grab.sweep_zombie_excel(log=log)
    excel_grab.guard_excel(log=log)
    top = ip.ensure_screen(scfg, log=log)
    ip.focus_window(top)
    time.sleep(0.5)

    dtps = ip.find_children_by_class(top, scfg["dtp_class"])
    if not dtps:
        raise RuntimeError("개시예정일(DateTimePicker)을 찾지 못했습니다 — "
                           "py probe_progress.py 로 화면 구성을 확인하세요.")
    if len(dtps) < 2 and not per.is_single_day:
        # ★ 기간을 넣을 칸이 없는데 월마감을 돌리면, 하루치를 한 달치로 착각해 저장한다.
        #   조용히 틀린 숫자를 쌓는 것보다 시작하지 않는 게 낫다.
        raise RuntimeError(
            f"이 화면의 날짜 칸이 1개뿐이라 기간({per.title}) 조회를 할 수 없습니다.\n"
            f"  일 단위로만 돌리거나, probe_progress.py 로 화면 구성을 확인하세요.")
    code_h, name_h, _edits = ip.resolve_team_fields(top, scfg, log=log)
    if code_h is None:
        raise RuntimeError(
            "작업반 입력칸을 찾지 못했습니다 — py probe_progress.py 로 정찰한 뒤 "
            "screen.json 의 team_field 를 채우세요.\n"
            "  (자동 추정은 automation_id 에 'DPT' 가 들어간 Edit 만 찾습니다)")
    ip.save_screen({k: scfg[k] for k in ("team_field", "team_name_field") if scfg.get(k)})

    # ★★ 조회조건은 **여기서 한 번만** 맞춘다(그리드가 비어 있을 때). 아래 순회에서는 안 건드린다.
    #   chkSELOPT1('작업완료 포함')이 꺼진 채로 조회하면 80작업완료가 통째로 빠져 진행률이
    #   매일 0% 로 나온다. 조건을 못 맞추면 **수집을 시작하지 않는다** — 조용히 틀린 숫자를
    #   쌓는 것이 아무것도 안 하는 것보다 나쁘다.
    if not ip.set_checkboxes(top, scfg, log=log):
        raise RuntimeError(
            "조회조건(체크박스)을 맞추지 못했습니다 — 이대로 받으면 80작업완료가 빠져\n"
            "  진행률이 0% 로 나옵니다. probe_progress.py 로 화면을 확인하세요.")

    if scfg.get("set_status"):
        ip.set_status(top, scfg, scfg.get("status_value", "ALL"), log=log)

    saved, empty, failed, names = [], [], [], {}
    start = time.monotonic()
    consec_fail = 0

    for i, t in enumerate(teams, 1):
        if should_stop and should_stop():
            raise Cancelled()
        if pump:
            pump()
        code = t["code"]
        out = run_dir / f"progress_{code}_{run_label}.xlsx"
        if out.exists():
            log(f"   ({i}/{len(teams)}) {code} · 이미 있음 — 건너뜀")
            saved.append(code)
            continue
        log(f"   ({i}/{len(teams)}) 작업반 {code} {t.get('name','')}")

        m0 = time.monotonic()
        try:
            import win32gui
            if not win32gui.IsWindow(top):        # 앞 작업반에서 화면이 죽었나(win32, 즉시)
                log("      ※ iERP 화면이 닫혔습니다 — 다시 엽니다")
                top = ip.ensure_screen(scfg, log=log)
                ip.focus_window(top)
                time.sleep(1.0)
                dtps = ip.find_children_by_class(top, scfg["dtp_class"])
                code_h, name_h, _ = ip.resolve_team_fields(top, scfg, log=log)
                ip.set_checkboxes(top, scfg, log=log)   # 화면을 새로 열면 기본값으로 돌아온다

            res, rows, nm = collect_one(top, scfg, dtps, code_h, name_h, code, per,
                                        out, log=log, should_stop=should_stop)
            if nm:
                names[code] = nm
                config.remember_team_name(code, nm)
            if res is True:
                saved.append(code)
                consec_fail = 0
                _log_row(run_dir, code, "ok", f"{rows}행")
            elif res is None:
                empty.append(code)
                consec_fail = 0
                log("      · 0건")
                _log_row(run_dir, code, "empty")
            else:
                failed.append(code)
                consec_fail += 1
                _log_row(run_dir, code, "fail")
        except Cancelled:
            raise
        except Exception as e:
            failed.append(code)
            consec_fail += 1
            _log_row(run_dir, code, "error", str(e))
            log(f"      ※ {code} 처리 실패: {e}")
        finally:
            c, q = excel_grab.cleanup_excel()
            if c or q:
                log(f"      (Excel 정리: 워크북 {c} 닫음, 인스턴스 {q} 종료)")

        # ★ 연속 실패가 쌓이면 화면이 죽은 것이다. 그대로 두면 남은 작업반이 전부
        #   몇 초짜리 실패로 소진된다(형제 프로젝트 실측).
        if consec_fail >= 2:
            log("      ※ 2개 연속 실패 — 화면을 다시 엽니다")
            try:
                top = ip.ensure_screen(scfg, log=log)
                ip.focus_window(top)
                time.sleep(1.0)
                dtps = ip.find_children_by_class(top, scfg["dtp_class"])
                code_h, name_h, _ = ip.resolve_team_fields(top, scfg, log=log)
                ip.set_checkboxes(top, scfg, log=log)   # 재접속 후에도 다시 맞춘다
                consec_fail = 0
            except Exception as e:
                log(f"      ※ 화면 재접속 실패({e}) — 중단합니다")
                break

        log(f"      ⏱ {_hms(time.monotonic() - m0)}")

    excel_grab.cleanup_excel()
    log(f"\n수집 완료: 저장 {len(saved)}, 0건 {len(empty)}, 실패 {len(failed)} "
        f"| 총 {_hms(time.monotonic() - start)}")
    if failed:
        log(f"※ 실패한 작업반: {', '.join(failed)} — 같은 날짜로 다시 실행하면 이어받습니다")

    return {"run_dir": run_dir, "period": per, "run_label": run_label,
            "ymd": per.label, "saved": saved, "empty": empty, "failed": failed,
            "names": names}
