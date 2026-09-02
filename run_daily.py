r"""
일·월 마감 원샷 실행 — **수집 → 집계 → 보고서** 한 번에. 스케줄러가 부르는 진입점.

  py run_daily.py                          # 전날치(일마감)
  py run_daily.py --date 2026-08-31        # 특정 하루
  py run_daily.py --month 2026-08          # ★ 그 달 전체(월마감)
  py run_daily.py --last-month             # ★ 전월 전체 — 월마감 스케줄이 쓰는 형태
  py run_daily.py --from 2026-08-01 --to 2026-08-15   # 직접 기간
  py run_daily.py --teams M2105,M1101      # 일부 작업반만(시험용)
  py run_daily.py --resume                 # 앞서 실패한 실행을 이어받기(새 폴더 대신)
  py run_daily.py --no-collect --month 2026-08        # 받아 둔 폴더로 집계·보고서만 다시
  py run_daily.py --auto                   # 스케줄러용(보고서 자동열기 안 함)
                                           #   실패하면 10분 뒤 이어받기로 재시도(최대 3회)
  py run_daily.py --auto --retry 5 --retry-wait 15   # 재시도 횟수/간격 조정

기간은 화면의 **개시예정일** 에 들어간다(작업일자·보고일자가 아니다).
결과는 실행할 때마다 **시분초가 붙은 새 폴더**에 쌓인다(같은 기간을 여러 번 돌려도 안 겹친다).
중간에 실패한 실행을 이어서 채우려면 `--resume` 을 붙인다.
--no-collect 는 WSL 에서도 돈다(수집만 Windows 전용).
"""
from __future__ import annotations

import datetime
import re
import sys
import time
import traceback
from pathlib import Path

import aggregate
import collect
import config
import period as P
import report

try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:
    pass


def make_logger(log_path: Path | None, echo=True):
    fh = None
    if log_path:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        fh = open(log_path, "a", encoding="utf-8")

    def log(msg=""):
        msg = str(msg)
        if echo:
            print(msg, flush=True)
        if fh:
            fh.write(f"{datetime.datetime.now():%H:%M:%S} {msg}\n")
            fh.flush()
    return log, fh


def run(per=None, cfg: dict | None = None, do_collect: bool = True,
        log=print, should_stop=None, pump=None, resume: bool = False) -> dict:
    """반환: {'ok':bool, 'report':Path|None, 'agg':dict|None, 'result':dict|None, 'summary':str}"""
    cfg = cfg or config.load()
    if per is None:
        per = P.previous_day(cfg.get("offset_days", 1))
    elif isinstance(per, str):
        per = P.parse(per)
    pretty = per.title

    result = None
    if do_collect:
        result = collect.run_collect(per, cfg, log=log, should_stop=should_stop,
                                     pump=pump, resume=resume)
        run_dir = result["run_dir"]
        run_label = result["run_label"]
    else:
        # 재집계는 **그 기간의 가장 최근 폴더**를 쓴다(실행마다 폴더가 갈리므로).
        run_dir = collect.find_latest_run_dir(cfg["export_dir"], per)
        if run_dir is None:
            raise RuntimeError(
                f"{per.title} 의 수집 폴더를 찾지 못했습니다.\n"
                f"  먼저 수집을 돌리거나 기간을 확인하세요"
                f"(찾은 위치: {config.to_local_path(cfg['export_dir'])}).")
        run_label = run_dir.name.split("_", 1)[1]
        log(f"■ 수집 생략 — 기존 폴더로 집계합니다: {run_dir}")
    if not run_dir.exists():
        raise RuntimeError(f"수집 폴더가 없습니다: {run_dir}")

    # ── 집계 + 보고서 ─────────────────────────────────────────────
    log("\n■ 집계 · 보고서 생성")
    agg = aggregate.aggregate_folder(run_dir, cfg["teams"])
    got = {t["team"] for t in agg["teams"]}
    kind = {"day": "일마감", "month": "월마감", "range": "기간"}[per.mode]
    meta = {
        "kind": kind,
        "failed": (result or {}).get("failed", []),
        # 파일이 아예 없는 작업반 — 0건이라 지운 것과 구분해 점검 시트에 남긴다
        "missing": [t["code"] for t in config.active_teams(cfg)
                    if t["code"] not in got
                    and t["code"] not in (result or {}).get("empty", [])
                    and t["code"] not in (result or {}).get("failed", [])],
    }
    out = run_dir / f"{kind}_공정진척_{run_label}.xlsx"
    if out.exists():
        # 지난 실행에서 열어 본 보고서가 Excel 에 물려 있으면 같은 이름으로 저장이 안 된다.
        try:
            import excel_grab
            excel_grab.close_workbook_at(out, log=log)
        except Exception:
            pass
    out = report.build(agg, pretty, out, meta)

    summary = aggregate.summary_text(agg, pretty, kind)
    log("")
    log(summary)
    log(f"\n보고서: {out}")
    return {"ok": True, "report": out, "agg": agg, "result": result,
            "summary": summary, "run_dir": run_dir, "period": per,
            "run_label": run_label, "ymd": per.label}


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    auto = "--auto" in argv
    do_collect = "--no-collect" not in argv
    resume = "--resume" in argv

    def opt(name):
        return argv[argv.index(name) + 1] if name in argv else None

    per = None
    try:
        if "--last-month" in argv:
            per = P.previous_month()
        elif "--this-month" in argv:
            per = P.this_month()
        elif opt("--month"):
            per = P.month(opt("--month"))
        elif opt("--from") or opt("--to"):
            a, b = opt("--from"), opt("--to")
            if not (a and b):
                print("--from 과 --to 는 함께 지정해야 합니다")
                return 2
            per = P.custom(a, b)
        elif opt("--date"):
            per = P.day(opt("--date"))
        elif opt("--period"):
            per = P.parse(opt("--period"))
    except (ValueError, IndexError) as e:
        print(f"기간 지정 오류: {e}")
        return 2

    def optint(name, default):
        try:
            return int(argv[argv.index(name) + 1]) if name in argv else default
        except (ValueError, IndexError):
            return default

    # ★ 스케줄 실행은 사람이 안 볼 때 돈다 — 한 번 삐끗했다고 그날치를 통째로 잃으면 안 된다.
    #   실패하거나 일부 작업반이 빠지면 잠시 뒤 **이어받기로** 다시 시도한다.
    max_try = optint("--retry", 3 if auto else 1)
    retry_wait = optint("--retry-wait", 10) * 60

    cfg = config.load()
    if "--teams" in argv:            # 시험용 — 일부 작업반만 (예: --teams M2105,M1101)
        want = {c.strip().upper() for c in argv[argv.index("--teams") + 1].split(",") if c.strip()}
        for t in cfg["teams"]:
            t["use"] = t["code"].upper() in want
        print(f"[일부 실행] 작업반 {sorted(want)} 만 처리합니다")
    if per is None:
        per = P.previous_day(cfg.get("offset_days", 1))
    log_path = (config.to_local_path(cfg["export_dir"]) / "logs"
                / f"run_{datetime.datetime.now():%Y%m%d_%H%M%S}.log")
    log, fh = make_logger(log_path, echo=True)
    kind = {"day": "일마감", "month": "월마감", "range": "기간"}[per.mode]
    log(f"=== {kind} 공정진척 자동 실행 시작 ({'스케줄' if auto else '수동'}) — {per.title} ===")
    if max_try > 1:
        log(f"    (실패하면 {retry_wait // 60}분 뒤 이어받기로 다시 시도, 최대 {max_try}회)")
    try:
        last_err = None
        for n in range(1, max_try + 1):
            if n > 1:
                log(f"\n=== {n}/{max_try}회차 재시도 (이어받기) ===")
            try:
                # 2회차부터는 **이어받기** — 앞서 받아 둔 작업반을 다시 받지 않는다.
                r = run(per, cfg, do_collect=do_collect, log=log,
                        resume=resume or n > 1)
                fails = (r["result"] or {}).get("failed", [])
                if not fails:
                    log(f"=== 완료 === 로그: {log_path}")
                    if not auto and cfg.get("open_report"):
                        try:
                            import os
                            os.startfile(str(r["report"]))     # Windows 에서만 존재
                        except Exception:
                            pass
                    return 0
                last_err = f"작업반 {len(fails)}개 실패: {', '.join(fails)}"
                log(f"※ {last_err}")
            except Exception as e:
                last_err = str(e)
                log(f"※ 실패: {e}")
                log(traceback.format_exc())

            if n < max_try:
                log(f"   {retry_wait // 60}분 대기 후 다시 시도합니다…")
                time.sleep(retry_wait)

        log(f"=== {max_try}회 시도 후에도 끝내지 못했습니다 === 로그: {log_path}")
        log(f"    마지막 사유: {last_err}")
        return 3 if "작업반" in (last_err or "") else 1
    finally:
        if fh:
            fh.close()


if __name__ == "__main__":
    raise SystemExit(main())
