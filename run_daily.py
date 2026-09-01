r"""
일마감 원샷 실행 — **수집 → 집계 → 보고서** 한 번에. 스케줄러가 부르는 진입점.

  py run_daily.py                    # 전날치 수집 + 보고서
  py run_daily.py --date 2026-08-31  # 특정일
  py run_daily.py --no-collect       # 이미 받아 둔 폴더로 **집계·보고서만** 다시
  py run_daily.py --auto             # 스케줄러용(로그 파일에 기록, 보고서 자동열기 안 함)

--no-collect 는 WSL 에서도 돈다(수집만 Windows 전용).
"""
from __future__ import annotations

import datetime
import re
import sys
import traceback
from pathlib import Path

import aggregate
import collect
import config
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


def run(ymd: str | None = None, cfg: dict | None = None, do_collect: bool = True,
        log=print, should_stop=None, pump=None) -> dict:
    """반환: {'ok':bool, 'report':Path|None, 'agg':dict|None, 'result':dict|None, 'summary':str}"""
    cfg = cfg or config.load()
    ymd = ymd or collect.target_date(cfg.get("offset_days", 1))
    pretty = f"{ymd[:4]}-{ymd[4:6]}-{ymd[6:]}"

    result = None
    if do_collect:
        result = collect.run_collect(ymd, cfg, log=log, should_stop=should_stop, pump=pump)
        run_dir = result["run_dir"]
    else:
        run_dir = collect.run_dir_for(cfg["export_dir"], ymd)
        log(f"■ 수집 생략 — 기존 폴더로 집계합니다: {run_dir}")
    if not run_dir.exists():
        raise RuntimeError(f"수집 폴더가 없습니다: {run_dir}\n"
                           f"  먼저 수집을 돌리거나 --date 를 확인하세요.")

    # ── 집계 + 보고서 ─────────────────────────────────────────────
    log("\n■ 집계 · 보고서 생성")
    agg = aggregate.aggregate_folder(run_dir, cfg["teams"])
    got = {t["team"] for t in agg["teams"]}
    meta = {
        "failed": (result or {}).get("failed", []),
        # 파일이 아예 없는 작업반 — 0건이라 지운 것과 구분해 점검 시트에 남긴다
        "missing": [t["code"] for t in config.active_teams(cfg)
                    if t["code"] not in got
                    and t["code"] not in (result or {}).get("empty", [])
                    and t["code"] not in (result or {}).get("failed", [])],
    }
    out = run_dir / f"일마감_공정진척_{ymd}.xlsx"
    if out.exists():
        # 지난 실행에서 열어 본 보고서가 Excel 에 물려 있으면 같은 이름으로 저장이 안 된다.
        try:
            import excel_grab
            excel_grab.close_workbook_at(out, log=log)
        except Exception:
            pass
    out = report.build(agg, pretty, out, meta)

    summary = aggregate.summary_text(agg, pretty)
    log("")
    log(summary)
    log(f"\n보고서: {out}")
    return {"ok": True, "report": out, "agg": agg, "result": result,
            "summary": summary, "run_dir": run_dir, "ymd": ymd}


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    auto = "--auto" in argv
    do_collect = "--no-collect" not in argv
    ymd = None
    if "--date" in argv:
        ymd = re.sub(r"\D", "", argv[argv.index("--date") + 1])
        if len(ymd) != 8:
            print("--date 는 YYYY-MM-DD 또는 YYYYMMDD 형식이어야 합니다")
            return 2

    cfg = config.load()
    if "--teams" in argv:            # 시험용 — 일부 작업반만 (예: --teams M2105,M1101)
        want = {c.strip().upper() for c in argv[argv.index("--teams") + 1].split(",") if c.strip()}
        for t in cfg["teams"]:
            t["use"] = t["code"].upper() in want
        print(f"[일부 실행] 작업반 {sorted(want)} 만 처리합니다")
    ymd = ymd or collect.target_date(cfg.get("offset_days", 1))
    log_path = (config.to_local_path(cfg["export_dir"]) / "logs"
                / f"run_{datetime.datetime.now():%Y%m%d_%H%M%S}.log")
    log, fh = make_logger(log_path, echo=True)
    log(f"=== 일마감 공정진척 자동 실행 시작 ({'스케줄' if auto else '수동'}) ===")
    try:
        r = run(ymd, cfg, do_collect=do_collect, log=log)
        fails = (r["result"] or {}).get("failed", [])
        log(f"=== 완료 === 로그: {log_path}")
        if not auto and cfg.get("open_report"):
            try:
                import os
                os.startfile(str(r["report"]))         # Windows 에서만 존재
            except Exception:
                pass
        return 3 if fails else 0                       # 3 = 일부 작업반 실패
    except Exception as e:
        log(f"※ 실패: {e}")
        log(traceback.format_exc())
        return 1
    finally:
        if fh:
            fh.close()


if __name__ == "__main__":
    raise SystemExit(main())
