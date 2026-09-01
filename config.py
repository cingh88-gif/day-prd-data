r"""
설정 로드/저장 — GUI 와 스케줄러가 같은 config.json 을 본다.

★ 이 파일은 **Windows/WSL 공용**이다. win32 를 import 하지 않는다.
  (수집은 Windows 전용이지만 집계·보고서는 WSL 에서도 돌려야 하므로)
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

# ★★ 콘솔 출력이 파이프/리다이렉트/작업스케줄러로 넘어가면 cp949 로 인코딩된다.
#   '—' 같은 글자 하나에서 UnicodeEncodeError 가 나 **수집이 통째로 죽는다**
#   (형제 프로젝트 ierp-prod-report 에서 실측된 사고. 2026-09-01 이 저장소에서도 재현됐다:
#    py.exe 로 scheduler.create 결과를 출력하다 '—' 에서 죽었다).
#   config 는 모든 진입점(gui/run_daily/collect/scheduler)이 가장 먼저 import 하므로
#   여기서 한 번만 막으면 전부 덮인다. 로그 한 줄 때문에 멈추지 않도록 대체 출력한다.
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:
        pass

HERE = Path(__file__).resolve().parent
CONFIG_PATH = HERE / "config.json"

IS_WINDOWS = sys.platform.startswith("win")

# 수집 결과 저장 위치. Windows 는 C 드라이브, WSL 은 /mnt/c 로 같은 곳을 본다.
WIN_EXPORT_DIR = r"C:\ierp_exports\day_prd"
WSL_EXPORT_DIR = "/mnt/c/ierp_exports/day_prd"


def default_export_dir() -> str:
    return WIN_EXPORT_DIR if IS_WINDOWS else WSL_EXPORT_DIR


def to_local_path(p: str) -> Path:
    r"""config 에 적힌 경로를 **지금 돌고 있는 OS 기준**으로 바꾼다.

    같은 config.json 을 Windows(수집)와 WSL(집계·확인)에서 함께 쓰기 때문에 필요하다.
    'C:\ierp_exports\day_prd' ↔ '/mnt/c/ierp_exports/day_prd'
    """
    p = (p or "").strip()
    if not p:
        return Path(default_export_dir())
    if IS_WINDOWS:
        if p.startswith("/mnt/") and len(p) > 6 and p[6] == "/":
            return Path(f"{p[5].upper()}:\\" + p[7:].replace("/", "\\"))
        return Path(p)
    # WSL 쪽
    if len(p) > 2 and p[1] == ":" and p[2] in "\\/":
        return Path(f"/mnt/{p[0].lower()}/" + p[3:].replace("\\", "/"))
    return Path(p)


# ── 대상 작업반 — 사용자가 지정한 12개 ─────────────────────────────────────
#   작업반명은 2026-09-01 첫 수집에서 iERP 가 돌려준 실측값이다(빈 이름은 수집 때 자동으로 채워진다).
DEFAULT_TEAMS = [
    {"code": "M2105", "name": "하이드로겔실", "use": True},
    {"code": "03160", "name": "세리안코스메틱", "use": True},
    {"code": "M2103", "name": "하이드로겔", "use": True},
    {"code": "M1103", "name": "시트(2사업장)", "use": True},
    {"code": "M2101", "name": "자동1", "use": True},
    {"code": "M2109", "name": "자동2(2사업장)", "use": True},
    {"code": "M2102", "name": "자동접지", "use": True},
    {"code": "J2001", "name": "진위_포장반", "use": True},
    {"code": "M2107", "name": "튜브", "use": True},
    {"code": "M1101", "name": "포장1", "use": True},
    {"code": "M1102", "name": "포장2", "use": True},
    {"code": "M1104", "name": "포장4", "use": True},
]

DEFAULTS = {
    "teams": DEFAULT_TEAMS,
    # 조회할 오더상태. 'ALL' 로 받아 온 뒤 **집계에서** 30 을 뺀다.
    #   ⚠️ 화면에서 미리 걸러 받으면 30 이 몇 건이었는지 알 수 없어 검증이 불가능해진다.
    "status_filter": "ALL",
    # 대상일자 = 오늘 - offset_days (1 = 전날)
    "offset_days": 1,
    "export_dir": default_export_dir(),
    "schedule": {
        "enabled": False,
        # 24시간제 HH:MM. 사용자가 11시/12시를 아직 못 정해서 GUI 에서 바꾸게 해 둔다.
        "time": "11:00",
        "task_name": "iERP_일마감_공정진척",
        # 월마감 — 매월 <month_day>일 <month_time> 에 **전월 전체**를 받는다.
        #   일마감과 작업 이름이 달라야 서로 덮어쓰지 않는다.
        "month_enabled": False,
        "month_time": "11:00",
        "month_day": 1,
        "month_task_name": "iERP_월마감_공정진척",
    },
    # 수집 후 보고서를 자동으로 열지 여부(수동 실행일 때만)
    "open_report": True,
}


def load() -> dict:
    cfg = json.loads(json.dumps(DEFAULTS))          # deep copy
    if CONFIG_PATH.exists():
        try:
            saved = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            saved = {}
        for k, v in saved.items():
            if k == "schedule" and isinstance(v, dict):
                cfg["schedule"].update(v)
            else:
                cfg[k] = v
    # 방어: teams 항목 형태 보정
    teams = []
    for t in cfg.get("teams") or []:
        if isinstance(t, str):
            t = {"code": t, "name": "", "use": True}
        teams.append({
            "code": str(t.get("code", "")).strip(),
            "name": str(t.get("name", "")).strip(),
            "use": bool(t.get("use", True)),
        })
    cfg["teams"] = [t for t in teams if t["code"]]
    if not cfg["teams"]:
        cfg["teams"] = json.loads(json.dumps(DEFAULT_TEAMS))
    return cfg


def save(cfg: dict) -> None:
    CONFIG_PATH.write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def active_teams(cfg: dict) -> list[dict]:
    return [t for t in cfg["teams"] if t.get("use", True)]


def team_name_map(cfg: dict) -> dict:
    return {t["code"]: t.get("name", "") for t in cfg["teams"]}


def remember_team_name(code: str, name: str) -> None:
    """수집 중 iERP 가 돌려준 작업반명을 config 에 적어 둔다(다음 실행부터 보고서에 표시)."""
    if not name:
        return
    cfg = load()
    changed = False
    for t in cfg["teams"]:
        if t["code"] == code and not t.get("name"):
            t["name"] = name
            changed = True
    if changed:
        save(cfg)
