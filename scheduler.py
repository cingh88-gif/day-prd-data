r"""
Windows 작업 스케줄러 등록/해제 — GUI 의 '매일 자동 실행' 버튼이 쓴다.

★ /IT 를 반드시 붙인다
  이 도구는 **GUI 자동화**다. 사용자가 로그온해 있고 대화형 데스크톱이 있어야 iERP 창을
  잡을 수 있다. /IT 없이 등록하면 세션 0(비대화형)에서 돌아 **창을 못 찾고 매일 조용히
  실패**한다. /RL LIMITED 도 함께 — iERP 가 일반 권한으로 뜨므로 무결성 레벨을 맞춘다.

★ WSL 에서도 동작한다
  schtasks.exe 를 상호 운용(interop)으로 부른다. 다만 **등록되는 작업은 Windows 쪽 사본**
  (C:\ierp_day_prd)을 가리킨다 — WSL 경로(\\wsl.localhost\...)는 cmd 의 작업 폴더가 될 수
  없어서, 형제 프로젝트들과 같은 규약(C 드라이브 복사본)을 따른다.
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

# import 만으로 콘솔 cp949 인코딩 방어가 걸린다(config.py 상단 주석 참조).
import config  # noqa: F401

IS_WINDOWS = sys.platform.startswith("win")

# Windows 쪽 배포 폴더 — 배포.sh 가 여기로 복사한다(형제 프로젝트와 같은 규약)
WIN_DIR = r"C:\ierp_day_prd"
BAT_NAME = "일일수집_자동.bat"
DEFAULT_TASK = "iERP_일마감_공정진척"

SCHTASKS = ("schtasks.exe" if IS_WINDOWS else "/mnt/c/Windows/System32/schtasks.exe")


def available() -> bool:
    return IS_WINDOWS or Path(SCHTASKS).exists()


def _run(args: list[str]) -> tuple[int, str]:
    """schtasks 실행. 출력은 cp949 라 디코딩해서 돌려준다."""
    try:
        p = subprocess.run([SCHTASKS] + args, capture_output=True, timeout=30)
    except FileNotFoundError:
        return 127, "schtasks.exe 를 찾지 못했습니다 (Windows 에서 실행하세요)"
    except subprocess.TimeoutExpired:
        return 124, "schtasks 응답 없음(30초 초과)"
    out = (p.stdout + p.stderr).decode("cp949", errors="replace").strip()
    return p.returncode, out


def valid_time(hhmm: str) -> bool:
    m = re.match(r"^([01]\d|2[0-3]):([0-5]\d)$", (hhmm or "").strip())
    return bool(m)


def describe_time(hhmm: str) -> str:
    """'11:00' → '오전 11:00' — 11시/12시가 오전인지 오후인지 헷갈리지 않게."""
    if not valid_time(hhmm):
        return hhmm
    h, m = (int(x) for x in hhmm.split(":"))
    ampm = "오전" if h < 12 else "오후"
    h12 = h if 1 <= h <= 12 else (h - 12 if h > 12 else 12)
    return f"{ampm} {h12}:{m:02d} ({hhmm})"


def target_command(win_dir: str = WIN_DIR) -> str:
    return f'"{win_dir}\\{BAT_NAME}"'


def create(hhmm: str, task_name: str = DEFAULT_TASK, win_dir: str = WIN_DIR
           ) -> tuple[bool, str]:
    """매일 hhmm 에 도는 작업을 등록(있으면 덮어씀). 반환 (성공, 메시지)."""
    if not valid_time(hhmm):
        return False, f"시간 형식이 잘못됐습니다: {hhmm!r} — 24시간제 HH:MM (예: 11:00, 23:30)"
    if not available():
        return False, "schtasks.exe 를 쓸 수 없습니다 — Windows 에서 실행하세요"
    rc, out = _run([
        "/Create", "/SC", "DAILY", "/ST", hhmm,
        "/TN", task_name,
        "/TR", target_command(win_dir),
        "/IT",              # ★ 로그온 상태에서 대화형으로만 실행 — GUI 자동화라 필수
        "/RL", "LIMITED",   # iERP 가 일반 권한이므로 맞춘다
        "/F",
    ])
    if rc == 0:
        return True, (f"등록됨 — 매일 {describe_time(hhmm)} 에 실행합니다.\n"
                      f"작업 이름: {task_name}\n실행 대상: {target_command(win_dir)}")
    return False, f"등록 실패(코드 {rc})\n{out}"


def delete(task_name: str = DEFAULT_TASK) -> tuple[bool, str]:
    if not available():
        return False, "schtasks.exe 를 쓸 수 없습니다 — Windows 에서 실행하세요"
    rc, out = _run(["/Delete", "/TN", task_name, "/F"])
    if rc == 0:
        return True, f"해제됨 — '{task_name}' 자동 실행을 껐습니다."
    return False, f"해제 실패(코드 {rc})\n{out}"


def query(task_name: str = DEFAULT_TASK) -> tuple[bool, str]:
    """등록돼 있는지 + 다음 실행 시각. 반환 (등록됨, 설명)."""
    if not available():
        return False, "schtasks.exe 를 쓸 수 없습니다 (WSL 에서는 조회만 안 됩니다)"
    rc, out = _run(["/Query", "/TN", task_name, "/FO", "LIST"])
    if rc != 0:
        return False, "등록돼 있지 않습니다"
    info = {}
    for line in out.splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            info[k.strip()] = v.strip()
    nxt = info.get("다음 실행 시간") or info.get("Next Run Time") or ""
    state = info.get("상태") or info.get("Status") or ""
    return True, f"등록됨 · 상태 {state} · 다음 실행 {nxt}".strip(" ·")


def run_now(task_name: str = DEFAULT_TASK) -> tuple[bool, str]:
    """등록된 작업을 지금 한 번 돌린다(스케줄 확인용)."""
    rc, out = _run(["/Run", "/TN", task_name])
    return rc == 0, out or ("실행 요청됨" if rc == 0 else f"실패(코드 {rc})")


if __name__ == "__main__":
    cfg = config.load()
    cmd = sys.argv[1] if len(sys.argv) > 1 else "query"
    name = cfg["schedule"].get("task_name", DEFAULT_TASK)
    if cmd == "create":
        t = sys.argv[2] if len(sys.argv) > 2 else cfg["schedule"]["time"]
        print(create(t, name)[1])
    elif cmd == "delete":
        print(delete(name)[1])
    elif cmd == "run":
        print(run_now(name)[1])
    else:
        print(query(name)[1])
