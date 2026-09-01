#!/bin/bash
# WSL 사본 → Windows 실행 폴더(C:\ierp_day_prd) 복사.
#
# 왜 복사하는가: cmd 는 \\wsl.localhost\... 를 작업 폴더로 쓰지 못해서, 작업 스케줄러가
# WSL 경로의 .bat 을 실행할 수 없다. 형제 프로젝트(C:\ierp_manhour, C:\ierp_prod_report)와
# 같은 규약으로 C 드라이브에 사본을 두고 거기서 돌린다.
#
# ★ config.json / screen.json 은 **덮어쓰지 않는다** — 사용자가 GUI 에서 정한 설정과
#   정찰로 얻은 컨트롤 ID 가 들어 있다.
set -eu
SRC="$(cd "$(dirname "$0")" && pwd)"
DST=/mnt/c/ierp_day_prd

mkdir -p "$DST"
cp -f "$SRC"/*.py "$DST"/
cp -f "$SRC"/*.bat "$DST"/
cp -f "$SRC"/requirements.txt "$SRC"/README.md "$DST"/ 2>/dev/null || true

for keep in config.json screen.json; do
    if [ -f "$SRC/$keep" ] && [ ! -f "$DST/$keep" ]; then
        cp -f "$SRC/$keep" "$DST/$keep"
        echo "  · $keep 최초 복사"
    elif [ -f "$DST/$keep" ]; then
        echo "  · $keep 은 Windows 쪽 것을 유지 (설정/정찰 결과 보존)"
    fi
done

echo "배포 완료 → C:\\ierp_day_prd"
echo "  Windows 에서:  C:\\ierp_day_prd\\실행.bat  (GUI)"
