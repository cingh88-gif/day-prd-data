r"""
조회 기간 — 일마감(하루) / 월마감(그 달 전체) / 직접지정. **순수 모듈**, WSL 에서도 돈다.

★ 왜 별도 모듈인가
  기간이 하루뿐일 때는 'YYYYMMDD' 문자열 하나면 됐지만, 월마감이 들어오면서 파일명·폴더명·
  보고서 제목·이어받기 키가 전부 기간에 묶인다. 한 군데서 만들지 않으면 서로 어긋나
  **같은 기간을 다른 폴더에 두 번 받는 사고**가 난다.

폴더/파일 이름 규약 (이어받기가 이 이름에 걸린다 — 함부로 바꾸면 기존 수집을 다시 받는다)
    일   day_20260831    / progress_M2105_20260831.xlsx
    월   month_202608    / progress_M2105_202608.xlsx
    직접 range_20260801-20260815 / progress_M2105_20260801-20260815.xlsx
"""
from __future__ import annotations

import calendar
import datetime
import re
from dataclasses import dataclass

DAY, MONTH, RANGE = "day", "month", "range"


def _digits(s) -> str:
    return re.sub(r"\D", "", str(s or ""))


@dataclass(frozen=True)
class Period:
    mode: str
    d_from: str          # YYYYMMDD — 화면 '개시예정일' 시작
    d_to: str            # YYYYMMDD — 화면 '개시예정일' 종료
    label: str           # 파일명에 쓰는 짧은 키
    title: str           # 사람이 읽는 제목 (보고서/로그)

    @property
    def dirname(self) -> str:
        return f"{self.mode}_{self.label}"

    @property
    def is_single_day(self) -> bool:
        return self.d_from == self.d_to

    def __str__(self) -> str:
        return self.title


def day(ymd: str) -> Period:
    """하루치(일마감). ymd = 'YYYYMMDD' 또는 'YYYY-MM-DD'."""
    d = _digits(ymd)
    if len(d) != 8:
        raise ValueError(f"일자 형식이 잘못됐습니다: {ymd!r} (YYYY-MM-DD)")
    datetime.date(int(d[:4]), int(d[4:6]), int(d[6:8]))       # 유효성
    return Period(DAY, d, d, d, f"{d[:4]}-{d[4:6]}-{d[6:8]}")


def month(ym: str) -> Period:
    """그 달 1일~말일(월마감). ym = 'YYYYMM' 또는 'YYYY-MM'. 말일은 달력에서 구한다(윤년 포함)."""
    m = _digits(ym)
    if len(m) == 8:                 # 'YYYYMMDD' 를 줘도 그 달로 받아들인다
        m = m[:6]
    if len(m) != 6:
        raise ValueError(f"연월 형식이 잘못됐습니다: {ym!r} (YYYY-MM)")
    y, mo = int(m[:4]), int(m[4:6])
    if not 1 <= mo <= 12:
        raise ValueError(f"월이 1~12 가 아닙니다: {ym!r}")
    last = calendar.monthrange(y, mo)[1]
    # title 에 '월마감' 을 넣지 않는다 — 구분(kind)은 호출부가 따로 붙인다(중복 표기 방지).
    return Period(MONTH, f"{y:04d}{mo:02d}01", f"{y:04d}{mo:02d}{last:02d}",
                  f"{y:04d}{mo:02d}",
                  f"{y:04d}-{mo:02d} ({y}-{mo:02d}-01 ~ {y}-{mo:02d}-{last:02d})")


def custom(a: str, b: str) -> Period:
    """직접 지정한 기간. 시작이 종료보다 뒤면 바꿔서 받아 준다."""
    x, y = _digits(a), _digits(b)
    for v, nm in ((x, "시작일"), (y, "종료일")):
        if len(v) != 8:
            raise ValueError(f"{nm} 형식이 잘못됐습니다: {a if v is x else b!r} (YYYY-MM-DD)")
        datetime.date(int(v[:4]), int(v[4:6]), int(v[6:8]))
    if x > y:
        x, y = y, x
    if x == y:
        return day(x)
    return Period(RANGE, x, y, f"{x}-{y}",
                  f"{x[:4]}-{x[4:6]}-{x[6:]} ~ {y[:4]}-{y[4:6]}-{y[6:]}")


# ── 상대 기간 ────────────────────────────────────────────────────────
def previous_day(offset_days: int = 1, today: datetime.date | None = None) -> Period:
    """오늘 - offset_days (기본 1 = 전날)."""
    d = (today or datetime.date.today()) - datetime.timedelta(days=int(offset_days))
    return day(d.strftime("%Y%m%d"))


def previous_month(today: datetime.date | None = None) -> Period:
    """전월 전체. 월마감은 보통 달이 바뀐 뒤 전월을 받는다."""
    t = today or datetime.date.today()
    y, m = (t.year - 1, 12) if t.month == 1 else (t.year, t.month - 1)
    return month(f"{y:04d}{m:02d}")


def this_month(today: datetime.date | None = None) -> Period:
    """이번 달 전체(1일~말일). 달 중간에 돌리면 아직 안 온 날짜까지 포함된다."""
    t = today or datetime.date.today()
    return month(f"{t.year:04d}{t.month:02d}")


def parse(spec: str) -> Period:
    """문자열 하나로 기간을 만든다 — GUI/CLI 공용.

      '2026-08-31' / '20260831'        → 하루
      '2026-08'    / '202608'          → 그 달 전체
      '2026-08-01~2026-08-15'          → 직접 기간 ('~' 또는 '..' 또는 ',')
      '전날' / '어제'                   → 전날
      '전월' / '지난달'                 → 전월
      '이번달' / '당월'                 → 이번 달
    """
    s = (spec or "").strip()
    if not s:
        raise ValueError("기간이 비어 있습니다")
    low = s.replace(" ", "")
    if low in ("전날", "어제", "yesterday"):
        return previous_day()
    if low in ("전월", "지난달", "lastmonth"):
        return previous_month()
    if low in ("이번달", "당월", "thismonth"):
        return this_month()
    for sep in ("~", "..", ","):
        if sep in s:
            a, b = s.split(sep, 1)
            return custom(a, b)
    d = _digits(s)
    if len(d) == 6:
        return month(d)
    if len(d) == 8:
        return day(d)
    raise ValueError(f"기간을 알아볼 수 없습니다: {spec!r}\n"
                     f"  예) 2026-08-31 (하루) / 2026-08 (한 달) / "
                     f"2026-08-01~2026-08-15 (기간) / 전날 / 전월")
