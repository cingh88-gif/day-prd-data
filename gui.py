r"""
공정진척현황(PM60250Rv3) 일마감 자동 수집 — GUI

  실행.bat 더블클릭  (또는  pyw gui.py)

하는 일
  · 대상 작업반 12개를 켜고 끄고 추가/삭제
  · 대상일자(기본 전날)로 수집 → 집계 → 보고서까지 한 번에
  · **매일 몇 시에 돌릴지** 정해서 Windows 작업 스케줄러에 등록/해제

⚠️ 수집은 Windows + iERP 로그인(iEMenu 떠 있음) 상태에서만 된다.
   WSL 에서 열면 수집 버튼이 잠기고 '집계·보고서만' 은 그대로 쓸 수 있다.
"""
from __future__ import annotations

import datetime
import os
import sys
import traceback

import config
import period as P
import scheduler

IS_WINDOWS = sys.platform.startswith("win")
FONT = ("맑은 고딕", 10)
FONT_B = ("맑은 고딕", 10, "bold")
FONT_MONO = ("D2Coding", 9) if IS_WINDOWS else ("monospace", 9)


def main():
    import tkinter as tk
    from tkinter import ttk, messagebox, simpledialog

    cfg = config.load()

    root = tk.Tk()
    root.title("iERP 일마감 공정진척 수집 (PM60250Rv3)")
    try:
        import sv_ttk
        sv_ttk.set_theme("light")
    except Exception:
        pass
    try:
        import tkinter.font as tkfont
        for fn in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
                   "TkLabelFont", "TkButtonFont", "TkEntryFont"):
            try:
                tkfont.nametofont(fn).configure(family="맑은 고딕", size=10)
            except Exception:
                pass
    except Exception:
        pass

    outer = ttk.Frame(root, padding=8)
    outer.pack(fill="both", expand=True)
    left = ttk.Frame(outer)
    left.pack(side="left", fill="y")
    right = ttk.Frame(outer)
    right.pack(side="left", fill="both", expand=True, padx=(10, 0))

    # ══════════════ 작업반 ══════════════
    f_team = ttk.LabelFrame(left, text="대상 작업반 (체크한 것만 수집)", padding=6)
    f_team.pack(fill="both", expand=True)

    canvas = tk.Canvas(f_team, width=250, height=300, highlightthickness=0)
    sb = ttk.Scrollbar(f_team, orient="vertical", command=canvas.yview)
    inner = ttk.Frame(canvas)
    inner.bind("<Configure>",
               lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=sb.set)
    canvas.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    team_vars: list[tuple[dict, "tk.BooleanVar"]] = []

    def redraw_teams():
        for w in inner.winfo_children():
            w.destroy()
        team_vars.clear()
        for i, t in enumerate(cfg["teams"]):
            v = tk.BooleanVar(value=t.get("use", True))
            label = f"{t['code']}" + (f"  {t['name']}" if t.get("name") else "")
            ttk.Checkbutton(inner, text=label, variable=v).grid(
                row=i, column=0, sticky="w", padx=4, pady=1)
            team_vars.append((t, v))

    def collect_team_state():
        for t, v in team_vars:
            t["use"] = bool(v.get())

    def add_team():
        code = simpledialog.askstring("작업반 추가", "작업반 코드 (예: M2106)", parent=root)
        if not code:
            return
        code = code.strip().upper()
        if any(t["code"].upper() == code for t in cfg["teams"]):
            messagebox.showinfo("추가", f"{code} 는 이미 목록에 있습니다.")
            return
        name = simpledialog.askstring("작업반 추가",
                                      f"{code} 의 이름(선택 — 비워도 됩니다)", parent=root) or ""
        collect_team_state()
        cfg["teams"].append({"code": code, "name": name.strip(), "use": True})
        redraw_teams()

    def del_team():
        collect_team_state()
        off = [t for t in cfg["teams"] if not t.get("use", True)]
        if not off:
            messagebox.showinfo("삭제", "지울 작업반의 체크를 먼저 해제하세요.\n"
                                        "(체크 해제된 작업반이 목록에서 빠집니다)")
            return
        if not messagebox.askyesno("삭제",
                                   f"체크 해제된 {len(off)}개를 목록에서 지울까요?\n"
                                   + ", ".join(t["code"] for t in off)):
            return
        cfg["teams"] = [t for t in cfg["teams"] if t.get("use", True)]
        redraw_teams()

    bar_t = ttk.Frame(left)
    bar_t.pack(fill="x", pady=(4, 0))
    ttk.Button(bar_t, text="＋ 추가", command=add_team, width=9).pack(side="left")
    ttk.Button(bar_t, text="－ 삭제", command=del_team, width=9).pack(side="left", padx=4)
    ttk.Button(bar_t, text="전체선택", width=9,
               command=lambda: [v.set(True) for _, v in team_vars]).pack(side="left")
    redraw_teams()

    # ══════════════ 실행 조건 ══════════════
    f_run = ttk.LabelFrame(right, text="실행 조건", padding=6)
    f_run.pack(fill="x")

    # ── 기간 모드: 일마감 / 월마감 / 직접지정 ──
    v_mode = tk.StringVar(value="day")
    v_a = tk.StringVar()      # 일마감=대상일자, 월마감=연월, 직접=시작일
    v_b = tk.StringVar()      # 직접지정 종료일
    v_auto = tk.BooleanVar(value=True)   # 실행 시점 기준 자동(전날/전월)

    row = ttk.Frame(f_run)
    row.grid(row=0, column=0, columnspan=4, sticky="w", padx=4, pady=(2, 4))
    for val, txt in (("day", "일마감 (하루)"), ("month", "월마감 (한 달)"),
                     ("range", "직접 지정")):
        ttk.Radiobutton(row, text=txt, value=val, variable=v_mode,
                        command=lambda: refresh_period()).pack(side="left", padx=(0, 14))

    lbl_a = ttk.Label(f_run, text="대상일자")
    lbl_a.grid(row=1, column=0, sticky="e", padx=4, pady=4)
    e_a = ttk.Entry(f_run, textvariable=v_a, width=13)
    e_a.grid(row=1, column=1, sticky="w", pady=4)
    lbl_b = ttk.Label(f_run, text="~ 종료일")
    e_b = ttk.Entry(f_run, textvariable=v_b, width=13)
    chk_auto = ttk.Checkbutton(f_run, text="실행 시점 기준 자동", variable=v_auto,
                               command=lambda: refresh_period())
    v_period_desc = tk.StringVar()
    ttk.Label(f_run, textvariable=v_period_desc, font=("맑은 고딕", 9),
              foreground="#356").grid(row=2, column=0, columnspan=4, sticky="w",
                                      padx=6, pady=(0, 4))

    def current_period():
        """지금 화면 설정으로 기간 객체를 만든다. 잘못된 입력이면 ValueError."""
        m = v_mode.get()
        if v_auto.get() and m == "day":
            return P.previous_day(cfg.get("offset_days", 1))
        if v_auto.get() and m == "month":
            return P.previous_month()
        if m == "day":
            return P.day(v_a.get())
        if m == "month":
            return P.month(v_a.get())
        return P.custom(v_a.get(), v_b.get())

    def refresh_period(*_):
        m = v_mode.get()
        lbl_b.grid_forget(); e_b.grid_forget(); chk_auto.grid_forget()
        if m == "range":
            lbl_a.configure(text="시작일")
            lbl_b.grid(row=1, column=2, sticky="e", padx=4)
            e_b.grid(row=1, column=3, sticky="w", padx=4)
            if not v_a.get():
                v_a.set(P.previous_month().d_from)
            if not v_b.get():
                v_b.set(P.previous_month().d_to)
        else:
            lbl_a.configure(text="대상일자" if m == "day" else "대상 연월")
            chk_auto.grid(row=1, column=2, sticky="w", padx=8)
        if v_auto.get() and m != "range":
            auto = P.previous_day(cfg.get("offset_days", 1)) if m == "day" \
                else P.previous_month()
            v_a.set(auto.label if m == "month" else auto.d_from)
            e_a.configure(state="disabled")
        else:
            e_a.configure(state="normal")
        try:
            per = current_period()
            kind = {"day": "일마감", "month": "월마감", "range": "기간"}[per.mode]
            v_period_desc.set(f"→ {kind} · 개시예정일 {per.title}   |   폴더 {per.dirname}")
        except ValueError as e:
            v_period_desc.set(f"→ {e}".splitlines()[0])

    for var in (v_a, v_b):
        var.trace_add("write", refresh_period)
    refresh_period()

    ttk.Label(f_run, text="저장 폴더").grid(row=3, column=0, sticky="e", padx=4, pady=4)
    v_out = tk.StringVar(value=cfg["export_dir"])
    ttk.Entry(f_run, textvariable=v_out, width=42).grid(
        row=3, column=1, columnspan=3, sticky="we", pady=4)
    f_run.columnconfigure(3, weight=1)

    # ══════════════ 매일 자동 실행 ══════════════
    f_sch = ttk.LabelFrame(right, text="매일 자동 실행 (Windows 작업 스케줄러)", padding=6)
    f_sch.pack(fill="x", pady=8)

    ttk.Label(f_sch, text="실행 시각").grid(row=0, column=0, sticky="e", padx=4, pady=4)
    v_time = tk.StringVar(value=cfg["schedule"].get("time", "11:00"))
    ttk.Entry(f_sch, textvariable=v_time, width=8).grid(row=0, column=1, sticky="w", pady=4)
    v_time_desc = tk.StringVar()
    ttk.Label(f_sch, textvariable=v_time_desc, font=("맑은 고딕", 9)).grid(
        row=0, column=2, sticky="w", padx=6)

    def on_time_change(*_):
        t = v_time.get().strip()
        v_time_desc.set(f"→ 매일 {scheduler.describe_time(t)}"
                        if scheduler.valid_time(t) else "→ 24시간제 HH:MM 로 (예: 11:00, 23:30)")
    v_time.trace_add("write", on_time_change)
    on_time_change()

    v_sch_state = tk.StringVar(value="확인 중…")
    ttk.Label(f_sch, textvariable=v_sch_state, font=("맑은 고딕", 9)).grid(
        row=1, column=0, columnspan=4, sticky="w", padx=4, pady=(2, 4))

    def refresh_sched():
        ok, msg = scheduler.query(cfg["schedule"].get("task_name", scheduler.DEFAULT_TASK))
        v_sch_state.set(("● 일마감 " if ok else "○ 일마감 ") + msg)
        try:
            ok2, msg2 = scheduler.query(mtask())
            v_msch_state.set(("● 월마감 " if ok2 else "○ 월마감 ") + msg2)
        except NameError:
            pass

    def do_register():
        t = v_time.get().strip()
        if not scheduler.valid_time(t):
            messagebox.showwarning("시간", "24시간제 HH:MM 로 입력하세요. 예) 11:00, 12:00, 23:30")
            return
        save_cfg()
        ok, msg = scheduler.create(t, cfg["schedule"].get("task_name", scheduler.DEFAULT_TASK))
        refresh_sched()
        (messagebox.showinfo if ok else messagebox.showerror)("자동 실행", msg)
        if ok:
            log(f"[스케줄] {msg}")

    def do_unregister():
        ok, msg = scheduler.delete(cfg["schedule"].get("task_name", scheduler.DEFAULT_TASK))
        refresh_sched()
        (messagebox.showinfo if ok else messagebox.showerror)("자동 실행", msg)

    ttk.Button(f_sch, text="일마감 등록", command=do_register, width=12).grid(
        row=2, column=0, padx=4, pady=2)
    ttk.Button(f_sch, text="해제", command=do_unregister, width=8).grid(
        row=2, column=1, padx=2, pady=2, sticky="w")
    ttk.Button(f_sch, text="상태 새로고침", command=lambda: refresh_sched(), width=13).grid(
        row=2, column=2, padx=2, pady=2, sticky="w")

    ttk.Separator(f_sch, orient="horizontal").grid(
        row=3, column=0, columnspan=4, sticky="we", pady=6)

    # ── 월마감: 매월 N일에 전월 전체 ──
    ttk.Label(f_sch, text="월마감 — 매월").grid(row=4, column=0, sticky="e", padx=4, pady=4)
    mrow = ttk.Frame(f_sch); mrow.grid(row=4, column=1, columnspan=2, sticky="w")
    v_mday = tk.StringVar(value=str(cfg["schedule"].get("month_day", 1)))
    ttk.Spinbox(mrow, from_=1, to=28, textvariable=v_mday, width=4).pack(side="left")
    ttk.Label(mrow, text="일").pack(side="left", padx=(2, 10))
    v_mtime = tk.StringVar(value=cfg["schedule"].get("month_time", "11:00"))
    ttk.Entry(mrow, textvariable=v_mtime, width=8).pack(side="left")
    v_mdesc = tk.StringVar()
    ttk.Label(mrow, textvariable=v_mdesc, font=("맑은 고딕", 9)).pack(side="left", padx=8)

    def on_mtime_change(*_):
        t = v_mtime.get().strip()
        v_mdesc.set(f"→ {scheduler.describe_time(t)} 에 전월 전체"
                    if scheduler.valid_time(t) else "→ HH:MM")
    v_mtime.trace_add("write", on_mtime_change); on_mtime_change()

    v_msch_state = tk.StringVar(value="확인 중…")
    ttk.Label(f_sch, textvariable=v_msch_state, font=("맑은 고딕", 9)).grid(
        row=5, column=0, columnspan=4, sticky="w", padx=4, pady=(2, 4))

    def mtask():
        return cfg["schedule"].get("month_task_name", scheduler.DEFAULT_MONTH_TASK)

    def do_register_month():
        t = v_mtime.get().strip()
        try:
            d = int(v_mday.get())
        except ValueError:
            messagebox.showwarning("실행일", "실행일은 1~28 사이의 숫자여야 합니다."); return
        save_cfg()
        ok, msg = scheduler.create_monthly(t, d, mtask())
        refresh_sched()
        (messagebox.showinfo if ok else messagebox.showerror)("월마감 자동 실행", msg)
        if ok:
            log(f"[스케줄] {msg}")

    def do_unregister_month():
        ok, msg = scheduler.delete(mtask())
        refresh_sched()
        (messagebox.showinfo if ok else messagebox.showerror)("월마감 자동 실행", msg)

    ttk.Button(f_sch, text="월마감 등록", command=do_register_month, width=12).grid(
        row=6, column=0, padx=4, pady=2)
    ttk.Button(f_sch, text="해제", command=do_unregister_month, width=8).grid(
        row=6, column=1, padx=2, pady=2, sticky="w")

    ttk.Label(f_sch,
              text="※ 자동 실행은 PC 가 켜져 있고 로그온된 상태에서만 됩니다 "
                   "(iERP 창을 조작해야 하므로).\n"
                   "※ 월마감은 매월 지정일에 전월 한 달치를 받습니다. "
                   "29~31일은 없는 달이 있어 1~28일만 됩니다.",
              font=("맑은 고딕", 9), foreground="#666", justify="left").grid(
        row=7, column=0, columnspan=4, sticky="w", padx=4, pady=(4, 0))

    # ══════════════ 로그 ══════════════
    f_log = ttk.LabelFrame(right, text="진행", padding=4)
    f_log.pack(fill="both", expand=True)
    txt = tk.Text(f_log, height=16, width=76, font=FONT_MONO, wrap="none")
    sb2 = ttk.Scrollbar(f_log, orient="vertical", command=txt.yview)
    txt.configure(yscrollcommand=sb2.set, state="disabled")
    txt.pack(side="left", fill="both", expand=True)
    sb2.pack(side="right", fill="y")

    def log(msg=""):
        txt.configure(state="normal")
        txt.insert("end", str(msg) + "\n")
        txt.see("end")
        txt.configure(state="disabled")
        root.update_idletasks()

    # ══════════════ 실행 버튼 ══════════════
    stop_flag = {"stop": False}
    bar = ttk.Frame(right)
    bar.pack(fill="x", pady=(8, 0))
    btn_run = ttk.Button(bar, text="지금 실행 (수집 → 보고서)")
    btn_run.pack(side="left")
    btn_again = ttk.Button(bar, text="집계·보고서만 다시")
    btn_again.pack(side="left", padx=6)
    btn_stop = ttk.Button(bar, text="중단", state="disabled")
    btn_stop.pack(side="left")
    ttk.Button(bar, text="폴더 열기",
               command=lambda: _open(config.to_local_path(v_out.get()))).pack(side="right")

    def _open(p):
        try:
            if IS_WINDOWS:
                os.startfile(str(p))
            else:
                os.system(f'explorer.exe "$(wslpath -w "{p}")" 2>/dev/null &')
        except Exception as e:
            log(f"※ 폴더 열기 실패: {e}")

    def save_cfg():
        collect_team_state()
        cfg["export_dir"] = v_out.get().strip() or config.default_export_dir()
        cfg["schedule"]["time"] = v_time.get().strip()
        cfg["schedule"]["month_time"] = v_mtime.get().strip()
        try:
            cfg["schedule"]["month_day"] = int(v_mday.get())
        except ValueError:
            pass
        cfg["offset_days"] = 1
        config.save(cfg)

    def _do(do_collect: bool):
        import run_daily
        if do_collect and not IS_WINDOWS:
            messagebox.showwarning(
                "수집 불가",
                "수집은 Windows 에서만 됩니다 (iERP 창을 조작해야 하므로).\n"
                "Windows 쪽 사본(C:\\ierp_day_prd)의 실행.bat 으로 여세요.\n\n"
                "'집계·보고서만 다시' 는 여기서도 됩니다.")
            return
        save_cfg()
        if not config.active_teams(cfg):
            messagebox.showwarning("작업반", "수집할 작업반을 하나 이상 체크하세요.")
            return
        stop_flag["stop"] = False
        btn_run.configure(state="disabled")
        btn_again.configure(state="disabled")
        btn_stop.configure(state="normal" if do_collect else "disabled")
        try:
            per = current_period()
        except ValueError as e:
            messagebox.showwarning("기간", str(e))
            btn_run.configure(state="normal"); btn_again.configure(state="normal")
            btn_stop.configure(state="disabled")
            return
        kind = {"day": "일마감", "month": "월마감", "range": "기간"}[per.mode]
        log("=" * 60)
        log(f"[{datetime.datetime.now():%H:%M:%S}] 시작 — {kind} {per.title}"
            f"{' (수집 포함)' if do_collect else ' (집계·보고서만)'}")
        if do_collect and per.mode != "day":
            log(f"  ※ 기간이 { (int(per.d_to) - int(per.d_from)) and '하루가 아닙니다' or ''}"
                f" 한 번에 받으므로 조회가 오래 걸릴 수 있습니다.")
        try:
            # ※ pywinauto(uia)는 메인 스레드에서 돌아야 한다 → 별도 스레드를 쓰지 않고
            #    root.update 를 pump 로 넘겨 화면이 멈춘 것처럼 보이지 않게 한다.
            r = run_daily.run(per, cfg, do_collect=do_collect, log=log,
                              should_stop=lambda: stop_flag["stop"], pump=root.update)
            fails = (r["result"] or {}).get("failed", [])
            msg = f"보고서:\n{r['report']}"
            if fails:
                msg += f"\n\n※ 수집 실패 작업반: {', '.join(fails)}\n" \
                       "같은 기간으로 다시 실행하면 이어받습니다."
            if cfg.get("open_report"):
                _open(r["report"])
            (messagebox.showwarning if fails else messagebox.showinfo)("완료", msg)
        except Exception as e:
            from collect import Cancelled
            if isinstance(e, Cancelled):
                log("중단됨")
                messagebox.showinfo("중단", "중단했습니다. 같은 날짜로 다시 실행하면 이어받습니다.")
            else:
                log(f"※ 오류: {e}")
                log(traceback.format_exc())
                messagebox.showerror("오류", str(e))
        finally:
            btn_run.configure(state="normal")
            btn_again.configure(state="normal")
            btn_stop.configure(state="disabled")

    btn_run.configure(command=lambda: _do(True))
    btn_again.configure(command=lambda: _do(False))

    def on_stop():
        stop_flag["stop"] = True
        btn_stop.configure(state="disabled")
        log("중단 요청 — 현재 작업반이 끝나면 멈춥니다…")
    btn_stop.configure(command=on_stop)

    # ── 첫 안내 ─────────────────────────────────────────────
    log("집계 기준")
    log("  · 오더상태 30작업지시 = 제외")
    log("  · 모수 = 50작업진행 + 80작업완료")
    log("  · 일마감 진행률 = 80작업완료 ÷ 모수")
    log("")
    if IS_WINDOWS:
        log("수집 전 확인: iERP 에 로그인돼 iEMenu 가 떠 있어야 합니다.")
        log("수집 중에는 마우스·키보드를 만지지 마세요 (키 입력이 iERP 창으로 들어갑니다).")
    else:
        log("※ 지금은 WSL 입니다 — 수집은 잠겨 있고 '집계·보고서만 다시' 만 됩니다.")
    log("")
    refresh_sched()

    def on_close():
        try:
            save_cfg()
        except Exception:
            pass
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_close)

    root.update_idletasks()
    root.minsize(900, 560)
    root.mainloop()


if __name__ == "__main__":
    main()
