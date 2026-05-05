#!/usr/bin/env python3
"""Resume Analyzer — AI-Powered HR Screening Tool."""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import json
import re
from pathlib import Path

# ─── Text extraction ──────────────────────────────────────────────────────────

def extract_pdf(path: str) -> str:
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception as exc:
        return f"[PDF read error: {exc}]"


def extract_docx(path: str) -> str:
    try:
        from docx import Document
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as exc:
        return f"[DOCX read error: {exc}]"


def extract_pptx(path: str) -> str:
    try:
        from pptx import Presentation
        prs = Presentation(path)
        parts = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    parts.append(shape.text)
        return "\n".join(parts)
    except Exception as exc:
        return f"[PPTX read error: {exc}]"


def extract_txt(path: str) -> str:
    try:
        return Path(path).read_text(encoding="utf-8", errors="replace")
    except Exception as exc:
        return f"[TXT read error: {exc}]"


def extract_text(path: str) -> str:
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        return extract_pdf(path)
    if ext in {".doc", ".docx"}:
        return extract_docx(path)
    if ext in {".ppt", ".pptx"}:
        return extract_pptx(path)
    if ext == ".txt":
        return extract_txt(path)
    return f"[Unsupported file type: {ext}]"


# ─── LLM interaction ──────────────────────────────────────────────────────────

ANALYSIS_PROMPT = """\
You are an expert HR analyst. Analyse the resume against the job description below.
Respond with ONLY a valid JSON object — no markdown fences, no explanation, just raw JSON.

JOB DESCRIPTION:
{jd}

RESUME:
{resume}

Return this exact JSON (fill in real values):
{{
  "name": "<candidate full name>",
  "total_experience": {{"years": <int>, "months": <int 0-11>}},
  "relevant_experience": {{"years": <int>, "months": <int 0-11>}},
  "highlights": ["<strength 1>", "<strength 2>", "<strength 3>"],
  "lowlights": ["<gap 1>", "<gap 2>"],
  "score": <integer 1-10>,
  "decision": "<Select if score > 6, else Reject>"
}}

Rules:
- score is 1–10 based on how well the resume matches the job description
- decision MUST be "Select" when score > 6, otherwise "Reject"
- Extract the candidate name directly from the resume text
- Be specific in highlights and lowlights (3 highlights, 2 lowlights minimum)
"""


def _strip_fences(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```[a-z]*\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()


def call_llm(model: str, api_key: str, prompt: str) -> str:
    if "claude" in model.lower():
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model=model,
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        return msg.content[0].text
    else:
        import openai
        client = openai.OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1024,
        )
        return resp.choices[0].message.content


def analyse_resume(path: str, jd_text: str, model: str, api_key: str) -> dict:
    resume_text = extract_text(path)
    prompt = ANALYSIS_PROMPT.format(jd=jd_text, resume=resume_text)
    raw = call_llm(model, api_key, prompt)
    return json.loads(_strip_fences(raw))


# ─── Styling constants ────────────────────────────────────────────────────────

ACCENT       = "#1565C0"
ACCENT_HOVER = "#1976D2"
BG           = "#F0F4FF"
CARD         = "#FFFFFF"
SEL_BG       = "#C8E6C9"
REJ_BG       = "#FFCDD2"
ROW_EVEN     = "#EEF2FF"
FONT_BODY    = ("Segoe UI", 10)
FONT_BOLD    = ("Segoe UI", 10, "bold")
FONT_TITLE   = ("Segoe UI", 15, "bold")
FONT_SMALL   = ("Segoe UI", 9)
FONT_MONO    = ("Consolas", 9)


# ─── Helpers ──────────────────────────────────────────────────────────────────

def fmt_exp(exp: dict | None) -> str:
    if not exp:
        return "N/A"
    y, m = int(exp.get("years", 0)), int(exp.get("months", 0))
    return f"{y}y {m}m"


class ToolTip:
    def __init__(self, widget, text: str):
        self.tip = None
        widget.bind("<Enter>", lambda _: self._show(widget, text))
        widget.bind("<Leave>", lambda _: self._hide())

    def _show(self, widget, text):
        x = widget.winfo_rootx() + 20
        y = widget.winfo_rooty() + widget.winfo_height() + 4
        self.tip = tk.Toplevel(widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=text, bg="#FFFDE7", relief="solid",
                 borderwidth=1, font=FONT_SMALL, padx=4, pady=2).pack()

    def _hide(self):
        if self.tip:
            self.tip.destroy()
            self.tip = None


# ─── Main application ─────────────────────────────────────────────────────────

class ResumeAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Resume Analyzer — AI-Powered HR Tool")
        self.geometry("1280x820")
        self.minsize(960, 640)
        self.configure(bg=BG)
        self.resume_files: list[str] = []
        self.results: list[dict] = []
        self._build_ui()
        self._apply_style()

    # ── Style ─────────────────────────────────────────────────────────────────

    def _apply_style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TEntry", fieldbackground="white", relief="flat", padding=4)
        s.configure("TScrollbar", troughcolor=BG, background="#BDBDBD")
        s.configure("Results.Treeview",
                    rowheight=56, font=FONT_SMALL,
                    background="white", fieldbackground="white")
        s.configure("Results.Treeview.Heading",
                    font=FONT_BOLD, background=ACCENT, foreground="white",
                    relief="flat", padding=6)
        s.map("Results.Treeview.Heading", background=[("active", ACCENT_HOVER)])
        s.map("Results.Treeview", background=[("selected", "#BBDEFB")])
        s.configure("Accent.TButton", background=ACCENT, foreground="white",
                    font=FONT_BOLD, padding=8, relief="flat")
        s.map("Accent.TButton",
              background=[("active", ACCENT_HOVER), ("disabled", "#90A4AE")])

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Title bar
        hdr = tk.Frame(self, bg=ACCENT, height=58)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  Resume Analyzer", font=FONT_TITLE,
                 bg=ACCENT, fg="white").pack(side="left", pady=12)
        tk.Label(hdr, text="AI-Powered HR Screening Tool",
                 font=FONT_BODY, bg=ACCENT, fg="#90CAF9").pack(side="left", padx=6)

        # Scrollable main area
        canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        vscroll = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self._main = tk.Frame(canvas, bg=BG)
        win_id = canvas.create_window((0, 0), window=self._main, anchor="nw")

        def _on_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_resize(e):
            canvas.itemconfig(win_id, width=e.width)

        self._main.bind("<Configure>", _on_configure)
        canvas.bind("<Configure>", _on_canvas_resize)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        pad = {"padx": 16, "pady": 6}

        self._build_config_card(**pad)
        self._build_jd_card(**pad)
        self._build_resumes_card(**pad)
        self._build_action_row(**pad)
        self._build_results_card(padx=16, pady=(0, 16))

    def _card(self, title: str) -> tk.Frame:
        outer = tk.Frame(self._main, bg=CARD,
                         highlightbackground="#C5CAE9", highlightthickness=1)
        outer.pack(fill="x", padx=16, pady=6)
        tk.Label(outer, text=title, font=FONT_BOLD,
                 bg=CARD, fg=ACCENT).pack(anchor="w", padx=12, pady=(8, 2))
        ttk.Separator(outer, orient="horizontal").pack(fill="x", padx=12, pady=(0, 4))
        return outer

    def _flat_btn(self, parent, text, command, primary=True, **kw):
        bg = ACCENT if primary else "#E3E8F8"
        fg = "white" if primary else "#333"
        hv = ACCENT_HOVER if primary else "#C5CAE9"
        b = tk.Button(parent, text=text, command=command,
                      font=FONT_SMALL, bg=bg, fg=fg,
                      activebackground=hv, activeforeground=fg,
                      relief="flat", padx=10, pady=4, cursor="hand2", **kw)
        b.bind("<Enter>", lambda _: b.config(bg=hv))
        b.bind("<Leave>", lambda _: b.config(bg=bg))
        return b

    # ── Config card ───────────────────────────────────────────────────────────

    def _build_config_card(self, **pad):
        card = self._card("Configuration")
        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", padx=12, pady=8)

        tk.Label(row, text="LLM Model:", font=FONT_BOLD,
                 bg=CARD, width=12, anchor="w").pack(side="left")
        self.model_var = tk.StringVar(value="claude-sonnet-4-6")
        me = ttk.Entry(row, textvariable=self.model_var, width=30, font=FONT_BODY)
        me.pack(side="left", padx=(0, 24))
        ToolTip(me, "Claude: claude-sonnet-4-6, claude-opus-4-7\nOpenAI: gpt-4o, gpt-4-turbo")

        tk.Label(row, text="API Key:", font=FONT_BOLD,
                 bg=CARD, width=8, anchor="w").pack(side="left")
        self.apikey_var = tk.StringVar()
        self._key_entry = ttk.Entry(row, textvariable=self.apikey_var,
                                    width=48, show="•", font=FONT_MONO)
        self._key_entry.pack(side="left", padx=(0, 8))
        self._show_key = tk.BooleanVar(value=False)
        tk.Checkbutton(row, text="Show", variable=self._show_key,
                       bg=CARD, font=FONT_SMALL,
                       command=lambda: self._key_entry.config(
                           show="" if self._show_key.get() else "•")
                       ).pack(side="left")

    # ── JD card ───────────────────────────────────────────────────────────────

    def _build_jd_card(self, **pad):
        card = self._card("Job Description")
        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", padx=12, pady=8)

        self.jd_var = tk.StringVar()
        jd_entry = ttk.Entry(row, textvariable=self.jd_var, width=72, font=FONT_SMALL)
        jd_entry.pack(side="left", padx=(0, 8))
        ToolTip(jd_entry, "Supported: PDF, DOCX, PPTX, TXT")
        self._flat_btn(row, "Browse…", self._browse_jd).pack(side="left")

    # ── Resumes card ──────────────────────────────────────────────────────────

    def _build_resumes_card(self, **pad):
        card = self._card("Resumes")

        btn_row = tk.Frame(card, bg=CARD)
        btn_row.pack(fill="x", padx=12, pady=(4, 6))
        self._flat_btn(btn_row, "+ Add Resumes", self._add_resumes).pack(side="left", padx=(0, 8))
        self._flat_btn(btn_row, "Remove Selected", self._remove_selected,
                       primary=False).pack(side="left", padx=(0, 8))
        self._flat_btn(btn_row, "Clear All", self._clear_resumes,
                       primary=False).pack(side="left")
        self._count_lbl = tk.Label(btn_row, text="No files added",
                                   font=FONT_SMALL, bg=CARD, fg="#555")
        self._count_lbl.pack(side="left", padx=16)

        list_frame = tk.Frame(card, bg=CARD)
        list_frame.pack(fill="x", padx=12, pady=(0, 10))
        self._listbox = tk.Listbox(list_frame, height=5, font=FONT_SMALL,
                                   bg="#F8F9FE", relief="flat", borderwidth=1,
                                   selectmode="extended", activestyle="none",
                                   highlightbackground="#C5CAE9", highlightthickness=1)
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self._listbox.yview)
        self._listbox.configure(yscrollcommand=sb.set)
        self._listbox.pack(side="left", fill="x", expand=True)
        sb.pack(side="right", fill="y")

    # ── Action row ────────────────────────────────────────────────────────────

    def _build_action_row(self, **pad):
        row = tk.Frame(self._main, bg=BG)
        row.pack(fill="x", padx=16, pady=4)

        self._analyse_btn = tk.Button(
            row, text="▶  Provide Summary",
            font=("Segoe UI", 12, "bold"), bg=ACCENT, fg="white",
            activebackground=ACCENT_HOVER, activeforeground="white",
            relief="flat", padx=28, pady=10, cursor="hand2",
            command=self._start_analysis,
        )
        self._analyse_btn.pack(side="left", padx=(0, 16))

        progress_col = tk.Frame(row, bg=BG)
        progress_col.pack(side="left", fill="x", expand=True)
        self._progress_lbl = tk.Label(progress_col, text="",
                                      font=FONT_SMALL, bg=BG, fg="#444")
        self._progress_lbl.pack(anchor="w")
        self._progress = ttk.Progressbar(progress_col, mode="determinate", length=500)
        self._progress.pack(anchor="w", pady=(2, 0))

    # ── Results card ──────────────────────────────────────────────────────────

    def _build_results_card(self, **pad):
        card = self._card("Analysis Results")
        card.pack(fill="both", expand=True, **pad)

        export_row = tk.Frame(card, bg=CARD)
        export_row.pack(fill="x", padx=12, pady=(0, 6))
        self._flat_btn(export_row, "Export to Excel (.xlsx)",
                       self._export_excel).pack(side="right", padx=(8, 0))
        self._flat_btn(export_row, "Export to CSV",
                       self._export_csv, primary=False).pack(side="right")
        tk.Label(export_row,
                 text="Double-click a row to view full details",
                 font=FONT_SMALL, bg=CARD, fg="#777").pack(side="left")

        cols = ("#", "Name", "Total Exp.", "Relevant Exp.",
                "Highlights", "Lowlights", "Score", "Decision")
        widths = (36, 160, 90, 95, 280, 220, 52, 78)

        tframe = tk.Frame(card, bg=CARD)
        tframe.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        self.tree = ttk.Treeview(tframe, columns=cols, show="headings",
                                  style="Results.Treeview", selectmode="browse")
        for col, w in zip(cols, widths):
            anchor = "center" if col in ("#", "Score", "Decision") else "w"
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor=anchor,
                             stretch=(col in ("Highlights", "Lowlights")))

        vsb = ttk.Scrollbar(tframe, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tframe, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tframe.grid_rowconfigure(0, weight=1)
        tframe.grid_columnconfigure(0, weight=1)

        self.tree.tag_configure("select", background=SEL_BG)
        self.tree.tag_configure("reject", background=REJ_BG)
        self.tree.tag_configure("even",   background=ROW_EVEN)

        self.tree.bind("<Double-1>", self._show_detail)

    # ── Event handlers ────────────────────────────────────────────────────────

    def _browse_jd(self):
        path = filedialog.askopenfilename(
            title="Select Job Description File",
            filetypes=[("Supported Documents",
                        "*.pdf *.doc *.docx *.ppt *.pptx *.txt"),
                       ("All files", "*.*")],
        )
        if path:
            self.jd_var.set(path)

    def _add_resumes(self):
        paths = filedialog.askopenfilenames(
            title="Select Resume Files",
            filetypes=[("Supported Documents",
                        "*.pdf *.doc *.docx *.ppt *.pptx *.txt"),
                       ("All files", "*.*")],
        )
        for p in paths:
            if p not in self.resume_files:
                self.resume_files.append(p)
                self._listbox.insert("end", os.path.basename(p))
        self._update_count()

    def _clear_resumes(self):
        self.resume_files.clear()
        self._listbox.delete(0, "end")
        self._update_count()

    def _remove_selected(self):
        for idx in reversed(self._listbox.curselection()):
            self._listbox.delete(idx)
            self.resume_files.pop(idx)
        self._update_count()

    def _update_count(self):
        n = len(self.resume_files)
        self._count_lbl.config(
            text=f"{n} file{'s' if n != 1 else ''} added" if n else "No files added"
        )

    def _validate(self) -> bool:
        if not self.jd_var.get():
            messagebox.showwarning("Missing Input", "Please select a Job Description file.")
            return False
        if not os.path.isfile(self.jd_var.get()):
            messagebox.showwarning("File Not Found",
                                   f"Job Description file not found:\n{self.jd_var.get()}")
            return False
        if not self.resume_files:
            messagebox.showwarning("Missing Input",
                                   "Please add at least one resume file.")
            return False
        if not self.model_var.get().strip():
            messagebox.showwarning("Missing Input", "Please enter an LLM model name.")
            return False
        if not self.apikey_var.get().strip():
            messagebox.showwarning("Missing Input", "Please enter your API key.")
            return False
        return True

    def _start_analysis(self):
        if not self._validate():
            return
        self._analyse_btn.config(state="disabled")
        self.results.clear()
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        threading.Thread(target=self._run_analysis, daemon=True).start()

    def _run_analysis(self):
        try:
            jd_text = extract_text(self.jd_var.get())
            model   = self.model_var.get().strip()
            api_key = self.apikey_var.get().strip()
            total   = len(self.resume_files)

            for i, path in enumerate(self.resume_files):
                fname = os.path.basename(path)
                self._update_progress(i, total, f"Analysing {fname}…")
                try:
                    result = analyse_resume(path, jd_text, model, api_key)
                    result["_file"] = fname
                    self.results.append(result)
                    self._append_row(len(self.results), result)
                except json.JSONDecodeError as exc:
                    self._err(fname, f"LLM returned invalid JSON.\n\nDetails: {exc}")
                except Exception as exc:
                    self._err(fname, str(exc))

            self._update_progress(total, total,
                                  f"Complete — {total} resume(s) analysed.")
        finally:
            self.after(0, lambda: self._analyse_btn.config(state="normal"))

    def _update_progress(self, current: int, total: int, msg: str):
        pct = (current / total * 100) if total else 0
        def _do():
            self._progress.config(value=pct)
            self._progress_lbl.config(text=msg)
        self.after(0, _do)

    def _append_row(self, serial: int, r: dict):
        highlights = " | ".join(r.get("highlights", []))
        lowlights  = " | ".join(r.get("lowlights",  []))
        score      = r.get("score", 0)
        decision   = r.get("decision", "Reject")
        tag        = "select" if str(decision).strip().lower() == "select" else "reject"

        values = (
            serial,
            r.get("name", "Unknown"),
            fmt_exp(r.get("total_experience")),
            fmt_exp(r.get("relevant_experience")),
            highlights,
            lowlights,
            score,
            decision,
        )
        self.after(0, lambda v=values, t=tag: self.tree.insert("", "end", values=v, tags=(t,)))

    def _err(self, fname: str, msg: str):
        self.after(0, lambda: messagebox.showerror(
            "Analysis Error", f"Failed to analyse:\n{fname}\n\n{msg}"
        ))

    # ── Detail popup ──────────────────────────────────────────────────────────

    def _show_detail(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if idx >= len(self.results):
            return
        r = self.results[idx]

        win = tk.Toplevel(self)
        win.title(f"Detail — {r.get('name', 'Unknown')}")
        win.geometry("660x560")
        win.configure(bg=BG)
        win.grab_set()
        win.resizable(True, True)

        hdr = tk.Frame(win, bg=ACCENT, height=46)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"  {r.get('name', 'Unknown')}", font=FONT_BOLD,
                 bg=ACCENT, fg="white").pack(side="left", pady=10)

        content = tk.Frame(win, bg=BG)
        content.pack(fill="both", expand=True, padx=16, pady=12)

        def section(title: str, text: str):
            tk.Label(content, text=title, font=FONT_BOLD,
                     bg=BG, fg=ACCENT).pack(anchor="w", pady=(10, 2))
            lbl = tk.Label(content, text=text, font=FONT_BODY, bg=CARD,
                           wraplength=600, justify="left", anchor="nw",
                           relief="flat", padx=10, pady=6)
            lbl.pack(fill="x")

        section("File", r.get("_file", ""))
        section("Total Experience",    fmt_exp(r.get("total_experience")))
        section("Relevant Experience", fmt_exp(r.get("relevant_experience")))
        section("Highlights",
                "\n".join(f"• {h}" for h in r.get("highlights", ["—"])))
        section("Lowlights",
                "\n".join(f"• {l}" for l in r.get("lowlights",  ["—"])))

        score    = r.get("score", 0)
        decision = r.get("decision", "Reject")
        fg_dec   = "#1B5E20" if str(decision).strip().lower() == "select" else "#B71C1C"
        row_bot  = tk.Frame(content, bg=BG)
        row_bot.pack(fill="x", pady=(14, 0))
        tk.Label(row_bot, text=f"Score: {score} / 10",
                 font=("Segoe UI", 13, "bold"), bg=BG).pack(side="left", padx=(0, 24))
        tk.Label(row_bot, text=f"Decision: {decision}",
                 font=("Segoe UI", 13, "bold"), bg=BG, fg=fg_dec).pack(side="left")

        self._flat_btn(win, "Close", win.destroy).pack(pady=12)

    # ── Export ────────────────────────────────────────────────────────────────

    def _export_csv(self):
        if not self.results:
            messagebox.showinfo("No Data", "Run the analysis first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile="resume_analysis",
        )
        if not path:
            return
        import csv
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["#", "Name", "Total Experience", "Relevant Experience",
                        "Highlights", "Lowlights", "Score", "Decision"])
            for i, r in enumerate(self.results, 1):
                w.writerow([
                    i, r.get("name", ""),
                    fmt_exp(r.get("total_experience")),
                    fmt_exp(r.get("relevant_experience")),
                    " | ".join(r.get("highlights", [])),
                    " | ".join(r.get("lowlights", [])),
                    r.get("score", ""),
                    r.get("decision", ""),
                ])
        messagebox.showinfo("Exported", f"CSV saved to:\n{path}")

    def _export_excel(self):
        if not self.results:
            messagebox.showinfo("No Data", "Run the analysis first.")
            return
        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        except ImportError:
            messagebox.showerror("Error",
                                 "openpyxl not available. Please use CSV export.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="resume_analysis",
        )
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Resume Analysis"

        # Header style
        h_fill   = PatternFill("solid", fgColor="1565C0")
        h_font   = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        thin     = Side(style="thin", color="C5CAE9")
        border   = Border(left=thin, right=thin, top=thin, bottom=thin)
        c_align  = Alignment(horizontal="center", vertical="top", wrap_text=True)
        l_align  = Alignment(horizontal="left",   vertical="top", wrap_text=True)

        headers  = ["#", "Name", "Total Experience", "Relevant Experience",
                    "Highlights", "Lowlights", "Score", "Decision"]
        for col, h in enumerate(headers, 1):
            cell           = ws.cell(row=1, column=col, value=h)
            cell.fill      = h_fill
            cell.font      = h_font
            cell.alignment = c_align
            cell.border    = border

        sel_fill = PatternFill("solid", fgColor="C8E6C9")
        rej_fill = PatternFill("solid", fgColor="FFCDD2")
        alt_fill = PatternFill("solid", fgColor="EEF2FF")

        for ri, r in enumerate(self.results, 2):
            decision = str(r.get("decision", "Reject")).strip().lower()
            if decision == "select":
                row_fill = sel_fill
            elif ri % 2 == 0:
                row_fill = alt_fill
            else:
                row_fill = rej_fill

            values = [
                ri - 1,
                r.get("name", ""),
                fmt_exp(r.get("total_experience")),
                fmt_exp(r.get("relevant_experience")),
                "\n".join(f"• {h}" for h in r.get("highlights", [])),
                "\n".join(f"• {l}" for l in r.get("lowlights",  [])),
                r.get("score", ""),
                r.get("decision", ""),
            ]
            for ci, val in enumerate(values, 1):
                cell           = ws.cell(row=ri, column=ci, value=val)
                cell.fill      = row_fill
                cell.border    = border
                cell.alignment = c_align if ci in (1, 3, 4, 7, 8) else l_align

        # Column widths
        for col, width in zip("ABCDEFGH", [5, 24, 16, 16, 44, 34, 8, 12]):
            ws.column_dimensions[col].width = width

        # Freeze header row
        ws.freeze_panes = "A2"

        wb.save(path)
        messagebox.showinfo("Exported", f"Excel file saved to:\n{path}")


# ─── Entry point ──────────────────────────────────────────────────────────────

def main():
    app = ResumeAnalyzerApp()
    app.mainloop()


if __name__ == "__main__":
    main()