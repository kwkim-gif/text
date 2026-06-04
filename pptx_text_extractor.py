# ============================================================
# PPTX Utility - 텍스트 추출 + 보안 PDF 변환 (Windows 11 GUI)
# ============================================================
# pip install python-pptx tkinterdnd2 pywin32 Pillow
# ============================================================

import os
import sys
import queue
import threading
import tempfile
import shutil
import subprocess
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


# ── 색상 팔레트 ──────────────────────────────────────────────
COLOR_BG        = "#F3F3F3"
COLOR_CARD      = "#FFFFFF"
COLOR_ACCENT    = "#0067C0"
COLOR_ACCENT_H  = "#005BA4"
COLOR_BORDER    = "#E0E0E0"
COLOR_TEXT      = "#1A1A1A"
COLOR_SUBTEXT   = "#6B6B6B"
COLOR_SUCCESS   = "#107C41"
COLOR_ERROR     = "#C42B1C"
COLOR_WARN      = "#CA5010"
COLOR_DROPZONE  = "#EBF3FB"
COLOR_DROP_ACT  = "#D0E8F8"
COLOR_ROW_ODD   = "#FFFFFF"
COLOR_ROW_EVEN  = "#F7F9FC"
COLOR_ROW_DONE  = "#EAF5EE"
COLOR_ROW_ERR   = "#FDF0EE"
COLOR_ROW_SEL   = "#D6E8FB"
COLOR_TAB_ACT   = "#1A1A1A"   # 활성 탭 배경
COLOR_TAB_INACT = "#F3F3F3"   # 비활성 탭 배경

ST_WAITING = "대기 중"
ST_RUNNING = "처리 중..."
ST_DONE    = "완료"
ST_ERROR   = "오류"


# ── 텍스트 추출 로직 ─────────────────────────────────────────

def _extract_shape_text(shape):
    lines = []
    if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        for row in shape.table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                lines.append("\t".join(cells))
    elif shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            text = "".join(r.text for r in para.runs).strip()
            if text:
                lines.append(text)
    return lines


def _extract_slide_text(slide):
    all_lines = []
    def process(shapes):
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                process(shape.shapes)
            else:
                all_lines.extend(_extract_shape_text(shape))
    process(slide.shapes)
    return all_lines


def extract_pptx_to_txt(pptx_path, progress_callback=None):
    base      = os.path.splitext(os.path.basename(pptx_path))[0]
    out_path  = os.path.join(os.path.dirname(pptx_path), f"{base}_오탈자검토.txt")
    prs       = Presentation(pptx_path)
    total     = len(prs.slides)
    empty_cnt = 0
    lines     = []

    for idx, slide in enumerate(prs.slides, start=1):
        if progress_callback:
            progress_callback(idx, total)
        lines += ["=========================================",
                  f"[Slide {idx}]",
                  "========================================="]
        texts = _extract_slide_text(slide)
        if texts:
            lines.extend(texts)
        else:
            lines.append(f"[Slide {idx}] 추출된 텍스트가 없습니다.")
            empty_cnt += 1
        lines.append("")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return out_path, total, empty_cnt


# ── PPTX → 이미지 → PDF 변환 로직 ───────────────────────────

def _make_ascii_tempdir():
    """
    COM Export 는 경로에 한글 등 비ASCII 문자가 있으면 실패.
    %TEMP% 가 한글 사용자명 폴더일 수 있으므로,
    ASCII 경로만 후보로 순서대로 시도하여 임시 폴더를 생성.
    """
    candidates = [
        r"C:\Temp",
        r"C:\Windows\Temp",
        r"C:\ProgramData\pptx_tmp",
        os.path.join(os.environ.get("SYSTEMROOT", r"C:\Windows"), "Temp"),
    ]
    for base in candidates:
        try:
            base.encode("ascii")          # ASCII 경로인지 확인
            os.makedirs(base, exist_ok=True)
            td = tempfile.mkdtemp(prefix="pptxpdf_", dir=base)
            return td
        except (UnicodeEncodeError, OSError):
            continue

    # 최후 수단: 기본 임시 폴더 (경고만 하고 진행)
    return tempfile.mkdtemp(prefix="pptxpdf_")


def convert_pptx_to_pdf(pptx_path, dpi, progress_callback=None):
    """
    PowerShell COM 으로 PPTX → 중간 PDF → 이미지 → 이미지 전용 PDF.

    win32com 은 PyInstaller --onefile 환경에서 CoInitialize/등록 문제로
    불안정. Windows 내장 PowerShell 로 COM 을 대신 호출해 우회.
    """
    import fitz
    from PIL import Image as PILImage

    base     = os.path.splitext(os.path.basename(pptx_path))[0]
    out_path = os.path.join(os.path.dirname(pptx_path), f"{base}_보안변환.pdf")

    temp_dir  = _make_ascii_tempdir()
    temp_pptx = os.path.join(temp_dir, "input.pptx")
    temp_pdf  = os.path.join(temp_dir, "intermediate.pdf")
    shutil.copy2(pptx_path, temp_pptx)

    try:
        # ── Step 1: PowerShell .ps1 파일로 COM 자동화
        # -Command 인라인 방식은 따옴표 이스케이프 문제로 불안정 →
        # .ps1 파일에 저장 후 -File 로 실행하는 방식이 안정적
        ps_file = os.path.join(temp_dir, "convert.ps1")
        ps_content = f"""$ErrorActionPreference = 'Stop'

$ppt = $null
$createdNew = $false
try {{
    $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
}} catch {{
    $ppt = New-Object -ComObject PowerPoint.Application
    $createdNew = $true
}}

$ppt.Visible = $true
try {{
    $prs = $ppt.Presentations.Open('{temp_pptx}', $true, $false, $false)
    # ppSaveAsPDF = 32  (ExportAsFixedFormat 보다 버전 호환성 높음)
    $prs.SaveAs('{temp_pdf}', 32)
    $prs.Close()
}} finally {{
    if ($createdNew) {{ $ppt.Quit() }}
}}
"""
        # UTF-8 BOM 포함 저장 (PowerShell이 한글 경로 읽을 때 필요)
        with open(ps_file, "w", encoding="utf-8-sig") as f:
            f.write(ps_content)

        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-File", ps_file],
            capture_output=True, text=True, timeout=300,
            creationflags=0x08000000,  # CREATE_NO_WINDOW
        )

        if result.returncode != 0 or not os.path.exists(temp_pdf):
            detail = (result.stderr or result.stdout or "no output").strip()
            raise Exception(detail[:120])

        # ── Step 2: PDF 페이지 → 이미지 (텍스트 레이어 제거, 무손실 PNG)
        zoom  = dpi / 72
        doc   = fitz.open(temp_pdf)
        total = len(doc)
        img_paths = []

        for i, page in enumerate(doc):
            if progress_callback:
                progress_callback(i + 1, total)
            pix      = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            img_file = os.path.join(temp_dir, f"slide_{i+1:04d}.png")
            pix.save(img_file)
            img_paths.append(img_file)
        doc.close()

        # ── Step 3: 이미지 → 이미지 전용 PDF
        images = [PILImage.open(p).convert("RGB") for p in img_paths]
        if images:
            images[0].save(
                out_path, save_all=True,
                append_images=images[1:], resolution=dpi)
        for img in images:
            img.close()

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    return out_path, total


# ── 공용: 파일 항목 데이터 ───────────────────────────────────

class FileItem:
    def __init__(self, path):
        self.path    = path
        self.name    = os.path.basename(path)
        self.status  = ST_WAITING
        self.slides  = ""
        self.message = ""
        self.output  = ""


# ── 공용: 파일 목록 테이블 위젯 ──────────────────────────────

class FileListWidget(tk.Frame):
    COLS = [
        ("번호",    40,  "center"),
        ("파일명",  270, "w"),
        ("슬라이드", 70,  "center"),
        ("상태",    90,  "center"),
        ("결과",    130, "w"),
    ]

    def __init__(self, parent, on_select=None, **kw):
        super().__init__(parent, bg=COLOR_CARD, **kw)
        self._on_select = on_select
        self._items: list[FileItem] = []
        self._selected_idx = -1
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=COLOR_ACCENT)
        hdr.pack(fill="x")
        for name, width, anchor in self.COLS:
            tk.Label(hdr, text=name, bg=COLOR_ACCENT, fg="white",
                     font=("Segoe UI", 9, "bold"),
                     width=width // 7, anchor=anchor,
                     padx=6, pady=6).pack(side="left")

        wrap = tk.Frame(self, bg=COLOR_CARD)
        wrap.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(wrap, bg=COLOR_CARD,
                                 highlightthickness=0, bd=0)
        sb = ttk.Scrollbar(wrap, orient="vertical",
                           command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._lf = tk.Frame(self._canvas, bg=COLOR_CARD)
        self._win = self._canvas.create_window((0, 0), window=self._lf, anchor="nw")
        self._lf.bind("<Configure>",
                      lambda e: self._canvas.configure(
                          scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
                          lambda e: self._canvas.itemconfig(self._win, width=e.width))
        self._canvas.bind("<MouseWheel>",
                          lambda e: self._canvas.yview_scroll(
                              -1 * (e.delta // 120), "units"))
        self._render_empty()

    def _render_empty(self):
        for w in self._lf.winfo_children():
            w.destroy()
        tk.Label(self._lf,
                 text="PPTX 파일을 드래그하거나 [파일 추가] 버튼으로 추가하세요",
                 bg=COLOR_CARD, fg=COLOR_SUBTEXT,
                 font=("Segoe UI", 10), pady=30).pack(fill="x")

    def refresh(self):
        for w in self._lf.winfo_children():
            w.destroy()
        if not self._items:
            self._render_empty()
            return

        ST_COLOR = {ST_WAITING: COLOR_SUBTEXT, ST_RUNNING: COLOR_ACCENT,
                    ST_DONE: COLOR_SUCCESS, ST_ERROR: COLOR_ERROR}

        for i, item in enumerate(self._items):
            bg = COLOR_ROW_ODD if i % 2 == 0 else COLOR_ROW_EVEN
            if item.status == ST_DONE:  bg = COLOR_ROW_DONE
            if item.status == ST_ERROR: bg = COLOR_ROW_ERR
            if i == self._selected_idx: bg = COLOR_ROW_SEL

            row = tk.Frame(self._lf, bg=bg, cursor="hand2")
            row.pack(fill="x")
            st_color = ST_COLOR.get(item.status, COLOR_TEXT)

            for j, (val, width, anchor) in enumerate([
                (str(i + 1),   40,  "center"),
                (item.name,    270, "w"),
                (item.slides,  70,  "center"),
                (item.status,  90,  "center"),
                (item.message, 130, "w"),
            ]):
                color = st_color if j == 3 else COLOR_TEXT
                tk.Label(row, text=val, bg=bg, fg=color,
                         font=("Segoe UI", 9),
                         width=width // 7, anchor=anchor,
                         padx=6, pady=7).pack(side="left")

            for w in [row] + row.winfo_children():
                w.bind("<Button-1>", lambda e, n=i: self._click(n))

        tk.Frame(self._lf, bg=COLOR_BORDER, height=1).pack(fill="x")

    def _click(self, idx):
        self._selected_idx = idx
        self.refresh()
        if self._on_select:
            self._on_select(self._items[idx])

    def add_paths(self, paths):
        existing = {it.path for it in self._items}
        added = 0
        for p in paths:
            p = p.strip().strip("{}")
            if p.lower().endswith(".pptx") and p not in existing:
                self._items.append(FileItem(p))
                existing.add(p)
                added += 1
        if added:
            self.refresh()
        return added

    def remove_selected(self):
        if 0 <= self._selected_idx < len(self._items):
            self._items.pop(self._selected_idx)
            self._selected_idx = min(self._selected_idx, len(self._items) - 1)
            self.refresh()

    def clear_all(self):
        self._items.clear()
        self._selected_idx = -1
        self.refresh()

    def update_item(self, idx, **kw):
        if 0 <= idx < len(self._items):
            for k, v in kw.items():
                setattr(self._items[idx], k, v)
            self.refresh()

    @property
    def items(self):
        return self._items


# ── 공용: 탭 기본 프레임 ─────────────────────────────────────

class BaseTabFrame(tk.Frame):
    """두 탭이 공유하는 드롭존 / 파일목록 / 진행바 / 버튼 뼈대"""

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=COLOR_BG, **kw)
        self._last_output_folder = ""
        self._is_running         = False
        self._queue              = queue.Queue()

    # ── 공통 UI 빌더 헬퍼 ────────────────────────────────────

    def _card(self, parent, title, expand=False):
        outer = tk.Frame(parent, bg=COLOR_BORDER)
        outer.pack(fill="both" if expand else "x",
                   expand=expand, pady=(0, 10))
        inner = tk.Frame(outer, bg=COLOR_CARD, padx=14, pady=12)
        inner.pack(fill="both", expand=expand, padx=1, pady=1)
        if title:
            tk.Label(inner, text=title, bg=COLOR_CARD, fg=COLOR_SUBTEXT,
                     font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 8))
        return inner

    def _accent_btn(self, parent, text, cmd, big=False, **kw):
        btn = tk.Button(parent, text=text,
                        bg=COLOR_ACCENT, fg="white",
                        activebackground=COLOR_ACCENT_H,
                        activeforeground="white",
                        font=("Segoe UI", 10 if not big else 11, "bold"),
                        relief="flat", bd=0,
                        padx=12, pady=6 if not big else 9,
                        cursor="hand2", command=cmd, **kw)
        btn.bind("<Enter>", lambda e: btn.configure(
            bg=COLOR_ACCENT_H if str(btn["state"]) != "disabled" else COLOR_BORDER))
        btn.bind("<Leave>", lambda e: btn.configure(
            bg=COLOR_ACCENT if str(btn["state"]) != "disabled" else COLOR_BORDER))
        return btn

    def _ghost_btn(self, parent, text, cmd, **kw):
        btn = tk.Button(parent, text=text,
                        bg=COLOR_CARD, fg=COLOR_SUBTEXT,
                        activebackground=COLOR_BORDER,
                        activeforeground=COLOR_TEXT,
                        font=("Segoe UI", 9),
                        relief="flat", bd=0,
                        padx=10, pady=6,
                        cursor="hand2", command=cmd, **kw)
        btn.bind("<Enter>", lambda e: btn.configure(bg=COLOR_BORDER, fg=COLOR_TEXT))
        btn.bind("<Leave>", lambda e: btn.configure(bg=COLOR_CARD,   fg=COLOR_SUBTEXT))
        return btn

    def _build_drop_zone(self, parent, hint_text):
        outer = tk.Frame(parent, bg=COLOR_BORDER)
        outer.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(outer, bg=COLOR_DROPZONE)
        inner.pack(fill="x", padx=1, pady=1)
        self._drop_zone = tk.Frame(inner, bg=COLOR_DROPZONE, pady=16)
        self._drop_zone.pack(fill="x")
        tk.Label(self._drop_zone, text=hint_text,
                 bg=COLOR_DROPZONE, fg=COLOR_ACCENT,
                 font=("Segoe UI", 11, "bold")).pack()
        tk.Label(self._drop_zone,
                 text="여러 파일을 한 번에 드롭하거나, 아래 [파일 추가] 버튼으로 선택할 수 있습니다",
                 bg=COLOR_DROPZONE, fg=COLOR_SUBTEXT,
                 font=("Segoe UI", 9)).pack(pady=(4, 0))
        if not DND_AVAILABLE:
            tk.Label(self._drop_zone, text="⚠  tkinterdnd2 미설치 — 드래그 앤 드롭 비활성",
                     bg=COLOR_DROPZONE, fg=COLOR_WARN,
                     font=("Segoe UI", 8)).pack(pady=(4, 0))
        self._drop_outer = outer
        if DND_AVAILABLE:
            for w in [outer, inner, self._drop_zone] + self._drop_zone.winfo_children():
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>",      self._on_drop)
                    w.dnd_bind("<<DragEnter>>", self._on_drag_enter)
                    w.dnd_bind("<<DragLeave>>", self._on_drag_leave)
                except Exception:
                    pass
        return outer

    def _build_file_list_area(self, parent):
        outer = tk.Frame(parent, bg=COLOR_BORDER)
        outer.pack(fill="both", expand=True, pady=(0, 10))
        inner = tk.Frame(outer, bg=COLOR_CARD)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        toolbar = tk.Frame(inner, bg=COLOR_CARD, pady=8, padx=10)
        toolbar.pack(fill="x")
        tk.Label(toolbar, text="파일 목록", bg=COLOR_CARD, fg=COLOR_SUBTEXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        self._lbl_count = tk.Label(toolbar, text="0개",
                                   bg=COLOR_CARD, fg=COLOR_ACCENT,
                                   font=("Segoe UI", 9, "bold"))
        self._lbl_count.pack(side="left", padx=(6, 0))

        self._btn_clear  = self._ghost_btn(toolbar, "전체 삭제", self._on_clear_all)
        self._btn_clear.pack(side="right", padx=(6, 0))
        self._btn_remove = self._ghost_btn(toolbar, "선택 삭제", self._on_remove)
        self._btn_remove.pack(side="right", padx=(6, 0))
        self._btn_add    = self._accent_btn(toolbar, "+ 파일 추가", self._on_add_files)
        self._btn_add.pack(side="right", padx=(6, 0))

        tk.Frame(inner, bg=COLOR_BORDER, height=1).pack(fill="x")

        self._file_list = FileListWidget(inner, height=180)
        self._file_list.pack(fill="both", expand=True)

        if DND_AVAILABLE:
            try:
                self._file_list.drop_target_register(DND_FILES)
                self._file_list.dnd_bind("<<Drop>>",      self._on_drop)
                self._file_list.dnd_bind("<<DragEnter>>", self._on_drag_enter)
                self._file_list.dnd_bind("<<DragLeave>>", self._on_drag_leave)
            except Exception:
                pass

    def _build_progress_area(self, parent):
        outer = tk.Frame(parent, bg=COLOR_BORDER)
        outer.pack(fill="x", pady=(0, 10))
        card = tk.Frame(outer, bg=COLOR_CARD, padx=14, pady=12)
        card.pack(fill="x", padx=1, pady=1)

        top = tk.Frame(card, bg=COLOR_CARD)
        top.pack(fill="x")
        self._lbl_status = tk.Label(top, text="파일을 추가하고 시작하세요.",
                                    bg=COLOR_CARD, fg=COLOR_SUBTEXT,
                                    font=("Segoe UI", 10))
        self._lbl_status.pack(side="left")
        self._lbl_prog_text = tk.Label(top, text="", bg=COLOR_CARD,
                                       fg=COLOR_SUBTEXT, font=("Segoe UI", 9))
        self._lbl_prog_text.pack(side="right")

        style = ttk.Style()
        style.configure("Accent.Horizontal.TProgressbar",
                        troughcolor=COLOR_BORDER, background=COLOR_ACCENT,
                        bordercolor=COLOR_BORDER,
                        lightcolor=COLOR_ACCENT, darkcolor=COLOR_ACCENT)
        self._progressbar = ttk.Progressbar(
            card, style="Accent.Horizontal.TProgressbar",
            orient="horizontal", mode="determinate")
        self._progressbar.pack(fill="x", ipady=3, pady=(8, 0))

    def _build_bottom_bar(self, parent, run_text, run_cmd):
        bar = tk.Frame(parent, bg=COLOR_BG)
        bar.pack(fill="x", pady=(0, 16))

        self._btn_open_folder = tk.Button(
            bar, text="📂  저장 폴더 열기",
            bg=COLOR_CARD, fg=COLOR_TEXT,
            font=("Segoe UI", 10), relief="flat", bd=0,
            padx=14, pady=8, cursor="hand2", state="disabled",
            activebackground=COLOR_BORDER,
            command=self._on_open_folder)
        self._btn_open_folder.pack(side="left")

        self._btn_run = self._accent_btn(bar, run_text, run_cmd, big=True)
        self._btn_run.pack(side="right")
        self._btn_run.configure(state="disabled")

    # ── 공통 이벤트 ──────────────────────────────────────────

    def _on_drag_enter(self, event):
        self._drop_zone.configure(bg=COLOR_DROP_ACT)
        self._drop_outer.configure(bg=COLOR_ACCENT)

    def _on_drag_leave(self, event):
        self._drop_zone.configure(bg=COLOR_DROPZONE)
        self._drop_outer.configure(bg=COLOR_BORDER)

    def _on_drop(self, event):
        self._drop_zone.configure(bg=COLOR_DROPZONE)
        self._drop_outer.configure(bg=COLOR_BORDER)
        paths = self._parse_drop(event.data)
        added = self._file_list.add_paths(paths)
        self._update_count()
        if added:
            self._sync_run_btn()

    def _parse_drop(self, raw):
        paths, raw, i = [], raw.strip(), 0
        while i < len(raw):
            if raw[i] == "{":
                end = raw.find("}", i)
                if end != -1:
                    paths.append(raw[i+1:end])
                    i = end + 1
                else:
                    break
            else:
                end = raw.find(" ", i)
                if end == -1:
                    paths.append(raw[i:])
                    break
                paths.append(raw[i:end])
                i = end + 1
        return [p.strip() for p in paths if p.strip()]

    def _on_add_files(self):
        paths = filedialog.askopenfilenames(
            title="PPTX 파일 선택 (다중 선택 가능)",
            filetypes=[("PowerPoint 파일", "*.pptx"), ("모든 파일", "*.*")])
        if paths:
            self._file_list.add_paths(list(paths))
            self._update_count()
            self._sync_run_btn()

    def _on_remove(self):
        self._file_list.remove_selected()
        self._update_count()
        self._sync_run_btn()

    def _on_clear_all(self):
        self._file_list.clear_all()
        self._update_count()
        self._sync_run_btn()
        self._lbl_status.configure(text="파일을 추가하고 시작하세요.", fg=COLOR_SUBTEXT)
        self._progressbar["value"] = 0
        self._lbl_prog_text.configure(text="")
        self._btn_open_folder.configure(state="disabled")

    def _on_open_folder(self):
        if self._last_output_folder and os.path.isdir(self._last_output_folder):
            subprocess.Popen(["explorer", self._last_output_folder])

    def _update_count(self):
        self._lbl_count.configure(text=f"{len(self._file_list.items)}개")

    def _sync_run_btn(self):
        state = "normal" if self._file_list.items and not self._is_running else "disabled"
        self._btn_run.configure(state=state)

    def _start_run(self, thread_target, targets):
        self._is_running = True
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        self._btn_run.configure(state="disabled", text="  처리 중...  ")
        self._btn_add.configure(state="disabled")
        self._btn_remove.configure(state="disabled")
        self._btn_clear.configure(state="disabled")
        self._btn_open_folder.configure(state="disabled")
        self._progressbar["value"] = 0
        self._last_output_folder = ""
        threading.Thread(target=thread_target, args=(targets,), daemon=True).start()
        self.after(50, self._poll_queue)

    def _poll_queue(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                if self._handle_msg(msg):
                    return   # all_done → 폴링 중단
        except queue.Empty:
            pass
        if self._is_running:
            self.after(50, self._poll_queue)

    def _handle_msg(self, msg):
        """메시지 처리. all_done 이면 True 반환."""
        kind = msg[0]
        if kind == "running":
            _, idx, jn, tf, name = msg
            self._file_list.update_item(idx, status=ST_RUNNING, message="")
            self._lbl_status.configure(
                text=f"[{jn}/{tf}]  {name}", fg=COLOR_ACCENT)

        elif kind == "progress":
            _, pct, cur, total, jn, tf = msg
            self._progressbar["value"] = pct
            self._lbl_prog_text.configure(
                text=f"파일 {jn}/{tf}  |  Slide {cur}/{total}  ({pct}%)")

        elif kind == "done":
            _, idx, slides, out = msg
            self._last_output_folder = os.path.dirname(out)
            self._file_list.update_item(
                idx, status=ST_DONE, slides=str(slides),
                message="저장 완료", output=out)

        elif kind == "error":
            _, idx, err = msg
            self._file_list.update_item(idx, status=ST_ERROR, message=err)

        elif kind == "all_done":
            _, tf = msg
            self._finish(tf)
            return True

        return False

    def _finish(self, total_files):
        self._is_running = False
        items = self._file_list.items
        done  = sum(1 for it in items if it.status == ST_DONE)
        err   = sum(1 for it in items if it.status == ST_ERROR)
        if err == 0:
            self._lbl_status.configure(
                text=f"✅  전체 완료 — {done}개 파일 처리 성공", fg=COLOR_SUCCESS)
        else:
            self._lbl_status.configure(
                text=f"⚠  완료 — 성공 {done}개 / 오류 {err}개", fg=COLOR_WARN)
        self._progressbar["value"] = 100
        self._lbl_prog_text.configure(text="")
        self._btn_run.configure(state="normal", text=self._run_label)
        self._btn_add.configure(state="normal")
        self._btn_remove.configure(state="normal")
        self._btn_clear.configure(state="normal")
        if self._last_output_folder:
            self._btn_open_folder.configure(state="normal")


# ── Tab 1: 텍스트 추출 ───────────────────────────────────────

class TextExtractTab(BaseTabFrame):
    _run_label = "  텍스트 추출 시작  "

    def __init__(self, parent):
        super().__init__(parent)
        self.pack(fill="both", expand=True)
        body = tk.Frame(self, bg=COLOR_BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        self._build_drop_zone(body, "여기에 PPTX 파일을 드래그 앤 드롭하세요")
        self._build_file_list_area(body)
        self._build_progress_area(body)
        self._build_bottom_bar(body, self._run_label, self._on_run)

    def _on_run(self):
        items = self._file_list.items
        if not items or self._is_running:
            return
        targets = [i for i, it in enumerate(items)
                   if it.status in (ST_WAITING, ST_ERROR)]
        if not targets:
            for it in items:
                it.status = ST_WAITING; it.slides = ""; it.message = ""
            targets = list(range(len(items)))
            self._file_list.refresh()
        self._start_run(self._thread, targets)

    def _thread(self, target_indices):
        items      = self._file_list.items
        total_files = len(target_indices)
        try:
            for jn, idx in enumerate(target_indices, start=1):
                item = items[idx]
                self._queue.put(("running", idx, jn, total_files, item.name))

                def make_cb(j, tf):
                    def cb(cur, tot):
                        pct = int(((j - 1) + cur / tot) / tf * 100)
                        self._queue.put(("progress", pct, cur, tot, j, tf))
                    return cb

                try:
                    out, slides, _ = extract_pptx_to_txt(
                        item.path, progress_callback=make_cb(jn, total_files))
                    self._queue.put(("done", idx, slides, out))
                except Exception as e:
                    self._queue.put(("error", idx, str(e)[:60]))
        except Exception as e:
            self._queue.put(("error", 0, f"Thread error: {e}"))
        self._queue.put(("all_done", total_files))


# ── Tab 2: PDF 변환 ──────────────────────────────────────────

class PDFConvertTab(BaseTabFrame):
    _run_label = "  PDF 변환 시작  "

    def __init__(self, parent):
        super().__init__(parent)
        self.pack(fill="both", expand=True)
        body = tk.Frame(self, bg=COLOR_BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        self._build_drop_zone(body, "여기에 PPTX 파일을 드래그 앤 드롭하세요")
        self._build_dpi_selector(body)
        self._build_file_list_area(body)
        self._build_progress_area(body)
        self._build_bottom_bar(body, self._run_label, self._on_run)

    def _build_dpi_selector(self, parent):
        card = self._card(parent, "이미지 해상도 (DPI) 선택")

        row = tk.Frame(card, bg=COLOR_CARD)
        row.pack(anchor="w")

        self._dpi_var = tk.IntVar(value=220)

        for dpi, label in [(96,  "96 dpi  — 빠른 변환 / 저용량"),
                           (220, "220 dpi — 권장 (균형)"),
                           (300, "300 dpi — 고품질 / 대용량")]:
            rb = tk.Radiobutton(
                row, text=label, variable=self._dpi_var, value=dpi,
                bg=COLOR_CARD, fg=COLOR_TEXT,
                activebackground=COLOR_CARD,
                font=("Segoe UI", 10),
                selectcolor=COLOR_CARD,
                cursor="hand2")
            rb.pack(side="left", padx=(0, 24))

        tk.Label(card,
                 text="⚠  변환에는 PC에 Microsoft PowerPoint가 설치되어 있어야 합니다.",
                 bg=COLOR_CARD, fg=COLOR_WARN,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 0))

    def _on_run(self):
        items = self._file_list.items
        if not items or self._is_running:
            return
        targets = [i for i, it in enumerate(items)
                   if it.status in (ST_WAITING, ST_ERROR)]
        if not targets:
            for it in items:
                it.status = ST_WAITING; it.slides = ""; it.message = ""
            targets = list(range(len(items)))
            self._file_list.refresh()
        self._start_run(self._thread, targets)

    def _thread(self, target_indices):
        items       = self._file_list.items
        total_files = len(target_indices)
        dpi         = self._dpi_var.get()
        try:
            for jn, idx in enumerate(target_indices, start=1):
                item = items[idx]
                self._queue.put(("running", idx, jn, total_files, item.name))

                def make_cb(j, tf):
                    def cb(cur, tot):
                        pct = int(((j - 1) + cur / tot) / tf * 100)
                        self._queue.put(("progress", pct, cur, tot, j, tf))
                    return cb

                try:
                    out, slides = convert_pptx_to_pdf(
                        item.path, dpi,
                        progress_callback=make_cb(jn, total_files))
                    self._queue.put(("done", idx, slides, out))
                except Exception as e:
                    self._queue.put(("error", idx, str(e)[:60]))
        except Exception as e:
            self._queue.put(("error", 0, f"Thread error: {e}"))
        self._queue.put(("all_done", total_files))


# ── 메인 앱 ──────────────────────────────────────────────────

class PPTXUtilityApp(TkinterDnD.Tk if DND_AVAILABLE else tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("PPTX Utility")
        self.geometry("780x700")
        self.minsize(660, 580)
        self.configure(bg=COLOR_BG)
        self.resizable(True, True)
        try:
            self.iconbitmap(default="")
        except Exception:
            pass

        self._active_tab = None
        self._tab_frames = {}
        self._tab_btns   = {}

        self._build_tab_bar()
        self._build_tabs()
        self._switch_tab("text")
        self._center_window()

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_tab_bar(self):
        bar = tk.Frame(self, bg=COLOR_TAB_ACT, height=48)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        self._tab_btn_bar = bar

        for key, label in [("text", "텍스트 추출"), ("pdf", "PDF 저장")]:
            btn = tk.Button(
                bar, text=label,
                font=("Segoe UI", 11, "bold"),
                relief="flat", bd=0,
                padx=28, pady=0,
                cursor="hand2",
                command=lambda k=key: self._switch_tab(k))
            btn.pack(side="left", fill="y")
            self._tab_btns[key] = btn

    def _build_tabs(self):
        container = tk.Frame(self, bg=COLOR_BG)
        container.pack(fill="both", expand=True)

        self._tab_frames["text"] = TextExtractTab(container)
        self._tab_frames["pdf"]  = PDFConvertTab(container)

        for f in self._tab_frames.values():
            f.place(relx=0, rely=0, relwidth=1, relheight=1)

    def _switch_tab(self, key):
        if self._active_tab == key:
            return
        self._active_tab = key
        for k, frame in self._tab_frames.items():
            if k == key:
                frame.lift()
            else:
                frame.lower()
        for k, btn in self._tab_btns.items():
            if k == key:
                btn.configure(bg=COLOR_CARD, fg=COLOR_TAB_ACT,
                              highlightbackground=COLOR_ACCENT,
                              highlightthickness=2)
            else:
                btn.configure(bg=COLOR_TAB_ACT, fg="#AAAAAA",
                              highlightthickness=0)


# ── 진입점 ────────────────────────────────────────────────────

if __name__ == "__main__":
    if sys.platform == "win32":
        import ctypes
        try:
            ctypes.windll.kernel32.FreeConsole()
        except Exception:
            pass

    app = PPTXUtilityApp()
    app.mainloop()
