# ============================================================
# PPTX 텍스트 추출기 - 오탈자 검토용 GUI 프로그램 (Windows 11)
# ============================================================
# [필요 패키지 설치]
# pip install python-pptx
# (tkinter는 Python 표준 라이브러리에 포함 - 별도 설치 불필요)
# ============================================================

import os
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, ttk, font
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


# ── 색상 팔레트 (Windows 11 스타일) ─────────────────────────
COLOR_BG       = "#F3F3F3"   # 앱 배경
COLOR_CARD     = "#FFFFFF"   # 카드 배경
COLOR_ACCENT   = "#0067C0"   # 강조색 (파란색)
COLOR_ACCENT_H = "#005BA4"   # 강조색 hover
COLOR_BORDER   = "#E0E0E0"   # 테두리
COLOR_TEXT     = "#1A1A1A"   # 기본 텍스트
COLOR_SUBTEXT  = "#6B6B6B"   # 보조 텍스트
COLOR_SUCCESS  = "#107C41"   # 성공 (초록)
COLOR_ERROR    = "#C42B1C"   # 오류 (빨강)
COLOR_DROPZONE = "#EBF3FB"   # 드롭존 배경


# ── 텍스트 추출 핵심 로직 ────────────────────────────────────

def extract_text_from_shape(shape):
    """단일 도형에서 텍스트 추출 (표, 일반 도형/텍스트 상자 모두 처리)"""
    lines = []

    if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        # 표: 각 행의 셀을 탭으로 구분해 한 줄로 표현
        for row in shape.table.rows:
            row_texts = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_texts:
                lines.append("\t".join(row_texts))

    elif shape.has_text_frame:
        # 일반 텍스트 상자 / 도형: 단락 단위로 추출
        for para in shape.text_frame.paragraphs:
            text = "".join(run.text for run in para.runs).strip()
            if text:
                lines.append(text)

    return lines


def extract_text_from_slide(slide):
    """슬라이드 한 장의 모든 도형 텍스트 추출 (그룹 도형 재귀 탐색)"""
    all_lines = []

    def process_shapes(shapes):
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                process_shapes(shape.shapes)  # 그룹 내부 재귀 탐색
            else:
                all_lines.extend(extract_text_from_shape(shape))

    process_shapes(slide.shapes)
    return all_lines


def extract_pptx_to_txt(pptx_path, progress_callback=None):
    """
    PPTX → TXT 변환 메인 함수.
    progress_callback(current, total, slide_num): 진행률 업데이트용 콜백
    반환값: (저장된 TXT 파일 경로, 총 슬라이드 수, 텍스트 없는 슬라이드 수)
    """
    base_name = os.path.splitext(os.path.basename(pptx_path))[0]
    output_dir = os.path.dirname(pptx_path)
    output_path = os.path.join(output_dir, f"{base_name}_오탈자검토.txt")

    prs = Presentation(pptx_path)
    total = len(prs.slides)
    empty_count = 0
    output_lines = []

    for idx, slide in enumerate(prs.slides, start=1):
        if progress_callback:
            progress_callback(idx, total, idx)

        output_lines.append("=========================================")
        output_lines.append(f"[Slide {idx}]")
        output_lines.append("=========================================")

        texts = extract_text_from_slide(slide)
        if texts:
            output_lines.extend(texts)
        else:
            output_lines.append(f"[Slide {idx}] 추출된 텍스트가 없습니다.")
            empty_count += 1

        output_lines.append("")  # 슬라이드 간 빈 줄

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))

    return output_path, total, empty_count


# ── GUI 앱 클래스 ─────────────────────────────────────────────

class PPTXExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PPTX 텍스트 추출기")
        self.geometry("680x560")
        self.minsize(580, 480)
        self.configure(bg=COLOR_BG)
        self.resizable(True, True)

        # Windows 11 스타일 아이콘 설정 시도 (없으면 무시)
        try:
            self.iconbitmap(default="")
        except Exception:
            pass

        self._selected_path = ""   # 선택된 파일 경로
        self._output_path   = ""   # 저장된 결과 파일 경로

        self._build_ui()
        self._center_window()

    def _center_window(self):
        """화면 정중앙에 창 배치"""
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── UI 구성 ───────────────────────────────────────────────

    def _build_ui(self):
        # 타이틀 헤더
        header = tk.Frame(self, bg=COLOR_ACCENT, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="📄  PPTX 텍스트 추출기",
            bg=COLOR_ACCENT, fg="white",
            font=("Segoe UI", 14, "bold"),
            padx=20,
        ).pack(side="left", fill="y")
        tk.Label(
            header,
            text="오탈자 검토용 TXT 변환",
            bg=COLOR_ACCENT, fg="#C8DFEF",
            font=("Segoe UI", 10),
            padx=4,
        ).pack(side="left", fill="y")

        # 본문 영역
        body = tk.Frame(self, bg=COLOR_BG)
        body.pack(fill="both", expand=True, padx=24, pady=20)

        # 1) 파일 선택 카드
        self._build_file_card(body)

        # 2) 진행률 카드
        self._build_progress_card(body)

        # 3) 결과 미리보기 카드
        self._build_preview_card(body)

        # 4) 하단 버튼 영역
        self._build_bottom_bar()

    def _card(self, parent, title):
        """둥근 테두리 카드 프레임 생성 헬퍼"""
        outer = tk.Frame(parent, bg=COLOR_BORDER, bd=0)
        outer.pack(fill="x", pady=(0, 12))
        inner = tk.Frame(outer, bg=COLOR_CARD, padx=16, pady=14)
        inner.pack(fill="x", padx=1, pady=1)
        if title:
            tk.Label(
                inner, text=title,
                bg=COLOR_CARD, fg=COLOR_SUBTEXT,
                font=("Segoe UI", 9, "bold"),
            ).pack(anchor="w", pady=(0, 8))
        return inner

    def _build_file_card(self, parent):
        card = self._card(parent, "파일 선택")

        # 드롭존 느낌의 파일 경로 표시 영역
        drop_frame = tk.Frame(card, bg=COLOR_DROPZONE, bd=1, relief="flat",
                              highlightbackground=COLOR_ACCENT,
                              highlightthickness=1)
        drop_frame.pack(fill="x", pady=(0, 10))

        inner = tk.Frame(drop_frame, bg=COLOR_DROPZONE, padx=12, pady=10)
        inner.pack(fill="x")

        tk.Label(inner, text="📂", bg=COLOR_DROPZONE,
                 font=("Segoe UI", 18)).pack(side="left")

        text_frame = tk.Frame(inner, bg=COLOR_DROPZONE)
        text_frame.pack(side="left", fill="x", expand=True, padx=10)

        self._lbl_hint = tk.Label(
            text_frame,
            text="PPTX 파일을 선택하세요",
            bg=COLOR_DROPZONE, fg=COLOR_SUBTEXT,
            font=("Segoe UI", 10),
            anchor="w",
        )
        self._lbl_hint.pack(anchor="w")

        self._lbl_path = tk.Label(
            text_frame,
            text="",
            bg=COLOR_DROPZONE, fg=COLOR_TEXT,
            font=("Segoe UI", 9),
            anchor="w", wraplength=420, justify="left",
        )
        self._lbl_path.pack(anchor="w")

        self._btn_select = self._accent_button(
            card, "파일 선택", self._on_select_file, width=14
        )
        self._btn_select.pack(anchor="e")

    def _build_progress_card(self, parent):
        card = self._card(parent, "진행 상태")

        self._lbl_status = tk.Label(
            card, text="대기 중...",
            bg=COLOR_CARD, fg=COLOR_SUBTEXT,
            font=("Segoe UI", 10),
        )
        self._lbl_status.pack(anchor="w", pady=(0, 6))

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor=COLOR_BORDER,
            background=COLOR_ACCENT,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_ACCENT,
            darkcolor=COLOR_ACCENT,
        )
        self._progressbar = ttk.Progressbar(
            card, style="Accent.Horizontal.TProgressbar",
            orient="horizontal", mode="determinate", length=100,
        )
        self._progressbar.pack(fill="x", ipady=3)

        self._lbl_slide_info = tk.Label(
            card, text="",
            bg=COLOR_CARD, fg=COLOR_SUBTEXT,
            font=("Segoe UI", 8),
        )
        self._lbl_slide_info.pack(anchor="e", pady=(4, 0))

    def _build_preview_card(self, parent):
        card = self._card(parent, "추출 결과 미리보기")
        card.pack_configure(expand=True, fill="both")  # 세로 확장

        text_frame = tk.Frame(card, bg=COLOR_CARD)
        text_frame.pack(fill="both", expand=True)

        self._txt_preview = tk.Text(
            text_frame,
            bg="#FAFAFA", fg=COLOR_TEXT,
            font=("Consolas", 9),
            relief="flat", bd=0,
            wrap="word",
            state="disabled",
            height=8,
        )
        scrollbar = ttk.Scrollbar(text_frame, command=self._txt_preview.yview)
        self._txt_preview.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self._txt_preview.pack(side="left", fill="both", expand=True)

        # 미리보기 텍스트 태그 (슬라이드 헤더 강조)
        self._txt_preview.tag_configure("header", foreground=COLOR_ACCENT,
                                        font=("Consolas", 9, "bold"))
        self._txt_preview.tag_configure("empty",  foreground=COLOR_SUBTEXT,
                                        font=("Consolas", 9, "italic"))

    def _build_bottom_bar(self):
        bar = tk.Frame(self, bg=COLOR_BG)
        bar.pack(fill="x", padx=24, pady=(0, 20))

        # 결과 파일 열기 버튼 (초기 비활성)
        self._btn_open = tk.Button(
            bar,
            text="📂  저장 폴더 열기",
            bg=COLOR_CARD, fg=COLOR_TEXT,
            font=("Segoe UI", 10),
            relief="flat", bd=0,
            padx=14, pady=8,
            cursor="hand2",
            state="disabled",
            activebackground=COLOR_BORDER,
            command=self._on_open_folder,
        )
        self._btn_open.pack(side="left")

        # 추출 실행 버튼
        self._btn_run = self._accent_button(
            bar, "  텍스트 추출 시작  ", self._on_run, width=18
        )
        self._btn_run.pack(side="right")
        self._btn_run.configure(state="disabled")

    # ── 이벤트 핸들러 ─────────────────────────────────────────

    def _on_select_file(self):
        """파일 선택 버튼 클릭"""
        path = filedialog.askopenfilename(
            title="PPTX 파일 선택",
            filetypes=[("PowerPoint 파일", "*.pptx"), ("모든 파일", "*.*")],
        )
        if not path:
            return

        self._selected_path = path
        self._output_path   = ""

        # UI 업데이트
        self._lbl_hint.configure(text=os.path.basename(path), fg=COLOR_TEXT,
                                 font=("Segoe UI", 10, "bold"))
        self._lbl_path.configure(text=os.path.dirname(path))
        self._lbl_status.configure(text="파일이 선택되었습니다. 추출을 시작하세요.",
                                   fg=COLOR_SUBTEXT)
        self._progressbar["value"] = 0
        self._lbl_slide_info.configure(text="")
        self._btn_run.configure(state="normal")
        self._btn_open.configure(state="disabled")
        self._clear_preview()

    def _on_run(self):
        """추출 시작 버튼 클릭 → 백그라운드 스레드로 실행 (UI 블로킹 방지)"""
        if not self._selected_path:
            return

        self._btn_run.configure(state="disabled")
        self._btn_select.configure(state="disabled")
        self._btn_open.configure(state="disabled")
        self._clear_preview()
        self._lbl_status.configure(text="추출 중...", fg=COLOR_ACCENT)
        self._progressbar["value"] = 0

        thread = threading.Thread(target=self._run_extraction, daemon=True)
        thread.start()

    def _run_extraction(self):
        """백그라운드 스레드: PPTX 추출 실행"""
        try:
            output_path, total, empty_count = extract_pptx_to_txt(
                self._selected_path,
                progress_callback=self._on_progress,
            )
            # 완료 후 UI 업데이트는 메인 스레드에서 실행해야 함
            self.after(0, self._on_done, output_path, total, empty_count)
        except Exception as e:
            self.after(0, self._on_error, str(e))

    def _on_progress(self, current, total, slide_num):
        """추출 진행률 콜백 → 메인 스레드에 UI 업데이트 요청"""
        pct = int(current / total * 100)
        self.after(0, self._update_progress, pct, current, total)

    def _update_progress(self, pct, current, total):
        self._progressbar["value"] = pct
        self._lbl_slide_info.configure(
            text=f"Slide {current} / {total}  ({pct}%)"
        )

    def _on_done(self, output_path, total, empty_count):
        """추출 완료 처리"""
        self._output_path = output_path
        self._progressbar["value"] = 100

        # 결과 통계 표시
        status_text = (
            f"✅  완료 — 총 {total}장 슬라이드, "
            f"텍스트 없음 {empty_count}장"
        )
        self._lbl_status.configure(text=status_text, fg=COLOR_SUCCESS)
        self._lbl_slide_info.configure(text=f"저장: {os.path.basename(output_path)}")

        # 미리보기 로드
        self._load_preview(output_path)

        # 버튼 복원
        self._btn_run.configure(state="normal")
        self._btn_select.configure(state="normal")
        self._btn_open.configure(state="normal")

    def _on_error(self, msg):
        """오류 처리"""
        self._lbl_status.configure(text=f"❌  오류: {msg}", fg=COLOR_ERROR)
        self._progressbar["value"] = 0
        self._btn_run.configure(state="normal")
        self._btn_select.configure(state="normal")

    def _on_open_folder(self):
        """결과 파일이 있는 폴더를 탐색기로 열기"""
        folder = os.path.dirname(self._output_path)
        if os.path.isdir(folder):
            subprocess.Popen(["explorer", folder])

    # ── 미리보기 ───────────────────────────────────────────────

    def _clear_preview(self):
        self._txt_preview.configure(state="normal")
        self._txt_preview.delete("1.0", "end")
        self._txt_preview.configure(state="disabled")

    def _load_preview(self, path):
        """TXT 파일 내용을 미리보기 영역에 색상 구분하여 표시"""
        self._txt_preview.configure(state="normal")
        self._txt_preview.delete("1.0", "end")

        with open(path, encoding="utf-8") as f:
            for line in f:
                line_stripped = line.rstrip("\n")
                if line_stripped.startswith("[Slide") and line_stripped.endswith("]"):
                    self._txt_preview.insert("end", line_stripped + "\n", "header")
                elif "추출된 텍스트가 없습니다" in line_stripped:
                    self._txt_preview.insert("end", line_stripped + "\n", "empty")
                else:
                    self._txt_preview.insert("end", line_stripped + "\n")

        self._txt_preview.configure(state="disabled")
        self._txt_preview.see("1.0")

    # ── 헬퍼: 강조 버튼 ───────────────────────────────────────

    def _accent_button(self, parent, text, command, width=None):
        btn = tk.Button(
            parent,
            text=text,
            bg=COLOR_ACCENT, fg="white",
            activebackground=COLOR_ACCENT_H, activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat", bd=0,
            padx=14, pady=8,
            cursor="hand2",
            command=command,
        )
        if width:
            btn.configure(width=width)

        # hover 효과
        btn.bind("<Enter>", lambda e: btn.configure(bg=COLOR_ACCENT_H))
        btn.bind("<Leave>", lambda e: btn.configure(
            bg=COLOR_ACCENT if btn["state"] != "disabled" else COLOR_BORDER
        ))
        return btn


# ── 진입점 ────────────────────────────────────────────────────

if __name__ == "__main__":
    # EXE(--windowed) 빌드 시 콘솔 창이 뜨는 것을 방지
    import sys
    if sys.platform == "win32":
        import ctypes
        # 이미 콘솔이 없는 경우(windowed 빌드)에만 무시됨
        try:
            ctypes.windll.kernel32.FreeConsole()
        except Exception:
            pass

    app = PPTXExtractorApp()
    app.mainloop()
