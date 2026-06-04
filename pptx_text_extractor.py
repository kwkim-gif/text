# ============================================================
# PPTX 텍스트 추출기 - 오탈자 검토용 TXT 파일 생성 프로그램
# ============================================================
# [필요 패키지 설치]
# pip install python-pptx
# (tkinter는 Python 표준 라이브러리에 포함되어 있으므로 별도 설치 불필요)
# ============================================================

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE


def extract_text_from_shape(shape):
    """
    단일 도형(Shape)에서 텍스트를 추출하는 함수.
    - 일반 텍스트 상자 및 도형 내부 텍스트를 처리
    - 표(Table) 내부 텍스트를 행/셀 단위로 처리
    """
    lines = []

    # 표(Table)인 경우: 행과 셀을 순회하며 텍스트 추출
    if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        table = shape.table
        for row in table.rows:
            row_texts = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text:
                    row_texts.append(cell_text)
            if row_texts:
                # 같은 행의 셀 텍스트는 탭으로 구분하여 한 줄로 표현
                lines.append("\t".join(row_texts))

    # 텍스트 프레임이 있는 도형(일반 텍스트 상자, 도형 등)인 경우
    elif shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            # 각 단락(paragraph)의 텍스트를 하나로 합치기
            para_text = "".join(run.text for run in paragraph.runs).strip()
            if para_text:
                lines.append(para_text)

    return lines


def extract_text_from_slide(slide):
    """
    슬라이드 한 장에서 모든 도형의 텍스트를 추출하는 함수.
    그룹 도형(GroupShape) 내부도 재귀적으로 탐색.
    """
    all_lines = []

    def process_shapes(shapes):
        """그룹 도형을 재귀적으로 탐색하는 내부 함수"""
        for shape in shapes:
            # 그룹 도형인 경우: 내부 도형들을 재귀 탐색
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                process_shapes(shape.shapes)
            else:
                extracted = extract_text_from_shape(shape)
                all_lines.extend(extracted)

    process_shapes(slide.shapes)
    return all_lines


def extract_pptx_to_txt(pptx_path):
    """
    PPTX 파일을 받아 슬라이드별 텍스트를 추출하고
    같은 경로에 TXT 파일로 저장하는 메인 함수.
    """
    # 출력 TXT 파일 경로 생성: 원본 파일명 + '_오탈자검토.txt'
    base_name = os.path.splitext(os.path.basename(pptx_path))[0]
    output_dir = os.path.dirname(pptx_path)
    output_path = os.path.join(output_dir, f"{base_name}_오탈자검토.txt")

    # PPTX 파일 열기
    prs = Presentation(pptx_path)

    output_lines = []

    # 슬라이드 순회 (1-based 번호 사용)
    for slide_num, slide in enumerate(prs.slides, start=1):
        # 슬라이드 구분선 및 번호 헤더
        output_lines.append("=========================================")
        output_lines.append(f"[Slide {slide_num}]")
        output_lines.append("=========================================")

        # 해당 슬라이드에서 텍스트 추출
        texts = extract_text_from_slide(slide)

        if texts:
            output_lines.extend(texts)
        else:
            # 텍스트가 없는 슬라이드 처리
            output_lines.append(f"[Slide {slide_num}] 추출된 텍스트가 없습니다.")

        # 슬라이드 사이 빈 줄 추가 (가독성)
        output_lines.append("")

    # TXT 파일로 저장 (UTF-8 인코딩)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))

    return output_path


def main():
    """
    프로그램 진입점.
    tkinter로 파일 선택 창을 띄우고, 추출 후 결과를 안내.
    """
    # tkinter 루트 창 생성 (파일 선택 창만 사용하므로 숨김 처리)
    root = tk.Tk()
    root.withdraw()

    # 파일 선택 팝업 창 열기 (PPTX 파일만 필터)
    pptx_path = filedialog.askopenfilename(
        title="PPTX 파일을 선택하세요",
        filetypes=[("PowerPoint 파일", "*.pptx"), ("모든 파일", "*.*")]
    )

    # 사용자가 파일을 선택하지 않고 창을 닫은 경우
    if not pptx_path:
        messagebox.showinfo("취소", "파일 선택이 취소되었습니다.")
        return

    try:
        # 텍스트 추출 및 TXT 저장 실행
        output_path = extract_pptx_to_txt(pptx_path)
        messagebox.showinfo(
            "완료",
            f"텍스트 추출이 완료되었습니다!\n\n저장 위치:\n{output_path}"
        )
    except Exception as e:
        # 오류 발생 시 사용자에게 알림
        messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{e}")


if __name__ == "__main__":
    main()
