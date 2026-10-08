import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="최종 컨택현황 변환",
    layout="wide"
)

st.title("최종 컨택현황 자동 변환")

uploaded_file = st.file_uploader(
    "컨택현황 엑셀 파일을 업로드하세요.",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    # --------------------------------------------------
    # 1. 엑셀 읽기
    # --------------------------------------------------
    raw = pd.read_excel(uploaded_file, header=None)

    # 실제 데이터는 3행부터 시작
    # 0행 : 빈 행
    # 1행 : 컨택현황
    # 2행 : 상위 헤더
    # 3행 : 실제 변수명
    df = raw.iloc[3:].copy()

    # 변수명 직접 지정
    columns = [
        "일련번호",
        "권역명",
        "시도명",
        "조사원ID",
        "조사원명",
        "방문구분명",
        "방문결과명",
        "방문차수",
        "비고",
        "등록일자",
        "등록시간",
        "사업체명",
        "종사자규모",
        "매출액규모",
        "연락처1",
        "연락처2",
        "휴대폰번호",
        "대표자성별",
        "주소(도로명)",
        "조사방식",
        "참여번호"
    ]

    df.columns = columns

    # --------------------------------------------------
    # 2. 불필요한 빈 행 제거
    # --------------------------------------------------
    df = df[df["일련번호"].notna()].copy()

    # 일련번호 문자열 처리
    df["일련번호"] = df["일련번호"].astype(str).str.strip()

    # 방문차수 숫자화
    df["방문차수"] = pd.to_numeric(
        df["방문차수"],
        errors="coerce"
    )

    # --------------------------------------------------
    # 3. 일련번호별 가장 최근 방문차수 선택
    # --------------------------------------------------
    result = (
        df.sort_values(
            ["일련번호", "방문차수"],
            ascending=[True, False]
        )
        .drop_duplicates(
            subset="일련번호",
            keep="first"
        )
        .copy()
    )

    # --------------------------------------------------
    # 4. 최종 컨택현황 생성
    # --------------------------------------------------
    def make_contact_status(row):

        visit_type = row["방문구분명"]
        visit_result = row["방문결과명"]

        if pd.notna(visit_result) and str(visit_result).strip() != "":
            return f"{visit_type} - {visit_result}"

        return visit_type

    result["최종 컨택현황"] = result.apply(
        make_contact_status,
        axis=1
    )

    # --------------------------------------------------
    # 5. 결과 확인
    # --------------------------------------------------
    st.success(
        f"변환 완료: {len(df):,}건 → {len(result):,}개 사업체"
    )

    st.dataframe(
        result,
        use_container_width=True
    )

    # --------------------------------------------------
    # 6. 엑셀 다운로드
    # --------------------------------------------------
    output = BytesIO()

    with pd.ExcelWriter(
            output,
            engine="openpyxl"
    ) as writer:

        result.to_excel(
            writer,
            index=False,
            sheet_name="최종 컨택현황"
        )

        wb = writer.book
        ws = writer.sheets["최종 컨택현황"]

        # ---------------------------------------
        # 헤더 스타일
        # ---------------------------------------
        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        header_font = Font(
            color="FFFFFF",
            bold=True
        )

        thin_border = Border(
            left=Side(style="thin", color="D9E1F2"),
            right=Side(style="thin", color="D9E1F2"),
            top=Side(style="thin", color="D9E1F2"),
            bottom=Side(style="thin", color="D9E1F2")
        )

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )
            cell.border = thin_border

        # ---------------------------------------
        # 데이터 영역
        # ---------------------------------------
        for row in ws.iter_rows(
                min_row=2,
                max_row=ws.max_row
        ):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(
                    vertical="center"
                )

        # ---------------------------------------
        # 열 너비 자동 조정
        # ---------------------------------------
        for column_cells in ws.columns:

            max_length = 0
            column_letter = get_column_letter(
                column_cells[0].column
            )

            for cell in column_cells:

                if cell.value is not None:
                    length = len(str(cell.value))

                    if length > max_length:
                        max_length = length

            ws.column_dimensions[
                column_letter
            ].width = min(max_length + 2, 40)

        # ---------------------------------------
        # 헤더 고정
        # ---------------------------------------
        ws.freeze_panes = "A2"

        # ---------------------------------------
        # 행 높이
        # ---------------------------------------
        ws.row_dimensions[1].height = 25
        ws.row_dimensions[3].height = 22

    output.seek(0)

    output.seek(0)

    st.download_button(
        label="최종 컨택현황 엑셀 다운로드",
        data=output,
        file_name="최종_컨택현황.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
