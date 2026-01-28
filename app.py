import streamlit as st
import google.generativeai as genai
import json
import io
import re
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from guide_data import MASTER_GUIDE_TEXT

# ==========================================
# 1. API 설정 및 모델 선언
# ==========================================
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    API_KEY = "YOUR_GEMINI_API_KEY" # 로컬 테스트용 키

genai.configure(api_key=API_KEY)

generation_config = {
    "temperature": 0.0,
    "top_p": 1,
    "top_k": 1,
    "max_output_tokens": 8000,
}

creative_config = {
    "temperature": 0.2, # 위험성평가 생성은 약간의 창의성이 필요하므로 0.2로 설정
    "top_p": 0.95,
    "top_k": 40,
    "max_output_tokens": 8000,
}

MODEL_ID = "models/gemini-2.5-flash"

# ==========================================
# 2. 엑셀 양식 생성 및 데이터 입력 함수
# ==========================================
def generate_excel_from_scratch(p_info, risk_data):
    """
    빈 엑셀이 아니라, 코드로 스타일(테두리, 색상)을 직접 그려서 
    완성된 형태의 엑셀 파일을 생성하는 함수
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "위험성평가서"

    # --- 스타일 정의 ---
    # 1. 테두리 스타일 (얇은 실선)
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))
    
    # 2. 헤더 스타일 (회색 배경, 굵은 글씨, 중앙 정렬)
    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    header_font = Font(bold=True, size=11)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 3. 제목 스타일
    title_font = Font(bold=True, size=16)

    # --- 1. 문서 제목 작성 ---
    ws.merge_cells('B2:F2')
    ws['B2'] = "공사 및 작업 안전보건 위험성평가서"
    ws['B2'].font = title_font
    ws['B2'].alignment = center_align

    # --- 2. 공사 개요 (표 상단) 작성 ---
    # 레이블 (B열)
    labels = ["공사명", "공사 장소", "공사 기간", "작업 내용"]
    keys = ["name", "loc", "period", "content"]
    
    start_row = 4
    for i, label in enumerate(labels):
        row = start_row + i
        # 레이블 셀 (B열)
        ws.cell(row=row, column=2, value=label).fill = header_fill
        ws.cell(row=row, column=2).font = header_font
        ws.cell(row=row, column=2).alignment = center_align
        ws.cell(row=row, column=2).border = thin_border
        
        # 데이터 셀 (C~F열 병합)
        ws.merge_cells(f'C{row}:F{row}')
        cell = ws.cell(row=row, column=3, value=p_info[keys[i]])
        cell.alignment = left_align
        cell.border = thin_border
        # 병합된 셀 테두리 적용을 위한 처리
        for col in range(3, 7):
            ws.cell(row=row, column=col).border = thin_border

    # --- 3. 위험성평가 표 헤더 작성 ---
    table_header_row = start_row + 5 # 개요 밑에 띄우고 시작
    headers = ["구분 (장비/작업)", "위험요인 (What)", "위험성", "안전대책 (How)", "담당자"]
    col_widths = [20, 40, 10, 50, 15] # 열 너비 설정

    for i, header in enumerate(headers):
        col_idx = i + 2 # B열(2)부터 시작
        cell = ws.cell(row=table_header_row, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        
        # 열 너비 조정
        ws.column_dimensions[get_column_letter(col_idx)].width = col_widths[i]

    # --- 4. AI 데이터 채우기 ---
    current_row = table_header_row + 1
    
    for item in risk_data:
        # 데이터 매핑
        row_data = [
            item.get('equipment', ''),
            item.get('risk_factor', ''),
            item.get('risk_level', ''),
            item.get('countermeasure', ''),
            item.get('manager', '')
        ]
        
        for i, val in enumerate(row_data):
            col_idx = i + 2
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = center_align if i != 3 else left_align # 대책만 왼쪽 정렬
            
            # 줄바꿈 허용 (내용이 길 경우)
            cell.alignment = Alignment(horizontal=cell.alignment.horizontal, 
                                     vertical="center", 
                                     wrap_text=True)
            
        current_row += 1

    # --- 5. 결재란 만들기 (선택사항) ---
    sign_row = current_row + 2
    ws.merge_cells(f'B{sign_row}:F{sign_row}')
    ws[f'B{sign_row}'] = "위와 같이 위험성평가를 실시하고 안전조치를 이행하겠습니다."
    ws[f'B{sign_row}'].alignment = center_align
    
    sign_row += 2
    ws.cell(row=sign_row, column=4, value="작성자(시공사): (인)")
    ws.cell(row=sign_row, column=6, value="확인자(감독자): (인)")

    # 파일 저장 (메모리)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. 메인 UI 구조 (대분류 -> 소분류)
# ==========================================
# 대분류 탭 생성
main_tab1, main_tab2 = st.tabs(["📑 적격수급업체 평가", "📊 위험성평가 관리"])

# ------------------------------------------------------------------------------
# [Main Tab 1] 적격수급업체 평가
# ------------------------------------------------------------------------------
with main_tab1:
    # 기존 Tab 1 기능 (안전보건관리계획서 정량 평가)
    st.header("1. 수급업체 안전보건관리계획서 정량 평가")
    st.info("AI가 가이드라인에 따라 점수를 산출합니다.")

    # 1. 모델 설정 (기존 Tab 1 코드 유지)
    eval_model = genai.GenerativeModel(
        model_name="models/gemini-2.5-flash",
        generation_config={
            "temperature": 0.0,
            "response_mime_type": "application/json",
        },
        system_instruction="당신은 창의성이 없는 '안전보건 점수 계산기'입니다. 문서를 해석하려 하지 말고, 텍스트에 키워드가 있는지만 확인하십시오."
    )

    user_file = st.file_uploader("업체 제출 계획서(PDF) 업로드", type=["pdf"], key="eval_upload_main")

    if st.button("평가 시작", key="eval_btn_main"):
        if not user_file:
            st.warning("파일을 업로드해 주세요.")
        else:
            with st.spinner("AI가 문서의 이미지와 내용을 정밀 분석 중..."):
                temp_path = "temp_eval.pdf"
                try:
                    with open(temp_path, "wb") as f:
                        f.write(user_file.getbuffer())
                    
                    uploaded_file = genai.upload_file(temp_path, mime_type="application/pdf")
                    while uploaded_file.state.name == "PROCESSING":
                        time.sleep(1)
                        uploaded_file = genai.get_file(uploaded_file.name)

                    # 기존 프롬프트 유지
                    prompt = f"""
                    [참조: 가이드라인]
                    {MASTER_GUIDE_TEXT}

                    [마스터 가이드라인]을 기준으로 수급업체 계획서를 채점하십시오.
                    변덕스러운 점수를 막기 위해, 각 항목별로 **반드시 PDF 내의 '증거 문장'을 먼저 찾고** 점수를 매기십시오.

                    [🚫 절대적 채점 규칙 (Tie-Breaker Rule)]
                    1. **증거 우선주의**: "잘 할 것으로 보임", "계획된 것으로 추정됨" 같은 추측은 절대 금지. PDF에 명시된 문구가 없으면 0점.
                    2. **하향 평가 원칙**: 
                       - 5점 줄까 3점 줄까 고민되면 -> **3점** 부여.
                       - 3점 줄까 1점 줄까 고민되면 -> **1점** 부여.
                       - **즉, 확실한 근거가 없는 한 높은 점수를 주지 마시오.**
                    3. **공종 일치성**: PDF 제목의 공사명과 본문의 작업 내용이 불일치(복사 붙여넣기 의심)하면 해당 항목 0점 처리.
                    4. **중대재해(17번)**: '해당없음' 또는 '무재해'라는 명확한 텍스트나 증명서가 없으면, 확인 불가로 간주하여 0점 처리.

                    [출력 형식]
                    [
                        {{
                            "item_no": 1,
                            "category": "항목명",
                            "score": 0,
                            "max_score": 5,
                            "evidence": "증거 내용",
                            "judgment": "등급"
                        }}
                    ]
                    """
                    
                    response = eval_model.generate_content([prompt, uploaded_file])
                    eval_data = json.loads(response.text)
                    
                    if isinstance(eval_data, list):
                        total_score = sum(item['score'] for item in eval_data)
                        st.markdown(f"## 🏆 종합 점수: **{total_score}점**")
                        
                        if total_score >= 90: st.success("✅ **[고위험군 / 일반군 모두 적격]**")
                        elif 80 <= total_score < 90: st.warning("⚠️ **[일반군 적격 / 고위험군 부적격]**")
                        elif 70 <= total_score < 80: st.error("❌ **[부적격]** (80점 미달)")
                        else: st.error("🚫 **[절대 선정 불가]** (70점 미만)")
                        
                        st.markdown("---")
                        display_data = [{"항목": f"{i['item_no']}. {i['category']}", "점수": f"{i['score']}/{i['max_score']}", "등급": i['judgment'], "근거": i['evidence']} for i in eval_data]
                        st.table(display_data)
                    else:
                        st.error("데이터 형식 오류")
                        st.json(eval_data)

                    genai.delete_file(uploaded_file.name)
                    if os.path.exists(temp_path): os.remove(temp_path)

                except Exception as e:
                    st.error(f"오류: {e}")
                    if os.path.exists(temp_path): os.remove(temp_path)

# ------------------------------------------------------------------------------
# [Main Tab 2] 위험성평가 관리 (소분류 Tab 2.1, 2.2)
# ------------------------------------------------------------------------------
with main_tab2:
    # 소분류 탭 생성
    sub_tab1, sub_tab2 = st.tabs(["📝 2-1. 직접 입력형 생성", "📑 2-2. PDF 기반 생성"])

    # [Sub Tab 2.1] 직접 입력형 (기존 Tab 2)
    with sub_tab1:
        st.subheader("2-1. 공사 내용 직접 입력")
        st.info("공사 내용을 입력하면 표준 위험성평가표 엑셀을 생성합니다.")

        with st.container(border=True):
            col1, col2 = st.columns([1, 1])
            with col1:
                p_name = st.text_input("공사명", placeholder="예: 3층 객실 리모델링")
                p_loc = st.text_input("장소", placeholder="예: 본관 3층")
                p_period = st.text_input("기간", placeholder="예: 26.02.01 ~ 02.15")
                p_content = st.text_area("작업 내용", height=100)
            with col2:
                risk_cols = st.columns(3)
                r_check = [
                    risk_cols[0].checkbox("🔥 화기"), risk_cols[0].checkbox("⚡ 전기"),
                    risk_cols[1].checkbox("🪜 고소"), risk_cols[1].checkbox("🏗️ 중량물"),
                    risk_cols[2].checkbox("☠️ 위험물"), risk_cols[2].checkbox("🕳️ 밀폐")
                ]
                selected_risks = [["화기","전기","고소","중량물","위험물","밀폐"][i] for i, v in enumerate(r_check) if v]
                st.markdown("---")
                gen_btn_manual = st.button("✨ 엑셀 생성 (입력형)", type="primary", use_container_width=True)

        if gen_btn_manual:
            if not p_name: st.warning("공사명을 입력하세요.")
            else:
                with st.spinner("AI 생성 중..."):
                    try:
                        risk_model = genai.GenerativeModel(MODEL_ID, generation_config=creative_config)
                        prompt = f"""
                        [공사정보] {p_name} / {p_content} / 위험요인: {", ".join(selected_risks)}
                        위험요인별 5~7개 항목 도출하여 JSON 출력:
                        [ {{ "equipment": "...", "risk_factor": "...", "risk_level": "...", "countermeasure": "...", "manager": "..." }} ]
                        """
                        response = risk_model.generate_content(prompt)
                        clean_text = re.sub(r"```json|```", "", response.text).strip()
                        risk_data = json.loads(clean_text)
                        
                        excel_byte = generate_excel_from_scratch({"name":p_name, "loc":p_loc, "period":p_period, "content":p_content}, risk_data)
                        st.success("완료!")
                        st.download_button("📥 엑셀 다운로드", excel_byte, f"위험성평가_{p_name}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e: st.error(f"오류: {e}")

    # [Sub Tab 2.2] PDF 기반 생성 (기존 Tab 3)
    with sub_tab2:
        st.subheader("2-2. 안전보건관리계획서(PDF) 기반 자동 생성")
        st.info("PDF 계획서를 분석하여 공사 개요와 위험요인을 스스로 추출합니다.")

        pdf_file = st.file_uploader("계획서(PDF) 업로드", type=["pdf"], key="risk_pdf_upload")
        
        if st.button("🚀 분석 및 엑셀 생성", key="pdf_risk_btn", type="primary"):
            if not pdf_file: st.warning("PDF를 업로드하세요.")
            else:
                with st.spinner("PDF 분석 중..."):
                    temp_pdf = "temp_plan.pdf"
                    try:
                        with open(temp_pdf, "wb") as f: f.write(pdf_file.getbuffer())
                        up_pdf = genai.upload_file(temp_pdf, mime_type="application/pdf")
                        while up_pdf.state.name == "PROCESSING": time.sleep(1); up_pdf = genai.get_file(up_pdf.name)

                        pdf_model = genai.GenerativeModel(MODEL_ID, generation_config=creative_config)
                        prompt = """
                        PDF를 분석하여 다음 두 가지를 JSON으로 추출하세요.
                        1. project_info: 공사명, 장소, 기간, 내용
                        2. risk_data: 작업 내용 기반 위험요인 7개 이상 (equipment, risk_factor, risk_level, countermeasure, manager)
                        출력 형식: { "project_info": {...}, "risk_data": [...] }
                        """
                        response = pdf_model.generate_content([prompt, up_pdf])
                        
                        raw_text = response.text
                        json_match = re.search(r'\{.*\}', raw_text, re.DOTALL)
                        
                        if json_match:
                            full_data = json.loads(json_match.group(0))
                            p_info = full_data.get("project_info", {})
                            r_data = full_data.get("risk_data", [])

                            st.success("분석 완료!")
                            with st.expander("추출된 개요 확인", expanded=True):
                                st.text(f"공사명: {p_info.get('name')}\n내용: {p_info.get('content')}")

                            excel_byte = generate_excel_from_scratch(p_info, r_data)
                            st.download_button("📥 엑셀 다운로드", excel_byte, f"위험성평가_{p_info.get('name','자동')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        else:
                            st.error("JSON 파싱 실패")

                        genai.delete_file(up_pdf.name)
                        if os.path.exists(temp_pdf): os.remove(temp_pdf)
                    except Exception as e:
                        st.error(f"오류: {e}")
                        if os.path.exists(temp_pdf): os.remove(temp_pdf)


