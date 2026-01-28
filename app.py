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
# 0. 페이지 설정 및 디자인 (샴페인 골드)
# ==========================================
st.set_page_config(page_title="호텔 안전보건 시스템", layout="wide")
LOGO_URL = "https://raw.githubusercontent.com/jonghyukkwon/Safety-Check-System/main/logo.png"

# 샴페인 골드 테마 & 다크 모드 호환 CSS
st.markdown(f"""
    <style>
        /* 상단 헤더 배경색 (샴페인 골드) */
        header[data-testid="stHeader"] {{
            background-color: #9F896C !important;
            
        }}

        /* 헤더 내부에 로고 강제 삽입 */
        header[data-testid="stHeader"]::before {{
            content: "";
            position: absolute;
            left: 20px;
            top: 50%;
            transform: translateY(-50%);
            width: 215px;
            height: 40px;
            background-image: url("{LOGO_URL}");
            background-size: contain;
            background-repeat: no-repeat;
            background-position: left center;
            z-index: 1;
        }}

        /* 아이콘에 마우스를 올렸을 때 배경색 (샴페인 골드와 어울리는 연한 흰색) */
        header[data-testid="stHeader"] button:hover {{
            background-color: rgba(255, 255, 255, 0.2) !important;
        }}

        /* 탭 선택 시 강조 색상 */
        .stTabs [data-baseweb="tab-highlight-indicator"] {{
            background-color: #9F896C !important;
        }}
        
        /* 버튼 스타일 */
        div.stButton > button:first-child {{
            background-color: #9F896C;
            color: white;
            border: none;
        }}
        div.stButton > button:hover {{
            background-color: #8A7558;
            color: white;
        }}
    </style>
    """, unsafe_allow_html=True)

st.title("🏨 호텔 안전보건 통합 관리 시스템")



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
# 3. 메인 UI 구조 (대분류 -> 소분류 구조 변경)
# ==========================================
# 대분류 탭 이름 변경
main_tab1, main_tab2 = st.tabs(["📑 안전보건관계서류 검토", "📊 위험성평가 관리"])

# ------------------------------------------------------------------------------
# [Main Tab 1] 안전보건관계서류 검토 (Sub Tab 1-1, 1-2)
# ------------------------------------------------------------------------------
with main_tab1:
    # 서브 탭 생성
    sub_tab1_1, sub_tab1_2 = st.tabs(["📝 1-1. 안전보건관리계획서 평가", "🔍 1-2. 위험성평가 적정성 평가"])

    # [Sub Tab 1-1] 기존 안전보건관리계획서 정량 평가 (기존 코드 이동)
    with sub_tab1_1:
        st.subheader("1-1. 안전보건관리계획서 정량 평가")
        st.info("AI가 '안전보건관리계획서 가이드라인'에 따라 점수를 산출합니다.")

        # 모델 설정 (엄격 모드)
        eval_model = genai.GenerativeModel(
            model_name=MODEL_ID,
            generation_config=generation_config, 
            safety_settings=safety_settings,
            system_instruction="당신은 감정이 없는 '안전보건 점수 계산기'입니다."
        )

        user_file = st.file_uploader("업체 제출 계획서(PDF) 업로드", type=["pdf"], key="eval_upload_1_1")

        if st.button("계획서 평가 시작", key="btn_eval_1_1"):
            if not user_file:
                st.warning("파일을 업로드해 주세요.")
            else:
                with st.spinner("AI가 계획서를 분석 중입니다..."):
                    temp_path = "temp_eval_plan.pdf"
                    try:
                        with open(temp_path, "wb") as f: f.write(user_file.getbuffer())
                        uploaded_file = genai.upload_file(temp_path, mime_type="application/pdf")
                        while uploaded_file.state.name == "PROCESSING": time.sleep(1); uploaded_file = genai.get_file(uploaded_file.name)

                        prompt = f"""
                        [참조: 가이드라인] {MASTER_GUIDE_TEXT}
                        [지침] 위 가이드라인을 기준으로 계획서를 채점하세요. 증거가 없으면 0점 처리하십시오.
                        [출력 형식] JSON 리스트: [ {{ "item_no": 1, "category": "항목명", "score": 0, "max_score": 5, "evidence": "...", "judgment": "..." }} ]
                        """
                        
                        response = eval_model.generate_content([prompt, uploaded_file])
                        eval_data = json.loads(response.text)
                        
                        if isinstance(eval_data, dict): eval_data = list(eval_data.values())[0]

                        if isinstance(eval_data, list):
                            total_score = sum(item['score'] for item in eval_data)
                            st.markdown(f"## 🏆 종합 점수: **{total_score}점**")
                            
                            # 등급 표시 로직
                            if total_score >= 90: st.success("✅ **[적격]**")
                            elif 70 <= total_score < 90: st.warning("⚠️ **[보완 필요]**")
                            else: st.error("❌ **[부적격]**")
                            
                            st.table([{"항목": f"{i['item_no']}. {i['category']}", "점수": f"{i['score']}", "근거": i['evidence']} for i in eval_data])
                        else:
                            st.error("데이터 형식 오류")

                        genai.delete_file(uploaded_file.name)
                        if os.path.exists(temp_path): os.remove(temp_path)
                    except Exception as e: st.error(f"오류: {e}")

    # [Sub Tab 1-2] 위험성평가 적정성 평가 (신규 기능)
    with sub_tab1_2:
        st.subheader("1-2. 위험성평가 적정성 검토")
        st.info("제출된 위험성평가서(PDF/Excel)가 '적정성 검토 가이드라인'에 부합하는지 분석합니다.")

        risk_file = st.file_uploader("위험성평가서 업로드 (PDF/Excel)", type=["pdf", "xlsx", "xls"], key="eval_upload_1_2")

        if st.button("위험성평가 검토 시작", key="btn_eval_1_2"):
            if not risk_file:
                st.warning("파일을 업로드해 주세요.")
            else:
                with st.spinner("위험성평가 내용을 분석 중입니다..."):
                    try:
                        # 파일 처리 로직
                        file_ext = risk_file.name.split('.')[-1].lower()
                        content_parts = []

                        # 1. 엑셀 파일일 경우: 텍스트로 변환하여 프롬프트에 삽입
                        if file_ext in ['xlsx', 'xls']:
                            import pandas as pd
                            df_dict = pd.read_excel(risk_file, sheet_name=None)
                            text_content = ""
                            for sheet, df in df_dict.items():
                                text_content += f"Sheet: {sheet}\n{df.to_string()}\n"
                            content_parts.append(text_content)
                        
                        # 2. PDF 파일일 경우: Gemini에 직접 업로드
                        elif file_ext == 'pdf':
                            temp_path = "temp_eval_risk.pdf"
                            with open(temp_path, "wb") as f: f.write(risk_file.getbuffer())
                            uploaded_file = genai.upload_file(temp_path, mime_type="application/pdf")
                            while uploaded_file.state.name == "PROCESSING": time.sleep(1); uploaded_file = genai.get_file(uploaded_file.name)
                            content_parts.append(uploaded_file)

                        # 평가 모델 호출
                        risk_eval_model = genai.GenerativeModel(
                            model_name=MODEL_ID,
                            generation_config=generation_config,
                            safety_settings=safety_settings
                        )

                        prompt = f"""
                        당신은 '위험성평가 검토 전문가'입니다.
                        제출된 문서를 아래 [위험성평가 가이드라인]에 따라 평가하십시오.

                        [위험성평가 가이드라인 (MASTER_GUIDE_TEXT2)]
                        {MASTER_GUIDE_TEXT2}

                        [평가 기준]
                        - 각 항목별로 구체적인 근거(문서 내 내용)를 찾아 평가할 것.
                        - 두루뭉술하거나 복사 붙여넣기 한 내용은 감점할 것.

                        [출력 형식]
                        반드시 아래 JSON 리스트 형태로 출력하세요.
                        [
                            {{
                                "category": "평가 항목명 (예: 위험요인 도출)",
                                "score": 25,
                                "max_score": 30,
                                "status": "양호/미흡",
                                "comment": "평가 의견 및 보완 필요 사항"
                            }}
                        ]
                        """
                        
                        content_parts.insert(0, prompt)
                        response = risk_eval_model.generate_content(content_parts)
                        result_data = json.loads(response.text)
                        
                        if isinstance(result_data, dict): result_data = list(result_data.values())[0]

                        # 결과 출력
                        st.markdown("### 📋 검토 결과 보고서")
                        if isinstance(result_data, list):
                            total_risk_score = sum(item['score'] for item in result_data)
                            st.markdown(f"#### 💯 종합 점수: **{total_risk_score}점**")
                            st.progress(total_risk_score / 100)
                            
                            st.markdown("---")
                            
                            # 카드 형태로 결과 보여주기
                            for item in result_data:
                                with st.container(border=True):
                                    c1, c2 = st.columns([8, 2])
                                    with c1:
                                        st.markdown(f"**📌 {item['category']}**")
                                        st.caption(f"의견: {item['comment']}")
                                    with c2:
                                        st.markdown(f"**{item['score']} / {item['max_score']}**")
                                        if item['status'] == "양호":
                                            st.success(item['status'])
                                        else:
                                            st.error(item['status'])
                        else:
                            st.error("분석 결과 형식이 올바르지 않습니다.")

                        # 뒷정리 (PDF인 경우만)
                        if file_ext == 'pdf':
                            genai.delete_file(uploaded_file.name)
                            if os.path.exists(temp_path): os.remove(temp_path)

                    except Exception as e:
                        st.error(f"분석 중 오류 발생: {e}")
