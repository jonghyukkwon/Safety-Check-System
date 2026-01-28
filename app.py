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
# 3. 메인 UI 구성
# ==========================================
st.set_page_config(page_title="호텔 안전보건 시스템", layout="wide")

# 상단 툴바 및 배경 스타일 커스텀
st.markdown("""
    <style>
        /* 상단 헤더 영역 배경색 설정 */
        header[data-testid="stHeader"] {
            background-color: #9F896C !important;
        }
        
        /* 헤더 내 아이콘 색상을 화이트로 변경 (골드 배경 대비) */
        header[data-testid="stHeader"] svg {
            fill: white !important;
        }

        /* 탭 바의 강조 라인 색상 */
        .stTabs [data-baseweb="tab-highlight-indicator"] {
            background-color: #9F896C !important;
        }
        
        /* 버튼 배경색 커스텀 */
        div.stButton > button:first-child {
            background-color: #9F896C;
            color: white;
            border: none;
        }
    </style>
    """, unsafe_allow_html=True)

st.title("🏨 호텔 안전보건 통합 관리 시스템")

tab1, tab2, tab3 = st.tabs(["📑 적격수급업체 평가", "📊 위험성평가 자동 생성", "📑 안전보건관리계획서 기반 위험성평가 생성"])

# --- TAB 1: 기존 코드 (유지) ---

# ------------------------------------------------------------------------------
# TAB 1: 적격수급업체 평가 (문법 오류 수정 완료)
# ------------------------------------------------------------------------------
with tab1:
    st.header("1. 수급업체 안전보건관리계획서 평가")
    st.info("AI가 가이드라인에 따라 점수를 산출합니다.")

    # 1. 모델 설정
    eval_model = genai.GenerativeModel(
        model_name="models/gemini-2.5-flash",
        generation_config={
            "temperature": 0.0,
            "response_mime_type": "application/json",
        },
        system_instruction=  "당신은 창의성이 없는 '안전보건 점수 계산기'입니다. "

            "문서를 해석하려 하지 말고, 텍스트에 키워드가 있는지만 확인하십시오. "

            "다음 '채점 알고리즘'을 엄격히 따르십시오: "

            "1. IF (해당 항목에 대한 구체적 계획 + 실행 증빙/사진) EXIST -> THEN [우수/만점] "

            "2. IF (계획은 있으나 증빙이 없거나 추상적임) -> THEN [보통/중간점수] "

            "3. IF (내용 누락 OR 다른 현장 복사/붙여넣기 흔적) -> THEN [미흡/최하점] "

            "4. DEFAULT (판단이 애매한 모든 경우) -> THEN [미흡/최하점] (점수 부풀리기 금지)"
    )

    user_file = st.file_uploader("업체 제출 계획서(PDF) 업로드", type=["pdf"], key="eval_upload")

    if st.button("평가 시작", key="eval_btn"):
        if not user_file:
            st.warning("파일을 업로드해 주세요.")
        else:
            with st.spinner("AI가 문서의 이미지와 내용을 정밀 분석 중..."):
                temp_path = "temp_eval.pdf"
                try:
                    # 파일 저장 및 업로드
                    with open(temp_path, "wb") as f:
                        f.write(user_file.getbuffer())
                    
                    uploaded_file = genai.upload_file(temp_path, mime_type="application/pdf")
                    
                    # 파일 처리 대기
                    while uploaded_file.state.name == "PROCESSING":
                        time.sleep(1)
                        uploaded_file = genai.get_file(uploaded_file.name)

                    # 프롬프트
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
                          ... (17번까지 반복)
                    ]
                    """
                    
                    # AI 실행
                    response = eval_model.generate_content([prompt, uploaded_file])
                    
                    # 결과 처리
                    eval_data = json.loads(response.text)
                    
                    # 데이터 유효성 검사 (리스트인지 확인)
                    if isinstance(eval_data, list):
                        total_score = sum(item['score'] for item in eval_data)
                        
                        st.markdown(f"## 🏆 종합 점수: **{total_score}점**")
                        st.markdown("---")
                        
                        # 자동 총평 출력
                        if total_score >= 90:
                            st.success("✅ **[고위험군 / 일반군 모두 적격]**")
                        elif 80 <= total_score < 90:
                            st.warning("⚠️ **[일반군 적격 / 고위험군 부적격]**")
                        elif 70 <= total_score < 80:
                            st.error("❌ **[부적격]** (80점 미달)")
                        else:
                            st.error("🚫 **[절대 선정 불가]** (70점 미만)")
                        
                        st.markdown("---")
                        
                        # 테이블 변환 및 출력
                        display_data = []
                        for item in eval_data:
                            display_data.append({
                                "항목": f"{item['item_no']}. {item['category']}",
                                "점수": f"{item['score']} / {item['max_score']}",
                                "등급": item['judgment'],
                                "근거": item['evidence']
                            })
                        st.table(display_data)

                    else:
                        # 리스트가 아닐 경우 에러 처리
                        st.error("AI가 올바른 리스트 형식의 데이터를 반환하지 않았습니다.")
                        st.json(eval_data) # 디버깅용 데이터 출력

                    # 뒷정리
                    genai.delete_file(uploaded_file.name)
                    if os.path.exists(temp_path): os.remove(temp_path)

                except Exception as e:
                    st.error(f"오류 발생: {e}")
                    if os.path.exists(temp_path): os.remove(temp_path)
                        
# --- TAB 2: 엑셀 자동 생성 (NEW) ---
with tab2:
    st.header("2. 공사 위험성평가 엑셀(Excel) 자동 작성")
    st.info("공사 내용을 입력하면, AI가 **표준 서식이 적용된 엑셀 파일**을 즉시 만들어줍니다.")

    with st.container(border=True):
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.subheader("📝 1. 공사 개요 입력")
            p_name = st.text_input("공사명", placeholder="예: 3층 객실 리모델링")
            p_loc = st.text_input("장소", placeholder="예: 본관 3층")
            p_period = st.text_input("기간", placeholder="예: 26.02.01 ~ 02.15")
            p_content = st.text_area("작업 상세 내용", height=120, placeholder="구체적인 작업 내용 입력...")

        with col2:
            st.subheader("⚠️ 2. 위험요인 선택")
            risk_cols = st.columns(3)
            r_check = [
                risk_cols[0].checkbox("🔥 화기"),
                risk_cols[0].checkbox("⚡ 전기"),
                risk_cols[1].checkbox("🪜 고소"),
                risk_cols[1].checkbox("🏗️ 중량물"),
                risk_cols[2].checkbox("☠️ 위험물"),
                risk_cols[2].checkbox("🕳️ 밀폐")
            ]
            risk_labels = ["화기", "전기", "고소", "중량물", "위험물", "밀폐"]
            selected_risks = [label for label, checked in zip(risk_labels, r_check) if checked]

            st.markdown("---")
            generate_btn = st.button("✨ 엑셀 파일 생성하기 (AI)", type="primary", use_container_width=True)

    if generate_btn:
        if not p_name or not selected_risks:
            st.warning("공사명과 최소 1개의 위험요인을 선택해주세요.")
        else:
            with st.spinner("AI가 위험요인을 분석하고 엑셀 서식을 그리는 중입니다..."):
                try:
                    # 1. 모델 호출 및 데이터 생성
                    model = genai.GenerativeModel(MODEL_ID)
                    
                    prompt = f"""
                    다음 공사 정보를 바탕으로 [위험성평가표]에 들어갈 내용을 JSON 형식으로 작성해줘.
                    
                    [공사 정보]
                    - 공사명: {p_name} / 내용: {p_content}
                    - 핵심 위험요인: {", ".join(selected_risks)}

                    [작성 규칙]
                    1. 선택된 위험요인과 관련된 구체적인 위험 항목을 5~7개 도출할 것.
                    2. 감소대책은 "안전모 착용" 처럼 짧게 쓰지 말고, "KCS 인증 안전모 착용 및 턱끈 체결 확인" 처럼 구체적으로 작성할 것.
                    3. **반드시 아래 JSON 구조만 출력할 것.**
                    
                    [
                        {{
                            "equipment": "작업명/장비 (예: 용접작업)",
                            "risk_factor": "위험요인 (예: 불티 비산)",
                            "risk_level": "상/중/하",
                            "countermeasure": "구체적 대책 내용을 길게 작성",
                            "manager": "안전담당자"
                        }}
                    ]
                    """
                    
                    response = model.generate_content(prompt)
                    
                    # 2. JSON 파싱
                    clean_text = re.sub(r"```json|```", "", response.text).strip()
                    risk_data_list = json.loads(clean_text)
                    
                    # 3. 엑셀 파일 생성 (스타일 적용)
                    p_info = {"name": p_name, "loc": p_loc, "period": p_period, "content": p_content}
                    excel_byte = generate_excel_from_scratch(p_info, risk_data_list)
                    
                    # 4. 다운로드 제공
                    st.success("✅ 엑셀 파일 생성이 완료되었습니다!")
                    st.download_button(
                        label="📥 엑셀 파일 다운로드 (.xlsx)",
                        data=excel_byte,
                        file_name=f"위험성평가_{p_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except json.JSONDecodeError:
                    st.error("AI 응답 처리 실패 (데이터 형식 오류). 다시 시도해주세요.")
                except Exception as e:

                    st.error(f"오류 발생: {e}")



# TAB 3: 위험성평가 엑셀 생성 (PDF 기반 - NEW)
# ------------------------------------------------------------------------------
with tab3:
    st.header("3. 안전보건관리계획서 기반 위험성평가 자동 생성")
    st.info("업로드한 계획서의 **공사 개요**와 **작업 내용**을 AI가 스스로 분석하여, 위험성평가 엑셀 파일을 만들어줍니다.")

    pdf_file = st.file_uploader("안전보건관리계획서(PDF) 업로드", type=["pdf"], key="risk_pdf_upload")
    
    if st.button("🚀 PDF 분석 및 엑셀 생성", key="pdf_risk_btn", type="primary"):
        if not pdf_file:
            st.warning("먼저 PDF 파일을 업로드해주세요.")
        else:
            with st.spinner("AI가 계획서를 분석하여 공사 정보와 위험요인을 추출하고 있습니다..."):
                temp_pdf_path = "temp_plan.pdf"
                try:
                    # 1. PDF 업로드
                    with open(temp_pdf_path, "wb") as f:
                        f.write(pdf_file.getbuffer())
                    
                    uploaded_pdf = genai.upload_file(temp_pdf_path, mime_type="application/pdf")
                    while uploaded_pdf.state.name == "PROCESSING":
                        time.sleep(1)
                        uploaded_pdf = genai.get_file(uploaded_pdf.name)

                    # 2. AI 분석 모델 호출
                    # (창의성 설정 적용: 위험요인 도출 시 유연성 필요)
                    pdf_risk_model = genai.GenerativeModel(MODEL_ID, generation_config=creative_config)

                    prompt = """
                    첨부된 [안전보건관리계획서] PDF를 정밀 분석하여, 아래 두 가지 정보를 JSON 형식으로 추출 및 생성하세요.

                    1. **공사 개요 정보 추출**: PDF 내에서 공사명, 현장 위치(장소), 공사 기간, 주요 작업 내용을 찾아내세요.
                       (만약 정확한 기간이나 장소가 명시되지 않았다면 'PDF 내 미기재'로 표기할 것)
                    
                    2. **위험성평가 데이터 생성**: 
                       - 분석된 '작업 내용'과 '현장 사진/도면' 등을 바탕으로 예상되는 주요 위험요인을 7가지 이상 도출하세요.
                       - 각 위험요인에 대해 구체적인 안전 대책을 수립하세요.

                    [필수 출력 형식 (JSON Only)]
                    반드시 아래 JSON 구조를 엄격히 지켜서 출력하세요. Markdown 코드는 제외하세요.

                    {
                        "project_info": {
                            "name": "공사명 추출 결과",
                            "loc": "장소 추출 결과",
                            "period": "기간 추출 결과",
                            "content": "작업 내용 요약"
                        },
                        "risk_data": [
                            {
                                "equipment": "작업단위 또는 사용장비 (예: 용접작업)",
                                "risk_factor": "구체적 위험요인 (예: 불티 비산에 의한 화재)",
                                "risk_level": "상/중/하",
                                "countermeasure": "구체적 안전 대책 (KCS 인증 보호구 착용 등)",
                                "manager": "안전담당자"
                            }
                        ]
                    }
                    """

                    response = pdf_risk_model.generate_content([prompt, uploaded_pdf])
                    
                    # 3. 데이터 파싱
                    raw_text = response.text
                    # JSON 블록 찾기 ({ ... })
                    json_match = re.search(r'\{.*\}', raw_text, re.DOTALL)
                    
                    if json_match:
                        full_data = json.loads(json_match.group(0))
                        
                        # 데이터 분리
                        extracted_info = full_data.get("project_info", {})
                        extracted_risks = full_data.get("risk_data", [])

                        # 4. 분석 결과 미리보기 (사용자 확인용)
                        st.success("✅ 분석 완료! 추출된 정보를 확인하세요.")
                        
                        with st.expander("📄 추출된 공사 개요 확인", expanded=True):
                            c1, c2 = st.columns(2)
                            c1.text_input("공사명", value=extracted_info.get("name", ""), disabled=True)
                            c1.text_input("장소", value=extracted_info.get("loc", ""), disabled=True)
                            c2.text_input("기간", value=extracted_info.get("period", ""), disabled=True)
                            c2.text_area("작업 내용", value=extracted_info.get("content", ""), disabled=True)

                        # 5. 엑셀 생성
                        excel_byte = generate_excel_from_scratch(extracted_info, extracted_risks)

                        st.markdown("---")
                        st.download_button(
                            label="📥 위험성평가 엑셀 다운로드 (.xlsx)",
                            data=excel_byte,
                            file_name=f"위험성평가_{extracted_info.get('name', '자동생성')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                        
                    else:
                        st.error("AI 응답에서 유효한 JSON 데이터를 찾지 못했습니다.")
                        st.text(raw_text)

                    # 파일 정리
                    genai.delete_file(uploaded_pdf.name)
                    if os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)

                except Exception as e:
                    st.error(f"분석 중 오류 발생: {e}")
                    if os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)





















