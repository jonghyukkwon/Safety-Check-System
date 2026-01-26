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
st.title("🏨 호텔 안전보건 통합 관리 시스템")

tab1, tab2 = st.tabs(["📑 적격수급업체 평가", "📊 위험성평가 엑셀 생성"])

# --- TAB 1: 기존 코드 (유지) ---

with tab1:
    st.header("1. 수급업체 안전보건관리계획서 적정성 검토")
    st.info("업체가 제출한 계획서를 가이드라인과 대조하여 분석합니다.")
    
    # 모델 설정 (평가용)
    eval_model = genai.GenerativeModel(
        model_name=MODEL_ID,
        system_instruction=(
            "당신은 감정이 없는 냉철한 대한민국 산업안전보건법 및 적격수급업체 평가 심사 전문가입니다. "
            "오직 가이드라인 기준표에 의거하여 기계적으로 점수를 매기십시오."
            "제공된 PDF 문서를 텍스트뿐만 아니라 시각적으로도 완벽히 분석하세요. "
            "특히 도장(직인), 서명, 현장 사진 증빙 등을 확인하여 가이드라인 준수 여부를 판정해야 합니다."
        )
    )

    user_file = st.file_uploader("업체 제출 계획서 업로드", type=["pdf"], key="eval_upload")

    if st.button("적정성 검토 시작", key="eval_btn"):
        if not user_file:
            st.warning("파일을 업로드해 주세요.")
        else:
            with st.spinner("AI가 문서의 이미지와 내용을 정밀 분석 중..."):
                try:
                    # 임시 파일 처리
                    temp_path = "temp_upload.pdf"
                    with open(temp_path, "wb") as f:
                        f.write(user_file.getbuffer())

                    uploaded_file = genai.upload_file(temp_path, mime_type="application/pdf")
                    
                    while uploaded_file.state.name == "PROCESSING":
                        time.sleep(1)
                        uploaded_file = genai.get_file(uploaded_file.name)

                    prompt = f"""
                    
                    [참조: 마스터 가이드라인 기준]
                    {MASTER_GUIDE_TEXT}
                    
                    *[분석 핵심 원칙: 맥락 일치성 검토]
                    **가장 중요:** 분석 시작 전, PDF에 명시된 '공사 제목'과 '공사 내용'을 통해 본 작업의 [주요 공종]을 확정하세요. 
                    만약 특정 항목에 텍스트가 입력되어 있더라도, 확정된 공사 내용과 전혀 관련 없는 엉뚱한 내용(예: 전기 공사인데 용접 작업 내용만 기재 등)이 기입되어 있다면 해당 항목은 0점 처리하고 코멘트에 '공종 부적합'임을 명시하세요.

                    [채점 원칙]
                    1. **증거 기반**: 반드시 PDF 내에서 발견된 근거(Evidence)를 찾아 점수를 부여하세요.
                    2. **점수 기준**: 가이드라인에 명시된 배점(예: 1, 3, 5, 10)을 우선적으로 부여하세요.
                    3. **중대재해(17번)**: 계획서에 중대재해 이력이 없으면 0점, 있으면 -40점입니다.
                    4. 가이드라인에 명시된 1번부터 17번까지의 모든 항목을 빠짐없이 분석할 것.
                    5. 각 항목별로 시각적 증빙(직인, 서명, 사진 등)과 텍스트 내용을 종합하여 점수를 부여할 것.
                    6. 표 하단에 '종합 등급(S~D)'과 '전체 총평' 및 '개선 권고 사항'을 요약하여 작성할 것.

                    [출력 형식]
                    결과는 **반드시 아래 JSON 리스트 형식**으로만 출력하세요. (Markdown 표 아님)
                    
                    [
                        {{
                            "item_no": 1,
                            "category": "안전보건관리 인력",
                            "score": 5,
                            "max_score": 5,
                            "evidence": "조직도상 본사 안전팀의 월 1회 기술지원 계획이 확인됨.",
                            "judgment": "우수"
                        }},
                        ... (17번까지 반복)
                    ]
             
                    """

                    # 3. AI 실행
                    response = eval_model.generate_content([prompt, uploaded_file])

                    # 4. JSON 파싱 및 결과 출력
                    raw_text = response.text
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)

                    if json_match:
                        eval_data = json.loads(json_match.group(0))
                        
                        # 점수 합계
                        total_score = sum(item['score'] for item in eval_data)
                        
                        # 등급 산정
                        st.markdown(f"## 🏆 종합 점수: **{total_score}점**")
                        

                        if total_score >= 90:
                            st.success(f"✅ **[고위험군 / 일반군 모두 적격]**\n\n이 업체는 **90점 이상**을 획득하여, 화재·폭발·밀폐작업 등 고위험 작업을 포함한 모든 도급 공사를 수행할 자격이 있습니다.")
                        
                        elif 80 <= total_score < 90:
                            st.warning(f"⚠️ **[조건부 적격]**\n\n- **일반군(일반 공사):** ✅ **적격** (80점 이상 충족)\n- **고위험군(화기/밀폐):** ❌ **부적격** (90점 미만)\n\n※ 화재, 폭발, 밀폐 작업이 포함된 공사에는 선정할 수 없습니다.")
                            
                        elif 70 <= total_score < 80:
                            st.error(f"❌ **[부적격]**\n\n이 업체는 **80점 미만**으로, 일반 도급 공사 선정 기준을 충족하지 못했습니다.")
                            
                        else: # 70점 미만
                            st.error(f"🚫 **[절대 선정 불가]**\n\n총점 **{total_score}점** (70점 미만)입니다. 규정에 따라 어떠한 경우에도 도급 업체로 선정할 수 없습니다.")
                        
                        st.markdown("---")
                    
                        # 테이블 데이터 변환
                        display_data = []
                        for item in eval_data:
                            display_data.append({
                                "항목": f"{item['item_no']}. {item['category']}",
                                "배점": f"{item['score']} / {item['max_score']}",
                                "등급": item['judgment'],
                                "판단 근거": item['evidence']
                            })
                        st.table(display_data)
                    else:
                        st.error("데이터 형식을 불러오지 못했습니다.")
                        st.text(raw_text)

                    # 뒷정리
                    genai.delete_file(uploaded_file.name)
                    if os.path.exists(temp_path): os.remove(temp_path)

                except Exception as e:
                    st.error(f"평가 중 오류 발생: {e}")

                    # 파일 정리
                    genai.delete_file(uploaded_file.name)
                    if os.path.exists(temp_path):
                        os.remove(temp_path)

                except Exception as e:
                    st.error(f"오류가 발생했습니다: {e}")


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





