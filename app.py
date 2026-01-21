import streamlit as st
import google.generativeai as genai
import time
import os
from guide_data import MASTER_GUIDE_TEXT

# ==========================================
# 1. API 설정 및 모델 선언
# ==========================================
API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=API_KEY)

# 최신 gemini-2.5-flash 모델 설정
# 멀티모달 분석을 위해 system_instruction에 시각적 검토 지침 강화
MODEL_ID = "models/gemini-2.5-flash" 

model = genai.GenerativeModel(
    model_name=MODEL_ID,
    system_instruction=(
        "당신은 대한민국 산업안전보건법 및 적격수급업체 평가 전문가입니다. "
        "제공된 PDF 문서를 텍스트뿐만 아니라 시각적으로도 완벽히 분석하세요. "
        "특히 도장(직인), 서명, 현장 사진 증빙, 안전장구 착용 사진 등을 확인하여 "
        "가이드라인 준수 여부를 판정해야 합니다."
    )
)

# ==========================================
# 2. UI 구성
# ==========================================
st.set_page_config(page_title="Gemini 2.5 멀티모달 점검", layout="wide")
st.title("🚀 Gemini 2.5 Flash 멀티모달 안전점검 시스템")
st.info("텍스트는 물론, 도장과 사진 증빙까지 AI가 직접 확인합니다.")

user_file = st.file_uploader("업체 제출 계획서(PDF)를 업로드하세요.", type=["pdf"])

if st.button("AI 시각적 적정성 검토 시작"):
    if not user_file:
        st.warning("파일을 업로드해 주세요.")
    else:
        with st.spinner("Gemini 2.5가 문서의 이미지와 내용을 정밀 분석 중..."):
            try:
                # 1. 임시 파일 저장 (Gemini File API 전달용)
                temp_path = "temp_upload.pdf"
                with open(temp_path, "wb") as f:
                    f.write(user_file.getbuffer())

                # 2. Gemini File API로 PDF 업로드 (멀티모달 처리를 위해 필수)
                # 이 방식은 서버에서 PDF를 이미지로 렌더링하여 AI에게 전달합니다.
                uploaded_file = genai.upload_file(temp_path, mime_type="application/pdf")
                
                # 파일 처리 대기 (ACTIVE 상태가 될 때까지)
                while uploaded_file.state.name == "PROCESSING":
                    time.sleep(1)
                    uploaded_file = genai.get_file(uploaded_file.name)

                # 3. 분석 프롬프트 작성
                prompt = f"""
                [참조: 마스터 가이드라인 기준]
                {MASTER_GUIDE_TEXT}

                [분석 요청 사항]
                첨부된 '수급업체 계획서' PDF를 다음 기준에 따라 분석 보고서를 작성하세요.

                1. 시각적 증빙 확인:
                   - 문서 마지막에 업체 직인(도장)이나 대표자 서명이 실제로 찍혀 있는가?
                   - 위험성평가나 안전교육 항목에 실제 현장 사진이나 교육 사진이 포함되어 있는가?
                   - 사진이 있다면, 가이드라인에서 요구하는 실제 작업 환경과 일치하는가?

                2. 텍스트 내용 검토:
                   - 가이드라인의 필수 항목(안전보건방침, 조직도, 비상대응 등)이 포함되었는가?
                   - 6대 유해위험요인에 대한 대책이 구체적인가?

                3. 종합 등급 판정:
                   - 가이드라인 배점표를 기준으로 예상 등급(S/A/B/C/D)을 산출하세요.
                   - 시각적 증빙(사진, 도장)이 누락되었다면 강력하게 감점 요인으로 명시하세요.

                **보고서는 깔끔한 표와 불렛포인트 형식으로 작성하세요.**
                """

                # 4. 멀티모달 생성 (텍스트 프롬프트 + PDF 파일 객체)
                response = model.generate_content([prompt, uploaded_file])

                # 5. 결과 출력
                st.success("분석 완료!")
                st.markdown("---")
                st.markdown(response.text)

                # 6. 파일 정리 (보안을 위해 업로드된 파일 삭제)
                genai.delete_file(uploaded_file.name)
                if os.path.exists(temp_path):
                    os.remove(temp_path)

            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
                if 'uploaded_file' in locals():

                    genai.delete_file(uploaded_file.name)
