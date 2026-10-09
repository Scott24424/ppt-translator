import streamlit as st
import io
from pptx import Presentation
from google import genai
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

st.set_page_config(page_title="주간보고 PPTX 자동 번역기", layout="centered")

st.title("📄 주간보고 PPTX 자동 번역기")
st.markdown("한국어 주간보고 PPTX를 업로드하면, 비즈니스 중국어로 완벽하게 번역하여 폰트(Microsoft YaHei)까지 자동 적용해 줍니다.")

# API Key 입력 (사용자 편의를 위해 UI에서 직접 입력)
api_key = st.text_input("🔑 Google Gemini API Key를 입력하세요", type="password")

def translate_to_chinese(text, client):
    if not text or not text.strip():
        return text
    
    prompt = f"""
    Translate the following Korean business weekly report text into professional, natural business Chinese. 
    Maintain the original tone, formatting, and any bullet points. Do not add any extra explanations.
    
    Korean Text:
    {text}
    """
    try:
        response = client.models.generate_content(
            model='gemini-3.8-flash',
            contents=prompt,
        )
        return response.text.strip()
    except Exception as e:
        st.error(f"번역 중 오류 발생: {e}")
        return text

def apply_chinese_font(run):
    run.font.name = 'Microsoft YaHei'
    rPr = run.font._element
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = parse_xml(r'<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="Microsoft YaHei"/>')
        rPr.append(ea)
    else:
        ea.set('typeface', 'Microsoft YaHei')

def process_text_frame(text_frame, client):
    for paragraph in text_frame.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs).strip()
        if full_text:
            translated_text = translate_to_chinese(full_text, client)
            if paragraph.runs:
                for i, run in enumerate(paragraph.runs):
                    if i == 0:
                        run.text = translated_text
                        apply_chinese_font(run)
                    else:
                        run.text = ""

uploaded_file = st.file_uploader("업로드할 PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    if st.button("🚀 번역 시작"):
        if not api_key:
            st.warning("API Key를 먼저 입력해 주세요!")
        else:
            client = genai.Client(api_key=api_key)
            
            with st.spinner("PPTX 파일을 분석하고 번역하는 중입니다. (약 1~2분 소요)"):
                try:
                    # 메모리상에서 PPTX 읽기
                    prs = Presentation(uploaded_file)
                    
                    if len(prs.slides) < 2:
                        st.error("오류: PPTX 파일은 최소 2장(1페이지 번역 대상, 2페이지 원본 유지)으로 구성되어야 합니다.")
                    else:
                        # 2페이지 번역
                        for shape in prs.slides[1].shapes:
                            if hasattr(shape, "text_frame") and shape.text_frame:
                                process_text_frame(shape.text_frame, client)
                            elif shape.has_table:
                                for row in shape.table.rows:
                                    for cell in row.cells:
                                        if hasattr(cell, "text_frame") and cell.text_frame:
                                            process_text_frame(cell.text_frame, client)
                        
                        # 3페이지 이후 번역
                        for slide_idx in range(2, len(prs.slides)):
                            for shape in prs.slides[slide_idx].shapes:
                                if hasattr(shape, "text_frame") and shape.text_frame:
                                    process_text_frame(shape.text_frame, client)
                                if shape.has_table:
                                    for row in shape.table.rows:
                                        for cell in row.cells:
                                            if hasattr(cell, "text_frame") and cell.text_frame:
                                                process_text_frame(cell.text_frame, client)
                        
                        # 번역 완료된 PPTX를 메모리에 저장
                        output_stream = io.BytesIO()
                        prs.save(output_stream)
                        output_stream.seek(0)
                        
                        st.success("✅ 번역이 완료되었습니다!")
                        
                        original_name = uploaded_file.name.replace(".pptx", "")
                        st.download_button(
                            label="📥 번역된 PPTX 다운로드",
                            data=output_stream,
                            file_name=f"{original_name}_CN.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                except Exception as e:
                    st.error(f"파일 처리 중 오류가 발생했습니다: {e}")
