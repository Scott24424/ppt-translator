import streamlit as st
import io
import json
import time
from pptx import Presentation
from google import genai
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

st.set_page_config(page_title="주간보고 PPTX 자동 번역기", layout="centered")

st.title("📄 주간보고 PPTX 자동 번역기")
st.markdown("한국어 주간보고 PPTX를 업로드하면, 비즈니스 중국어로 완벽하게 번역하여 폰트(Microsoft YaHei)까지 자동 적용해 줍니다.")

api_key = st.text_input("🔑 Google Gemini API Key를 입력하세요", type="password")

def batch_translate(texts, client, retries=5):
    """모든 텍스트를 모아서 단 1번의 API 호출로 번역합니다. (503 에러 시 자동 재시도 포함)"""
    if not texts:
        return []
        
    prompt = f"""
    Translate the following JSON array of Korean business weekly report texts into professional, natural business Chinese.
    Maintain the original tone, formatting, and any bullet points. Do not add any extra explanations.
    Return ONLY a valid JSON array of strings in the exact same order.
    
    Korean Texts:
    {json.dumps(texts, ensure_ascii=False)}
    """
    
    for attempt in range(retries):
        try:
            response = client.models.generate_content(
                model='gemini-3.8-flash',
                contents=prompt,
                config={"response_mime_type": "application/json"}
            )
            translated_list = json.loads(response.text)
            
            if len(translated_list) != len(texts):
                st.warning("경고: 원본 문장 수와 번역된 문장 수가 일치하지 않아 일부 텍스트가 누락될 수 있습니다.")
                return texts
                
            return translated_list
            
        except Exception as e:
            error_msg = str(e)
            if "503" in error_msg or "UNAVAILABLE" in error_msg:
                if attempt < retries - 1:
                    wait_time = (attempt + 1) * 3
                    st.toast(f"구글 서버 과부하(503) 발생. {wait_time}초 후 자동 재시도합니다... (시도 {attempt+1}/{retries})", icon="⏳")
                    time.sleep(wait_time)
                    continue
            st.error(f"일괄 번역 중 오류 발생: {e}")
            return texts
            
    return texts

def apply_chinese_font(run):
    run.font.name = 'Microsoft YaHei'
    rPr = run.font._element
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = parse_xml(r'<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="Microsoft YaHei"/>')
        rPr.append(ea)
    else:
        ea.set('typeface', 'Microsoft YaHei')

uploaded_file = st.file_uploader("업로드할 PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    if st.button("🚀 번역 시작"):
        if not api_key:
            st.warning("API Key를 먼저 입력해 주세요!")
        else:
            clean_api_key = api_key.strip()
            client = genai.Client(api_key=clean_api_key)
            
            with st.spinner("PPTX 파일을 분석하고 번역하는 중입니다. (단 1회의 통신으로 일괄 번역 진행 중...)"):
                try:
                    prs = Presentation(uploaded_file)
                    
                    if len(prs.slides) < 2:
                        st.error("오류: PPTX 파일은 최소 2장(1페이지 번역 대상, 2페이지 원본 유지)으로 구성되어야 합니다.")
                    else:
                        paragraphs_to_translate = []
                        original_texts = []
                        
                        # 2페이지
                        for shape in prs.slides[1].shapes:
                            if hasattr(shape, "text_frame") and shape.text_frame:
                                for p in shape.text_frame.paragraphs:
                                    t = "".join(r.text for r in p.runs).strip()
                                    if t:
                                        paragraphs_to_translate.append(p)
                                        original_texts.append(t)
                            elif shape.has_table:
                                for row in shape.table.rows:
                                    for cell in row.cells:
                                        if hasattr(cell, "text_frame") and cell.text_frame:
                                            for p in cell.text_frame.paragraphs:
                                                t = "".join(r.text for r in p.runs).strip()
                                                if t:
                                                    paragraphs_to_translate.append(p)
                                                    original_texts.append(t)
                        
                        # 3페이지 이후
                        for slide_idx in range(2, len(prs.slides)):
                            for shape in prs.slides[slide_idx].shapes:
                                if hasattr(shape, "text_frame") and shape.text_frame:
                                    for p in shape.text_frame.paragraphs:
                                        t = "".join(r.text for r in p.runs).strip()
                                        if t:
                                            paragraphs_to_translate.append(p)
                                            original_texts.append(t)
                                if shape.has_table:
                                    for row in shape.table.rows:
                                        for cell in row.cells:
                                            if hasattr(cell, "text_frame") and cell.text_frame:
                                                for p in cell.text_frame.paragraphs:
                                                    t = "".join(r.text for r in p.runs).strip()
                                                    if t:
                                                        paragraphs_to_translate.append(p)
                                                        original_texts.append(t)
                        
                        if original_texts:
                            translated_texts = batch_translate(original_texts, client)
                            
                            for p, t_text in zip(paragraphs_to_translate, translated_texts):
                                if p.runs:
                                    for i, run in enumerate(p.runs):
                                        if i == 0:
                                            run.text = t_text
                                            apply_chinese_font(run)
                                        else:
                                            run.text = ""
                        
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
