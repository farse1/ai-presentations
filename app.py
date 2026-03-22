import streamlit as st
from pptx import Presentation
import fitz  
from langchain_openai import ChatOpenAI
from langchain_community.tools.tavily_search import TavilySearchResults
import json
import os

st.set_page_config(page_title="AI Presentation Creator", page_icon="📊")

def extract_text_from_pptx(file):
    prs = Presentation(file)
    return "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])

def extract_text_from_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    return "\n".join([page.get_text() for page in doc])

def create_pptx(slides_data):
    prs = Presentation()
    for slide_info in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = slide_info.get("title", "Slide")
        slide.placeholders[1].text = slide_info.get("content", "")
    output_path = "output.pptx"
    prs.save(output_path)
    return output_path

st.title("📊 AI Presentation Generator")
st.info("Upload a reference and I will research the web to create up to 20 slides.")

with st.sidebar:
    o_api = st.text_input("OpenAI API Key", type="password")
    t_api = st.text_input("Tavily API Key", type="password")
    num_slides = st.slider("Max Slides", 5, 20, 10)

topic = st.text_input("Presentation Topic")
ref_file = st.file_uploader("Reference PDF or PPTX", type=["pdf", "pptx"])

if st.button("Generate"):
    if not o_api or not t_api or not topic:
        st.error("Missing API keys or Topic")
    else:
        with st.spinner("Generating..."):
            os.environ["TAVILY_API_KEY"] = t_api
            ref_content = ""
            if ref_file:
                ref_content = extract_text_from_pptx(ref_file) if ref_file.name.endswith("pptx") else extract_text_from_pdf(ref_file)
            
            search = TavilySearchResults(max_results=5)
            web_data = search.invoke(topic)
            
            llm = ChatOpenAI(model="gpt-4o", api_key=o_api)
            prompt = f"Topic: {topic}. Reference: {ref_content[:1500]}. Web: {web_data}. Create {num_slides} slides in JSON format: [{{'title':'', 'content':''}}]"
            
            res = llm.invoke(prompt)
            data = json.loads(res.content.replace("```json", "").replace("```", ""))
            path = create_pptx(data)
            
            with open(path, "rb") as f:
                st.download_button("Download Presentation", f, file_name="AI_Presentation.pptx")
