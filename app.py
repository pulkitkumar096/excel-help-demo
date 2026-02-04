import streamlit as st
from pptx import Presentation
import os

from openai import OpenAI
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
st.set_page_config(page_title="Excel Help Assistant", layout="centered")

st.title("Excel Help Assistant")
st.write("Ask any question about creating a Table in Excel")

# Load PPT content
ppt = Presentation("How to make a Table in Excel.pptx")
content = ""
for slide in ppt.slides:
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            content += shape.text + "\n"

question = st.text_input("Type your question")

if question:
    prompt = f"""
You are an Excel expert.
Use the instructions below to answer clearly.

Instructions:
{content}

Question:
{question}
"""

    response = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {"role": "system", "content": "You are an Excel expert."},
        {"role": "user", "content": prompt}
    ]
)

st.success(response.choices[0].message.content)
    st.video("How to make a Table in Excel.mp4")
