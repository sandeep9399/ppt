# Apollo PPT Enhancer with AI Suggestions

import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io
import requests

st.set_page_config(page_title="Apollo PPT Enhancer", layout="wide")
st.title("🚀 Apollo Slide Enhancer")

st.markdown("""
Upload your **old PowerPoint (.pptx)** file. You'll see AI-powered layout and design suggestions first.
Then, click the button to generate a professionally enhanced Apollo-branded presentation.
""")

uploaded_file = st.file_uploader("Upload old-format PPT", type=["pptx"])
apollo_logo_url = "https://upload.wikimedia.org/wikipedia/en/1/1e/Apollo_Hospitals_Logo.png"

def suggest_design_elements(text):
    text = text.lower()
    if "who" in text:
        return {"Layout": "Two-column", "Font": "Poppins + Segoe UI", "Color": "Blue-grey", "Visual": "WHO quote + doctor team image"}
    elif "components" in text:
        return {"Layout": "4-quadrant grid", "Font": "Arial Rounded", "Color": "Bright multicolor", "Visual": "Infographic with icons"}
    elif "india" in text:
        return {"Layout": "Map overlay", "Font": "Segoe UI", "Color": "Warm tones", "Visual": "India map infographic"}
    else:
        return {"Layout": "Standard", "Font": "Segoe UI", "Color": "Apollo theme", "Visual": "Photo + icon"}

if uploaded_file:
    old_ppt = Presentation(uploaded_file)
    preview_data = []

    for i, slide in enumerate(old_ppt.slides, start=1):
        text = " ".join([shape.text.strip() for shape in slide.shapes if shape.has_text_frame and shape.text.strip()])
        suggestion = suggest_design_elements(text)
        prompt = f"Design a slide using layout: {suggestion['Layout']}, color: {suggestion['Color']}, with {suggestion['Visual']}. Use Apollo University branding."
        preview_data.append({
            "Slide Number": i,
            "Content Preview": text[:80] + ("..." if len(text) > 80 else ""),
            "Layout": suggestion["Layout"],
            "Font": suggestion["Font"],
            "Color": suggestion["Color"],
            "Visual": suggestion["Visual"],
            "Prompt": prompt
        })

    df = pd.DataFrame(preview_data)
    st.markdown("## Step 1: AI Suggestions Table")
    st.dataframe(df, use_container_width=True)

    st.markdown("## 🔍 Detailed Slide-by-Slide Suggestions")
    for row in preview_data:
        st.markdown(f"**Slide {row['Slide Number']}**")
        st.markdown(f"• **Content Preview**: {row['Content Preview']}")
        st.markdown(f"• **Layout**: {row['Layout']}")
        st.markdown(f"• **Font**: {row['Font']}")
        st.markdown(f"• **Color Theme**: {row['Color']}")
        st.markdown(f"• **Visual Suggestion**: {row['Visual']}")
        st.markdown(f"• **Prompt**: {row['Prompt']}")
        st.markdown("---")

    if st.button("✨ Step 2: Generate Enhanced Apollo Slides"):
        new_ppt = Presentation()
        new_ppt.slide_width = Inches(13.33)
        new_ppt.slide_height = Inches(7.5)

        title_slide = new_ppt.slides.add_slide(new_ppt.slide_layouts[0])
        title_slide.shapes.title.text = "Apollo University"
        subtitle = title_slide.placeholders[1]
        subtitle.text = "Enhanced Slide Deck — AI Formatted"
        subtitle.text_frame.paragraphs[0].font.name = "Segoe UI"
        subtitle.text_frame.paragraphs[0].font.size = Pt(20)
        subtitle.text_frame.paragraphs[0].font.italic = True

        layout = new_ppt.slide_layouts[1]

        for row in preview_data:
            i = row["Slide Number"]
            old_slide = old_ppt.slides[i - 1]
            content_text = [shape.text.strip() for shape in old_slide.shapes if shape.has_text_frame and shape.text.strip()]
            split_required = len(content_text) > 5
            parts = [content_text[:len(content_text)//2], content_text[len(content_text)//2:]] if split_required else [content_text]

            new_slide = new_ppt.slides.add_slide(layout)
            try:
                old_title = old_slide.shapes.title.text.strip()
                new_slide.shapes.title.text = old_title
                title_para = new_slide.shapes.title.text_frame.paragraphs[0]
                title_para.font.name = "Poppins"
                title_para.font.size = Pt(32)
                title_para.font.bold = True
                title_para.font.color.rgb = RGBColor(0, 51, 102)
            except:
                new_slide.shapes.title.text = f"Slide {i}"

            content_box = new_slide.placeholders[1].text_frame
            content_box.clear()
            for line in parts[0]:
                para = content_box.add_paragraph()
                para.text = line
                para.font.size = Pt(18)
                para.font.name = "Segoe UI"
                para.font.color.rgb = RGBColor(0, 0, 0)

            layout_box = new_slide.shapes.add_textbox(Inches(0.5), Inches(6.0), Inches(5.5), Inches(0.7))
            layout_tf = layout_box.text_frame
            layout_tf.text = f"Layout: {row['Layout']} | Visual: {row['Visual']}
Prompt: {row['Prompt']}"

            notes_slide = new_slide.notes_slide
            notes_slide.notes_text_frame.text = row['Prompt']

            footer = new_slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(9), Inches(0.3))
            footer_tf = footer.text_frame
            footer_tf.text = "Powered by Apollo Knowledge"
            footer_para = footer_tf.paragraphs[0]
            footer_para.font.name = "Segoe UI"
            footer_para.font.size = Pt(10)
            footer_para.font.italic = True
            footer_para.font.color.rgb = RGBColor(100, 100, 100)

            try:
                response = requests.get(apollo_logo_url)
                if response.status_code == 200:
                    img_data = io.BytesIO(response.content)
                    new_slide.shapes.add_picture(img_data, Inches(8.2), Inches(6.7), width=Inches(1))
            except:
                pass

            if split_required:
                slide2 = new_ppt.slides.add_slide(layout)
                slide2.shapes.title.text = new_slide.shapes.title.text + " (contd)"
                indicator_box = slide2.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(6), Inches(0.3))
                indicator_tf = indicator_box.text_frame
                indicator_tf.text = "🔁 Continued from previous slide"
                indicator_tf.paragraphs[0].font.size = Pt(14)
                indicator_tf.paragraphs[0].font.color.rgb = RGBColor(90, 90, 90)
                part2_box = slide2.placeholders[1].text_frame
                part2_box.clear()
                for line in parts[1]:
                    para = part2_box.add_paragraph()
                    para.text = line
                    para.font.size = Pt(18)
                    para.font.name = "Segoe UI"
                    para.font.color.rgb = RGBColor(0, 0, 0)
                slide2.notes_slide.notes_text_frame.text = row['Prompt']

        ppt_io = io.BytesIO()
        new_ppt.save(ppt_io)
        ppt_io.seek(0)
        st.download_button(
            label="📥 Download Enhanced Apollo PPT",
            data=ppt_io,
            file_name="Apollo_AI_Enhanced_Slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )