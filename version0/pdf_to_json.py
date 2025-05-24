import os
import fitz  # PyMuPDF
import google.generativeai as genai
import pdfplumber
from docx import Document
import pandas as pd
import json
import time
from dotenv import load_dotenv
import zipfile
import shutil
import re
from uuid import uuid4

# Load environment variables
load_dotenv()
GOOGLE_API_KEY = os.getenv("GEMINI_API_KEY")
genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel("gemini-2.0-flash")

# Retry-safe Gemini content generation
def safe_generate_response(prompt, retries=3, delay=5):
    for attempt in range(retries):
        try:
            return model.generate_content(prompt)
        except Exception as e:
            print(f"⚠️ Gemini API error: {e} (attempt {attempt+1}/{retries})")
            time.sleep(delay)
    return None

# Extract text, tables, and images from PDF
# def extract_pdf_content(pdf_path, image_output_dir="pdf_images"):
#     os.makedirs(image_output_dir, exist_ok=True)
#     doc = fitz.open(pdf_path)
#     content_blocks = []

#     with pdfplumber.open(pdf_path) as plumber_pdf:
#         for page_num in range(len(doc)):
#             page = doc[page_num]
#             plumber_page = plumber_pdf.pages[page_num]

#             blocks = page.get_text("dict")["blocks"]
#             text_blocks = []

#             for b in blocks:
#                 if "lines" in b:
#                     y = b["bbox"][1]
#                     text = "\n".join(["".join([span["text"] for span in line["spans"]]) for line in b["lines"]])
#                     if text.strip():
#                         text_blocks.append({"type": "text", "content": text.strip(), "y": y})

#             tables = plumber_page.find_tables()
#             for table in tables:
#                 try:
#                     table_data = table.extract()
#                     if not table_data or all(not any(cell for cell in row) for row in table_data):
#                         continue
#                     df = pd.DataFrame(table_data[1:], columns=table_data[0]) if len(table_data) > 1 else pd.DataFrame(table_data)
#                     markdown_table = df.to_markdown(index=False)
#                     text_blocks.append({"type": "table", "content": markdown_table, "y": table.bbox[1]})
#                 except Exception:
#                     continue

#             text_blocks.sort(key=lambda b: b["y"])
#             combined_text = "\n\n".join([b["content"] for b in text_blocks])

#             image_path = None
#             images = page.get_images(full=True)
            
#             # Inside extract_pdf_content
#             if images:
#                 xref = images[0][0]
#                 pix = fitz.Pixmap(doc, xref)
#                 if pix.n > 4:
#                     pix = fitz.Pixmap(fitz.csRGB, pix)
#                 unique_id = str(uuid4())[:8]
#                 img_filename = f"page_{page_num+1}_{unique_id}.png"
#                 image_path = os.path.join(image_output_dir, img_filename)
#                 pix.save(image_path)
#                 pix = None

#             content_blocks.append({
#                 "text": combined_text.strip(),
#                 "image_path": image_path
#             })

#     return content_blocks

def extract_pdf_content(pdf_path, image_output_dir="pdf_images"):
    # Extract session_id from the path (assuming uploads/session_id/filename.pdf)
    parts = pdf_path.split(os.sep)
    session_id = parts[-2]  # Second-to-last part should be the session ID
    
    # Create session-specific directory for images
    session_image_dir = os.path.join(image_output_dir, session_id)
    os.makedirs(session_image_dir, exist_ok=True)
    
    doc = fitz.open(pdf_path)
    content_blocks = []

    with pdfplumber.open(pdf_path) as plumber_pdf:
        for page_num in range(len(doc)):
            page = doc[page_num]
            plumber_page = plumber_pdf.pages[page_num]

            blocks = page.get_text("dict")["blocks"]
            text_blocks = []

            for b in blocks:
                if "lines" in b:
                    y = b["bbox"][1]
                    text = "\n".join(["".join([span["text"] for span in line["spans"]]) for line in b["lines"]])
                    if text.strip():
                        text_blocks.append({"type": "text", "content": text.strip(), "y": y})

            tables = plumber_page.find_tables()
            for table in tables:
                try:
                    table_data = table.extract()
                    if not table_data or all(not any(cell for cell in row) for row in table_data):
                        continue
                    df = pd.DataFrame(table_data[1:], columns=table_data[0]) if len(table_data) > 1 else pd.DataFrame(table_data)
                    markdown_table = df.to_markdown(index=False)
                    text_blocks.append({"type": "table", "content": markdown_table, "y": table.bbox[1]})
                except Exception:
                    continue

            text_blocks.sort(key=lambda b: b["y"])
            combined_text = "\n\n".join([b["content"] for b in text_blocks])

            image_path = None
            images = page.get_images(full=True)
            
            # Store images in session-specific directory with UUID
            if images:
                xref = images[0][0]
                pix = fitz.Pixmap(doc, xref)
                if pix.n > 4:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                unique_id = str(uuid4())[:8]  # Keep using UUID for uniqueness
                img_filename = f"page_{page_num+1}_{unique_id}.png"
                image_path = os.path.join(session_image_dir, img_filename)
                pix.save(image_path)
                pix = None

            content_blocks.append({
                "text": combined_text.strip(),
                "image_path": image_path
            })

    return content_blocks

# Extract text, tables, and images from PDF
def extract_docx_content(docx_path, output_dir="docx_extracted"):
    # Setup paths
    image_output_dir = os.path.join(output_dir, "images")
    os.makedirs(image_output_dir, exist_ok=True)

    # Unzip DOCX to get images
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        docx_zip.extractall(output_dir)

    media_folder = os.path.join(output_dir, "word", "media")
    image_rel_map = {}

    # Load document relationships to map image rId to filenames
    rels_path = os.path.join(output_dir, "word", "_rels", "document.xml.rels")
    if os.path.exists(rels_path):
        with open(rels_path, "r", encoding="utf-8") as f:
            rels_data = f.read()
            for match in re.findall(r'Id="(rId\d+)"[^>]+Target="media/([^"]+)"', rels_data):
                image_rel_map[match[0]] = os.path.join("word", "media", match[1])

    # Load document content
    doc = Document(docx_path)
    content_blocks = []

    # For extracting embedded image relationship IDs
    for para in doc.paragraphs:
        block = {}
        text = para.text.strip()

        # Check for image in runs
        for run in para.runs:
            if "graphic" in run._element.xml:
                embed_match = re.search(r'embed="(rId\d+)"', run._element.xml)
                if embed_match:
                    r_id = embed_match.group(1)
                    if r_id in image_rel_map:
                        source_img_path = os.path.join(output_dir, image_rel_map[r_id])
                        img_filename = os.path.basename(source_img_path)
                        target_img_path = os.path.join(image_output_dir, img_filename)

                        # Copy image to output/images folder
                        os.makedirs(os.path.dirname(target_img_path), exist_ok=True)
                        shutil.copy(source_img_path, target_img_path)

                        block["image_path"] = os.path.relpath(target_img_path, start=output_dir).replace("\\", "/")

        if text:
            block["text"] = text

        if block:
            content_blocks.append(block)

    return content_blocks

# Convert Gemini response to JSON
def convert_gemini_response_to_list(response):
    try:
        text_content = ""
        if hasattr(response, 'candidates'):
            for candidate in response.candidates:
                if hasattr(candidate, 'content') and candidate.content.parts:
                    for part in candidate.content.parts:
                        if hasattr(part, 'text'):
                            text_content += part.text

        if text_content:
            if "```json" in text_content:
                json_string = text_content.split("```json")[1].split("```")[0].strip()
            else:
                json_string = text_content.strip()
            return json.loads(json_string)

    except Exception as e:
        print(f"Error parsing Gemini response: {e}")
    return None

# Break PDF content into chunks
def chunk_content(content_blocks, chunk_size=3):
    return [content_blocks[i:i + chunk_size] for i in range(0, len(content_blocks), chunk_size)]

# Generate slides from chunks
# def generate_slide_data(content_blocks):
#     all_slides = []
#     chunks = chunk_content(content_blocks)

#     for idx, chunk in enumerate(chunks):
#         sections_text = "\n\n".join([
#             f"Page {i + 1}:\n{block['text']}" +
#             (f"\n[IMAGE_PATH: {block['image_path']}]" if block['image_path'] else "")
#             for i, block in enumerate(chunk)
#         ])

#         prompt_template = """
#         You are a presentation expert.

#         Convert the following document into a Microsoft PowerPoint presentation using various layouts based on content type.

#         Supported layouts:
#         - Title Only
#         - Title Slide
#         - Section Header
#         - Title and Content
#         - Two Content
#         - Comparison
#         - Content/Image with Caption
#         - Title with Table

#         Return a JSON list of slides like this:
#         [
#           {{
#             "layout": "title_only",
#             "title": "max 60 characters"
#           }},
#           {{
#             "layout": "title_slide",
#             "title": "max 60 char",
#             "sub-heading": "sub-heading(max 250 characters)"
#           }},
#           {{
#             "layout": "title_and_content",
#             "title": "title(max 65 characters)",
#             "content": "approax 1250 character, add \\n for new line"
#           }},
#           {{
#             "layout": "two_content",
#             "title": "max 65 characters",
#             "content": "450 to 460 max character, add \\n for new line"
#           }},
#           {{
#             "layout": "section_header",
#             "title": "max 60 characters",
#             "sub_heading": "270 character, add \\n for new line"
#           }},
#           {{
#             "layout": "comparison",
#             "title": "max 60 characters",
#             "left_content": {{
#               "title": "max 36 characters",
#               "content": "max 360 characters, add \\n for new line"
#             }},
#             "right_layout": {{
#               "title": "max 36 characters",
#               "content": "max 360 characters, add \\n for new line"
#             }}
#           }},
#           {{
#             "layout": "content_with_caption",
#             "content": {{
#               "title": "max 60 chracters if the sentence exceed 30 character enter new line",
#               "content": "max 630 characters if sentence exceed 45 charcters use new line"
#             }},
#             "chart/smart3D_icon": "properties of this thing goes here"
#           }},
#           {{
#             "layout": "image_with_caption",
#             "image_path": "image_path_goes_here",
#             "title": "max 60 chracters",
#             "content": "max 250 characters, add \\n for new line"
#           }},
#           {{
#             "layout": "title_with_table",
#             "title": "Budget Status",
#             "table": {{
#               "headers": ["Category", "Budgeted Amount", "Actual Spend", "Variance"],
#               "rows": [
#                 ["Hardware", "$150,000", "$140,000", "$10,000"],
#                 ["Software", "$100,000", "$95,000", "$5,000"],
#                 ["Labor", "$200,000", "$210,000", "-$10,000"],
#                 ["Training", "$50,000", "$45,000", "$5,000"]
#               ]
#             }}
#           }}
#         ]

#         Only include image_path if it was explicitly mentioned as [IMAGE_PATH: ...] in the source.

#         Here is the source:
#         \"\"\"{content}\"\"\"
# """

#         prompt = prompt_template.format(content=sections_text.replace("{", "{{").replace("}", "}}"))
#         prompt = prompt.replace("\n", " ")
#         print(f"🔹 Processing chunk {idx + 1}/{len(chunks)}...")
#         response = safe_generate_response(prompt)
#         slides = convert_gemini_response_to_list(response)
#         if slides:
#             all_slides.extend(slides)

#     return all_slides
def generate_slide_data(content_blocks):
    all_slides = []
    chunks = chunk_content(content_blocks)

    for idx, chunk in enumerate(chunks):
        sections_text = ""
        
        for i, block in enumerate(chunk):
            # Check if 'text' exists in the block
            if "text" in block:
                sections_text += f"Page {i + 1}:\n{block['text']}"
            
            # Check if 'image_path' exists in the block
            if "image_path" in block and block['image_path']:
                sections_text += f"\n[IMAGE_PATH: {block['image_path']}]"

        prompt_template = """
        You are a presentation expert.

        Convert the following document into a Microsoft PowerPoint presentation using various layouts based on content type.

        Supported layouts:
        - Title Only
        - Title Slide
        - Section Header
        - Title and Content
        - Two Content
        - Comparison
        - Content/Image with Caption
        - Title with Table

        Return a JSON list of slides like this:
        [
          {{
            "layout": "title_only",
            "title": "max 60 characters"
          }},
          {{"layout": "title_slide", "title": "max 60 char", "sub-heading": "sub-heading(max 250 characters)"}},
          {{"layout": "title_and_content", "title": "title(max 65 characters)", "content": "approax 1250 character, add \\n for new line"}},
          {{"layout": "two_content", "title": "max 65 characters", "content": "450 to 460 max character, add \\n for new line"}},
          {{"layout": "section_header", "title": "max 60 characters", "sub_heading": "270 character, add \\n for new line"}},
          {{"layout": "comparison", "title": "max 60 characters", "left_content": {{"title": "max 36 characters", "content": "max 360 characters, add \\n for new line"}}, "right_layout": {{"title": "max 36 characters", "content": "max 360 characters, add \\n for new line"}}}},
          {{"layout": "content_with_caption", "content": {{"title": "max 60 chracters if the sentence exceed 30 character enter new line", "content": "max 630 characters if sentence exceed 45 charcters use new line"}}}},
          {{"layout": "image_with_caption", "image_path": "image_path_goes_here", "title": "max 60 chracters", "content": "max 250 characters, add \\n for new line"}},
          {{"layout": "title_with_table", "title": "Budget Status", "table": {{"headers": ["Category", "Budgeted Amount", "Actual Spend", "Variance"], "rows": [["Hardware", "$150,000", "$140,000", "$10,000"], ["Software", "$100,000", "$95,000", "$5,000"], ["Labor", "$200,000", "$210,000", "-$10,000"], ["Training", "$50,000", "$45,000", "$5,000"]]}}}}
        ]

        Only include image_path if it was explicitly mentioned as [IMAGE_PATH: ...] in the source.

        Here is the source:
        \"\"\"{content}\"\"\" 
        """

        prompt = prompt_template.format(content=sections_text.replace("{", "{{").replace("}", "}}"))
        prompt = prompt.replace("\n", " ")
        print(f"🔹 Processing chunk {idx + 1}/{len(chunks)}...")
        response = safe_generate_response(prompt)
        slides = convert_gemini_response_to_list(response)
        if slides:
            all_slides.extend(slides)

    return all_slides

# Main pipeline to generate JSON
# def convert_pdf_to_slide_json(pdf_path, output_json_path="slides.json"):
#     print("📄 Extracting text & images...")
#     content = extract_pdf_content(pdf_path)

#     print("🧠 Asking Gemini to build slides...")
#     slide_data = generate_slide_data(content)

#     print("💾 Saving slide data to JSON...")
#     with open(output_json_path, "w", encoding="utf-8") as f:
#         json.dump(slide_data, f, indent=2, ensure_ascii=False)

#     print("✅ Done.")
def convert_file_to_slide_json(file_path, output_json_path="slides.json"):
    print("📄 Extracting content...")
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        content = extract_pdf_content(file_path)
    elif ext == ".docx":
        content = extract_docx_content(file_path)  # DOCX content extraction
    else:
        raise ValueError("Unsupported file format. Only PDF and DOCX are supported.")

    print("🧠 Asking Gemini to build slides...")
    slide_data = generate_slide_data(content)

    print("💾 Saving slide data to JSON...")
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(slide_data, f, indent=2, ensure_ascii=False)

    print("✅ Done.")

# if __name__ == "__main__":
#     file_path = "D:\\temp\\uploads\\Assignment_IPDC.docx"  # or a DOCX file path
#     convert_file_to_slide_json(file_path)
