import os
import fitz  # PyMuPDF
import google.generativeai as genai
from pdfminer.high_level import extract_text
import re
import json

# STEP 1: Setup Gemini Flash
GOOGLE_API_KEY = "AIzaSyAkAm--EIB-5jgoFJecpZ9iQslbXFRllUQ"
genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel("gemini-2.0-flash")

# === SAFE GEMINI RESPONSE WITH RETRIES ===
import time

def safe_generate_response(prompt, retries=3, delay=5):
    for attempt in range(retries):
        try:
            return model.generate_content(prompt)
        except Exception as e:
            print(f"⚠️ Gemini API error: {e} (attempt {attempt+1}/{retries})")
            time.sleep(delay)
    return None


# === STEP 2: Extract text & images from PDF ===
def extract_pdf_content(pdf_path, image_output_dir="pdf_images"):
    os.makedirs(image_output_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    combined_pages = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text().strip()
        image_path = None

        images = page.get_images(full=True)
        if images:
            xref = images[0][0]  # First image on page
            pix = fitz.Pixmap(doc, xref)
            if pix.n > 4:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            img_filename = f"page_{page_num+1}.png"
            img_full_path = os.path.join(image_output_dir, img_filename)
            pix.save(img_full_path)
            pix = None
            image_path = img_full_path

        combined_pages.append({
            "text": text,
            "image_path": image_path
        })

    return combined_pages

#! content extraction from the response of the json
def convert_gemini_response_to_list(response):
    try:
        # First, extract the text content from the GenerateContentResponse
        text_content = ""
        
        # Check if the response has candidates
        if hasattr(response, 'candidates') and response.candidates:
            for candidate in response.candidates:
                if hasattr(candidate, 'content') and candidate.content.parts:
                    for part in candidate.content.parts:
                        if hasattr(part, 'text'):
                            text_content += part.text
        
        # If we have text content, parse it as JSON
        if text_content:
            # Find JSON content if wrapped in code blocks
            if "```json" in text_content:
                json_start = text_content.find("```json") + 7
                json_end = text_content.rfind("```")
                json_string = text_content[json_start:json_end].strip()
            else:
                # Assume the entire text is JSON
                json_string = text_content.strip()
            
            # Parse JSON into Python objects
            json_data = json.loads(json_string)
            print(json_data)
            # return json.loads(json_string)
            return json_data
        
        else:
            print("No text content found in the response")
            return None
            
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")
        return None
    except Exception as e:
        print(f"Unexpected error: {e}")
        return None

# === STEP 3: Ask Gemini to generate rich slide layouts ===
# def generate_slide_data(content_blocks):
#     # Combine text per section, keeping track of image context
#     sections = []
#     for i, block in enumerate(content_blocks):
#         if block["text"]:
#             entry = {
#                 "page": i + 1,
#                 "text": block["text"],
#                 "image_path": block["image_path"]
#             }
#             sections.append(entry)

#     # Build prompt for Gemini
#     sections_text = "\n\n".join([
#         f"Page {sec['page']}:\n{sec['text']}" +
#         (f"\n[IMAGE_PATH: {sec['image_path']}]" if sec['image_path'] else "")
#         for sec in sections
#     ])

#     prompt = f"""
# You are a presentation expert.

# Convert the following document into a Microsoft PowerPoint presentation using various layouts based on content type.

# Supported layouts:
# - Title Slide
# - Section Header
# - Title and Content
# - Two Content
# - Comparison
# - Picture with Caption
# - Title and Table

# Return a JSON list of slides like this:
# {{
#   "layout": "Title and Content",
#   "title": "Blockchain Basics",
#   "content": ["Definition", "How it works"],
#   "image_path": "/path/to/image.png"  ← Only include this if image was present on that page
# }}

# Only include image_path if it was explicitly mentioned as [IMAGE_PATH: ...] in the source.

# Here is the source:
# \"\"\"{sections_text}\"\"\"
# """
#     response = model.generate_content(prompt)
#     print(type(response))
#     convert_gemini_response_to_list(response)
#     print(type(convert_gemini_response_to_list(response)))
#     return response.text

def chunk_content(content_blocks, chunk_size=3):
    return [content_blocks[i:i + chunk_size] for i in range(0, len(content_blocks), chunk_size)]

def generate_slide_data(content_blocks):
    all_slides = []
    chunks = chunk_content(content_blocks, chunk_size=3)  # Tune size if needed

    for idx, chunk in enumerate(chunks):
        sections_text = "\n\n".join([
            f"Page {i + 1}:\n{block['text']}" +
            (f"\n[IMAGE_PATH: {block['image_path']}]" if block['image_path'] else "")
            for i, block in enumerate(chunk)
        ])

        prompt = f"""
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
{
  {
    "layout": "title_only",
    "title": "max 60 characters"
  },
  {
    "layout": "title_slide",
    "title": "max 60 char",
    "sub-heading": "sub-heading(max 250 characters)"
  },
  {
    "layout": "title_and_content",
    "title": "title(max 65 characters)",
    "content": "approax 1250 character if sentence cross the more than 80 character enter new line"
  },
  {
    "layout": "two_content",
    "title": "max 65 characters",
    "content": "450 to 460 max character if sentence cross the more than 37 character enter new line"
  },
  {
    "layout": "section_header",
    "title": "max 60 char",
    "sub_heading": "270 character if sentence exceed the 90 character use new line"
  },
  {
    "layout": "comparison",
    "title": "max 60 char",
    "left_content": {
      "title": "max 36 characters",
      "content": "max 360 characters if sentence exceed 30 charcters use new line"
    },
    "right_layout": {
      "title": "max 36 characters",
      "content": "max 360 characters if sentence exceed 30 charcters use new line"
    }
  },
  {
    "layout": "content_with_caption",
    "content": {
      "title": "max 60 chracters if the sentence exceed 30 character enter new line",
      "content": "max 630 characters if sentence exceed 45 charcters use new line"
    },
    "chart/smart3D_icon": "properties of this thing goes here"
  },
  {
    "layout": "image_with_caption",
    "image_path": "image_path_goes_here",
    "title": "max 60 characters",
    "content": "max 300 characters"
  },
  {
    "layout": "title_with_table",
    "title": "Budget Status",
    "table": {
      "headers": ["Category", "Budgeted Amount", "Actual Spend", "Variance"],
      "rows": [
        ["Hardware", "$150,000", "$140,000", "$10,000"],
        ["Software", "$100,000", "$95,000", "$5,000"],
        ["Labor", "$200,000", "$210,000", "-$10,000"],
        ["Training", "$50,000", "$45,000", "$5,000"]
      ]
    }
  }
}

Only include image_path if it was explicitly mentioned as [IMAGE_PATH: ...] in the source.

Here is the source:
\"\"\"{sections_text}\"\"\"
"""

        print(f"🔹 Processing chunk {idx + 1}/{len(chunks)}...")
        # response = model.generate_content(prompt)
        response = safe_generate_response(prompt)
        slides = convert_gemini_response_to_list(response)

        if slides:
            all_slides.extend(slides)

    return all_slides

# === STEP 4: Run full pipeline ===
def convert_pdf_to_slide_json(pdf_path, output_json_path="slides.json"):
    print("📄 Extracting text & images...")
    content = extract_pdf_content(pdf_path)

    print("🧠 Asking Gemini to build slides...")
    slide_data = generate_slide_data(content)

    print("💾 Saving slide data to JSON...")
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(slide_data, f, indent=2, ensure_ascii=False)

    print("✅ Done.")

# === USAGE ===
pdf_path = "pdf1.pdf"
convert_pdf_to_slide_json(pdf_path)