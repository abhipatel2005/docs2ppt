# 🎯 docs2ppt — From Document to Stunning PPT in Seconds

Transform your boring PDFs and DOCX files into **polished, presentation-ready slides** using AI.  
Powered by **Google Gemini** for intelligent summarization and **python-pptx** for sleek deck creation.  
Built with ❤️ for educators, students, and fast-moving teams.

---

## 🚀 What It Does

1. 📄 **Reads PDFs or DOCX files** and extracts the content (text, images, tables)
2. 🧠 **Summarizes** the content into slide-friendly sections using **Google Gemini**
3. 🛠️ Converts summaries into structured `JSON` format (slide-wise)
4. 🎨 Generates a clean, professional **PowerPoint presentation**
   - Auto-handles long text (splits slides)
   - Applies **custom themes per layout** (title, content, comparison, image, table, etc.)
   - Maintains full control over fonts, positioning, and aesthetics

---

## 🧰 Tech Stack

| Tool           | Purpose                            |
|----------------|------------------------------------|
| `python-pptx`  | Generate PowerPoint presentations  |
| `Google Gemini`| Summarize and chunk content smartly|
| `pdfminer.six` | Extract text from PDFs             |
| `python-docx`  | Extract text from DOCX files       |
| `json`         | Intermediate format for slides     |

---

## 🖼️ Sample Workflow

```bash
# Step 1: Extract & summarize
python app.py

# Step 2: Generate slides
python demo.py slides.json --output my_presentation.pptx
