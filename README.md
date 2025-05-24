# 🎯 Docs2PPT — From Document to Stunning PPT in Seconds with AI

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
   - Maintains full control over fonts, positioning, and aesthetics(for now it is only one)

---

## 🧰 Tech Stack

| Tool                  | Purpose                                 |
| --------------------- | --------------------------------------- |
| `python-pptx`         | Generate PowerPoint presentations       |
| `google-generativeai` | Summarize and chunk content smartly     |
| `python-docx`         | Extract text from DOCX files            |
| `PyMuPDF`             | Extract text/images from pdf files      |
| `pdfplumber`          | Extractr tables from the pdf files      |
| `dotenv`              | For securing the api keys and variables |
| `json`                | Intermediate format for slides          |

---

## 🖼️ Sample Workflow

### Go to `version0` folder and do the following steps

```bash
# Download the required packages in virtual environment(use python10 for better compatibility with packages)
pip install -r requirements.txt

# Step 2: Generate slides(local host)
python app.py
```

### Leave your thought here...

- https://www.linkedin.com/in/abhipatel2005/
