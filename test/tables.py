import re

def clean_and_convert_to_markdown(raw_text):
    lines = raw_text.strip().splitlines()
    tables = []
    current_table = []
    current_section = ""
    in_table = False

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Detect new section title
        if re.search(r'(Short Questions|Descriptive Questions)', line, re.IGNORECASE):
            if current_table:
                tables.append((current_section, current_table))
                current_table = []
            current_section = line
            in_table = True
            continue

        # Check for a new question row
        if re.match(r'^\d+[\).]?\s+', line):
            current_table.append(line)
        elif current_table:
            # Likely a continuation of the last question
            current_table[-1] += " " + line

    if current_table:
        tables.append((current_section, current_table))

    markdown_tables = []

    for section, rows in tables:
        markdown = f"\n### {section.strip()}\n\n"
        markdown += "| Sr. No | Question | Marks |\n"
        markdown += "|--------|----------|-------|\n"

        for row in rows:
            # Attempt to extract Sr. No, Question, Marks
            match = re.match(r'^(\d+)[\).]?\s*(.+?)\s*(?:[-–—]\s*(\d{1,2}))?$', row)
            if match:
                sr_no = match.group(1)
                question = match.group(2).strip()
                marks = match.group(3) or "01"
                markdown += f"| {sr_no} | {question} | {marks} |\n"
            else:
                # Fallback: put the whole line as question
                markdown += f"| ? | {row.strip()} | ? |\n"

        markdown_tables.append(markdown)

    return "\n".join(markdown_tables)


if __name__ == "__main__":
    with open("input.txt", "r", encoding="utf-8") as f:
        raw = f.read()

    markdown = clean_and_convert_to_markdown(raw)

    with open("output.md", "w", encoding="utf-8") as f:
        f.write(markdown)

    print("✅ Clean Markdown exported to output.md")