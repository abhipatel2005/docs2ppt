import argparse
import json
from pptx import Presentation
from layout import (
    add_title_slide, add_title_only_slide, add_title_and_content_slide,
    add_section_header_slide, add_two_content_slide, add_comparison_slide,
    add_content_with_caption_slide, add_image_with_caption_slide, add_table_slide, MAX_BULLET_POINTS_PER_SLIDE
)

def create_presentation_from_json(json_data):
    """Create a presentation from JSON data."""
    prs = Presentation()
    
    for slide_data in json_data:
        layout = slide_data.get("layout", "")
        
        if layout == "title_only":
            add_title_only_slide(prs, slide_data)
        elif layout == "title_slide":
            add_title_slide(prs, slide_data)
        elif layout == "title_and_content":
            add_title_and_content_slide(prs, slide_data)
        elif layout == "two_content":
            add_two_content_slide(prs, slide_data)
        elif layout == "section_header":
            add_section_header_slide(prs, slide_data)
        elif layout == "comparison":
            add_comparison_slide(prs, slide_data)
        elif layout in "content_with_caption":
            add_content_with_caption_slide(prs, slide_data)
        elif layout in "image_with_caption":
            add_image_with_caption_slide(prs, slide_data)
        elif layout in ["title_with_table", "chart", "other_multi_media"]:
            add_table_slide(prs, slide_data)
    
    return prs

def main():
    parser = argparse.ArgumentParser(description='Generate PowerPoint from JSON')
    parser.add_argument('json_file', help='Path to the JSON file')
    parser.add_argument('--output', default='presentation.pptx', help='Output file name')
    parser.add_argument('--max-bullets', type=int, default=MAX_BULLET_POINTS_PER_SLIDE, 
                        help=f'Maximum number of bullet points per slide (default: {MAX_BULLET_POINTS_PER_SLIDE})')
    
    args = parser.parse_args()
    
    # Update global max bullets if specified
    # global MAX_BULLET_POINTS_PER_SLIDE
    # if args.max_bullets:
    #     MAX_BULLET_POINTS_PER_SLIDE = args.max_bullets
    
    try:
        with open(args.json_file, 'r', encoding='utf-8') as file:
            json_data = json.load(file)
        
        prs = create_presentation_from_json(json_data)
        prs.save(args.output)
        
        print(f"Presentation created successfully: {args.output}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()