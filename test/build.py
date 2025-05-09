#!/usr/bin/env python3
import os
import sys
import json
import argparse
import subprocess
from pathlib import Path

def log(msg):
    print(f"[LOG] {msg}")

def error(msg):
    print(f"[ERROR] {msg}", file=sys.stderr)
    return 1

def extract_schema(args):
    """Extract layout schema from a PowerPoint template"""
    if not os.path.exists(args.pptx):
        return error(f"Template file not found: {args.pptx}")
    
    # Check if python-pptx is installed
    try:
        import pptx
    except ImportError:
        return error("python-pptx is not installed. Please run: pip install python-pptx")
    
    # Import the extract_layout_schema function
    try:
        from extract_layout_schema import extract_layout_schema
        extract_layout_schema(args.pptx, args.output)
        log(f"✅ Schema extracted to {args.output}")
        return 0
    except Exception as e:
        return error(f"Failed to extract schema: {e}")

def convert_schema(args):
    """Convert between schema formats"""
    if not os.path.exists(args.input):
        return error(f"Input schema file not found: {args.input}")
    
    try:
        from updated_schema_converter import convert_business_schema
        convert_business_schema(args.input, args.output)
        log(f"✅ Schema converted to {args.output}")
        return 0
    except Exception as e:
        return error(f"Failed to convert schema: {e}")

def create_sample_content(args):
    """Create a sample content JSON file"""
    output_path = args.output
    
    sample_content = [
        {
            "layout": "title_slide",
            "title": "Presentation Title",
            "subtitle": "Your Name or Company"
        },
        {
            "layout": "section_header",
            "title": "Section Title"
        },
        {
            "layout": "title_and_content",
            "title": "Title and Content",
            "content": "• First bullet point\n• Second bullet point\n• Third bullet point"
        },
        {
            "layout": "two_content",
            "title": "Two Column Content",
            "left_content": "Left column content",
            "right_content": "Right column content"
        }
    ]
    
    with open(output_path, 'w') as f:
        json.dump(sample_content, f, indent=2)
    
    log(f"✅ Sample content created at {output_path}")
    return 0

def build_presentation(args):
    """Build a presentation from schema and content JSON"""
    if not os.path.exists(args.schema):
        return error(f"Schema file not found: {args.schema}")
    
    if not os.path.exists(args.content):
        return error(f"Content file not found: {args.content}")
    
    # Check if python-pptx is installed
    try:
        import pptx
    except ImportError:
        return error("python-pptx is not installed. Please run: pip install python-pptx")
    
    try:
        from enhanced_presentation_builder import build_enhanced_presentation
        success = build_enhanced_presentation(
            args.schema,
            args.content,
            args.output,
            args.images
        )
        if success:
            log(f"✅ Presentation built successfully: {args.output}")
            return 0
        else:
            return error("Failed to build presentation")
    except Exception as e:
        return error(f"Failed to build presentation: {e}")

def setup_environment(args):
    """Setup the required environment (install dependencies)"""
    dependencies = ["python-pptx"]
    
    log("Setting up environment...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + dependencies)
        log("✅ Dependencies installed successfully")
        return 0
    except subprocess.CalledProcessError as e:
        return error(f"Failed to install dependencies: {e}")

def main():
    # Create the top-level parser
    parser = argparse.ArgumentParser(description="PowerPoint Presentation CLI Tool")
    subparsers = parser.add_subparsers(dest="command", help="Command to execute")
    
    # Parser for the "extract" command
    extract_parser = subparsers.add_parser("extract", help="Extract layout schema from a PowerPoint template")
    extract_parser.add_argument("--pptx", required=True, help="Path to the PowerPoint template")
    extract_parser.add_argument("--output", default="template_schema.json", help="Output JSON file path")
    
    # Parser for the "convert" command
    convert_parser = subparsers.add_parser("convert", help="Convert between schema formats")
    convert_parser.add_argument("--input", required=True, help="Input schema file")
    convert_parser.add_argument("--output", required=True, help="Output schema file")
    
    # Parser for the "sample" command
    sample_parser = subparsers.add_parser("sample", help="Create a sample content JSON file")
    sample_parser.add_argument("--output", default="sample_content.json", help="Output JSON file path")
    
    # Parser for the "build" command
    build_parser = subparsers.add_parser("build", help="Build a presentation from schema and content")
    build_parser.add_argument("--schema", required=True, help="Schema JSON file")
    build_parser.add_argument("--content", required=True, help="Content JSON file")
    build_parser.add_argument("--output", default="output.pptx", help="Output presentation file")
    build_parser.add_argument("--images", default="assets", help="Directory containing images")
    
    # Parser for the "setup" command
    setup_parser = subparsers.add_parser("setup", help="Setup the required environment")
    
    args = parser.parse_args()
    
    # Execute the appropriate command
    if args.command == "extract":
        return extract_schema(args)
    elif args.command == "convert":
        return convert_schema(args)
    elif args.command == "sample":
        return create_sample_content(args)
    elif args.command == "build":
        return build_presentation(args)
    elif args.command == "setup":
        return setup_environment(args)
    else:
        parser.print_help()
        return 0

if __name__ == "__main__":
    sys.exit(main())