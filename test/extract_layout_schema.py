import json
import argparse

def convert_business_schema(input_path, output_path):
    """
    Convert the business_schema.json format to a simplified format
    that works better with the presentation generator
    """
    # Load the input schema
    with open(input_path, 'r') as f:
        input_schema = json.load(f)
    
    # Create the converted schema
    output_schema = []
    
    for layout in input_schema:
        layout_index = layout["layout_index"]
        layout_name = layout.get("layout", f"layout_{layout_index}")
        
        # Create the layout entry
        layout_entry = {
            "layout_index": layout_index,
            "layout": layout_name,
            "placeholders": layout["placeholders"]
        }
        
        output_schema.append(layout_entry)
    
    # Save to output file
    with open(output_path, 'w', encoding="utf-8") as f:
        json.dump(output_schema, f, indent=2)
    
    print(f"Converted schema saved to {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert business schema to presentation schema")
    parser.add_argument("--input", required=True, help="Path to the business schema JSON file")
    parser.add_argument("--output", required=True, help="Path for the output schema JSON file")
    
    args = parser.parse_args()
    convert_business_schema(args.input, args.output)