import pandas as pd
import json
import os
import sys
import xml.etree.ElementTree as ET

def excel_to_json(excel_file, sheet_name, index_column, output_file):
    """
    Reads an Excel file, sets the index column, and exports the data to a JSON file.

    Args:
        excel_file (str): Path to the Excel file.
        sheet_name (str): Name of the sheet to read.
        index_column (str): Name of the column to use as the index.
        output_file (str): Name of the output JSON file.
    """
    if not os.path.exists(excel_file):
        print(f"Error: Excel file '{excel_file}' not found.")
        sys.exit(1)

    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    try:
        df = df.set_index(index_column)
    except KeyError:
        print(f"Error: Column '{index_column}' not found in sheet '{sheet_name}'.")
        sys.exit(1)
    except Exception as e:
        print(f"Error setting index: {e}")
        sys.exit(1)

    try:
        data = df.to_json(orient='records', indent=4)  # Use records for easier handling
    except Exception as e:
        print(f"Error converting to JSON: {e}")
        sys.exit(1)

    try:
        with open(output_file, 'w') as f:
            f.write(data)
        print(f"Successfully exported to {output_file}")
    except Exception as e:
        print(f"Error writing JSON file: {e}")
        sys.exit(1)
    return output_file #returning the output file


def create_pnml_files(json_file, template_file="src/templates/cargo_template.pnml", output_dir="src/cargo"):
    """
    Reads a JSON file, merges its contents into a PNML template, and saves each output as a PNML file
    in its own folder.

    Args:
        json_file (str): Path to the JSON file.
        template_file (str, optional): Path to the PNML template file. Defaults to "src/templates/cargo_template.pnml".
        output_dir (str, optional): Path to the directory where the PNML files should be created.
            Defaults to "src/cargo".
    """
    if not os.path.exists(json_file):
        print(f"Error: JSON file '{json_file}' not found.")
        sys.exit(1)

    if not os.path.exists(template_file):
        print(f"Error: PNML template file '{template_file}' not found.")
        sys.exit(1)

    try:
        with open(json_file, 'r') as f:
            data = json.load(f)
    except Exception as e:
        print(f"Error reading JSON file: {e}")
        sys.exit(1)

    try:
        with open(template_file, 'r') as f:
            template_content = f.read()
    except Exception as e:
        print(f"Error reading PNML template file: {e}")
        sys.exit(1)

    # Create the output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    for item_data in data:
        # Use folder field for the folder name
        if 'folder' not in item_data:
            print(f"Error: 'folder' key not found in data: {item_data}")
            sys.exit(1)
        folder_name = item_data['folder'].replace(" ", "_")
        target_folder = os.path.join(output_dir, folder_name)
        os.makedirs(target_folder, exist_ok=True)

        try:
            root = ET.fromstring(template_content)
        except Exception as e:
            print(f"Error parsing PNML template as XML.  Template content:\n{template_content}\nError:", e)
            sys.exit(1)

        def update_xml_template(element, data_dict):
            """Recursively updates the text of XML elements based on the data dictionary."""
            for child in element:
                if child.text is not None:
                    for key, val in data_dict.items():
                         if key in child.tag:
                            child.text = str(val)
                update_xml_template(child, data_dict)

        # Create a copy of the template XML for each item
        updated_tree = ET.ElementTree(ET.fromstring(template_content))
        updated_root = updated_tree.getroot()
        update_xml_template(updated_root, item_data)

        # Write the updated XML to a new PNML file, using the folder name
        pnml_file_path = os.path.join(target_folder, f"{folder_name}.pnml")
        updated_tree.write(pnml_file_path, encoding="UTF-8", xml_declaration=True)
        print(f"Successfully created PNML file: {pnml_file_path}")



if __name__ == "__main__":
    excel_file = 'docs/otis.xlsx'
    sheet_name = 'cargo'
    index_column = 'cargo_item_name'
    output_file = 'lib/cargo.json'

    json_file = excel_to_json(excel_file, sheet_name, index_column, output_file)  # Get the JSON file path
    create_pnml_files(json_file)
