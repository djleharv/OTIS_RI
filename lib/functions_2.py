#from lib import dictionaries
import pandas as pd
import json, copy, itertools, codecs, os, shutil

import json

def ExportToJSON(dictionary, target_file):
    """
    Exports a Python dictionary to a JSON file with indentation for readability.

    Args:
        dictionary (dict): The dictionary to be exported.
        target_file (str): The path to the JSON file where the dictionary will be saved.
    """
    try:
        with open(target_file, 'w', encoding='utf-8') as fp:
            json.dump(dictionary, fp, indent=4, ensure_ascii=False)
        print(f"JSON file created at {target_file}")
    except TypeError as e:
        print(f"Error: Could not serialize dictionary to JSON. {e}")
    except IOError as e:
        print(f"Error: Could not write to file '{target_file}'. {e}")
        
import json

def LoadJSON(target_file):
    """
    Loads data from a JSON file into a Python dictionary.

    Args:
        target_file (str): The path to the JSON file to be loaded.

    Returns:
        dict: The dictionary loaded from the JSON file, or None if an error occurs.
    """
    try:
        with open(target_file, 'r', encoding='utf-8') as file:
            data = json.load(file)
        return data
    except FileNotFoundError:
        print(f"Error: File not found at '{target_file}'.")
        return None
    except json.JSONDecodeError as e:
        print(f"Error: Could not decode JSON from '{target_file}'. {e}")
        return None
    except IOError as e:
        print(f"Error: Could not open or read file '{target_file}'. {e}")
        return None
        
def CreateCargoJSON():
    """
    Reads cargo data from an Excel spreadsheet, transforms it into a dictionary,
    and exports it as a JSON file.
    """
    excel_filepath = 'docs/otis.xlsx'
    sheet_name = 'cargo'
    output_filepath = 'lib/cargo.json'

    try:
        # Check if the Excel file exists
        if not os.path.exists(excel_filepath):
            print(f"Error: Excel file not found at '{excel_filepath}'.")
            return

        # Convert excel spreadsheet into dataframe
        df_cargo = pd.read_excel(excel_filepath, sheet_name=sheet_name)

        # Check if the expected column exists
        if 'cargo_item_name' not in df_cargo.columns:
            print(f"Error: Column 'cargo_item_name' not found in sheet '{sheet_name}' of '{excel_filepath}'.")
            return

        cargo = df_cargo.set_index('cargo_item_name').T.to_dict('dict')

        # Ensure the output directory exists
        output_dir = os.path.dirname(output_filepath)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        ExportToJSON(cargo, output_filepath)

    except FileNotFoundError:
        print(f"Error: Excel file not found at '{excel_filepath}'.")
    except KeyError:
        print(f"Error: Sheet '{sheet_name}' not found in '{excel_filepath}'.")
    except ValueError as e:
        print(f"Error reading Excel file '{excel_filepath}': {e}")
    except Exception as e:
        print(f"An unexpected error occurred during CreateCargoJSON: {e}")
    
def CreateCargoPNMLs():
    cargo = LoadJSON('lib/cargo.json')
    if cargo is None:
        return

    active_cargo_data = {name: data for name, data in cargo.items() if data["include"] == True}
    active_cargo_folders = {data["folder"] for data in active_cargo_data.values()}

    folder = './src/cargo/'
    if os.path.exists(folder):
        shutil.rmtree(folder)
    os.makedirs(folder)

    try:
        with open("./src/templates/cargo_template.pnml", "rt") as template_file:
            template_content = template_file.readlines()
    except FileNotFoundError:
        print("Error: Template file not found.")
        return

    for cargo_name, cargo_data in cargo.items():
        cargo_folder = os.path.join("./src/cargo", cargo_name)
        os.makedirs(cargo_folder, exist_ok=True)
        pnml_filepath = os.path.join(cargo_folder, f"{cargo_name}.pnml")

        try:
            with open(pnml_filepath, "wt") as current_cargo_file:
                for line in template_content:
                    current_cargo_file.write(line.replace('_name_', f'_{cargo_name}'))

            with open(pnml_filepath, 'r+') as pnml_file:
                data = pnml_file.read()
                replacements = {
                    '_cargo_icon_x_': str(cargo_data["cargo_icon_x"]),
                    '_cargo_icon_y_': str(cargo_data["cargo_icon_y"]),
                    '_cargo_ID_': str(cargo_data["cargo_ID"]),
                    '_cargo_colour_number_': str(cargo_data["cargo_colour_number"]),
                    '_town_growth_effect_': str(cargo_data["town_growth_effect"]),
                    '_town_growth_multiplier_': str(cargo_data["town_growth_multiplier"]),
                    '_is_freight_': str(cargo_data["is_freight"]),
                    '_string_': str(cargo_data["string"]),
                    '_cargo_label_': str(cargo_data["cargo_label"]),
                    '_capacity_multiplier_': str(cargo_data["capacity_multiplier"]),
                    '_cargo_weight_': str(cargo_data["cargo_weight"]),
                    '_cargo_classes_': str(cargo_data["cargo_classes"]),
                    '_penalty_lower_bound_': str(cargo_data["penalty_lower_bound"]),
                    '_single_penalty_length_': str(cargo_data["single_penalty_length"]),
                    '_price_factor_': str(cargo_data["price_factor"]),
                }
                for old, new in replacements.items():
                    data = data.replace(old, new)
                pnml_file.seek(0)
                pnml_file.write(data)
                pnml_file.truncate()

            with open(pnml_filepath, 'r') as read_file:
                lines = [line for line in read_file if 'none' not in line]
            with open(pnml_filepath, 'w') as write_file:
                write_file.writelines(lines)

        except IOError as e:
            print(f"Error processing cargo item {cargo_name}: {e}")
        except KeyError as e:
            print(f"Error: Missing key '{e}' in cargo data for {cargo_name}")

    # MERGE THE ITEMS
    cargo_pnml_path = "./src/cargo.pnml"
    output_content = ['// Cargo pnml files']
    output_content.extend([f'#include "src/cargo/{folder}/{folder}.pnml"' for folder in active_cargo_folders])
    output_content.append('')

    for cargo_name in active_cargo_data.keys():  
        filepath = os.path.join("./src/cargo", cargo_name, f"{cargo_name}.pnml")
        try:
            with open(filepath, 'r', encoding='utf8') as cargo_pnml:
                output_content.append(cargo_pnml.read())
        except FileNotFoundError:
            print(f"Warning: File not found during merge: {filepath}")
        except IOError as e:
            print(f"Error reading file during merge {filepath}: {e}")

    try:
        with open(cargo_pnml_path, 'w', encoding='utf-8') as processed_pnml_file:
            processed_pnml_file.write('\n'.join(output_content))
        print("Cargo PNMLs Created")
    except IOError as e:
        print(f"Error writing final merged file: {e}")
            
def CreateCargoLangFiles():
    cargo = LoadJSON('lib/cargo.json')
    if cargo is None:
        return

    active_cargo_data = {name: data for name, data in cargo.items() if data["include"] == True}
    active_cargo_folders = {data["folder"] for data in active_cargo_data.values()}

    folder = './src/cargo/'
    if not os.path.exists(folder):
        os.makedirs(folder)

    try:
        with open("./src/templates/cargo_lang_template.lng", "rt") as template_file:
            template_content = template_file.readlines()
    except FileNotFoundError:
        print("Error: Template file not found.")
        return

    for cargo_name, cargo_data in cargo.items():
        cargo_folder = os.path.join("./src/cargo", cargo_name)
        os.makedirs(cargo_folder, exist_ok=True)
        lng_filepath = os.path.join(cargo_folder, f"{cargo_name}.lng")

        try:
            with open(lng_filepath, "wt") as current_cargo_file:
                for line in template_content:
                    modified_line = line.replace('_name_', f'_{cargo_name.upper()}')
                    current_cargo_file.write(modified_line)  # Use the modified_line here

            try:
                with open(lng_filepath, 'r+') as lng_file:
                    data = lng_file.read()
                    replacements = {
                        '_string_': str(cargo_data["string"]),
                        '_str_cargo name_': str(cargo_data["str_cargo_name"]),
                        '_str_cargo_CID_': str(cargo_data["str_cargo_CID"]),
                        '_str_cargo_units_': str(cargo_data["str_cargo_units"]),
                        '_str_cargo_short_units_': str(cargo_data["str_cargo_short_units"]),
                    }
                    for old, new in replacements.items():
                        data = data.replace(old, new)
                    lng_file.seek(0)
                    lng_file.write(data)
                    lng_file.truncate()
                    
            except IOError as e:
                print(f"Error processing cargo item {cargo_name}: {e}")
            except KeyError as e:
                print(f"Error: Missing key '{e}' in cargo data for {cargo_name}")

            with open(lng_filepath, 'r') as read_file:
                lines = [line for line in read_file if 'none' not in line]
            with open(lng_filepath, 'w') as write_file:
                write_file.writelines(lines)
        
        except IOError as e: # Catch errors from the first with open
            print(f"Error during initial write to {lng_filepath}: {e}")
        except Exception as e: # Catch any other potential errors in the outer try
            print(f"An unexpected error occurred while processing {cargo_name}: {e}")
       

    # MERGE THE ITEMS
    cargo_lng_path = "./src/cargo.lng"
    output_content = ['# Cargo lng files']
    output_content.extend([f'#include "src/cargo/{folder}/{folder}.lng"' for folder in active_cargo_folders])
    output_content.append('')

    for cargo_name in active_cargo_data.keys():
        filepath = os.path.join("./src/cargo", cargo_name, f"{cargo_name}.lng")
        try:
            with open(filepath, 'r', encoding='utf8') as cargo_lng:
                output_content.append(cargo_lng.read())
        except FileNotFoundError:
            print(f"Warning: File not found during merge: {filepath}")
        except IOError as e:
            print(f"Error reading file during merge {filepath}: {e}")

    try:
        with open(cargo_lng_path, 'w', encoding='utf-8') as processed_lng_file:
            processed_lng_file.write('\n'.join(output_content))
        print("Cargo Lang File created")
    except IOError as e:
        print(f"Error writing final merged file: {e}")
        
