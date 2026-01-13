import pandas as pd
import os 
import shutil 
import re
import json 

def CreateCargoTable(
    excel_filepath="docs/otis.xlsx",
    template_path="src/templates/cargo_table_template.pnml",
    output_file_path="src/cargo_table.pnml",
    sheet_name="cargo"
):
    """
    Reads cargo data from an Excel spreadsheet, formats it, and saves it
    into a PNML file using a template.

    Args:
        excel_filepath (str): The path to the Excel file.
        template_path (str): The path to the PNML template file.
        output_file_path (str): The path to the output PNML file.
        sheet_name (str): The name of the sheet in the Excel file to read.
    """
    try:
        # Read the Excel file using pandas
        xls = pd.ExcelFile(excel_filepath)
        df = xls.parse(sheet_name)

        # Filter data based on the 'include' column
        filtered_data = df[df['include'].astype(str).str.lower() == 'true']

        # Format the cargo labels as strings, adding four tabs for indentation
        cargo_labels = [f'\t\t\t\t{label}' for label in filtered_data['cargo_label'].astype(str)]

        # Join the cargo labels with commas and newlines
        cargo_table_data = ",\n".join(cargo_labels)

        # Read the template file
        try:
            with open(template_path, "r") as template_file:
                template_content = template_file.read()
        except FileNotFoundError:
            print(f"Error: Template file not found at {template_path}")
            return

        # Replace the placeholder in the template with the formatted data
        modified_content = template_content.replace("_cargo_table_", cargo_table_data)

        # Write the modified content to the output file
        os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
        with open(os.path.join("src", "cargo_table.pnml"), "w", encoding="utf-8") as outfile: # Changed output path
            outfile.write(modified_content)
        print(f"Successfully created cargo table file: {os.path.join('src', 'cargo_table.pnml')}") # changed output path in print

    except FileNotFoundError:
        print(f"Error: Excel file not found at {excel_filepath}")
        return
    except KeyError:
        print(
            f"Error: 'include' or 'cargo_label' column not found in sheet '{sheet_name}'."
        )
        return
    except Exception as e:
        print(f"An error occurred: {e}")
        return

def CreateCargoPNMLs():
    """
    Reads a specific Excel spreadsheet ('docs/otis.xlsx', sheet 'cargo'),
    filters cargo records based on an 'include' column, and merges data
    into a PNML template ('src/templates/cargo_template.pnml'), saving
    each record to a separate file in individual folders within 'src/cargo_individual/',
    and also saves the processed content for all included records into a single
    'cargo.pnml' file in the 'src/' directory.
    """
    
    # --- Integrated Helper Function ---
    def format_value_for_template(value):
        """
        Checks if a numeric value is a whole number (e.g., 123.0) and formats it
        as an integer string ("123") to remove trailing decimals.
        Retains original string representation for non-whole numbers and non-numeric types.
        (This function is nested for local use within CreateCargoPNMLs).
        """
        # 1. Check if the value is numeric and not a missing value (NaN/NaT)
        if isinstance(value, (int, float)):
            if pd.isna(value):
                return "" # Treat NaN/missing as an empty string
                
            # 2. Check if the float value is equivalent to a whole number
            # Example: 123.0 == int(123.0) is True
            # Example: 123.4 == int(123.4) is False (123.4 != 123)
            if value == int(value):
                return str(int(value)) # Convert 123.0 -> 123 -> "123"
                
        # 3. For all other cases (non-whole floats, strings, dates, etc.), convert directly to string
        return str(value)
    # --- End of Integrated Helper Function ---
    
    # Define paths and settings
    spreadsheet_path = 'docs/otis.xlsx'
    template_path = 'src/templates/cargo_template.pnml'
    include_column = 'include'
    sheet_name = 'cargo'
    output_combined_file = 'src/cargo.pnml'
    output_individual_dir = 'src/cargo/'

    # Setup directories
    if os.path.exists(output_individual_dir):
        shutil.rmtree(output_individual_dir)
    os.makedirs(output_individual_dir, exist_ok=True)
    
    # Read the spreadsheet data
    try:
        # NOTE: We do not use 'dtype' here because we want to preserve flexibility
        # and handle the formatting later, which is safer when column types vary.
        df = pd.read_excel(spreadsheet_path, sheet_name=sheet_name)
        data = df.to_dict(orient='records')
    except FileNotFoundError:
        print(f"Error: Spreadsheet not found at {spreadsheet_path}")
        return
    except ValueError:
        print(f"Error: Sheet '{sheet_name}' not found in the spreadsheet.")
        return

    # Filter data based on the 'include' column
    filtered_data = [record for record in data if str(record.get(include_column)).lower() == 'true']

    # Read the PNML template
    try:
        with open(template_path, 'r') as f:
            template_content = f.read()
    except FileNotFoundError:
        print(f"Error: Template not found at {template_path}")
        return

    combined_output = ""

    # Process each filtered record
    for record in filtered_data:
        item_name = str(record.get('cargo_item_name', 'default_item'))
        folder_name = item_name.replace(" ", "_")
        file_name = f"{item_name.replace(' ', '_')}.pnml"
        output_folder = os.path.join(output_individual_dir, folder_name)
        output_path = os.path.join(output_folder, file_name)

        os.makedirs(output_folder, exist_ok=True)

        modified_content = template_content
        
        # Apply data to the template with the new formatting logic
        for key, value in record.items():
            placeholder = f"_{key}_"
            
            # Use the dedicated formatting function to handle whole numbers correctly
            formatted_value = format_value_for_template(value)
            
            # Replace the placeholder in the template content
            modified_content = modified_content.replace(placeholder, formatted_value)

        # Write the individual file
        try:
            with open(output_path, 'w') as outfile:
                outfile.write(modified_content)
            print(f"Processed and saved: {output_path}")
        except Exception as e:
            print(f"Error writing to individual file {output_path}: {e}")

        combined_output += modified_content + "\n\n"

    # Write the combined file
    try:
        with open(output_combined_file, 'w') as outfile:
            outfile.write(combined_output.strip())
        print(f"\nCombined processed content into: {output_combined_file}")
    except Exception as e:
        print(f"Error writing combined file {output_combined_file}: {e}")

    print("Cargo PNML creation complete (both individual and combined files).")
    
def CreateCargoLNGs():
    """
    Reads a specific Excel spreadsheet ('docs/otis.xlsx', sheet 'cargo'),
    filters cargo records based on the 'include' column, and merges data
    into an LNG template ('src/templates/cargo_lang_template'), saving
    the processed content for all included records into a single
    'cargo.lng' file in the 'src/' directory, and also saves individual
    LNG files into subfolders within 'src/cargo/'.
    """

    spreadsheet_path = 'docs/otis.xlsx'
    template_path_lng = 'src/templates/cargo_lang_template.lng'
    include_column = 'include'
    sheet_name = 'cargo'
    output_combined_file_lng = 'src/cargo_lang.lng'
    output_individual_dir = 'src/cargo/'

    # Ensure the individual output directory exists
    os.makedirs(output_individual_dir, exist_ok=True)

    try:
        df = pd.read_excel(spreadsheet_path, sheet_name=sheet_name)
        data = df.to_dict(orient='records')
    except FileNotFoundError:
        print(f"Error: Spreadsheet not found at {spreadsheet_path}")
        return
    except ValueError:
        print(f"Error: Sheet '{sheet_name}' not found in the spreadsheet.")
        return

    filtered_data = [record for record in data if str(record.get(include_column)).lower() == 'true']

    try:
        with open(template_path_lng, 'r') as f:
            template_content_lng = f.read()
    except FileNotFoundError:
        print(f"Error: LNG Template not found at {template_path_lng}")
        return

    combined_output_lng = ""

    for record in filtered_data:
        item_name = str(record.get('cargo_item_name', 'default_item'))
        folder_name = item_name.replace(" ", "_")
        file_name_lng = f"{item_name.replace(' ', '_')}.lng"
        output_folder = os.path.join(output_individual_dir, folder_name)
        output_path_lng = os.path.join(output_folder, file_name_lng)

        os.makedirs(output_folder, exist_ok=True)  # Ensure folder exists

        modified_content_lng = template_content_lng
        for key, value in record.items():
            placeholder = f"_{key}_"
            modified_content_lng = modified_content_lng.replace(placeholder, str(value))

        # Write the individual LNG file
        try:
            with open(output_path_lng, 'w') as outfile:
                outfile.write(modified_content_lng)
            print(f"Processed and saved: {output_path_lng}")
        except Exception as e:
            print(f"Error writing to individual LNG file {output_path_lng}: {e}")

        combined_output_lng += modified_content_lng + "\n\n"

    # Write the combined LNG file
    try:
        with open(output_combined_file_lng, 'w') as outfile:
            outfile.write(combined_output_lng.strip())
        print(f"\nCombined processed content into: {output_combined_file_lng}")
    except Exception as e:
        print(f"Error writing combined LNG file {output_combined_file_lng}: {e}")

    print("Cargo LNG creation complete (both individual and combined files).")
 

def CreateIndustries(excel_filepath='docs/otis.xlsx', base_folder='src/industries'):
    """
    Extracts data from an Excel spreadsheet, creates JSON files, linking rows
    from the 'industries' sheet with corresponding rows from industry-specific sheets.
    Also creates a .pnml file for each industry, replacing placeholders with data from the JSON.
    The PNML file is generated by replacing placeholders in the template.  Finally,
    combines all individual industry PNML files into a single 'industries.pnml' file in src/.

    Args:
        excel_filepath (str): The path to the Excel file.
        base_folder (str): The base folder where industry folders will be created.
    """
    try:
        # Read the Excel file using pandas
        xls = pd.ExcelFile(excel_filepath)

        # Read the 'industries' sheet
        df_industries = xls.parse('industries')

        # Read all other sheets into a dictionary
        industry_sheets = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names if sheet_name != 'industries'}

        # Define the PNML template file.  This will now be dynamic.
        pnml_template_file = 'src/templates/industry_template.pnml'  # Default, will be overridden


        # Ensure the base folder exists exists, and clean it if it does
        if os.path.exists(base_folder):
            print(f"Cleaning directory: {base_folder}")
            shutil.rmtree(base_folder)  # Remove the entire directory tree
        os.makedirs(base_folder, exist_ok=True)  # Create the base folder

        # Store industry data for later processing
        industry_data = {}

        # First, process all data and store it in the industry_data dictionary
        for index, industry_row in df_industries.iterrows():
            # Check the 'include' column (case-insensitive)
            if 'include' in industry_row and str(industry_row['include']).lower() == 'true':
                industry_name = industry_row['industry_item_name']
                industry_data[industry_name] = industry_row.to_dict() # Store row

                # Create a folder for the industry
                industry_folder = os.path.join(base_folder, industry_name)
                os.makedirs(industry_folder, exist_ok=True)


                # Get the DataFrame for the industry-specific sheet
                if industry_name in industry_sheets:
                    df_industry_data = industry_sheets[industry_name]
                    accept_cargo_list = []
                    produce_cargo_list = []
                    for cargo_index, cargo_row in df_industry_data.iterrows():
                        row_data = cargo_row.to_dict()
                        # Include accept_cargo and produce_cargo only if they have non-null values
                        if 'accept_cargo' in row_data and pd.notna(row_data['accept_cargo']):
                            accept_cargo_list.append({
                                    'accept_cargo': row_data['accept_cargo'],
                                    'accept_cargo_type': row_data.get('accept_cargo_type'),
                                    'stock_num': int(row_data.get('stock_num')) if pd.notna(row_data.get('stock_num')) and isinstance(row_data.get('stock_num'), (int, float)) else None,
                                    'cons_num': int(row_data.get('cons_num')) if pd.notna(row_data.get('cons_num'))and isinstance(row_data.get('cons_num'), (int, float)) else None
                                })
                        if 'produce_cargo' in row_data and pd.notna(row_data['produce_cargo']):
                            produce_cargo_list.append({
                                    'produce_cargo': row_data['produce_cargo'],
                                    'produce_cargo_type': row_data.get('produce_cargo_type'),
                                    'prod_num': int(row_data.get('prod_num')) if pd.notna(row_data.get('prod_num')) and isinstance(row_data.get('prod_num'), (int, float)) else None,
                                    'demand_num': int(row_data.get('demand_num')) if pd.notna(row_data.get('demand_num')) and isinstance(row_data.get('prod_num'), (int, float)) else None,
                                    'bias_num': int(row_data.get('bias_num')) if pd.notna(row_data.get('bias_num')) and isinstance(row_data.get('bias_num'), (int, float)) else None
                                })

                    industry_data[industry_name]['accept_cargo_list'] = accept_cargo_list
                    industry_data[industry_name]['produce_cargo_list'] = produce_cargo_list



        # Second, calculate demand_customers *after* all industry data is processed
        for industry_name, data in industry_data.items():
            produce_cargo_list = data.get('produce_cargo_list', [])
            demand_customers = []
            for item in produce_cargo_list:
                target_cargo = item["produce_cargo"]
                demand_num = item.get("demand_num")
                if pd.notna(demand_num):
                    accepting_industries = []
                    for other_industry_name, other_data in industry_data.items():
                        if other_industry_name != industry_name:
                            other_accept_cargo_list = other_data.get('accept_cargo_list', [])
                            for other_item in other_accept_cargo_list:
                                if other_item['accept_cargo'] == target_cargo:
                                    # Check industry_pack here
                                    if data.get('industry_pack') == other_data.get('industry_pack'):
                                        accepting_industries.append(other_industry_name)
                                    break
                    demand_customers.append({
                            'produce_cargo': target_cargo,
                            'accepted_by': accepting_industries, # Store the list of accepting industries
                            'demand_num': demand_num  # Include demand_num
                    })
            industry_data[industry_name]['demand_customers'] = demand_customers

        # Third, write the JSON files
        for industry_name, data in industry_data.items():
            json_filepath = os.path.join(base_folder, industry_name, f'{industry_name}.json')
            try:
                with open(json_filepath, 'w', encoding='utf-8') as jsonfile:
                    json.dump(data, jsonfile, indent=4)
                print(f"Processed and Created JSON: {json_filepath}")
            except Exception as e:
                print(f"  Error writing JSON file: {json_filepath} - e")

        # Fourth, Create the PNML files and combine them
        combined_pnml_content = ""
        for industry_name, data in industry_data.items(): # use the updated industry_data
            pnml_filepath = os.path.join(base_folder, industry_name, f'{industry_name}.pnml')
            # Use Industry type to pick template
            industry_type = data.get('industry_type', 'industry')  # default to 'industry'
            pnml_template_file = f'src/templates/{industry_type}_industry_template.pnml'
            if not os.path.exists(pnml_template_file):
                pnml_template_file = 'src/templates/industry_template.pnml' # Fallback
            try:
                with open(pnml_template_file, 'r') as template_file, open(pnml_filepath, 'w') as pnml_file:
                    pnml_content = template_file.read()
                    # Replace placeholders with data from json_data
                    for key, value in data.items():
                        # Convert key to underscore format
                        placeholder = f'_{key}_'
                        if isinstance(value, (int, float, str, bool)):
                            pnml_content = pnml_content.replace(placeholder, str(value))

                    # Handle accept_cargo_list
                    accept_cargo_str = ""
                    accept_cargo_list = data.get('accept_cargo_list', [])
                    for item in accept_cargo_list:
                        accept_cargo_str += f'\t\t\t\taccept_cargo("{item["accept_cargo"]}"),\n'
                    pnml_content = pnml_content.replace('_accept_cargo_list_', accept_cargo_str)

                    # Handle produce_cargo_list
                    produce_cargo_str = ""
                    produce_cargo_list = data.get('produce_cargo_list', [])
                    for i, item in enumerate(produce_cargo_list):
                        produce_cargo_str += f'\t\t\t\tproduce_cargo("{item["produce_cargo"]}",0)'
                        if i < len(produce_cargo_list) - 1:
                            produce_cargo_str += ',\n'
                    pnml_content = pnml_content.replace('_produce_cargo_list_', produce_cargo_str)

                    # Handle cargo_stockpiles
                    cargo_stockpiles_str = ""
                    for item in accept_cargo_list:
                        stock_num = item.get("stock_num")
                        if pd.isna(stock_num):
                            cargo_stockpiles_str += f'\t\t\t\tSTORE_PERM (incoming_cargo_waiting("{item["accept_cargo"]}"),),\n'
                        else:
                            cargo_stockpiles_str += f'\t\t\t\tSTORE_PERM (incoming_cargo_waiting("{item["accept_cargo"]}"),{item["stock_num"]}),\n'
                    pnml_content = pnml_content.replace('_cargo_stockpiles_', cargo_stockpiles_str)

                    # Handle cargo_consumption
                    cargo_consumption_str = ""
                    execute_consumption_str = ""
                    for item in accept_cargo_list:
                        cons_num = item.get("cons_num")
                        if pd.isna(cons_num):
                            execute_consumption_str += f'\t\t\t\t{item["accept_cargo"]}: LOAD_PERM();\n'
                        else:
                            execute_consumption_str += f'\t\t\t\t{item["accept_cargo"]}: LOAD_PERM({item["cons_num"]});\n'
                    pnml_content = pnml_content.replace('_execute_consumption_', execute_consumption_str)
                    for item in accept_cargo_list:
                        cons_num = item.get("cons_num")
                        if pd.isna(cons_num):
                            cargo_consumption_str += f'\t\t\t\tSTORE_PERM ({data["industry_type"]}_{item["accept_cargo_type"]}_cargo_consumption_{item["stock_num"]}(),),\n'
                        else:
                            cargo_consumption_str += f'\t\t\t\tSTORE_PERM ({data["industry_type"]}_{item["accept_cargo_type"]}_cargo_consumption_{item["stock_num"]}(),{item["cons_num"]}),\n'
                    pnml_content = pnml_content.replace('_cargo_consumption_', cargo_consumption_str)

                    # Handle cargo_production
                    cargo_production_str = ""
                    execute_production_str = ""
                    produce_cargo_list = data.get('produce_cargo_list', [])
                    for item in produce_cargo_list:
                        prod_num = item.get("prod_num")
                        if pd.isna(prod_num):
                            execute_production_str += f'\t\t\t\t{item["produce_cargo"]}: LOAD_PERM();\n'
                        else:
                            execute_production_str += f'\t\t\t\t{item["produce_cargo"]}: LOAD_PERM({item.get("prod_num")});\n'
                    pnml_content = pnml_content.replace('_execute_production_', execute_production_str)
                    cargo_production_str = ""
                    for item in produce_cargo_list:
                        prod_num = item.get("prod_num")
                        if pd.isna(prod_num):
                            cargo_production_str += f'\t\t\t\tSTORE_PERM ({data["industry_type"]}_{item.get("produce_cargo_type")}_cargo_production(),),\n'
                        else:
                            cargo_production_str += f'\t\t\t\tSTORE_PERM ({data["industry_type"]}_{item.get("produce_cargo_type")}_cargo_production_{item["prod_num"]}(),{item["prod_num"]}),\n'
                    pnml_content = pnml_content.replace('_cargo_production_', cargo_production_str)

                    # Handle supply competition
                    pnml_content = pnml_content.replace('_supply_competition_', f'\t\t\t\tSTORE_PERM (industry_count({industry_name}),2),')

                    # Handle demand_customers
                    demand_customers_str = ""
                    demand_customers_list = data.get('demand_customers', []) # Get the list
                    for item in demand_customers_list:
                        accepted_by_list = item['accepted_by']
                        demand_num = item['demand_num']
                        if accepted_by_list: # Check if the list is not empty
                            accepted_by_list_str = " + ".join([f'industry_count({industry})' for industry in accepted_by_list])
                            demand_customers_str += f'\t\t\t\tSTORE_PERM ({accepted_by_list_str},\n\t\t\t\t\t\t\t{demand_num}),\n'
                        # else:  # Removed the else condition, so nothing is added if the list is empty.
                    pnml_content = pnml_content.replace('_demand_customers_', demand_customers_str)

                    # Handle production bias
                    production_bias_str = ""
                    for item in produce_cargo_list:
                        bias_num = item.get("bias_num")
                        if  pd.isna(bias_num):  # Check if bias_num exists and is not None
                            production_bias_str = production_bias_str  # Do not add a string if bias_num is NaN
                        else:
                            production_bias_str += f'\t\t\t\tSTORE_PERM (transported_last_month_pct("{item["produce_cargo"]}"),{item["bias_num"]}),\n'
                    pnml_content = pnml_content.replace('_production_bias_', production_bias_str)

                    # Write to the individual PNML file
                    pnml_file.write(pnml_content) # Write the content to the file.

                    combined_pnml_content += pnml_content + "\n\n" # Add content of current file

                print(f"Created PNML file: {pnml_filepath}")
            except Exception as e:
                print(f"  Error creating PNML file: {pnml_filepath} - e")

        # Write the combined PNML to src/industries.pnml
        combined_pnml_filepath = os.path.join('src', 'industries.pnml')
        try:
            with open(combined_pnml_filepath, 'w') as combined_file:
                combined_file.write(combined_pnml_content)
            print(f"Combined processed content into: {combined_pnml_filepath}")
        except Exception as e:
            print(f"  Error writing combined PNML file: {combined_pnml_filepath} - {e}")

        print(f"Industry PNML creation complete (both individual and combined files).")

    except FileNotFoundError:
        print(f"Error: File not found at {excel_filepath}")
        return
    except KeyError as e:
        print(f"Error: Column not found: {e}.  Make sure 'industry_item_name' and 'include' exist in the 'industries' sheet.")
        return
    except Exception as e:
        print(f"An error occurred: {e}")
        return

def CreateIndustryLNGs(industries_data_path='src/industries'):
    """
    Reads industry data from JSON files (created by CreateIndustries),
    merges data into an LNG template, and saves the processed content
    into individual and a combined LNG file.  JSON files are
    expected to be in subdirectories of industries_data_path,
    with each subdirectory named after the industry_item_name.

    Args:
        industries_data_path (str): The path to the directory containing the industry JSON files.
    """
    template_path_lng = 'src/templates/industry_lang_template.lng'
    output_combined_file_lng = 'src/industries_lang.lng'

    # Ensure the output directory exists
    os.makedirs(industries_data_path, exist_ok=True)

    try:
        with open(template_path_lng, 'r') as f:
            template_content_lng = f.read()
    except FileNotFoundError:
        print(f"Error: LNG Template not found at {template_path_lng}")
        return

    combined_output_lng = ""
    processed_count = 0  # Keep track of the number of files processed

    # Iterate through the subdirectories in the industries_data_path
    for industry_folder in os.listdir(industries_data_path):
        industry_folder_path = os.path.join(industries_data_path, industry_folder)
        if os.path.isdir(industry_folder_path):  # Only process directories
            json_filepath = os.path.join(industry_folder_path, f'{industry_folder}.json') # Correct json file path
            if os.path.exists(json_filepath):
                try:
                    with open(json_filepath, 'r', encoding='utf-8') as json_file:
                        record = json.load(json_file)
                        # print(f"Loaded data from: {json_filepath}")  # Debug
                except Exception as e:
                    print(f"Error reading JSON file: {json_filepath} - {e}")
                    continue  # Skip to the next file

                item_name = record.get('industry_item_name', 'default_item')  # Get industry name from JSON
                folder_name = item_name.replace(" ", "_")  # redundant
                file_name_lng = f"{item_name.replace(' ', '_')}.lng"
                output_folder = industry_folder_path
                output_path_lng = os.path.join(output_folder, file_name_lng)

                os.makedirs(output_folder, exist_ok=True)

                modified_content_lng = template_content_lng
                for key, value in record.items():
                    placeholder = f"_{key}_"
                    if isinstance(value, (int, float, str, bool, list, dict)):  # Add more datatypes if needed
                        modified_content_lng = modified_content_lng.replace(placeholder, str(value))

                # Write the individual LNG file
                try:
                    with open(output_path_lng, 'w', encoding='utf-8') as outfile:
                        outfile.write(modified_content_lng)
                    print(f"Processed and saved: {output_path_lng}")
                    processed_count += 1
                except Exception as e:
                    print(f"Error writing to individual LNG file {output_path_lng}: {e}")

                combined_output_lng += modified_content_lng + "\n\n"

    # Write the combined LNG file
    try:
        with open(output_combined_file_lng, 'w', encoding='utf-8') as outfile:
            outfile.write(combined_output_lng.strip())
        print(f"\nCombined processed content into: {output_combined_file_lng}")
    except Exception as e:
        print(f"Error writing combined LNG file {output_combined_file_lng}: {e}")

    print(f"Industry LNG creation complete (both individual and combined files).")
    if processed_count == 0:
        print(
            f"Warning: No LNG files were created.  Check if JSON files were generated in {industries_data_path} and if the 'industry_item_name' key exists in the JSON data, and that directories match the industry names."
        )

def CreateIndustryHelpText(excel_filepath='docs/otis.xlsx', base_folder='src/industries', output_file_path='src/helptext.pnml'):
    """
    Creates .pnml help text files for each industry, using the appropriate template
    based on the industry_type.  Files are saved in subfolders of the base_folder,
    with the subfolder name matching the industry_item_name.  The output file
    name is "[industry_item_name]_help.pnml".  Placeholders in the template
    are replaced with data from the 'industries' sheet in the Excel file.
    Finally, all generated files are combined into a single output file.

    Args:
        excel_filepath (str): The path to the Excel file.
        base_folder (str): The base folder where industry folders are located.
        output_file_path (str): The path to the final combined output file.
    """
    try:
        # Read the Excel file using pandas
        xls = pd.ExcelFile(excel_filepath)
        df = xls.parse('industries')  # Parse the 'industries' sheet

        # List to store paths of generated individual files
        generated_files = []

        # Iterate over the rows of the DataFrame
        for index, row in df.iterrows():
            # Check the 'include' column (case-insensitive)
            if 'include' in row and str(row['include']).lower() == 'true':
                industry_name = row['industry_item_name']
                industry_type = row.get('industry_type', 'generic')  # Default to 'generic' if missing

                # Construct the template file name.
                template_file = f'src/templates/{industry_type}_industry_help_template.pnml'

                # Construct the output file path.
                output_folder = os.path.join(base_folder, industry_name)
                individual_output_file = os.path.join(output_folder, f'{industry_name}_help.pnml')  # Changed output filename

                # Ensure the output directory exists.
                os.makedirs(output_folder, exist_ok=True)

                # Check if the template file exists
                if not os.path.exists(template_file):
                    print(f"Warning: Template file not found: {template_file}. Skipping {industry_name}.")
                    continue  # Skip to the next industry

                # Read the template and write to the output file.
                try:
                    with open(template_file, 'r', encoding='utf-8') as infile, open(individual_output_file, 'w', encoding='utf-8') as outfile:
                        template_content = infile.read()
                        # Replace placeholders with data from the row
                        for column, value in row.items():
                            placeholder = f'_{column}_'  # Placeholders are column names
                            if isinstance(value, (int, float, str, bool)):
                                template_content = template_content.replace(placeholder, str(value))
                        outfile.write(template_content)
                    print(f"Created help text file: {individual_output_file}")
                    generated_files.append(individual_output_file) #stores the file path
                except Exception as e:
                    print(f"Error creating help text file: {individual_output_file} - {e}")

        # Combine all generated files into a single output file
        try:
            with open(output_file_path, 'w', encoding='utf-8') as outfile:
                for file_path in generated_files:
                    with open(file_path, 'r', encoding='utf-8') as infile:
                        outfile.write(infile.read())
                        outfile.write("\n\n")  # Add separators between files
            print(f"Successfully combined all help text files into: {output_file_path}")
        except Exception as e:
            print(f"Error combining help text files: {e}")

    except FileNotFoundError:
        print(f"Error: File not found at {excel_filepath}")
        return
    except KeyError as e:
        print(f"Error: Column not found: {e}.  Make sure 'industry_item_name', 'industry_type' and 'include' exist in the 'industries' sheet.")
        return
    except Exception as e:
        print(f"An error occurred: {e}")
        return
        
def CreateIndustryHelpTextsLNGs(excel_filepath='docs/otis.xlsx', base_folder='src/industries', output_file_path='src/helptext_lang.lng'):
    """
    Creates .lng help text files for each industry, using the appropriate template
    based on the industry_type. Files are saved in subfolders of the base_folder,
    with the subfolder name matching the industry_item_name. The output file
    name is "[industry_item_name]_help.lng". Placeholders in the template
    are replaced with data from the 'industries' sheet in the Excel file,
    and with cargo data from the industry-specific sheets and the 'cargo' sheet.
    Finally, all generated files are combined into a single output file.

    Args:
        excel_filepath (str): The path to the Excel file.
        base_folder (str): The base folder where industry folders are located.
        output_file_path (str): The path to the final combined output file.
    """
    try:
        # Read the Excel file using pandas
        xls = pd.ExcelFile(excel_filepath)
        df_industries = xls.parse('industries')  # Parse the 'industries' sheet
        df_cargo = xls.parse('cargo')  # Parse the 'cargo' sheet
        cargo_label_to_str_cargo_name = dict(zip(df_cargo['cargo_label'].astype(str), df_cargo['str_cargo_name'].astype(str)))

        # List to store paths of generated individual files
        generated_files = []

        # Iterate over the rows of the DataFrame
        for index, row in df_industries.iterrows():
            # Check the 'include' column (case-insensitive)
            if 'include' in row and str(row['include']).lower() == 'true':
                industry_name = row['industry_item_name']
                industry_type = row.get('industry_type', 'generic')  # Default to 'generic' if missing

                # Construct the template file name.
                template_file = f'src/templates/{industry_type}_industry_lang_template.lng'

                # Construct the output file path for the individual file.
                output_folder = os.path.join(base_folder, industry_name)
                os.makedirs(output_folder, exist_ok=True)
                individual_output_file = os.path.join(output_folder, f'{industry_name}_help.lng')


                # Ensure the output directory exists.
                os.makedirs(output_folder, exist_ok=True)

                # Check if the template file exists
                if not os.path.exists(template_file):
                    print(f"Warning: Template file not found: {template_file}. Skipping {industry_name}.")
                    continue  # Skip to the next industry

                # Read the template and write to the output file.
                try:
                    with open(template_file, 'r', encoding='utf-8') as infile, open(individual_output_file, 'w', encoding='utf-8') as outfile:
                        template_content = infile.read()

                        # 1. Replace placeholders from the 'industries' sheet
                        for column, value in row.items():
                            placeholder = f'_{column}_'  # Placeholders are column names
                            if isinstance(value, (int, float, str, bool)):
                                template_content = template_content.replace(placeholder, str(value))

                        # 2. Handle cargo-related placeholders from the industry-specific sheet
                        try:
                            df_industry = xls.parse(industry_name)  # Parse the industry-specific sheet
                            accept_cargo_replacements = {}
                            for _, cargo_row in df_industry.iterrows():
                                accept_cargo = cargo_row.get('accept_cargo')
                                accept_cargo_type = cargo_row.get('accept_cargo_type')
                                if pd.notna(accept_cargo) and pd.notna(accept_cargo_type):
                                    cargo_name = cargo_label_to_str_cargo_name.get(str(accept_cargo), f"Cargo Label '{accept_cargo}' not found")
                                    placeholder_type = accept_cargo_type.lower()  # строчные буквы
                                    if placeholder_type not in accept_cargo_replacements:
                                        accept_cargo_replacements[placeholder_type] = []
                                    accept_cargo_replacements[placeholder_type].append(cargo_name)

                            for cargo_type, cargo_names in accept_cargo_replacements.items():
                                placeholder = f'_{cargo_type}_cargo_'
                                replacement_text = ", ".join(cargo_names)
                                template_content = template_content.replace(placeholder, replacement_text)
                            # Replace any remaining placeholders with "N/A"
                            placeholders = [f'_{column}_cargo_' for column in ['primary', 'secondary', 'support', 'supply']]  # Changed tertiary to support and quaternary to supply
                            for placeholder in placeholders:
                                if placeholder in template_content:
                                    template_content = template_content.replace(placeholder, "n/a")

                        except KeyError:
                            print(f"  Warning: Sheet '{industry_name}' not found in Excel file. Skipping cargo replacement for this industry.")
                        except Exception as e:
                            print(f"  Error processing cargo data for industry '{industry_name}': {e}")
                        # Replace any remaining placeholders with "N/A"
                        placeholders = [f'_{column}_' for column in df_industries.columns]
                        for placeholder in placeholders:
                            if placeholder in template_content:
                                template_content = template_content.replace(placeholder, "N/A")
                        outfile.write(template_content)
                    print(f"Created help text file: {individual_output_file}")
                    generated_files.append(individual_output_file) #stores the file path
                except Exception as e:
                    print(f"Error creating help text file: {individual_output_file} - {e}")

        # Combine all generated files into a single output file
        try:
            with open(output_file_path, 'w', encoding='utf-8') as outfile:
                for file_path in generated_files:
                    with open(file_path, 'r', encoding='utf-8') as infile:
                        outfile.write(infile.read())
                        outfile.write("\n\n")  # Add separators between files
                print(f"Successfully combined all help text files into: {output_file_path}")
        except Exception as e:
            print(f"Error combining help text files: {e}")

    except FileNotFoundError:
        print(f"Error: File not found at {excel_filepath}")
        return
    except KeyError as e:
        print(f"Error: Column not found: {e}.  Make sure 'industry_item_name', 'industry_type' and 'include' exist in the 'industries' sheet.")
        return
    except Exception as e:
        print(f"An error occurred: {e}")
        return
        
def CreateLNGFile(
    header_lang_path="src/header_lang.lng",
    cargo_lang_path="src/cargo_lang.lng",
    industries_lang_path="src/industries_lang.lng",
    helptext_lang_path="src/helptext_lang.lng",
    output_file_path="src/lang/english.lng",
):
    """
    Combines the content of  "header_lang.lng", "cargo_lang.lng", "industries_lang.lng" and "helptext_lang.lng"
    into a single english.lng file.

    Args:
        header_lang_path (str): Path to the header_lang.lng file.
        cargo_lang_path (str):  Path to cargo_lang.lng.
        industries_lang_path (str): Path to industries_lang.lng
        helptext_lang_path (str): Path to helptext_lang.lng
        output_file_path (str): Path to the output english.lng file.
    """

    combined_content = ""

    # Function to read and append LNG file content, with error handling
    def append_lng_content(file_path, description):
        nonlocal combined_content
        try:
            with open(file_path, "r", encoding="utf-8") as infile:
                combined_content += infile.read() + "\n\n"
            print(f"  Successfully read {description}: {file_path}")
        except FileNotFoundError:
            print(f"  Warning: {description} file not found: {file_path}")
        except Exception as e:
            print(f"  Error reading {description} file: {file_path} - {e}")

    # 1. Read and append header_lang.lng
    append_lng_content(header_lang_path, "Header LNG")

    # 2. Read and append cargo_lang.lng
    append_lng_content(cargo_lang_path, "Cargo LNG")

    # 3. Read and append industries_lang.lng
    append_lng_content(industries_lang_path, "Industries LNG")

    # 4. Read and append helptext_lang.lng
    append_lng_content(helptext_lang_path, "Helptext LNG")

    # 5. Write the combined content to english.lng
    try:
        os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
        with open(output_file_path, "w", encoding="utf-8") as outfile:
            outfile.write(combined_content.strip())
        print(f"Successfully created combined LNG file: {output_file_path}")
    except Exception as e:
        print(f"Error writing combined LNG file: {output_file_path} - {e}")