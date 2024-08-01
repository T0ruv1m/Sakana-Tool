import os
import re
import xml.etree.ElementTree as ET
import pandas as pd

# Function to extract data from the XML file
def extract_data_from_xml(xml_file):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Define namespaces to handle XML namespaces correctly
        namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

        # Find the chNFe element and extract its text
        chNFe_element = root.find('.//nfe:protNFe/nfe:infProt/nfe:chNFe', namespaces)
        chNFe_text = chNFe_element.text if chNFe_element is not None else None

        # Find the xMun element under enderDest and extract its text
        xMun_element = root.find('.//nfe:dest/nfe:enderDest/nfe:xMun', namespaces)
        xMun_text = xMun_element.text if xMun_element is not None else None

        # Find the infCpl element and extract its text
        infCpl_element = root.find('.//nfe:infAdic/nfe:infCpl', namespaces)
        infCpl_text = infCpl_element.text if infCpl_element is not None else ''

        # Extract the 44-digit number using regex
        number_match = re.search(r'\b\d{44}\b', infCpl_text)
        extracted_number = number_match.group(0) if number_match else None

        # Extract the text after the last semicolon in infCpl
        last_semicolon_text = infCpl_text.split(';')[-1].strip() if ';' in infCpl_text else None

        # Find the vProd element within ICMSTot and extract its text
        vProd_element = root.find('.//nfe:total/nfe:ICMSTot/nfe:vProd', namespaces)
        vProd_text = vProd_element.text if vProd_element is not None else None

        return extracted_number, chNFe_text, xMun_text, last_semicolon_text, vProd_text

    except Exception as e:
        print(f"Error processing file {xml_file}: {e}")
    return None, None, None, None, None

# Get the current script's directory
pato = os.path.dirname(os.path.abspath(__file__))

# Define the XML directory relative to the script's directory
xml_directory = os.path.join(pato, '../xmldata')

# List to store extracted data
extracted_data = []
print('Processing started...')

# Iterate over all XML files in the directory and its subdirectories
for root_dir, dirs, files in os.walk(xml_directory):
    for filename in files:
        if filename.endswith('.xml'):
            file_path = os.path.join(root_dir, filename)
            extracted_number, chNFe, xMun, last_text, vProd = extract_data_from_xml(file_path)
            if extracted_number and chNFe and xMun and last_text and vProd:
                extracted_data.append((filename, extracted_number, chNFe, xMun, last_text, vProd))

print('Processing finished.')

# Create a DataFrame from the extracted data
df = pd.DataFrame(extracted_data, columns=['Filename', 'chNFe', 'chNTR', 'xMun', 'FOR', 'vProd'])

# Define the output Excel file path
output_excel_file = '../xlsx/extracted_data.xlsx'
dataframe_dir = os.path.join(pato, output_excel_file)

# Create the directory for the output Excel file if it doesn't exist
output_dir = os.path.dirname(dataframe_dir)
os.makedirs(output_dir, exist_ok=True)

# Save the DataFrame to an Excel file
df.to_excel(dataframe_dir, index=False)

print(f"Data has been written to {output_excel_file}")
print(df)
