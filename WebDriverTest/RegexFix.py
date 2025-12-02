import pandas as pd
import re
import glob
import os

def remove_duplicate(text):
    text = str(text)
    pattern = r'(.{5,}?)(\1)+'
    cleaned_text = re.sub(pattern, r'\1', text, flags=re.DOTALL)
    if cleaned_text == 'nan':
        return float('nan')
    return cleaned_text

list_of_files = glob.glob(r'Z:\PR\Daily\GestaoPedidos\*.xlsx')
latest_file = max(list_of_files, key=os.path.getmtime)

df = pd.read_excel(latest_file)

output_directory = r"Z:\PR\15Min\GestaoPedidosRotas"
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

if df.shape[1] >= 37:
    df.iloc[:, 36] = df.iloc[:, 36].apply(remove_duplicate)

    # Extract only the file name, not the full path
    file_name = os.path.basename(latest_file)
    # Create a new file name with "_cleaned" appended before the extension
    base_name, extension = os.path.splitext(file_name)
    new_file_name = f"{base_name}_cleaned{extension}"
    # Correctly join the new file name with the output directory
    full_path = os.path.join(output_directory, new_file_name)
    
    df.to_excel(full_path, index=False)
else:
    print("Falha no excel.")
