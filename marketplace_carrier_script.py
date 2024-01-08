import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def process_file(file_path):
    try:
        excel_data = pd.ExcelFile(file_path)
        vinted_marketplace_df = pd.read_excel(excel_data, sheet_name='Vinted marketplace invoice July')
        weight_listing_df = pd.read_excel(excel_data, sheet_name='Weight listing')
        carrier_price_list_df = pd.read_excel(excel_data, sheet_name='Carrier price list')

        # Create a dictionary for weight range mapping
        weight_mapping = dict(zip(weight_listing_df['Weight Range (A)'], weight_listing_df['Matched Description (F)']))
        
        def map_weight_range(weight_range):
            return weight_mapping.get(weight_range, weight_range)

        def format_carrier_id(carrier_id):
            if pd.notna(carrier_id):
                return str(int(carrier_id))
            return ''

        vinted_marketplace_df['Corrected_weight_range'] = vinted_marketplace_df['weight_range'].apply(map_weight_range)
        vinted_marketplace_df['UniqueID_Marketplace'] = (
            vinted_marketplace_df['seller_country'].astype(str) + 
            vinted_marketplace_df['buyer_country'].astype(str) + 
            vinted_marketplace_df['first_mile_carrier_id'].apply(format_carrier_id) + 
            vinted_marketplace_df['last_mile_carrier_id'].apply(format_carrier_id) + 
            vinted_marketplace_df['Corrected_weight_range'] + 
            vinted_marketplace_df['selling_price_currency']
        )

        def extract_package_size(vinted_package_size):
            match = re.search(r'(\d+\.\d+[-]\d+\.\d+ kg)', vinted_package_size)
            return match.group(1) if match else None

        carrier_price_list_df['Package size range'] = carrier_price_list_df['Vinted Package size'].apply(extract_package_size)
        carrier_price_list_df['UniqueID_Carrierlist'] = (
            carrier_price_list_df['From Country'].astype(str) + 
            carrier_price_list_df['To Country'].astype(str) + 
            carrier_price_list_df['First mile carrier ID'].astype(str) + 
            carrier_price_list_df['Last mile carrier ID'].astype(str) + 
            carrier_price_list_df['Package size range'] + 
            carrier_price_list_df['Rates Currency']
        )

        unique_id_carrierlist_set = set(carrier_price_list_df['UniqueID_Carrierlist'])

        def check_for_match(unique_id):
            return 'Match' if unique_id in unique_id_carrierlist_set else 'No match'

        vinted_marketplace_df['Match'] = vinted_marketplace_df['UniqueID_Marketplace'].apply(check_for_match)

        def determine_unmatched_mapping(row):
            if row['Match'] == 'No match':
                if row['transaction_count'] < 20:
                    return "3.5"
                else:
                    carrier_currency = carrier_price_list_df.loc[carrier_price_list_df['Service Provider Legal Name'] == row['carriers_code'], 'Rates Currency']
                    if not carrier_currency.empty and row['selling_price_currency'] != carrier_currency.iloc[0]:
                        return "Currency difference"
                    else:
                        return "Other reasons"
            return None

        vinted_marketplace_df['unmatched mapping'] = vinted_marketplace_df.apply(
            lambda row: determine_unmatched_mapping(row), axis=1
        )

        def mapping_all(row):
            if row['Match'] == 'Match':
                return row['UniqueID_Marketplace']
            elif row['Match'] == 'No match' and row['transaction_count'] < 20:
                return '3.5'
            elif row['unmatched mapping'] == 'Currency difference':
                correct_currency = carrier_price_list_df.loc[
                    carrier_price_list_df['Service Provider Legal Name'] == row['carriers_code'],
                    'Rates Currency'
                ].values[0] if not carrier_price_list_df.loc[
                    carrier_price_list_df['Service Provider Legal Name'] == row['carriers_code'],
                    'Rates Currency'
                ].empty else 'Unknown Currency'

                return (
                    row['seller_country']
                    + row['buyer_country']
                    + str(row['first_mile_carrier_id'])
                    + str(row['last_mile_carrier_id'])
                    + row['Corrected_weight_range']
                    + correct_currency
                )
            return 'Other reasons'

        vinted_marketplace_df['Mapping all'] = vinted_marketplace_df.apply(
            lambda row: mapping_all(row), axis=1
        )

        directory = os.path.dirname(file_path)
        results_file_path = os.path.join(directory, "Results.xlsx")
        writer = pd.ExcelWriter(results_file_path, engine='xlsxwriter')
        vinted_marketplace_df.to_excel(writer, sheet_name='Vinted marketplace invoice July', index=False)
        carrier_price_list_df.to_excel(writer, sheet_name='Carrier price list', index=False)
        writer.close()

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_file_dialog():
    file_path = filedialog.askopenfilename(title="Select file", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        process_file(file_path)

root = tk.Tk()
root.title('Excel File Processor')

open_button = tk.Button(root, text="Open Excel File", command=open_file_dialog)
open_button.pack()

root.mainloop()
