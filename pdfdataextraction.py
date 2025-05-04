import pandas as pd
from pdfminer.high_level import extract_text
import re

def remove_empty_spaces(input_list):
    return [item for item in input_list if item.strip()]

# Function to process the text and extract customer details
def extract_customer_details(text):
    # Split the text into lines and clean it
    lines = text.split('\n')
    cleaned_lines = remove_empty_spaces(lines)

    # Initialize a list to store customer selections
    customer_selections = []
    current_customer = None

    # Regex patterns for detecting selections, customer name, location, and phone number
    customer_pattern = re.compile(r"Customer Name:\s*(.+)", re.IGNORECASE)
    location_pattern = re.compile(r"Location:\s*(.+)", re.IGNORECASE)
    phone_pattern = re.compile(r"Phone Number:\s*(\d+)", re.IGNORECASE)
    tick_pattern = re.compile(r'\[✓\]\s*(.+)', re.IGNORECASE)

    # Process each line
    for line in cleaned_lines:
        # Check if the line indicates a new customer
        customer_match = customer_pattern.search(line)
        if customer_match:
            if current_customer:
                customer_selections.append(current_customer)
            # Initialize a new customer
            current_customer = {
                "Customer Name": customer_match.group(1).strip(),
                "Location": "Unknown",
                "Phone Number": "Unknown",
                "Brick Material": None,
                "Cement Company": None,
                "Wood Type": None,
                "Chimney": None
            }
        # Extract location
        elif current_customer and location_pattern.search(line):
            current_customer["Location"] = location_pattern.search(line).group(1).strip()
        # Extract phone number
        elif current_customer and phone_pattern.search(line):
            current_customer["Phone Number"] = phone_pattern.search(line).group(1).strip()
        # Detect ticked options
        elif current_customer:
            match = tick_pattern.search(line)
            if match:
                option = match.group(1).strip()
                if "bricks" in option.lower():
                    current_customer["Brick Material"] = option
                elif "cement" in option.lower():
                    current_customer["Cement Company"] = option
                elif "wood" in option.lower():
                    current_customer["Wood Type"] = option
                elif "available" in option.lower():
                    current_customer["Chimney"] = option

    # Append the last customer
    if current_customer:
        customer_selections.append(current_customer)

    return customer_selections

def save_to_excel(customers, filename):
    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(customers)
    
    # Save the DataFrame to an Excel file
    df.to_excel(filename, index=False, engine='openpyxl')

# Extract text from the PDF
pdf_path = "C:\\Users\\AM.L.KANTH REDDY\\Documents\\hack python\\ADI.pdf"  # Update this path to your PDF file
text = extract_text(pdf_path)

# Extract customer details from the text
customer_details = extract_customer_details(text)

# Save customer details to an Excel file
excel_filename = "customer_details.xlsx"
save_to_excel(customer_details, excel_filename)

print(f"Customer details have been saved to {excel_filename}.")