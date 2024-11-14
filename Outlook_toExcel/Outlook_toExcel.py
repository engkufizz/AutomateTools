import extract_msg
import pandas as pd

# Function to extract data from the .msg file
def extract_data_from_msg(file_path):
    msg = extract_msg.Message(file_path)
    msg_body = msg.body
    msg.close()

    # Print the body content for debugging
    print("Message Body:\n")
    print(msg_body)

    # List the headers without repeats
    headers = [
        "FCN Docket No", "Circuit ID", "Date", "Time Failure", 
        "Time Restore", "Nature of failure", "Cause of the failure", 
        "Action performed to restore the Services", "Ownership of the incident", 
        "Location of the incident"
    ]

    # Prepare to store data
    data = {}
    date_positions = ["Date (Start)", "Date (Restore)"]
    date_index = 0

    lines = [line.strip() for line in msg_body.splitlines() if line.strip()]  # Remove empty lines and trim spaces

    # Process each line
    i = 0
    while i < len(lines):
        line = lines[i]
        header = None
        
        if line.startswith("Date:") and date_index < len(date_positions):
            header = date_positions[date_index]
            date_index += 1
        else:
            for h in headers:
                if line.lower().startswith(h.lower()):
                    header = h.replace(":", "").strip()
                    break

        if header:
            # Assume details are on the next line
            if i + 1 < len(lines):
                detail = lines[i + 1].strip()
                data[header] = detail
        i += 1

    return data

# Save extracted data to an Excel file
def save_data_to_excel(data, output_path):
    df = pd.DataFrame(list(data.items()), columns=['Item Type', 'Detail'])
    df.to_excel(output_path, index=False, engine='openpyxl')

# Main function
def main():
    file_path = 'Data.msg'  # Path to your .msg file
    output_path = 'extracted_data.xlsx'  # Path to save the Excel file

    data = extract_data_from_msg(file_path)
    if data:
        save_data_to_excel(data, output_path)
        print(f"Data extracted and saved to {output_path}")
    else:
        print("No data found or failed to extract data.")

if __name__ == "__main__":
    main()
