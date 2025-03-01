import requests
import base64
import pandas as pd

# Freshservice API details
url = "https://xxxxx" # your freshservice url
api_key = "add your admin api here"  # Add your Admin API
auth_header = base64.b64encode(f"{api_key}:X".encode()).decode()

headers = {
    "Content-Type": "application/json",
    "Authorization": f"Basic {auth_header}"
}

# Load Excel file
file_path = "Your excel file path" # the excel file path - the excel that needs to be converted into a dataframe in python
df = pd.read_excel(file_path, engine='openpyxl')


# Converting the date column into a string (example : Tuesday , Febuary, 25, 2025)
df["date"] = pd.to_datetime(df["date"], errors='coerce')
df["date"] = df["date"].apply(lambda x: x.strftime('%A, %B %d, %Y') if pd.notna(x) else "Unknown Date")

# Loop through each row in the Excel file and create tickets
for index, row in df.iterrows():
    try:
        # Extract details from Excel
        netid = str(row["netid"]).strip()
        subject = str(row["title"]).strip()
        description = str(row["description"]).strip()
        date = row["date"] 
        full_description = f"Date: {date}\n\n\n:  {description}"  # Append date to description
        resolution = str(row["resolution"]).strip()
        service = str(row["service"]).strip()
        responder_id = int(row["responder_id"]) if pd.notna(row["responder_id"]) else None
        group_id = int(row["group_id"]) if pd.notna(row["group_id"]) else None
        status = int(row["status"]) if pd.notna(row["status"]) else 2  # Read 'status' from Excel, default to 'In Progress'
        source = int(row["source"]) if pd.notna(row["source"]) else 3  # Read 'source' from Excel, default to 'Phone'

        # Prepare the ticket payload
        create_payload = {
            "ticket": {
                "subject": subject,
                "description": full_description,
                "email": f"{netid}@colostate.edu",  # Assuming email format from NetID
                "status": 2,  # Start with 'In Progress' before updating
                "priority": 1,
                "source": source,  # Read from Excel
                "group_id": group_id,
                "responder_id": responder_id,
                "custom_fields": {
                    "csu_id": netid,
                    "issue_field_test": service,
                    "resolution": resolution
                }
            }
        }

        # Send request to create the ticket
        response = requests.post(url, json=create_payload, headers=headers)

        if response.status_code == 201:
            ticket_id = response.json()["ticket"]["id"]
            print(f"✅ Ticket Created Successfully: {ticket_id} for {netid}")

            # Step 2: Update the ticket status to the one from Excel
            update_url = f"https://csusystem.freshservice.com/api/v2/tickets/{ticket_id}"
            update_payload = {"ticket": {"status": status}}  # Set status from Excel

            update_response = requests.put(update_url, json=update_payload, headers=headers)

            if update_response.status_code == 200:
                print(f"✅ Ticket {ticket_id} Updated to Status {status}")
            
            else:
                print(f"⚠️ Failed to update ticket {ticket_id}: {update_response.status_code}", update_response.text)
        else:
            print(f" Error Creating Ticket for {netid}: {response.status_code}", response.text)

    except Exception as e:
        print(f" Error processing row {index + 1}: {e}")

print("\n Final message - All tickets have been processed !")