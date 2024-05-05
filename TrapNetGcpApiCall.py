import requests
import win32com.client

class EmailInspector:
    def __init__(self):
        # Connect to Outlook via COM
        self.outlook = win32com.client.Dispatch("Outlook.Application")

    def check_email(self):
        """Checks the body of the currently selected email in Outlook."""
        try:
            explorer = self.outlook.ActiveExplorer()
            selection = explorer.Selection
            if selection.Count > 0:
                # Assumes the first item in the selection is an email.
                # Adjust accordingly if your selection might include other item types.
                item = selection.Item(1)
                if item.MessageClass == "IPM.Note":
                    email_subject = item.Subject  # Capture the email's subject
                    email_body = item.Body[:512]  # Truncate the email body to fit the model's requirements
                    
                    print(f"Checking email: {email_subject}")
                    
                    # Define the GCP endpoint URL
                    API_URL = "https://us-central1-aiplatform.googleapis.com/v1/projects/tidy-reporter-417118/locations/us-central1/endpoints/your-endpoint-id:predict"
                    
                    # Prepare the data payload and headers for the POST request to the API
                    payload = {"instances": [{"text": email_body}]}
                    headers = {"Authorization": "Bearer $(gcloud auth print-access-token)"}
                    
                    # Make a POST request to the GCP endpoint
                    response = requests.post(API_URL, headers=headers, json=payload)
                    
                    # Print the API's response
                    print("API Result:", response.json())
                else:
                    print("The selected item is not an email.")
            else:
                print("No email selected.")
        except Exception as e:
            print("Error:", str(e))

if __name__ == "__main__":
    inspector = EmailInspector()
    inspector.check_email()
