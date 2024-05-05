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
                    
                    # Define API URL and headers
                    API_URL = "https://api-inference.huggingface.co/models/ealvaradob/bert-finetuned-phishing"
                    headers = {"Authorization": "Bearer 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"}
                    
                    # Prepare payload and make a POST request to the API
                    payload = {"inputs": email_body}
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
