Outlook Email Automation with UiPath
Project Description
This UiPath project automates the process of sending personalized emails from Outlook to multiple recipients. The recipient details, including email addresses and personalized messages, are stored in an Excel file. The automation reads the data from the Excel file, sends emails accordingly, and updates the Excel file with the status of each email (whether it was successfully sent or failed).

Features
Automated Email Sending: Sends emails from Outlook based on the details provided in the Excel file.
Error Handling: Catches exceptions and records the status of each email (success or failure) in the Excel file.
Personalization: Allows for personalized email content using data from the Excel file.
Logging: Updates the Excel file with the status of each email after the sending process.
Prerequisites
Before running this project, ensure that you have the following software installed:

UiPath Studio: Latest version is recommended.
Microsoft Outlook: Installed and configured with the email account from which emails will be sent.
Excel Application: Used to create and manage the recipient details and log the email status.
Setup Instructions
1. Clone the Repository
Clone this repository to your local machine:

bash
Copy code
git clone https://github.com/yourusername/outlook-email-automation.git
2. Open the Project in UiPath Studio
Open UiPath Studio.
Select "Open a Local Project."
Navigate to the cloned repository and select the project file (project.json).
3. Configure the Excel File
The Excel file should have the following columns:

Email: The recipient's email address.
Message: The personalized message for the recipient.
Status: This column will be automatically updated by the automation with "Success" or "Failure."
Place the Excel file in the project directory or update the path in the Excel Process Scope activity.

4. Run the Automation
Open the Main.xaml file.
Run the project from UiPath Studio.
The automation will read the Excel file, send emails, and update the status in the Excel file.
Project Structure
Main.xaml: The main workflow file that contains the logic for reading the Excel file, sending emails, and updating the status.
Excel Process Scope: Contains activities for reading and writing data to and from the Excel file.
Try-Catch Block: Handles any errors that occur during the email sending process and logs the failure in the Excel file.
Error Handling
Try-Catch Block: Ensures that the automation continues running even if an error occurs while sending an email. The error is logged in the Excel file under the "Status" column.
Troubleshooting
Outlook Configuration: Ensure that Outlook is properly configured and connected to the internet.
Excel Path: Verify that the correct path to the Excel file is set in the Excel Process Scope activity.
SMTP Settings: If you encounter issues with sending emails, double-check your Outlook account settings and SMTP configuration.
