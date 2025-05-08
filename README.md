# FuelEU Outreach Automation

This project automates the process of creating personalized email drafts for outreach campaigns using Microsoft Graph API. The script reads contact data from an Excel file, generates customized email content for each contact, and creates email drafts in the sender's Microsoft Outlook account.

## Features

- **Authentication**: Uses Microsoft Graph API's Client Credentials Flow to authenticate and acquire an access token.
- **Excel Data Processing**: Reads contact information from an Excel file, validates required columns, and cleans up missing or invalid data.
- **Email Personalization**: Generates personalized email content using a predefined HTML email template.
- **Draft Email Creation**: Asks the developer to confirm the action of creating x drafts. Creates email drafts in the sender's Outlook account with options for CC and BCC recipients.
- **Error Handling**: Includes error handling for authentication, Excel file reading, and email draft creation.
- **Progress Tracking**: Displays progress using the `tqdm` library and provides a summary of successful and failed operations.

## Prerequisites

- Python 3.8 or higher
- Required Python libraries:
  - `pandas`
  - `requests`
  - `openpyxl`
  - `msal`
  - `tqdm`
- Microsoft Azure App Registration with the following details:
  - **Client ID**
  - **Tenant ID**
  - **Client Secret**
- An Excel file containing contact data with the following required columns:
  - `Company Name`
  - `First Name`
  - `Last Name`
  - `Job Title`
  - `Email`

## Configuration

Update the `CONFIG` dictionary in the script with the following details:

- `client_id`: Your Azure App's Client ID.
- `tenant_id`: Your Azure Tenant ID.
- `client_secret`: Your Azure App's Client Secret.
- `bcc_email`: Email address for BCC recipients. - this email creates a Hubspot Profile for the contacted companies. 
- `cc_email`: Email address for CC recipients.
- `drive_id`, `site_id`, `file_id`: SharePoint/OneDrive file details (if applicable).
- `hubspot_api_key`: API key for HubSpot integration (if used).

## Usage

1. **Install Dependencies**:
   Install the required Python libraries using `pip`:
   ```bash
   pip install pandas requests openpyxl msal tqdm
   ```

2. **Prepare Contact Data**:
    To avoid making a complex code (authorization-acces wise) and have more control over the file, i just manually add the contacts we want from the One Drive to the xlsx of this project.

   Ensure the Excel file containing contact data is in the same directory as the script. The file should have the required columns (`Company Name`, `First Name`, `Last Name`, `Job Title`, `Email`).
   Batch1_Guido.xlsx - contains the list for Guido to be the main sender of. (Meesz on cc)
   Batch1_Meesz.xlsx - contains the list for Meesz to be the main sender of. (Guido on cc)

3. **Run the Script**:
   Execute the script using Python:
   Option1
   ```bash
   python Meesz_outreach.ipynb
   python Guido_outreach.ipynb
   ```
   Option2: Ctrl + Enter each cell

4. **Follow Prompts**:
   - The script will authenticate with Microsoft Graph API.
   - It will read the contact data from the Excel file.
   - You will be prompted to confirm the creation of email drafts.
   - The script will process each contact and create email drafts in the sender's Outlook account.

5. **Review Results**:
   After execution, the script will display a summary of the operation, including the number of successful and failed email drafts.

## Email Template

The email content is defined in the `DEFAULT_EMAIL` variable. It uses placeholders (e.g., `{first_name}`) for personalization. You can customize the template as needed.

## Error Handling

- If the script encounters missing columns in the Excel file, it will display an error message and exit.
- If an error occurs during email draft creation, the script will log the error and continue processing the remaining contacts.

## Notes

- Ensure your Azure App has the necessary permissions to access Microsoft Graph API.
- Avoid sharing sensitive information (e.g., client secrets, API keys) publicly.

## License

This project is proprietary and intended for internal use by Carbon Leap. Unauthorized distribution or use is prohibited.

## Contact

Added contact info in signature details