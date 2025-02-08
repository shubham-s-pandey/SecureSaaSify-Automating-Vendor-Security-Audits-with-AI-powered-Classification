# SecureSaaSify: Automating SaaS Security with AI-powered Classification

## Features:

- **Automated SaaS Security Classification**: Uses AI to determine whether vendor remarks are compliant or non-compliant with SaaS security standards.
- **Vendor Remark Analysis**: Evaluates each vendor's response to audit questions on topics such as data encryption, access control, compliance with regulations (e.g., GDPR, SOC2), and more.
- **Detailed Report Generation**: The classification, along with reasons, is saved in a tabular format in a CSV file, which can be easily analyzed.
- **Custom Row Range Processing**: Users can specify a row range in the Excel sheet to selectively process vendor remarks.

## Installation:

To run this project, ensure you have the following dependencies installed:

    Python 3.x
    openpyxl (For working with Excel files)
    ollama (For AI-based classification)
    csv (For writing results to CSV)

## How to Use:

1. **Prepare your Excel file**:  
    The input Excel file should contain vendor audit questions in column A and their corresponding remarks in column B. Ensure the sheet name is **"Vendor checklist"**.

2. **Run the Script**:  
    - Provide the file path to the Excel sheet.
    - Specify the start and end rows to selectively process the questions or leave it blank to process the entire sheet.
    - The results will be saved in a CSV file containing the audit question, vendor remark, classification (compliant/non-compliant), and a summarized reason for the classification.

## Sample Output:

The program will output the classified results as a CSV file with the following columns:
- **Audit Question**
- **Vendor Remark**
- **Classification** (Compliant/Non-Compliant)
- **Reason** (A brief summary or explanation for the classification)

### Sample Input:

| Audit Question                                            | Vendor Remark                                                            |
|-----------------------------------------------------------|---------------------------------------------------------------------------|
| Does the vendor encrypt all sensitive data at rest?       | The vendor uses AES-256 encryption for sensitive data.                     |
| Does the vendor provide audit logs for all system access?| The vendor does not provide detailed logs for system access.             |

### Sample Output (CSV):

```csv
Audit Question,Vendor Remark,Classification,Reason
"Does the vendor encrypt all sensitive data at rest?","The vendor uses AES-256 encryption for sensitive data.","Compliant","The vendor uses AES-256 encryption, which is a widely accepted standard for data protection."
"Does the vendor provide audit logs for all system access?","The vendor does not provide detailed logs for system access.","Non-Compliant","The vendor does not provide detailed audit logs, which is a security gap that needs addressing."
```
## Bugs and Feature Requests

Please raise an issue if you encounter a bug or have a feature request. 

## Contributing

If you want to contribute to a project and make it better, your help is very welcome.
