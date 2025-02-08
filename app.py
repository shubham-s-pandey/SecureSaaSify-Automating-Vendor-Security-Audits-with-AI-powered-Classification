import openpyxl
import ollama
import time
import csv
from tqdm import tqdm

def print_ascii_header():
    ascii_art = r"""


 _____                          _____             _____ _  __       
/  ___|                        /  ___|           /  ___(_)/ _|      
\ `--.  ___  ___ _   _ _ __ ___\ `--.  __ _  __ _\ `--. _| |_ _   _ 
 `--. \/ _ \/ __| | | | '__/ _ \`--. \/ _` |/ _` |`--. \ |  _| | | |
/\__/ /  __/ (__| |_| | | |  __/\__/ / (_| | (_| /\__/ / | | | |_| |
\____/ \___|\___|\__,_|_|  \___\____/ \__,_|\__,_\____/|_|_|  \__, |
                                                               __/ |
                                                              |___/ 
                                                                                                  
                                     Author: @shubham-s-pandey
    """
    print(ascii_art)

def read_vendor_remarks(file_path, start_row=2, end_row=None):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook['Vendor checklist']  # Load the "Vendor checklist" sheet
    
    vendor_remarks = []
    for row in sheet.iter_rows(min_row=start_row, max_row=end_row, values_only=True): 
        audit_question = row[0]  # Column A (Audit Question)
        vendor_remark = row[1]  # Column B (Vendor Remark)
        vendor_remarks.append((audit_question, vendor_remark))
    
    return vendor_remarks, workbook, sheet

def classify_vendor_remarks(vendor_remarks):
    classified_results = []
    
    for audit_question, remark in tqdm(vendor_remarks, desc="Classifying remarks", unit="remark"):
        prompt = f"As an expert auditor, review the following vendor remark in relation to SaaS security standards. " \
                 f"Classify the remark as either 'Compliant' or 'Non-Compliant' in one word. Then, provide a detailed " \
                 f"reasoning for your classification. Your response should include both a one-word classification and a detailed reason.\n\n" \
                 f"Vendor Remark: {remark}\nAudit Question: {audit_question}\n"
        
        response = ollama.chat(
            model="deepseek-r1:1.5b",
            messages=[{"role": "user", "content": prompt}]
        )
        
        result = response["message"]["content"].strip()

        if ":" in result:
            classification, reason = result.split(":", 1)
            classification = classification.strip()
            reason = reason.strip()
        else:
            classification = "Non-Compliant"
            reason = result
        
        reason = f"{classification}: {reason}"
        
        classified_results.append((audit_question, remark, reason))
        
        print(f"Processed: {audit_question} -> {reason}")
        
        time.sleep(1)  
        
    return classified_results

def write_classified_results_to_csv(classified_results, csv_file_path):
    with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["Audit Question", "Vendor Remark", "Reason"]) 
        
        for audit_question, remark, reason in classified_results:
            writer.writerow([audit_question, remark, reason])
    
    print(f"Results have been saved to {csv_file_path}")

def run_analysis(file_path, start_row=2, end_row=None, csv_file_path="classified_saas_security_remarks.csv"):
    print_ascii_header()
    
    vendor_remarks, workbook, sheet = read_vendor_remarks(file_path, start_row, end_row)
    
    classified_results = classify_vendor_remarks(vendor_remarks)
    
    write_classified_results_to_csv(classified_results, csv_file_path)

file_path = "~/SAAS_vendor_checklist.xlsx"  

start_row = int(input("Enter the start row (default is 2): ") or 2)
end_row = input("Enter the end row (leave blank for the entire sheet): ")
end_row = int(end_row) if end_row else None

csv_file_path = "~/classified_vendor_remarks_saas_security.csv"
run_analysis(file_path, start_row, end_row, csv_file_path)
