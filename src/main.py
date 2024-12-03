import os
from presidio_analyzer import AnalyzerEngine
from pdfminer.high_level import extract_text as extract_pdf_text
import docx
import pandas as pd

# Initialize the analyzer engine
analyzer = AnalyzerEngine()

def extract_text_from_txt(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as file:
            return file.read()
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return ''

def extract_text_from_pdf(filepath):
    try:
        text = extract_pdf_text(filepath)
        return text
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return ''

def extract_text_from_docx(filepath):
    try:
        doc = docx.Document(filepath)
        return '\n'.join([para.text for para in doc.paragraphs])
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return ''

def extract_text_from_csv(filepath):
    try:
        df = pd.read_csv(filepath, encoding='utf-8', errors='ignore')
        text = df.to_string(index=False)
        return text
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return ''

def extract_text_from_excel(filepath):
    try:
        df = pd.read_excel(filepath, sheet_name=None)  # Read all sheets
        text = ''
        for sheet_name, sheet_df in df.items():
            text += f"Sheet: {sheet_name}\n"
            text += sheet_df.to_string(index=False)
            text += '\n'
        return text
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return ''

def extract_text(filepath):
    if filepath.endswith('.txt'):
        return extract_text_from_txt(filepath)
    elif filepath.endswith('.pdf'):
        return extract_text_from_pdf(filepath)
    elif filepath.endswith('.docx'):
        return extract_text_from_docx(filepath)
    elif filepath.endswith('.csv'):
        return extract_text_from_csv(filepath)
    elif filepath.endswith('.xlsx'):
        return extract_text_from_excel(filepath)
    else:
        return ''

def scan_text_for_pii(text):
    results = analyzer.analyze(text=text, entities=[], language='en')
    return results

def scan_directory(directory):
    report = {}
    file_texts = {}
    for root, _, files in os.walk(directory):
        for filename in files:
            filepath = os.path.join(root, filename)
            text = extract_text(filepath)
            if text:
                findings = scan_text_for_pii(text)
                if findings:
                    report[filepath] = findings
                    file_texts[filepath] = text
    return report, file_texts

def generate_report(report, file_texts):
    for filepath, findings in report.items():
        text = file_texts[filepath]
        print(f"File: {filepath}")
        for result in findings:
            entity = result.entity_type
            start = result.start
            end = result.end
            score = result.score
            detected_text = text[start:end]
            print(f"  Entity: {entity}, Text: '{detected_text}', Position: {start}-{end}, Confidence Score: {score:.2f}")
        print('-' * 40)

def main():
    directory = input("Enter the directory to scan: ")
    report, file_texts = scan_directory(directory)
    if report:
        generate_report(report, file_texts)
    else:
        print("No PII found in the scanned directory.")

if __name__ == "__main__":
    main()
