import os
from presidio_analyzer import AnalyzerEngine, RecognizerResult, DataRecognizerResult
from presidio_analyzer.nlp_engine import NlpEngineProvider
import pandas as pd
import docx
from pdfminer.high_level import extract_text as extract_pdf_text

def get_analyzer_engine():
    nlp_configuration = {
        "nlp_engine_name": "spacy",
        "models": [
            {
                "lang_code": "en",
                "model_name": "en_core_web_lg"
            }
        ]
    }
    provider = NlpEngineProvider(nlp_configuration=nlp_configuration)
    nlp_engine = provider.create_engine()
    analyzer = AnalyzerEngine(
        nlp_engine=nlp_engine,
        supported_languages=["en"]
    )
    return analyzer

analyzer = get_analyzer_engine()

def scan_dataframe_for_pii(df):
    results = []
    # Iterate over each column and row
    for column in df.columns:
        for index, cell_value in df[column].iteritems():
            if pd.isnull(cell_value):
                continue
            cell_text = str(cell_value)
            cell_results = analyzer.analyze(text=cell_text, entities=[], language='en')
            if cell_results:
                results.append({
                    'row': index + 1,  # 1-based index
                    'column': column,
                    'text': cell_text,
                    'pii_entities': cell_results
                })
    return results

def extract_and_scan_excel(filepath):
    try:
        xls = pd.ExcelFile(filepath)
        file_findings = []
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            findings = scan_dataframe_for_pii(df)
            if findings:
                file_findings.append({
                    'sheet': sheet_name,
                    'findings': findings
                })
        return file_findings
    except Exception as e:
        print(f"Error processing {filepath}: {e}")
        return []

def extract_and_scan_csv(filepath):
    try:
        df = pd.read_csv(filepath, encoding='utf-8', errors='ignore')
        findings = scan_dataframe_for_pii(df)
        return findings
    except Exception as e:
        print(f"Error processing {filepath}: {e}")
        return []

def extract_text_from_txt(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as file:
            text = file.read()
            return [text]
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return []

def extract_text_from_pdf(filepath):
    try:
        text = extract_pdf_text(filepath)
        return [text]
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return []

def extract_text_from_docx(filepath):
    try:
        doc = docx.Document(filepath)
        text = '\n'.join([para.text for para in doc.paragraphs])
        return [text]
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return []

def scan_text_for_pii(text_chunks):
    all_results = []
    for text in text_chunks:
        try:
            results = analyzer.analyze(text=text, entities=[], language='en')
            if results:
                all_results.append((text, results))
        except Exception as e:
            print(f"Error during PII analysis: {e}")
            continue
    return all_results

def extract_text(filepath):
    if filepath.endswith('.txt'):
        return extract_text_from_txt(filepath)
    elif filepath.endswith('.pdf'):
        return extract_text_from_pdf(filepath)
    elif filepath.endswith('.docx'):
        return extract_text_from_docx(filepath)
    else:
        return []

def scan_directory(directory):
    report = {}
    for root, _, files in os.walk(directory):
        for filename in files:
            filepath = os.path.join(root, filename)
            if filepath.endswith(('.xlsx', '.xls')):
                findings = extract_and_scan_excel(filepath)
                if findings:
                    report[filepath] = findings
            elif filepath.endswith('.csv'):
                findings = extract_and_scan_csv(filepath)
                if findings:
                    report[filepath] = findings
            else:
                # Process unstructured files
                text_chunks = extract_text(filepath)
                if text_chunks:
                    findings = scan_text_for_pii(text_chunks)
                    if findings:
                        report[filepath] = findings
    return report

def generate_report(report):
    for filepath, findings in report.items():
        print(f"File: {filepath}")
        if isinstance(findings, list) and findings and 'sheet' in findings[0]:
            # Structured data (Excel files)
            for sheet_info in findings:
                sheet_name = sheet_info['sheet']
                sheet_findings = sheet_info['findings']
                print(f"  Sheet: {sheet_name}")
                for finding in sheet_findings:
                    row = finding['row']
                    column = finding['column']
                    text = finding['text']
                    for entity in finding['pii_entities']:
                        entity_type = entity.entity_type
                        score = entity.score
                        print(f"    Cell ({row}, {column}): '{text}'")
                        print(f"      Entity: {entity_type}, Confidence Score: {score:.2f}")
                print('-' * 40)
        elif isinstance(findings, list) and findings and 'row' in findings[0]:
            # Structured data (CSV files)
            for finding in findings:
                row = finding['row']
                column = finding['column']
                text = finding['text']
                for entity in finding['pii_entities']:
                    entity_type = entity.entity_type
                    score = entity.score
                    print(f"  Row {row}, Column '{column}': '{text}'")
                    print(f"    Entity: {entity_type}, Confidence Score: {score:.2f}")
            print('-' * 40)
        else:
            # Unstructured data
            for text, results in findings:
                for result in results:
                    entity = result.entity_type
                    start = result.start
                    end = result.end
                    score = result.score
                    detected_text = text[start:end]
                    print(f"  Entity: {entity}, Text: '{detected_text}', Position: {start}-{end}, Confidence Score: {score:.2f}")
            print('-' * 40)

def main():
    directory = input("Enter the directory to scan: ")
    report = scan_directory(directory)
    if report:
        generate_report(report)
    else:
        print("No PII found in the scanned directory.")

if __name__ == "__main__":
    main()
