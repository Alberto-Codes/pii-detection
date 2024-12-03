import os
from presidio_analyzer import AnalyzerEngine
from presidio_structured import StructuredEngine, StructuredAnalysis, PandasDataProcessor
import pandas as pd

# Initialize the AnalyzerEngine
analyzer_engine = AnalyzerEngine()

# Initialize the StructuredEngine
structured_engine = StructuredEngine(analyzer_engine=analyzer_engine, data_processor=PandasDataProcessor())

def scan_dataframe_for_pii(df):
    # Automatically detect entity types for all columns
    entity_mapping = {column: None for column in df.columns}
    analysis_definition = StructuredAnalysis(entity_mapping=entity_mapping)
    results = structured_engine.analyze(df, analysis_definition)
    
    findings = []
    for result in results:
        if result.recognizer_results:
            findings.append({
                'row': result.row_index + 1,
                'column': result.field_name,
                'text': result.field_value,
                'pii_entities': result.recognizer_results
            })
    return findings

def extract_and_scan_excel(filepath):
    try:
        df = pd.read_excel(filepath, sheet_name=None)
        findings = []
        for sheet_name, sheet_df in df.items():
            sheet_findings = scan_dataframe_for_pii(sheet_df)
            if sheet_findings:
                findings.append({
                    'sheet': sheet_name,
                    'findings': sheet_findings
                })
        return findings
    except Exception as e:
        print(f"Error processing {filepath}: {e}")
        return []

def extract_and_scan_csv(filepath):
    try:
        df = pd.read_csv(filepath, encoding='utf-8', errors='ignore')
        findings = scan_dataframe_for_pii(df)
        if findings:
            return [{'findings': findings}]
        else:
            # Handle unstructured files if needed
            pass
    except Exception as e:
        print(f"Error processing {filepath}: {e}")
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
            elif filepath.endswith(('.csv', '.txt')):
                findings = extract_and_scan_csv(filepath)
                if findings:
                    report[filepath] = findings
    return report

if __name__ == '__main__':
    directory = input("Enter the directory to scan: ")
    report = scan_directory(directory)
    
    if report:
        generate_report(report)
    else:
        print("No PII found in the scanned directory.")
