import os
from config.config import Config
from src.core.document_processor import DocumentProcessor


def main():
    listpage_url = Config.LISTPAGE_TEMPLATE_URL
    maininfo_url = Config.MAIN_INFO_TEMPLATE_URL 
    
    template_path = 'data/template.json'
    
    doc_processor = DocumentProcessor()
    output_title = "Combined Document"
    result_url = doc_processor.merge_google_docs(listpage_url, maininfo_url, output_title, template_path)
    
    if result_url:
        print(f"You can access the new document here: {result_url}")
    else:
        print("\nFailed to merge documents. Please check the error messages above.")


if __name__ == '__main__':
    main()
