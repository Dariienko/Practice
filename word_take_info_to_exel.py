import docx
import pandas as pd
import os
import glob
import re

def is_regex(keyword):
    # Heuristic: treat as regex if contains regex special characters
    regex_chars = set('.^$*+?{}[]\\|()')
    return any(c in regex_chars for c in keyword)

def extract_from_word(docx_path, keywords):
    doc = docx.Document(docx_path)
    info = {k: '' for k in keywords}
    for para in doc.paragraphs:
        text = para.text.strip()
        for keyword in keywords:
            if is_regex(keyword):
                if re.search(keyword, text) and not info[keyword]:
                    info[keyword] = text
            else:
                if keyword.lower() in text.lower() and not info[keyword]:
                    info[keyword] = text
    return info

if __name__ == "__main__":
    folder = os.path.dirname(os.path.abspath(__file__))
    keywords = ['Голова комітету:', 'Заступник Голови комітету:', 'Члени комітету:',
                'Секретар комітету (без права голосу):', 'Відсутні:', 'Запрошені:',
                'Кворум:', 'Порядок прийняття рішень:','ПОРЯДОК ДЕННИЙ:',
                r'Питання\s+[\d]+:', 'Виступив:', 'Голосували:', 'Вирішили:', ]  # Replace with your keywords
    
    all_results = []
    file_list = glob.glob(os.path.join(folder, '*.docx'))
    for idx, file_path in enumerate(file_list, 1):
        info = extract_from_word(file_path, keywords)
        row = {'№': idx, 'File': os.path.basename(file_path)}
        row.update(info)
        all_results.append(row)
    
    df = pd.DataFrame(all_results)
    df.to_excel('extracted_info.xlsx', index=False)
    print("Extraction complete. Results saved to extracted_info.xlsx.")