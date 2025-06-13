import docx
import pandas as pd
import os
import glob
import re
import numpy as np

def get_text(docx_path):
    doc = docx.Document(docx_path)
    return "\n".join([p.text for p in doc.paragraphs])

def extract_meeting_info(text):
    def extract_after(label, text, multiline=False):
        pattern = rf"{re.escape(label)}\s*(.*)"
        if multiline:
            # Get all lines after the label until a blank line or next label
            lines = text.split('\n')
            start = None
            for i, line in enumerate(lines):
                if line.strip().startswith(label):
                    start = i + 1
                    break
            if start is not None:
                values = []
                for line in lines[start:]:
                    if not line.strip() or ':' in line:
                        break
                    values.append(line.strip())
                return values
            return []
        else:
            m = re.search(pattern, text)
            return m.group(1).strip() if m else ''
    
    info = {}
    info['Голова комітету'] = extract_after('Голова комітету:', text)
    info['Заступник Голови комітету'] = extract_after('Заступник Голови комітету:', text)
    info['Секретар комітету'] = extract_after('Секретар комітету (без права голосу):', text)
    info['Відсутні'] = extract_after('Відсутні:', text)
    info['Запрошені'] = extract_after('Запрошені:', text)
    info['Кворум'] = extract_after('Кворум:', text)
    info['Порядок прийняття рішень'] = extract_after('Порядок прийняття рішень:', text)
    # Agenda
    agenda_match = re.search(r'ПОРЯДОК ДЕННИЙ:(.*?)Питання', text, re.DOTALL)
    if agenda_match:
        agenda = [line.strip(" .\t") for line in agenda_match.group(1).split('\n') if line.strip()]
        info['ПОРЯДОК ДЕННИЙ'] = " | ".join(agenda)
    else:
        info['ПОРЯДОК ДЕННИЙ'] = ''
    # Members
    members = extract_after('Члени комітету:', text, multiline=True)
    info['Члени комітету'] = members
    return info

def extract_questions(text):
    # Find all "Питання N:" blocks
    questions = []
    for match in re.finditer(r'Питання\s*\d+:(.*?)(?=Питання\s*\d+:|ПІБ|\Z)', text, re.DOTALL):
        block = match.group(1)
        # Agenda item
        agenda_match = re.search(r'^\s*(.*?)\n', block)
        agenda = agenda_match.group(1).strip() if agenda_match else ''
        # Speaker
        speaker_match = re.search(r'Виступив:([^\n]*)', block)
        speaker = speaker_match.group(1).strip() if speaker_match else ''
        # Votes
        votes_match = re.search(r'Голосували:(.*?)Вирішили:', block, re.DOTALL)
        votes = {'за': '', 'проти': '', 'утримались': ''}
        if votes_match:
            for v in votes:
                m = re.search(rf'{v}\s*-\s*(\d+)', votes_match.group(1))
                if m:
                    votes[v] = m.group(1)
        # Decision
        decision_match = re.search(r'Вирішили:(.*)', block)
        decision = decision_match.group(1).strip() if decision_match else ''
        questions.append({
            'agenda': agenda,
            'speaker': speaker,
            'votes': votes,
            'decision': decision
        })
    return questions

def extract_vote_table(text):
    # Find the table at the end (ПІБ, за, проти, утримались, примітка)
    table = []
    table_match = re.search(r'ПІБ\s+за\s+проти\s+утримались\s+примітка(.*)', text, re.DOTALL)
    if not table_match:
        return table
    lines = table_match.group(1).split('\n')
    current_name = None
    for line in lines:
        if not line.strip():
            continue
        # If line contains only a name
        if re.match(r'^[А-ЯІЇЄҐ][а-яіїєґ]+\s+[А-ЯІЇЄҐ][а-яіїєґ]+\s+[А-ЯІЇЄҐ][а-яіїєґ]+$', line.strip()):
            current_name = line.strip()
            continue
        # If line contains vote marks
        cells = line.strip().split('\t')
        if len(cells) >= 4 and current_name:
            table.append({
                'ПІБ': current_name,
                'за': cells[0].strip(),
                'проти': cells[1].strip(),
                'утримались': cells[2].strip(),
                'примітка': cells[3].strip() if len(cells) > 3 else ''
            })
            current_name = None
        elif len(cells) >= 4:
            # Sometimes name and votes are on the same line
            table.append({
                'ПІБ': cells[0].strip(),
                'за': cells[1].strip(),
                'проти': cells[2].strip(),
                'утримались': cells[3].strip(),
                'примітка': cells[4].strip() if len(cells) > 4 else ''
            })
    return table

def extract_doc_number_and_date(text):
    # Extract document number (handles underscores and spaces)
    number_match = re.search(r'ПРОТОКОЛ\s*№\s*_?(\d+)_?', text, re.IGNORECASE)
    number = number_match.group(1) if number_match else ''
    # Extract date (handles underscores, spaces, and Ukrainian months)
    date_match = re.search(
        r'«\s*_?(\d{1,2})_?\s*»\s*_+([а-яА-Яіїєґ]+)_+\s*(\d{4})\s*року', text)
    month_map = {
        'січня': '01', 'лютого': '02', 'березня': '03', 'квітня': '04',
        'травня': '05', 'червня': '06', 'липня': '07', 'серпня': '08',
        'вересня': '09', 'жовтня': '10', 'листопада': '11', 'грудня': '12'
    }
    if date_match:
        day = date_match.group(1)
        month_ua = date_match.group(2).strip('_').lower()
        year = date_match.group(3)
        month = month_map.get(month_ua, '')
        if month:
            date = f"{int(day):02d}.{month}.{year}"
        else:
            date = f"{day} {month_ua} {year}"
    else:
        date = ''
    return number, date

if __name__ == "__main__":
    folder = os.path.dirname(os.path.abspath(__file__))
    file_list = glob.glob(os.path.join(folder, '*.docx'))
    all_rows = []

    for file_path in file_list:
        filename = os.path.basename(file_path)
        text = get_text(file_path)
        info = extract_meeting_info(text)
        questions = extract_questions(text)
        vote_table = extract_vote_table(text)
        number, date = extract_doc_number_and_date(text)
        with open("test_output.txt", "w", encoding="utf-8") as file:
            file.write(text)
        # For each agenda item (Питання), create a block of rows
        for q_idx, q in enumerate(questions):
            # First row: meeting info + first member
            first_member = vote_table[0] if vote_table else {'ПІБ': np.nan, 'за': '', 'проти': '', 'утримались': '', 'примітка': ''}
            meeting_row = {
                'Файл': filename,
                'Дата': date,
                'Номер': number,
                'Голова комітету': info['Голова комітету'],
                'Заступник Голови комітету': info['Заступник Голови комітету'],
                'Члени комітету': info['Члени комітету'][0] if info['Члени комітету'] else np.nan,
                'Секретар комітету': info['Секретар комітету'],
                'Відсутні': info['Відсутні'],
                'Запрошені': info['Запрошені'],
                'Кворум': info['Кворум'],
                'Порядок прийняття рішень': info['Порядок прийняття рішень'],
                'ПОРЯДОК ДЕННИЙ': q['agenda'],
                'Виступив:': q['speaker'],
                'Голосували: за': q['votes']['за'],
                'Голосували: проти': q['votes']['проти'],
                'Голосували: утримались': q['votes']['утримались'],
                'Вирішили:': q['decision'],
                'ПІБ': first_member['ПІБ'],
                'за': first_member['за'],
                'проти': first_member['проти'],
                'утримались': first_member['утримались'],
                'примітка': first_member['примітка']
            }
            all_rows.append(meeting_row)
            # Next rows: for each additional member
            for member in vote_table[1:]:
                member_row = {k: np.nan for k in meeting_row}
                member_row['Члени комітету'] = member['ПІБ']
                member_row['ПІБ'] = member['ПІБ']
                member_row['за'] = member['за']
                member_row['проти'] = member['проти']
                member_row['утримались'] = member['утримались']
                member_row['примітка'] = member['примітка']
                all_rows.append(member_row)

    df = pd.DataFrame(all_rows)
    df.to_excel('extracted_info_flat.xlsx', index=False)
    print("Extraction complete. Results saved to extracted_info_flat.xlsx.")