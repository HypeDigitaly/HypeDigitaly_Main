import requests
import os
import json
from datetime import datetime
import openpyxl  # Přidejte tento import na začátek souboru
from openpyxl.styles import Font, Alignment
import re
from openpyxl.chart import PieChart, Reference
from openpyxl.styles import PatternFill
import random
import matplotlib.pyplot as plt
from io import BytesIO
from openpyxl.drawing.image import Image
import numpy as np

# Funkce pro načtení konfigurace ze souboru
def load_config():
    config_file = 'ExportConvos_config.txt'
    config = {}
    with open(config_file, 'r') as config_file:
        for line in config_file:
            key, value = line.strip().split('=', 1)
            if key == 'CATEGORIES':
                config[key] = value.strip('[]').split(',')
            else:
                config[key] = value.strip()
    return config

# Načtení konfigurace
config = load_config()

# Použití načtených hodnot
AUTH_TOKEN = config['AUTH_TOKEN']
PROJECT_ID = config['PROJECT_ID']
START_DATE = config['START_DATE']
END_DATE = config['END_DATE']
OUTPUT_DIRECTORY = config['OUTPUT_DIRECTORY']
CATEGORIES = config['CATEGORIES']

# Constants
BASE_URL = "https://api.voiceflow.com/v2/transcripts"
HEADERS = {
    "Authorization": AUTH_TOKEN,
    "accept": "application/json"
}

def create_output_directory():
    if not os.path.exists(OUTPUT_DIRECTORY):
        os.makedirs(OUTPUT_DIRECTORY)
        print(f"Created directory: {OUTPUT_DIRECTORY}")
    else:
        print(f"Directory already exists: {OUTPUT_DIRECTORY}")

def get_transcript_ids():
    url = f"{BASE_URL}/{PROJECT_ID}"
    params = {
        "startDate": START_DATE,
        "endDate": END_DATE
    }
    response = requests.get(url, headers=HEADERS, params=params)
    response.raise_for_status()
    return [transcript["_id"] for transcript in response.json()]

def get_transcript_dialog(transcript_id):
    url = f"{BASE_URL}/{PROJECT_ID}/{transcript_id}"
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    return response.json()

def extract_messages(dialog):
    messages = []
    for turn in dialog:
        if turn["type"] == "debug" and "payload" in turn and "payload" in turn["payload"]:
            debug_message = turn["payload"]["payload"].get("message", "")
            if "CategoryFilter" in debug_message:
                messages.append({
                    "role": "DEBUG",
                    "content": debug_message,
                    "timestamp": turn.get("startTime", "")
                })
        elif turn["type"] == "request":
            if "payload" in turn and "query" in turn["payload"].get("payload", {}):
                messages.append({
                    "role": "HUMAN",
                    "content": turn["payload"]["payload"]["query"],
                    "timestamp": turn.get("startTime", "")
                })
        elif turn["type"] == "text" and "payload" in turn and "message" in turn["payload"].get("payload", {}):
            messages.append({
                "role": "BOT",
                "content": turn["payload"]["payload"]["message"],
                "timestamp": turn.get("startTime", "")
            })
    return messages

def save_transcript_to_txt(transcript_id, messages):
    filename = os.path.join(OUTPUT_DIRECTORY, f"transcript_{transcript_id}.txt")
    with open(filename, 'w', encoding='utf-8') as file:
        for message in messages:
            if message['role'] == 'DEBUG':
                file.write(f"DEBUG: {message['content']}\n")
            else:
                file.write(f"{message['role']}: {message['content']}\n")
            file.write("----------\n")  # Přidá oddělovací čáru
    print(f"Transkript uložen do {filename}")

def count_human_occurrences():
    human_count = 0
    for filename in os.listdir(OUTPUT_DIRECTORY):
        if filename.startswith("transcript_") and filename.endswith(".txt"):
            file_path = os.path.join(OUTPUT_DIRECTORY, filename)
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                human_count += content.count("HUMAN:")
    return human_count

def count_category_occurrences():
    category_counts = {category: 0 for category in CATEGORIES}
    
    for filename in os.listdir(OUTPUT_DIRECTORY):
        if filename.startswith("transcript_") and filename.endswith(".txt"):
            file_path = os.path.join(OUTPUT_DIRECTORY, filename)
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                for category in CATEGORIES:
                    exact_match = f'\\"{category}\\"'
                    category_counts[category] += content.count(exact_match)
    
    return category_counts

def create_pie_chart(category_counts):
    labels = list(category_counts.keys())
    sizes = list(category_counts.values())
    
    plt.figure(figsize=(10, 8))
    plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
    plt.axis('equal')
    plt.title('Rozložení kategorií')
    
    # Uložení grafu jako obrázek
    img_path = os.path.join(OUTPUT_DIRECTORY, 'category_distribution.png')
    plt.savefig(img_path)
    plt.close()
    
    return img_path

def update_excel_report(human_count, category_counts):
    excel_filename = os.path.join(OUTPUT_DIRECTORY, f"{START_DATE} to {END_DATE}_Report.xlsx")
    
    # Kontrola, zda soubor existuje, pokud ne, vytvoříme nový
    if not os.path.exists(excel_filename):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"
        
        # Základní struktura reportu
        ws['A1'] = f"REPORT [{START_DATE} - {END_DATE}]"
        ws['A3'] = "CELKOVÝ POČET AI ODPOVĚDÍ ZA DANÉ OBDOBÍ:"
        ws['A5'] = "KATEGORIE"
        ws['B5'] = "POČET"
        ws['C5'] = "PROCENTO"
        
        # Uložíme nově vytvořený soubor
        wb.save(excel_filename)
        print(f"Vytvořen nový soubor: {excel_filename}")
    
    # Nyní otevřeme existující soubor
    wb = openpyxl.load_workbook(excel_filename)
    ws = wb.active
    
    ws['B3'] = human_count
    
    print("\nVýsledky hledání kategorií:")
    print(f"Celkový počet HUMAN odpovědí: {human_count}")
    print("Počty výskytů jednotlivých kategorií:")
    
    total_category_count = sum(category_counts.values())
    
    # Seřazení kategorií sestupně podle počtu
    sorted_categories = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)
    
    for i, (category, count) in enumerate(sorted_categories, start=6):
        ws[f'A{i}'] = category
        ws[f'B{i}'] = count
        ws[f'C{i}'] = f'{count / total_category_count:.2%}'
        print(f"{category}: {count} ({count / total_category_count:.2%})")
    
    # Vytvoření koláčového grafu pomocí matplotlib
    plt.figure(figsize=(10, 8))
    labels = [cat for cat, count in sorted_categories if count > 0]
    sizes = [count for _, count in sorted_categories if count > 0]
    colors = plt.cm.Set3(np.linspace(0, 1, len(labels)))
    
    plt.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
    plt.axis('equal')
    plt.title('Rozložení kategorií')
    plt.legend(labels, loc="center left", bbox_to_anchor=(1, 0.5))
    
    # Uložení grafu jako obrázek
    img_path = os.path.join(OUTPUT_DIRECTORY, 'category_distribution.png')
    plt.savefig(img_path, bbox_inches='tight')
    plt.close()
    
    # Vložení obrázku do Excel souboru
    img = Image(img_path)
    img.width = 500
    img.height = 400
    ws.add_image(img, 'E5')
    
    wb.save(excel_filename)
    print(f"\nReport aktualizován a uložen do souboru: {excel_filename}")
    print(f"Obrázek grafu uložen do: {img_path}")

def save_report_to_excel(human_count):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    
    # Nadpis
    ws.merge_cells('A1:C1')
    ws['A1'] = f"REPORT [{START_DATE} - {END_DATE}]"
    ws['A1'].font = Font(size=16, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center')
    
    # Hlavní informace
    ws['A3'] = "CELKOVÝ POČET AI ODPOVĚDÍ ZA DANÉ OBDOBÍ:"
    ws['B3'] = human_count
    ws['A3'].font = Font(size=14)
    ws['B3'].font = Font(size=14, bold=True)
    
    # Přidání nové tabulky s kategoriemi
    ws['A5'] = "KATEGORIE"
    ws['B5'] = "POČET"
    ws['A5'].font = Font(bold=True)
    ws['B5'].font = Font(bold=True)
    
    for i, category in enumerate(CATEGORIES, start=6):
        ws[f'A{i}'] = category
        ws[f'B{i}'] = 0
    
    excel_filename = f"{OUTPUT_DIRECTORY}_Report.xlsx"
    wb.save(excel_filename)
    print(f"Report uložen do souboru: {excel_filename}")

def print_summary(human_count, category_counts):
    print("\nSouhrn zpracování:")
    print(f"1) Počet HUMAN výskytů: {human_count}")
    print("\n2) Počty výskytů kategorií:")
    for category, count in category_counts.items():
        print(f"   {category}: {count}")

def main():
    create_output_directory()
    transcript_ids = get_transcript_ids()
    
    total_human_count = 0
    
    for transcript_id in transcript_ids:
        dialog = get_transcript_dialog(transcript_id)
        messages = extract_messages(dialog)
        save_transcript_to_txt(transcript_id, messages)
        
        for message in messages:
            if message['role'] == 'HUMAN':
                total_human_count += 1
    
    category_counts = count_category_occurrences()
    update_excel_report(total_human_count, category_counts)
    
    print("Zpracování dokončeno.")

if __name__ == "__main__":
    main()