import tkinter as tk
from tkinter import filedialog, messagebox, Text
import pandas as pd
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
from tkinter.scrolledtext import ScrolledText

# Init tkinter
root = tk.Tk()
root.title("KWS - Generator XML")
root.geometry("1200x675")
root.configure(bg="#2E2E2E")

# Header and Input Labels
header_font = ("Arial", 14, "bold")
label_font = ("Arial", 10)
text_color = "#FFFFFF"

# Title Label
tk.Label(root, text="Tytuł wyborów", font=label_font, fg=text_color, bg="#2E2E2E").grid(row=0, column=0, padx=10, pady=10, sticky="w")
title_entry = tk.Entry(root, width=40, font=("Arial", 10))
title_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")

# Vacancies Label
tk.Label(root, text="Ilość wakatów", font=label_font, fg=text_color, bg="#2E2E2E").grid(row=1, column=0, padx=10, pady=10, sticky="w")
vacancies_entry = tk.Entry(root, width=40, font=("Arial", 10))
vacancies_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")

# Upload Button
candidate_list = []

def upload_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        try:
            df = pd.read_excel(filepath, header=None)
            global candidate_list
            candidate_list = df.iloc[:, 0].tolist()
            messagebox.showinfo("Success", "Załadowano listę!")
        except Exception as e:
            messagebox.showerror("Error", f"Błąd w ładowaniu pliku: {e}")

upload_button = tk.Button(root, text="Wczytaj listę kandydatów (.xlsx)", command=upload_file, font=label_font)
upload_button.grid(row=2, column=0, columnspan=2, pady=10)

# Output Text (XML)
output_text = ScrolledText(root, wrap='word', font=("Consolas", 10), width=105, height=40, bg="#1E1E1E", fg="#D4D4D4", insertbackground="white")
output_text.grid(row=0, column=2, rowspan=10, padx=10, pady=10, sticky="n")

# Generating XML function
def create_xml(title, vacancies, candidates):
    questionnaire = ET.Element("questionnaire")
    questionnaire.set("xsi:noNamespaceSchemaLocation", "questionnaire.xsd")
    questionnaire.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

    # Section 2 - For Selection
    page1 = ET.SubElement(questionnaire, "page", id="Page1")
    header1 = ET.SubElement(page1, "header")
    header1.text = title
    questions1 = ET.SubElement(page1, "questions")
    multi = ET.SubElement(questions1, "multi", id="pytanieZa", maxAnswers=str(vacancies), required="hard", showID="false", showTick="false", naLabel="Nie głosuję pozytywnie za żadną z tych osób")
    multi_header = ET.SubElement(multi, "header")
    multi_header.text = "Poniżej zaznacz osoby za którymi głosujesz pozytywnie."
    answers = ET.SubElement(multi, "answers")

    # Add candidates to "for selection" answers
    for idx, candidate in enumerate(candidates, start=1):
        ET.SubElement(answers, "textitem", code=str(idx), value=candidate)

    # Section 3 - Against and Abstain
    page2 = ET.SubElement(questionnaire, "page", id="Page2")
    header2 = ET.SubElement(page2, "header")
    questions2 = ET.SubElement(page2, "questions")
    info = ET.SubElement(questions2, "information", id="Informacja", showID="false", showTick="false")
    info_header = ET.SubElement(info, "header")
    info_header.text = "Poniżej znajdują się osoby na które nie zagłosowałeś/aś pozytywnie. Określ czy Wstrzymujesz się, czy jesteś Przeciwko ich kandydaturze."

    # Subsections for each candidate in "against/abstain"
    for idx, candidate in enumerate(candidates, start=1):
        single = ET.SubElement(questions2, "single", id=f"pytanieInne{idx}", required="hard", showID="false")
        header = ET.SubElement(single, "header")
        header.text = candidate
        filter_elem = ET.SubElement(single, "filter")
        not_elem = ET.SubElement(filter_elem, "not")
        ET.SubElement(not_elem, "condition", aid=f"{idx}@pytanieZa", value="1")
        answers2 = ET.SubElement(single, "answers")
        ET.SubElement(answers2, "textitem", code="1", value="Przeciw")
        ET.SubElement(answers2, "textitem", code="2", value="Wstrzymany")

    rough_string = ET.tostring(questionnaire, encoding="utf-8")
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

# Highlight XML
def highlight_xml():
    output_text.tag_configure("tag", foreground="orange")
    output_text.tag_configure("attr", foreground="skyblue")
    output_text.tag_configure("string", foreground="lightgreen")

    for tag, color in [("<", "tag"), ("=", "attr"), ('"', "string")]:
        start = 1.0
        while True:
            start = output_text.search(tag, start, tk.END)
            if not start:
                break
            end = f"{start}+{len(tag)}c"
            output_text.tag_add(color, start, end)
            start = end

# Link UI and XML generation
def generate_xml():
    title = title_entry.get()
    try:
        vacancies = int(vacancies_entry.get())
    except ValueError:
        vacancies = 0

    if not title or vacancies == 0 or not candidate_list:
        messagebox.showwarning("Missing Data", "Nie uzupełniono wszystkich wymaganych danych.")
        return

    xml_data = create_xml(title, vacancies, candidate_list)
    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, xml_data)
    highlight_xml()  # Apply syntax highlighting after generating XML

# Generate button
generate_button = tk.Button(root, text="Generuj XML", command=generate_xml, font=label_font)
generate_button.grid(row=4, column=0, columnspan=2, pady=10)

root.mainloop()
