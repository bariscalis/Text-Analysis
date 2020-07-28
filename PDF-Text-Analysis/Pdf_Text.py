# Text analyzing in PDF via GUI
import PyPDF2
import re
import tkinter as tk
from tkinter import filedialog as fd
import openpyxl
from itertools import tee, islice, chain


# To control words combination is trying to be able to find previous, current and next words
def previous_and_next(some_iterable):
    prevs, items, nexts = tee(some_iterable, 3)
    prevs = chain([None], prevs)
    nexts = chain(islice(nexts, 1, None), [None])
    return zip(prevs, items, nexts)


def search(myDict, lookup):
    for key, value in myDict.items():
        if lookup.lower() == value.lower():
            return key


# Integer is finding in text to sort descending
def sort_fn(x):
    y = x.find(":")
    z = x.find(" %")
    return int(x[y + 2:z])  # After : character rest of ist numeric


def SrtFn(x):
    y = x.find(":")
    return int(x[y + 3:-1])   # After : character rest of ist numeric


def remove_values_from_list(main_list, omitted_list):
    return [value for value in main_list if value.strip() not in omitted_list]


def dialog_window():
    entry.delete(0, fd.END)
    global filename
    filename = fd.askopenfilename(title="Select A File", filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*")))
    entry.insert(0, filename)
    return filename


def run():
    n_page = 0
    Uniqgroup = []
    UniqList = []
    t_Word = []


    # Twits' words are being compared with main category in excel
    wb = openpyxl.load_workbook('Word_Glossory.xlsx')
    sh = wb['Sheet1']
    glos = []

    file = open(filename, 'rb')  # Name of the PDF
    p_count = PyPDF2.PdfFileReader(file).numPages

    while n_page < p_count:
        p_Word = PyPDF2.PdfFileReader(file).getPage(n_page).extractText().split()
        t_Word.extend(p_Word)
        n_page += 1

    for item in t_Word:
        if re.search(r'\(|\)|/|[^a-zA-ZäÄöÖüÜß]', item):
            loc = t_Word.index(item)
            t_Word.remove(item)
            t_Word.insert(loc, re.sub(r'\(+|\)+|/+|[^a-zA-ZäÄöÖüÜß]+', "", item))

    # transform lowercase
    t_Word = [x.lower() for x in t_Word]

    stop_words = ["a", "about", "above", "after", "again", "against", "ain", "all", "am", "an", "and", "any", "are",
                  "aren", "aren't", "as", "at", "be", "because", "been", "before", "being", "below", "between", "both",
                  "but", "by", "can", "couldn", "couldn't", "d", "did", "didn", "didn't", "do", "does", "doesn",
                  "doesn't", "doing", "don", "don't", "down", "during", "each", "few", "for", "from", "further",
                  "had", "hadn", "hadn't", "has", "hasn", "hasn't", "have", "haven", "haven't", "having", "he", "her",
                  "here", "hers", "herself", "him", "himself", "his", "how", "i", "if", "in", "into", "is", "isn",
                  "isn't", "it", "it's", "its", "itself", "just", "ll", "m", "ma", "me", "mightn", "mightn't", "more",
                  "most", "mustn", "mustn't", "my", "myself", "needn", "needn't", "no", "nor", "not", "now", "o", "of",
                  "off", "on", "once", "only", "or", "other", "our", "ours", "ourselves", "out", "over", "own", "re",
                  "s", "same", "shan", "shan't", "she", "she's", "should", "should've", "shouldn", "shouldn't", "so",
                  "some", "such", "t", "than", "that", "that'll", "the", "their", "theirs", "them", "themselves",
                  "then", "there", "these", "they", "this", "those", "through", "to", "too", "under", "until", "up",
                  "ve", "very", "was", "wasn", "wasn't", "we", "were", "weren", "weren't", "what", "when", "where",
                  "which", "while", "who", "whom", "why", "will", "with", "won", "won't", "wouldn", "wouldn't", "y",
                  "you", "you'd", "you'll", "you're", "you've", "your", "yours", "yourself", "yourselves", "could",
                  "he'd", "he'll", "he's", "here's", "how's", "i'd", "i'll", "i'm", "i've", "let's", "ought", "she'd",
                  "she'll", "that's", "there's", "they'd", "they'll", "they're", "they've", "we'd", "we'll", "we're",
                  "we've", "what's", "when's", "where's", "who's", "why's", "would", ""]
    
    # Remove stopwords from main list
    t_Word = remove_values_from_list(t_Word, stop_words)

    # Excel columns was added in a dictionary to be able to search and to find word category
    group_dict = dict((g_word.value, word.value) for g_word, word in sh['A1':'B' + str(len(sh['B']))])

    # To look words combination previous, current and next words combine with each other
    for prev, current, nxt in previous_and_next(t_Word):
        try:
            if search(group_dict, current):
                glos += [search(group_dict, current)]

            if search(group_dict, prev + ' ' + current):
                glos += [search(group_dict, prev + ' ' + current)]

            if search(group_dict, prev + ' ' + current + ' ' + nxt):
                glos += [search(group_dict, prev + ' ' + current + ' ' + nxt)]

            if not (search(group_dict, current) or search(group_dict, prev + ' ' + current) or search(group_dict, prev + ' ' + current + ' ' + nxt)):
                glos += ["Words not yet grouped"]
        except TypeError:
            continue

    # All occurrence are being counted and being calculated frequency
    for item in glos:
        num = glos.count(item)
        density = item + "\t" * 3 + ": " + str(round((num / len(glos)) * 100)) + ' %' + ' (' + str(num) + ' times)'

        # get rid of from duplicated data
        if density not in Uniqgroup:
            Uniqgroup.append(density)

    # Sort as descending
    checkList = sorted(Uniqgroup, key=sort_fn, reverse=True)

    for item in t_Word:
        num = t_Word.count(item)
        denst = item.lower() + "\t\t : (" + str(num) + ")"

        if denst not in UniqList:   # get rid of from duplicated data
            UniqList.append(denst)

    srtList = sorted(UniqList, key=SrtFn, reverse=True)

    title = " WORDS GROUP" + " " * 22 + "RATE" + "\n" + "-" * 70 + "\n"

    scrollbar = tk.Scrollbar(lower_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    textbox = tk.Text(lower_frame)
    textbox.place(relwidth=0.95, relheight=1)
    textbox["font"] = ("Arial", "10")
    textbox.insert(1.0, title + "\n".join(checkList) + "\n" + "-"*70 + "\n\t\t ALL WORDS \n" + "-"*70 + "\n" + "\n".join(srtList))

    textbox.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=textbox.yview)


# For Window
root = tk.Tk()
root.title("Report")
root.geometry("650x500")

frame = tk.Frame(root, bg='#DDB6C6', bd=5)
frame.place(relx=0.55, rely=0.05, relwidth=0.85, relheight=0.1, anchor='n')

frame_file = tk.Frame(root, bg='#DDB6C6', bd=5)
frame_file.place(relx=0.075, rely=0.05, relwidth=0.1, relheight=0.1, anchor='n')

entry = tk.Entry(frame, font=("Arial", 10))
entry.place(relwidth=0.680, relheight=1)

button = tk.Button(frame, text='Get Result', bd=2.5, fg='#484C7F', command=lambda: run())
button.place(relx=0.7, relheight=1, relwidth=0.3)

lower_frame = tk.Frame(root, bg='#DDB6C6', bd=7.5)
lower_frame.place(relx=0.5, rely=0.175, relwidth=0.95, relheight=0.75, anchor='n')

label_text = tk.Label(lower_frame, text="")
label_text.place(relwidth=1, relheight=1)

# label2 = tk.Label(root, text="URL :")
# label2.place(relx=0.06, rely=0.12)

button_file = tk.Button(frame_file, text="File  :  ", bd=2.5, fg='#484C7F', command=lambda: dialog_window())
button_file.place(relheight=1)


root.mainloop()
