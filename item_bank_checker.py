#!/usr/bin/python3
# Fergus McLellan - 21/04/2020

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re
import pandas as pd
import nltk
from bs4 import BeautifulSoup
import spacy
from spacy.matcher import Matcher
from spacy.matcher import PhraseMatcher

print("Please wait whilst spaCy language library is loaded...")
nlp = spacy.load('en_core_web_md')

"""
//////////////////////////////////////////////////////
Change global values for bad words here
//////////////////////////////////////////////////////
"""
BAD_STEM_WORDS_LIST =  ["you", "option", "accurate", "correct", "true", "can be", "only",
                        "statement" ]
BAD_OPTION_WORDS_LIST = ["only", "statement", "all of the above" ]

# Create spaCy PhraseMatchers (lowercase for case-insensitivity)
dnd_matcher = PhraseMatcher(nlp.vocab, attr="LOWER")
dnd_term = ["Drag and drop the"]
dnd_patterns = [nlp.make_doc(text) for text in dnd_term]
dnd_matcher.add("TerminologyList", None, *dnd_patterns)

canbe_matcher = PhraseMatcher(nlp.vocab, attr="LOWER")
canbe_term = ["can be"]
canbe_patterns = [nlp.make_doc(text) for text in canbe_term]
canbe_matcher.add("TerminologyList", None, *canbe_patterns)

negative_matcher = Matcher(nlp.vocab)
negative_matcher.add("NegativeList", None, [{'POS': 'VERB'}, {'DEP': 'neg'}],
                     [{'POS': 'AUX'}, {'DEP': 'neg'}])

class ItemCheckerapp(tk.Tk):

    def __init__(self, *args, **kwargs):

        tk.Tk.__init__(self, *args, **kwargs)
        tk.Tk.wm_title(self, "Item Bank Checker")

        self.ExcelFilenameVar = tk.StringVar()
        self.VocabFilenameVar = tk.StringVar()
        self.NaughtyStemWordsVar = tk.StringVar()
        self.NaughtyOptionsWordsVar = tk.StringVar()

        str1 = ", "
        self.NaughtyStemWordsVar.set(str1.join(BAD_STEM_WORDS_LIST))
        self.NaughtyOptionsWordsVar.set(str1.join(BAD_OPTION_WORDS_LIST))

        # Select Excel custom report file labels and options
        self.excel_filename_prompt_label = ttk.Label(self,
                                                     text="Select ED Custom Report Excel File:")
        self.excel_filename_prompt_label.grid(row=0, column=0, sticky='E')

        self.select_excel_file_location_button = tk.Button(self, text="SELECT EXCEL FILE",
                                                           command=self.show_excel_file_dialog)
        self.select_excel_file_location_button.grid(row=0, column=1, sticky='W')

        self.excel_filename_label = ttk.Label(self, text="Filename (path .xls):",
                                              background='white')
        self.excel_filename_label.grid(row=2, column=0, sticky='E')

        self.excel_filename = tk.Entry(self, textvariable=self.ExcelFilenameVar, width=50)
        self.excel_filename.grid(row=2, column=1)

        # Select Vocabulary text file (for topic specific spell checking) labels and options
        self.vocab_filename_prompt_label = ttk.Label(self,
        text="Select vocabulary list for spell checking:")
        self.vocab_filename_prompt_label.grid(row=3, column=0)

        self.select_vocab_file_location_button = tk.Button(self, text="SELECT VOCAB FILE",
                                                           command=self.show_vocab_file_dialog)
        self.select_vocab_file_location_button.grid(row=3, column=1, sticky='W')

        self.vocab_filename_label = ttk.Label(self, text="Filename (path .txt):",
                                              background='white')
        self.vocab_filename_label.grid(row=4, column=0, sticky='E')

        self.vocab_filename = tk.Entry(self, textvariable=self.VocabFilenameVar, width=50)
        self.vocab_filename.grid(row=4, column=1)

        # Naughty stem words label and text box
        self.stem_naughty_vocab_label = ttk.Label(self, text="Naughty list - stem words",
                                                  background='white')
        self.stem_naughty_vocab_label.grid(row=5, column=0, sticky='E')

        self.stem_naughty_vocab_text_entry = tk.Text(self, bg='white', borderwidth=2,
                                                     relief=tk.SUNKEN, height=4, width=50)
        self.stem_naughty_vocab_text_entry.grid(row=5, column=1, sticky='W')
        self.stem_naughty_vocab_text_entry.insert(0.0, self.NaughtyStemWordsVar.get())

        # Naughty options words label and text box
        self.options_naughty_vocab_label = ttk.Label(self, text="Naughty list - options words",
                                                     background='white')
        self.options_naughty_vocab_label.grid(row=6, column=0, sticky='E')

        self.options_naughty_vocab_text_entry = tk.Text(self, bg='white', borderwidth=2,
                                                        relief=tk.SUNKEN, height=4, width=50)
        self.options_naughty_vocab_text_entry.grid(row=6, column=1, sticky='W')
        self.options_naughty_vocab_text_entry.insert(0.0, self.NaughtyOptionsWordsVar.get())

        # Process Excel custom report button
        self.process_excel_button = tk.Button(self, text="PROCESS EXCEL FILE",
                                              command=self.process_wrapper)
        self.process_excel_button.grid(row=10, column=1)


    def show_excel_file_dialog(self):
        filename_and_path = filedialog.askopenfilename(initialdir='~',
                                                       title = "Select file",
                                                       filetypes = (("Excel files","*.xls"),
                                                       ("all files","*.*")))
        self.ExcelFilenameVar.set(filename_and_path)

    def show_vocab_file_dialog(self):
        filename_and_path = filedialog.askopenfilename(initialdir='~',
                                                       title = "Select file",
                                                       filetypes = (("Text files","*.txt"),
                                                       ("all files","*.*")))
        self.VocabFilenameVar.set(filename_and_path)

    def process_wrapper(self):
        excel_file_name=self.excel_filename.get()

        if not len(excel_file_name) > 0:
            messagebox.showerror(title="Error!",
                message="Please specify ED Custom Report Excel file to use for processing.")

        output_filename = re.sub(r'.xls', '_error_summary.xlsx', excel_file_name)
        print(output_filename)

        vocab_file_name = self.vocab_filename.get()
        # Create the vocab_list as a global so that it is available to all functions
        global vocab_list
        if len(vocab_file_name) > 1:
            vocab_list = text_file_to_list(vocab_file_name)
        else:
            vocab_list = []

        global naughty_stem_text_list
        global naughty_options_text_list
        naughty_stem_text_list = (self.stem_naughty_vocab_text_entry.get("1.0",
                                  'end-1c')).split(',')
        naughty_options_text_list = (self.options_naughty_vocab_text_entry.get("1.0",
                                     'end-1c')).split(',')

        # trim any leading/trailing spaces from naughty word lists
        naughty_stem_text_list = [item.strip() for item in naughty_stem_text_list]
        naughty_options_text_list = [item.strip() for item in naughty_options_text_list]

        process_excel_file(excel_file_name, output_filename)

def process_excel_file(excel_file_name, output_filename):

    # Create DataFrame as a global so that it is available to all functions
    global ITEMS_DF

    # Open Excel in DataFrame
    pd.set_option('max_colwidth', None)
    ITEMS_DF = pd.read_excel(excel_file_name, header=1)

    # drop all columns except for 'Question Number', 'Type', 'Stem Text', 'Option Text'
    ITEMS_DF = ITEMS_DF[['Question Number', 'Type', 'Stem Text', 'Option Text']]

    # create new stem_cleaned and option_cleaned columns, with html tags removed
    ITEMS_DF['stem_cleaned'] = ITEMS_DF['Stem Text'].apply(lambda stem: clean_html_tags(stem))
    ITEMS_DF['option_cleaned'] = ITEMS_DF['Option Text'].apply(lambda option: clean_html_tags(option))

    # Tokenize stems by word
    ITEMS_DF['stem_tokenized'] = ITEMS_DF['stem_cleaned'].apply(lambda stem: nltk.word_tokenize(stem))

    # Apply all checks against stem text
    ITEMS_DF['stem_errors'] = (ITEMS_DF['stem_cleaned'].apply(lambda stem: basic_stem_wording_checks(stem)) +
    ITEMS_DF['Question Number'].apply(lambda item_number: check_stem_img_cue(item_number)) +
    ITEMS_DF['Question Number'].apply(lambda item_number: check_stem_double_cue(item_number)) +
    ITEMS_DF['Question Number'].apply(lambda item_number: check_stem_dnd_wording(item_number)) +
    ITEMS_DF['Question Number'].apply(lambda item_number: check_stem_negative(item_number)) +
    ITEMS_DF['Question Number'].apply(lambda item_number: check_stem_spelling(item_number)))

    # Apply basic checks against options text
    ITEMS_DF['options_errors'] = (ITEMS_DF['option_cleaned'].apply(lambda options: basic_options_wording_checks(options)))

    # Create new error dataframe for output
    errors_df = (ITEMS_DF.where(((ITEMS_DF['stem_errors'].str.len() > 1) | (ITEMS_DF['options_errors'].str.len() > 1)), inplace=False)).dropna()

    if errors_df.shape[0] > 0:
        errors_df.drop(columns=['Type', 'stem_tokenized'], inplace=True)
        errors_df.to_excel(output_filename, index=False)
        messagebox.showinfo(title="Completed", message=("Error summary file created: " + \
            output_filename))
    else:
        messagebox.showinfo(title="Completed",
                            message=("No errors found. No output file produced."))

def text_file_to_list(filename):
    """
    open text file, and output contents as a list
    """
    input_file = open(filename, "r")
    vocab = []
    for line in input_file:
        stripped_line = line.strip()
        vocab.append(stripped_line)
    input_file.close()
    return vocab


def clean_html_tags(block_of_html_text):
    """
    extract text from stem and options (remove any html tags in the raw/source text)
    """
    # check that text has been passed to function
    if isinstance(block_of_html_text, str):
        # The bullet points produced by Exam Developer result in some weird text - clean manually
        block_of_html_text = re.sub(r'&bull;', ',', block_of_html_text)
        soup = BeautifulSoup(block_of_html_text, "html.parser")
        clean_text = soup.get_text().lstrip()
    else:
        clean_text = ""

    return clean_text


def basic_stem_wording_checks(text):
    # Basic stem wording checks

    # Check for multiple full stops
    error_text = ""
    if re.search(r'\.\.', text):
        error_text = "Multiple full stops found."
    # Check for multiple spaces
    if re.search(r'\s\s', text):
        error_text = error_text + "Multiple spaces found."
    # Check for bad words
    for bad_word in naughty_stem_text_list:
        if re.search(rf'{bad_word}', text.lower()):
            error_text = error_text + "Stem includes the word " + bad_word + "."
    return error_text


def basic_options_wording_checks(text):
    """
    Basic options wording check
    """
    # Check for multiple full stops
    error_text = ""
    if re.search(r'\.\.', text):
        error_text = "Multiple full stops found."
    # Check for multiple spaces
    if re.search(r'\s\s', text):
        error_text = error_text + "Multiple spaces found."
    # Check for bad words
    for bad_word in naughty_options_text_list:
        if re.search(rf'{bad_word}', text.lower()):
            error_text = error_text + "Option includes the word " + bad_word + "."
    return error_text


def check_stem_img_cue(item_number):
    """
    Stem wording check: if item contains an image
    1) stem should start with "Refer to the exhibit"
    2) image should be centered
    """
    error_text = ""
    raw_stem_text = str(ITEMS_DF.loc[ITEMS_DF['Question Number'] == item_number]['Stem Text'])
    if re.search(r'img', raw_stem_text):
        tokens = ITEMS_DF.loc[ITEMS_DF['Question Number'] == item_number]['stem_tokenized'].tolist()
        first_word = tokens[0][0]
        second_word = tokens[0][1]
        third_word = tokens[0][2]
        fourth_word = tokens[0][3]
        if not (first_word == "Refer" and second_word == "to" and third_word == "the" and "exhibit" in fourth_word):
            error_text = "Item has an image, but does not start with: Refer to the exhibit."
        if not (re.search(r'text-align: center', raw_stem_text) or re.search(r'text-align:center', raw_stem_text) or
                re.search(r'align=.center', raw_stem_text)):
            error_text = error_text + "Item has an image, but it does not appear to be centered."
    return error_text


def check_stem_double_cue(item_number):
    """
    Stem wording check: if item type is McqMultiple
    1) stem should end with "(Choose two.)"
    2) stem should include the word "two" ("Which two statements...", etc.)
    """
    error_text = ""
    item_type = (ITEMS_DF.loc[ITEMS_DF['Question Number'] == item_number]['Type'])
    if 'McqMultiple' in str(item_type):
        tokens = ITEMS_DF.loc[ITEMS_DF['Question Number'] == item_number]['stem_tokenized'].tolist()
        last_token = tokens[0][-1]
        second_last_token = tokens[0][-2]
        third_last_token = tokens[0][-3]
        fourth_last_token = tokens[0][-4]
        fifth_last_token = tokens[0][-5]

        if not (last_token == ")" and second_last_token == "." and third_last_token == "two"
                and fourth_last_token == "Choose" and fifth_last_token == "("):
            error_text = "Item is McqMultiple, but does not end with (Choose two.)"
        # check first part of stem for the word two
        if "two" not in str(tokens[0][:-5]).lower():
            error_text = error_text + "Item is McqMultiple, but does not appear to have double cue."
    return error_text


def check_stem_dnd_wording(item_number):
    """
    Stem wording check: check that stem includes "Drag and drop the"
    """
    error_text = ""
    item_type = (ITEMS_DF.loc[ITEMS_DF['Question Number'] == item_number]['Type'])
    if 'EnhancedMatching' in str(item_type):
        dnd_text = nlp(ITEMS_DF.loc[ITEMS_DF['Question Number'] == item_number]['stem_cleaned'].to_string())
        dnd_matches = dnd_matcher(dnd_text)

        if len(dnd_matches) > 0:
            # dnd Phrase matches found
            error_text = ""
        else:
            error_text = "Item is EnhancedMatching, but does not contain Drag and drop the ..."

    return error_text


def check_stem_negative(item_number):
    """
    Check stem for negatives wording: checks the last sentence in the stem for the
    use of a verb (or auxiliary verb) followed by a negative word, like "not"
    """
    error_text = ""
    stem_text = nlp(ITEMS_DF.loc[ITEMS_DF['Question Number'] == item_number]['stem_cleaned'].to_string())

    negative_matches = []
    number_of_sentences_in_text = len(list(stem_text.sents))
    item_type = (ITEMS_DF.loc[ITEMS_DF['Question Number'] == item_number]['Type'])
    if 'McqMultiple' in str(item_type):
        # Ignore the last sentence, as should be "(Choose two.)". Check 2nd last sentence for negative wording
        number_of_sentences_in_text = number_of_sentences_in_text - 1
    sentence_counter = 0
    for sentence in stem_text.sents:
        sentence_counter += 1
        # Check if this is the last sentence
        if sentence_counter == number_of_sentences_in_text:

            this_sentence_tokens = nlp(sentence.text)
            negative_matches = negative_matcher(this_sentence_tokens)
            if len(negative_matches) > 0:
                # print(sentence)
                # negative Phrase matches found
                error_text = error_text + "Negative words found at end of stem."
    return error_text


def check_stem_spelling(item_number):
    # Stem spellcheck
    error_text = ""
    stem_text = nlp(ITEMS_DF.loc[ITEMS_DF['Question Number'] == item_number]['stem_cleaned'].to_string())
    skip_next_token = 0
    for word in stem_text:
        word_text = word.text
        if skip_next_token > 0:
            skip_next_token -= 1
            continue

        if not len(word_text) > 0:
            # token is empty - ignore
            continue
        elif any(char.isdigit() for char in word_text):
            """
            this "word" contains digits, so not a proper word - could be a network name,
            password, etc., so ignore
            """
            continue
        elif not word_text.find('\"') or not word_text.find("\'"):
            # This token is an inverted comma, so might indicate a username, filename etc,
            # Ignore this token, next token (password), and
            # 2nd next token (closing inverted comma)
            skip_next_token = 2
            continue
        else:
            if "/" in word_text:
                # might be an API call - ignore
                continue
            else:
                if word_text not in nlp.vocab and word_text not in vocab_list:
                    """Try removing any residual punctuation or non alpha-numeric characters and
                    try again"""
                    cleaned_word_text = re.sub('[^A-Za-z0-9]+', '', word_text)
                    if cleaned_word_text not in nlp.vocab and cleaned_word_text not in vocab_list:
                        error_text = error_text + "Unrecognized spelling of word: " + word_text + "."
    return error_text


APP = ItemCheckerapp()
APP.mainloop()
