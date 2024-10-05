import openpyxl
import random
import string
import re

def read_terms_from_excel(file_name):
    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active
    terms = []

    for row in sheet.iter_rows(values_only=True):
        term = ''
        for cell in row:
            if cell is not None:
                term += str(cell) + ' '

        terms.append(term.strip())
    
    return terms

def create_question(term):
    words = term.split()
    valid_words = []

    for word in words:
        if len(word) > 3:
            valid_words.append(word)
    
    if not valid_words:
        return term, None

    correct_answer = random.choice(valid_words)
    missing_word_index = words.index(correct_answer)
    words[missing_word_index] = '_____'
    question = ' '.join(words)

    return question, correct_answer

def strip_punctuation(text): #remove punctation and parentheses 
    return text.translate(str.maketrans('', '', string.punctuation + '()')) 

def normalize_answer(text): #transform answer to lower-case, remove non-alpanumeric and space
    return re.sub(r'[^a-zA-Z0-9\s]', '', text).strip().lower()

def start_quiz(terms, num_questions, show_answer):
    score = 0
    used_terms = []
    
    for question_number in range(1, num_questions + 1):
        term = random.choice(terms)
        while term in used_terms:
            term = random.choice(terms)
        
        used_terms.append(term)
        question, missing_word = create_question(term)
        
        if missing_word is None:
            print("No valid word to blank out. Skipping this question...")
            continue
        
        print(f"Question {question_number}: Fill in the blank: {question}")
        answer = input("Your answer: ").strip()
        
        normalized_missing_word = normalize_answer(missing_word)
        normalized_answer = normalize_answer(answer)
        
        if normalized_answer == normalized_missing_word:
            print("Correct!")
            score += 1
        else:
            if show_answer == 'Y':
                print(f"The correct answer is: '{missing_word}'")
            else:
                print(f"Wrong!")
    
    print(f"\nQuiz Over! Your score: {score}/{num_questions}")

file_name = 'terms.xlsx'
try:
    terms = read_terms_from_excel(file_name)
except PermissionError:
    print("\nClose the terms.xlsx file to proceed\n")
    quit()

print(f"Available terms: {len(terms)}")

num_questions = int(input("How many questions would you like? "))

while num_questions > len(terms) or num_questions <= 0:
    print(f"Please enter a valid number (1 to {len(terms)}): ")
    num_questions = int(input("How many questions would you like? "))

# Input to show the correct answer for every wrong answer
show_answer = ''
while show_answer not in ['Y', 'N']:
    show_answer = input('Would you like to show the right answer for every wrong answer? (Y/N)\n').strip().upper()
    if show_answer not in ['Y', 'N']:
        print("Please enter 'Y' or 'N'.")

start_quiz(terms, num_questions, show_answer)
