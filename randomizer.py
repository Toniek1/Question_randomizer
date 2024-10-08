import pandas as pd
import random
from fpdf import FPDF

# 0# Prerequisite(
questions_file = input("Podaj nazwę pliku z pytaniami wraz z rozszerzeniem (np. pytania.xlsx): ")

# 1# Prerequisite: name test / describe test
test_title = input("Jak chcesz nazwać ten test?: ")

# 2# Prerequisite: count rows
df = pd.read_excel(questions_file)
num_rows = df.shape

# 3# Prerequisite: number of questions in test
while True:
    input_number_questions = input(f"Z ilu pytań ma się składać test (podaj liczbę całkowitą w przedziale 1-{num_rows}): ")
    
    if input_number_questions.isdigit():
        input_number_questions = int(input_number_questions)
        
        if 1 <= input_number_questions <= 30:
            print(f"Test będzie składał się z {input_number_questions} pytań.")
            break  # Break of the loop if the number is correct
        else:
            print("Liczba pytań musi być w przedziale 1-30.")
    else:
        print("Proszę podać poprawną liczbę całkowitą.")

# Step 1: Read questions from the Excel file
def load_questions_from_excel(file_path):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path)
    return df

# Step 2: Randomize the questions
def randomize_questions(df, num_questions=None):
    df_randomized = df.sample(n=num_questions)
    return df_randomized

# Step 3: Export the questions to a PDF

def export_questions_to_pdf(questions, output_file):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)  # Pobierz i załaduj czcionkę


    
    # Add title
    #pdf.set_font("DejaVuSans-Bold", 'B', 16)
    pdf.set_font("DejaVu", size=16)
    pdf.cell(200, 10, txt=test_title, ln=True, align='C')
    pdf.ln(10)

    #Add desc

    # Add questions and answers (a, b, c, d)
    pdf.set_font("DejaVu", size=10)
    for i, row in enumerate(questions.iterrows(), start=1):  # Start numbering from 1
        question_text = row[1]['Question']  # Excel column must be 'Question'
        answer_a = row[1]['Answer A']  # Column with answer A
        answer_b = row[1]['Answer B']  # Column with answer B
        answer_c = row[1]['Answer C']  # Column with answer C
        answer_d = row[1]['Answer D']  # Column with answer D
        
        # add question
        pdf.multi_cell(0, 10, f"{i}. {question_text}")
        pdf.ln(2)
        
        # add answers a, b, c, d
        pdf.multi_cell(0, 10, f"a) {answer_a}")
        pdf.multi_cell(0, 10, f"b) {answer_b}")
        pdf.multi_cell(0, 10, f"c) {answer_c}")
        pdf.multi_cell(0, 10, f"d) {answer_d}")
        pdf.ln(10)  # Odstęp między pytaniami

    # Save the PDF
    pdf.output(output_file)

# Main function to execute the script
def main():
    # File path to your Excel file with questions
    excel_file_path = questions_file  # Replace with the path to your file
    output_pdf_path = f"{test_title}.pdf"  # Output PDF file name
    
    # Load questions from Excel
    questions_df = load_questions_from_excel(excel_file_path)
    
    # Randomize the questions (you can specify the number of questions)
    randomized_questions = randomize_questions(questions_df, num_questions=input_number_questions)  # Change the number as needed
    
    # Export the randomized questions to a PDF
    export_questions_to_pdf(randomized_questions, output_pdf_path)
    print(f"PDF saved successfully as {output_pdf_path}")

if __name__ == "__main__":
    main()
