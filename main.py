import random
import docx
from docx.shared import Inches, Cm
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ALIGN_VERTICAL


def get_random_bingo_section():
    bingo_answers = "bingoBois.txt"  # Replace with your file path
    with open(bingo_answers, "r", encoding="utf-8") as file:
        answer = file.readlines()
    file.close()

    bingo_boi = random.choice(answer)
    return bingo_boi


def make_doc_table(table_info, docs_number, title):

    document = docx.Document()

    heading = document.add_heading(title, level=0)
    heading.alignment = 1

    sections = document.sections
    for section in sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1)

    rows = 5
    cols = 5
    table_info.insert(12, 'FREE')

    table = document.add_table(rows=rows, cols=cols)
    table.style = 'Table Grid'

    # print(table_info)
    j = 0
    print(table_info)
    for row in range(rows):
        row_cells = table.rows[row].cells
        for col in range(cols):
            cell = row_cells[col]
            cell.text = table_info[j]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.paragraphs[0].alignment = WD_ALIGN_VERTICAL.CENTER
            cell.width = Inches(1.5)
            cell.height = Inches(1)
            j = j + 1

    for row in table.rows:
        row.height = Inches(1.5)

    document.save(f'bingo_card{docs_number}.docx')


def main():
    title = input("What is the bingo game's title? Ex Lawrence Bingo: ")
    docs_number = input("How many Bingo Cards do you want: ")


    for num in range(int(docs_number)):
        bingo_card = []

        full_card = False

        card_size = 1

        while not full_card:
            current_line = get_random_bingo_section().rsplit('\n')[0]

            if current_line not in bingo_card:

                bingo_card.append(current_line)

                card_size = card_size + 1
                if card_size == 25:
                    full_card = True

        make_doc_table(bingo_card, num + 1, title)


if __name__ == '__main__':
    main()
