import math

from docx import Document
import csv


def open_file(path):
    doc = Document(path)
    return doc


if __name__ == '__main__':
    PATH = r'C:\Users\Chend\PycharmProjects\DictionaryQuiz\english_test.docx'
    document = open_file(PATH)
    i = 0
    with open('english.csv', 'w', encoding='utf-8') as file:
        fieldnames = ['word', 'meaning']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for table in document.tables:
            iterator = iter(table.rows)
            while i < len(table.rows):
                tmp = next(iterator)
                leftHand = tmp.cells[0].text
                rightHand = tmp.cells[1].text
                leftHand = leftHand.split('\n')
                rightHand = rightHand.split('\n')
                for j in range(len(leftHand)):
                    try:
                        if not rightHand[j] or not leftHand[j] or rightHand[j] == " ":
                            continue
                        writer.writerow(
                            {'word': str(leftHand[j]), 'meaning': str(rightHand[j])})
                    except IndexError:
                        continue
                percentage = round((i / len(table.rows)), 2)
                i += 1
                print("Progress: " + str(percentage) + "\n")
