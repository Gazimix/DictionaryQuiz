import math

from docx import Document
import csv



def open_file(path):
    doc = Document(path)
    return doc


def parse_docx_file(path, name):
    """
    Parses the docx file at the first parameter, and outputs a csv file of the word and it's
    meaning
    :param path: the path for the docx file to parse
    :param name: the output file
    """
    document = open_file(path)
    i = 0
    with open(name, 'w', encoding='utf-8') as file:
        fieldnames = ['word', 'meaning']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for table in document.tables:
            iterator = iter(table.rows)
            while i < len(table.rows):
                extract_row(iterator, writer)
                i += 1
                print_percentage(i, table)


def extract_row(iterator, writer):
    """
    Parses a row of the docx file and extracts it's data into the csv file.
    :param iterator: the iterator object
    :param writer: where to write to
    """
    tmp = next(iterator)
    leftHand = tmp.cells[0].text
    rightHand = tmp.cells[1].text
    leftHand = leftHand.split('\n')
    rightHand = rightHand.split('\n')
    k = 0
    for j in range(len(rightHand)):
        try:
            while leftHand[k] == "":
                k += 1
            if rightHand[j] == "":
                continue
            if not rightHand[j] or not leftHand[k] or rightHand[j] == " ":
                continue
            writer.writerow(
                {'word': str(leftHand[k]), 'meaning': str(rightHand[j])})
            k += 1
        except IndexError:
            continue


def print_percentage(i, table):
    percentage = round(i / len(table.rows), 2)
    print("Progress: " + str(percentage) + "\n")


if __name__ == '__main__':
    parse_docx_file(r'C:\Users\Chend\PycharmProjects\DictionaryQuiz\test.docx',
                    "tst.csv")
