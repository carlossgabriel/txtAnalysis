import os

import spacy
import xlsxwriter
from collections import Counter

nlp = spacy.load('pt')

# TODO run the following: pip install xlsxwriter
# TODO iterate through text files
# TODO get date from file name
# TODO create a list with the name of everyone that wrote in the chat
# TODO count how many times the person has send a message
# TODO generate file

def print_txt(single_file):
    pessoas = []
    with open(single_file, encoding="utf8") as f:
        my_list = [line.rstrip('\n') for line in f]
        # print(my_list)
        for line in my_list:
            if line == '' or line == '\n':
                my_list.remove('')
            if ':' in line and (not line.split(":")[0].isnumeric() or line.split(":")[0] == ''):
                pessoas.append(line.split(":")[0])
        persons_dict = Counter(pessoas)
        return persons_dict.items()


def iterate_through_txts():
    directory = r'D:\Users\carlo\PycharmProjects\pythonOpenTxt\txts'
    workbook = xlsxwriter.Workbook('dataSheet.xlsx')
    for singleFile in os.scandir(directory):
        date = str(singleFile).split("(")[1].split(" ")[0]
        worksheet = workbook.add_worksheet(str(date))
        worksheet.write(0, 0, "NAME")
        worksheet.write(0, 1, 'COUNT')
        result = print_txt(singleFile)
        for index, value in enumerate(result, start=2):
            # x, y, data[0]=key, data=[1]=value
            worksheet.write('A'+str(index), value[0])
            worksheet.write('B'+str(index), value[1])
    workbook.close()


iterate_through_txts()

# print_txt()
