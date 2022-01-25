#!usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import re
from nltk.stem import SnowballStemmer
snowball = SnowballStemmer(language = 'english')
import xlwt
from xlwt import Workbook
regex = re.compile(r'\b[a-zA-Z0-9]+\b')

class WordCounter:

    def __init__(self, excel_file_name, resume_file_name, job_description_file_name):
        self._excel_file_name = excel_file_name
        self._resume_file_name = resume_file_name
        self._job_description_file_name = job_description_file_name
        path = r"C:\Users\test\"
        self.TxtFiles(path)

    def TxtFiles(self, path):
        resume = self._resume_file_name
        job_des = self._job_description_file_name
        excel = self._excel_file_name

        files = [resume, job_des]
        returned_data = []
        count = 0
        for file in files:
            if 'txt' in file or 'rtf' in file:
                with open(file,"r",encoding='utf-8') as f:
                    file_txt = f.read()
                    file_words = re.findall(regex, file_txt)
                    stemmed_words = self.Stemming(file_words, count)
                    count += 1
                    returned_data.append(stemmed_words)
        self.ExcelFunct(excel, returned_data[0], returned_data[1])

    def Stemming(self, text, count):
        tokenized_words = [snowball.stem(token.lower()) for token in text]
        return self.MakeDict(tokenized_words, count)

    def MakeDict(self, words, count):
        bad_words = ['and', '&', 'move', 'pounds', 'walk', 'handshake', 'acronym', 'gone', 'isn\'t', 'summer', 'their', 'through', 'do', 'not', 'such', 'within', 'well', 'likely', 'i.e.', 'dreamers', 'doers', 'pretty', 'vietnam', 'out', 'at', 'or', 'the', 'is', 'an', 'to', 'as', 'about', 'of', 'this', 'for', 'while', 'in', 'by', 'you', 'with', 'when', 'are', 'will', 'and/or', 'who', 'must', 'off', 'its', 'whereve', 'what', 'that', 'looking', 'on', 'all', 'a', 'be', 'ed', 'ing', 'tion', 'ion', 'also', 'they', 'from', '_', '|', '*', '-']
        resume_dict = {}
        job_description_dict = {}

        if count == 0:
            for word in words:
                if word not in bad_words and word not in resume_dict:
                    resume_dict[word] = 1
                elif word not in bad_words and word in resume_dict:
                    resume_dict[word] += 1
        if count == 1:
            for word in words:
                if word not in bad_words and word not in job_description_dict:
                    job_description_dict[word] = 1
                elif word not in bad_words and word in job_description_dict:
                    job_description_dict[word] += 1

        if count == 0:
            return resume_dict
        if count == 1:
            return job_description_dict
    
    def ExcelFunct(self, excel, resume_dict, job_des_dict):
        wb = Workbook()
        sheet1 = wb.add_sheet('Sheet 1')
        counter = 0

        for key, value in resume_dict.items():
            if counter == 0:
                sheet1.write(counter, 0, 'Resume Word')
                sheet1.write(counter, 1, 'Resume Word Count')
            else:
                sheet1.write(counter, 0, key)
                sheet1.write(counter, 1, value)
            counter += 1
        
        count = 0
        for key, value in job_des_dict.items():
            if count == 0:
                sheet1.write(count, 2, 'Job Description Word')
                sheet1.write(count, 3, 'Job Description Word Count')
            else:
                sheet1.write(count, 2, key)
                sheet1.write(count, 3, value)
            count += 1
        
        wb.save(excel + ".xls")

new_excel_file_name = input('Name of new excel file: ')
resume_file = input('Name of resume file: ')
job_file = input('Name of job description file: ')
WordCounter(new_excel_file_name, resume_file, job_file)