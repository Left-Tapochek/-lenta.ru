import os
import re
import json
import nltk
import pymorphy2
import datetime
from pathlib import Path
from nltk.corpus import stopwords
import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS
from tkinter import filedialog
from tkinter import *
import tkinter.ttk as ttk
from PIL import Image
import ctypes.wintypes

def program_analiz_text():
    path_file = filename.replace('/', '\\\\')
    nltk.download('stopwords')
    
    def find_date(path):
        return re.findall('\d{4}_\d{2}_\d{2}', path)[-1]
    
    def parse_date(year, month, day):
        return datetime.date(year, month, day).strftime("%B")
    
    def normalization_of_words(path):
        morph = pymorphy2.MorphAnalyzer()
        with open(path, 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
            p = morph.parse(data['text'])[0]
            return p.normal_form

    def deleting_stop_words(text):
        stop_words = set(stopwords.words('russian'))
        lst_text = text.split()
        new_text = ''
        for i in lst_text:
            if i not in stop_words:
                new_text += i + ' '
        return new_text

    def delt_special_signs(text):
        return re.sub('\W+',' ', text)

    def app_month(month, dict_month):
        if not month in dict_month:
            dict_month[month] = ''
        return dict_month

    def cheсk_month(file_month, dict_month, text):
        for key_month in dict_month.keys():
            if file_month == key_month:
                dict_month[key_month] += text
        return dict_month
    
    def progress_bar_build(maxx, x_bar, y_bar, x_label, y_label, text_bar, window):
        progress_bar = ttk.Progressbar(window, length=510, orient="horizontal",
                                        mode="determinate", maximum=maxx, value=0) 
        label_1 = Label(text=text_bar).place(x=x_label, y=y_label)
        progress_bar.place(x = x_bar, y = y_bar)
        window.update()
        progress_bar['value'] = 0
        window.update()
        
        return progress_bar   
    
    def count_file(path_file):
        count = 0
        base_path = Path(path_file)
        for path_folders in base_path.iterdir():
            path = str(path_folders)
            base_path = Path(path_file + '\\\\' + path[-14:len(path)])
            for path_files in base_path.iterdir():
                count += 1
        return count
    
    def file_crawling(path_file):
        base_path = Path(path_file)
        for path_folders in base_path.iterdir():
            path = str(path_folders)
            base_path = Path(path_file + '\\\\' + path[-14:len(path)])
            for path_files in base_path.iterdir():
                yield str(path_files)
       
    
    maxx = count_file(path_file)       
    x_bar = 20
    y_bar = 110
    x_label = 30
    y_label = 80
    text_bar = 'Чтение файлов'
    progress_bar = progress_bar_build(maxx, x_bar, y_bar, x_label, y_label, text_bar, window)
    dict_month = {} 
    for path in file_crawling(path_file):
        date = find_date(path)
        par_date = parse_date(int(date[:4]), int(date[5:7]), int(date[8:10]))
        dict_month = app_month(par_date, dict_month)
        norm_text = normalization_of_words(path)
        text_without_stopwords = deleting_stop_words(norm_text)
        result_text = delt_special_signs(text_without_stopwords)
        result_dict = cheсk_month(par_date, dict_month, result_text)
        progress_bar['value'] += 1
        window.update()
        
    def count_month(result_dict):
        count = 0
        for i in result_dict.keys():
            count += 1
        return count
    
    maxx = count_month(result_dict)
    count = 0
    x_bar = 20
    y_bar = 180
    x_label = 30
    y_label = 150
    text_bar = 'Построение wordcloud и графика биграмм'
    progress_bar = progress_bar_build(maxx, x_bar, y_bar, x_label, y_label, text_bar, window)
    for text in result_dict.values():
        wordcloud = WordCloud(width= 3000, height = 2000, random_state=1, background_color='darkblue', colormap='cubehelix', collocations=False, stopwords = STOPWORDS).generate(text)
        wordcloud.to_file(f'картинка {count + 1}.png')
        progress_bar['value'] += 1
        window.update()
        count += 1
     
    def separation_bigrams(bigram):
        for word in bigram:
            yield word

    def count_word_in_text(word_bigram, text):
        dict_word = {}
        if not word_bigram in text:
            dict_word[word_bigram] = 0
        else:
            for word in text:
                if word_bigram == word:
                    if word_bigram in dict_word:
                        dict_word[word_bigram] += 1
                    else:
                        dict_word[word_bigram] = 1
        return dict_word

    def append_count_word(dict_word, count_word):
        for i in dict_word.values():
            count_word.append(i)

    def create_list_month(result_dict):
        month = []
        for mon in result_dict.keys():
            month.append(mon)
        return month
 
    def create_grafic_word(month, count_word, bigram):
        plt.plot(month, count_word)
        plt.title('График биграмм')
        plt.legend(bigram)
        plt.grid(True)
        plt.savefig('График биграмм.png')
                   
    bigram = 'год интернет'
    
    for word in separation_bigrams(bigram.split()):
        list_count_word = []
        for month_text in result_dict.values():
            text = month_text.split()
            dict_word = count_word_in_text(word, text)
            append_count_word(dict_word, list_count_word)
        month = create_list_month(result_dict)
        list_count_word = create_grafic_word(month, list_count_word, bigram.split())
        
    def get_path():
        CSIDL_PERSONAL= 40    
        SHGFP_TYPE_CURRENT= 0   

        buf= ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(0, CSIDL_PERSONAL, 0, SHGFP_TYPE_CURRENT, buf)
        return str(buf.value)
    
    def open_file(path):
        path = os.path.realpath(path)
        os.startfile(path)
    
    maxx = count_month(result_dict)
    x_bar = 20
    y_bar = 250
    x_label = 30
    y_label = 220
    text_bar = 'Создание GIF'
    progress_bar = progress_bar_build(maxx - 1, x_bar, y_bar, x_label, y_label, text_bar, window)
    
    frames = []
    
    for frame_number in range(1, maxx):
        progress_bar['value'] += 1
        window.update()
        frame = Image.open(f'картинка {frame_number}.png')
        frames.append(frame)
     
    frames[0].save(
        'wordcloud_gif.gif',
        save_all=True,
        append_images=frames[1:],  
        optimize=True,
        duration=1500,
        loop=0
    )
    
    path = get_path()
    path_gif = f'{path}\wordcloud_gif.gif'
    open_file(path_gif)
    path_grafic = f'{path}\График биграмм.png'
    open_file(path_grafic)
        
def browse_button():
    global folder_path, filename
    filename = filedialog.askdirectory()
    folder_path.set(filename)
            
window = Tk()
window.title('Приложение для анализа текста')
window.geometry('550x290')
label = Label(text="Выберите папку с данными", font='Arial 12', width=30, height=2).grid(row=0, column=0)
folder_path = StringVar()
entry = Entry(textvariable=folder_path, width=60).place(x=20, y=45)
button_1 = Button(text = 'Обзор', command=browse_button).place(x=390, y=41)
button_2 = Button(text = 'Начать анализ', command=program_analiz_text).place(x=440, y=41)
mainloop()

