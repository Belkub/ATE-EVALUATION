import streamlit as st
import numpy as np
import math
import pandas as pd

import pickle
#from pydub import AudioSegment
import cloudpickle
from datetime import datetime
import openpyxl as xl


import speech_recognition as sr

try:

    def voice():
        def speech_to_text(audio_file):
            recognizer = sr.Recognizer()
            with sr.AudioFile(audio_file) as source:
                audio = recognizer.record(source)
                try:
                    text = recognizer.recognize_google(audio, language="ru-RU")
                    return text
                except sr.UnknownValueError:
                    return "Could not understand audio"
                except sr.RequestError as e:
                    return f"Error: {str(e)}"
        recognizer = sr.Recognizer()   
        audio_value = st.audio_input("Record a voice message")
        if audio_value:
            st.audio(audio_value)
            text = speech_to_text(audio_value)
            return text

    wb=xl.Workbook()
    st.header("chart_mobile/sampling")

    df_c = pd.read_excel('company_1.xlsx')
    df_f = pd.read_excel('Example.xlsx')
    #df_f['Дата теста/отправки'] = pd.to_datetime(df_f['Дата теста/отправки'])
    if 'Unnamed: 0' in df_c.columns:
        df_c = df_c.drop('Unnamed: 0', axis = 1)
    if 'Unnamed: 0' in df_f.columns:
        df_f = df_f.drop('Unnamed: 0', axis = 1)
    df_c = df_c.dropna(subset=['Компания'])
    df_c = df_c.drop_duplicates()


    ss = st.sidebar.checkbox('Изменить имя образца/удалить')
    pp = st.sidebar.checkbox('Сформировать сводную таблицу образцов')
    if pp == True and ss == False:
        st.title('Создание сводной таблицы')
        bn = st.checkbox('Вывести таблицу и ограничить число полей')
        if bn:
            number = df_f.columns
            num = st.multiselect('Выбрать значимые поля', list(df_f.columns))
            if num:
                number = num
            bnn = st.checkbox('Записать сводную таблицу в файл .xlsx')
            if bnn:
                st.warning('Создайте папку Data на диске С для записи данных')
                if st.button('Показать свод и записать в файл C:\Data\table.xlsx'):
                    st.dataframe(df_f[number]) 
                    df_f[number].to_excel("C:\\Data\\table.xlsx")
            else:
                if st.button('Показать таблицу'): 
                    st.dataframe(df_f[number]) 
        else:
            number = list(df_f.columns)
        
        spp = float(st.number_input('Установить число полей группировки: ', min_value = 1, max_value = 3, value = 1, step = 1))
        #new_row_1 = {'Компания':['Компания'],'Имя образца':['Имя образца'],'Дата теста/отправки':['Дата теста/отправки'], 'Производитель/поставщик':['Производитель/поставщик'], 'Тестировщик':['Тестировщик'], 'Отрасль':['Отрасль'], 'Дисперсионная среда':['Дисперсионная среда'], 'Состав системы':['Состав системы'], 'Методика':['Методика'], 'Результат':['Результат'], 'Примечания':['Примечания'], 'Финальное решение':['Финальное решение']}
        new_row_1 = {}
        for i in number:
            new_row_1[i] = i
        p_1 = st.selectbox('Выбрать поле_1: ', list(new_row_1))
        p_11 = st.multiselect(f'Выбрать позиции поля {p_1}', list(df_f[p_1].unique()))
        if p_11:
            df_ff = df_f.loc[df_f[p_1].isin(p_11)]
            if spp == 1:
                ta = st.checkbox('Записать таблицу в файл .xlsx')
                if ta:
                    st.warning('Создайте папку Data на диске С для записи данных')
                    if st.button('Показать и записать таблицу в файл C:\Data\table.xlsx'):
                        st.dataframe(df_ff[number]) 
                        df_ff[number].to_excel("C:\\Data\\table.xlsx")
                else:
                    if st.button('Показать сводную таблицу'): 
                        st.dataframe(df_ff[number]) 
            else:
                del new_row_1[p_1]
                p_2 = st.selectbox('Добавить поле_2: ', list(new_row_1))
                p_22 = st.multiselect(f'Выбрать позиции поля {p_2}', list(df_ff[p_2].unique()))
                if p_22:
                    df_ff = df_ff.loc[df_ff[p_2].isin(p_22)]
                    if spp == 2:
                        ta = st.checkbox('Записать таблицу в файл .xlsx')
                        if ta:
                            st.warning('Создайте папку Data на диске С для записи данных')
                            if st.button('Показать и записать таблицу в файл C:\Data\table.xlsx'):
                                st.dataframe(df_ff[number]) 
                                df_ff[number].to_excel("C:\\Data\\table.xlsx")
                        else:
                            if st.button('Показать сводную таблицу'): 
                                st.dataframe(df_ff[number]) 
                    else:
                        del new_row_1[p_2]
                        p_3 = st.selectbox('Добавить поле_3: ', list(new_row_1))
                        p_33 = st.multiselect(f'Выбрать позиции поля {p_3}', list(df_ff[p_3].unique()))
                        if p_33:
                            df_ff = df_ff.loc[df_ff[p_3].isin(p_33)]
                            if spp == 3:
                                ta = st.checkbox('Записать таблицу в файл .xlsx')
                                if ta:
                                    st.warning('Создайте папку Data на диске С для записи данных')
                                    if st.button('Показать и записать таблицу в файл C:\Data\table.xlsx'):
                                        st.dataframe(df_ff[number]) 
                                        df_ff[number].to_excel("C:\\Data\\table.xlsx")
                                else:
                                    if st.button('Показать сводную таблицу'): 
                                        st.dataframe(df_ff[number]) 

                
        
    elif ss == True and pp == False:
        a_1 = {}
        for i in list(df_f['Компания']):
            a_1[i] = i
        t_2 = st.selectbox('Выбрать компанию: ', list(a_1))
        df_e1 = df_f.loc[df_f['Компания'] == t_2]
        b_1 = {}
        for i in list(df_e1['Имя образца']):
            b_1[i] = i
        e_2 = st.selectbox('Выбрать образец: ', list(b_1))
        num = list(df_e1.loc[df_e1['Имя образца'] == e_2, 'Номер'])[0]
        print(num)
        ss_1 = st.sidebar.checkbox('Изменить имя образца')
        dd = st.sidebar.checkbox('Заполнить/изменить поля')

        if ss_1 == True and dd == False:
            st.title('Изменение имени/удаление образца')
            with open("file_1.txt", "w") as file:  
                file.write(e_2)
                file.close()
            with open("file_1.txt", "r") as file:  
                example_1 = file.read()
                file.close()
            rr_1 = st.checkbox('Глосовой ввод')
            if rr_1:
                text = voice()
                if text:
                    with open("file_1.txt", "w") as file:  
                        file.write(text)
                        file.close()
                    example_1 = text
            example = st.text_input('Название компании: ', value = example_1)
            with open("file_1.txt", "w") as file:  
                file.write(example)
                file.close()
            with open("file_1.txt", "r") as file:  
                text = file.read()
                file.close()
            if st.button('Записать новое имя'):
                df_f = df_f.dropna(subset=['Имя образца'])
                df_f = df_f.drop_duplicates(subset=['Компания','Имя образца'])
                df_f.loc[(df_f['Компания'] == t_2) & (df_f['Номер'] == num), 'Имя образца'] = text
                df_f.to_excel('Example.xlsx')
            ss_2 = st.checkbox('Удалить образец')
            if ss_2:     
                if st.button('Удалить информацию об образце'):
                    df_f = df_f.loc[df_f['Имя образца'] != e_2]
                    df_f.to_excel('Example.xlsx')
    ##            
        elif dd == True and ss_1 == False:
            st.title('Заполнение/изменение полей')
            now = datetime.now()
            new_row_1 = {'Дата теста/отправки':['Дата теста/отправки'], 'Производитель/поставщик':['Производитель/поставщик'], 'Тестировщик':['Тестировщик'], 'Отрасль':['Отрасль'], 'Дисперсионная среда':['Дисперсионная среда'], 'Состав системы':['Состав системы'], 'Методика':['Методика'], 'Результат':['Результат'], 'Примечания':['Примечания'], 'Финальное решение':['Финальное решение']}
            s_2 = st.selectbox('Выбрать поле: ', list(new_row_1))
    ##        comp = list(df_c['Компания'])
            with open("file_1.txt", "w") as file:  
                file.write(str(list(df_e1.loc[df_e1['Имя образца'] == e_2, s_2])[0]))
                file.close()    

            def fr(type):           
                rr_11 = st.checkbox('Глосовой ввод')
                if rr_11:
                    text_22 = voice()
                    if text_22:
                        with open("file_1.txt", "w") as file:  
                            file.write(text_22)
                            file.close()
                with open("file_1.txt", "r") as file:  
                    text_11 = file.read()
                    file.close()
                change = st.text_input('Содержание поля: ', value = text_11)
                ddd = st.checkbox(f'Выбрать содержание поля {s_2}')
                if ddd:
                    t_3 = st.selectbox('Выбрать содержание: ', list(type))
                    change = st.text_input('Новое содержание поля: ', value = t_3)
                return change
            if s_2 in ['Производитель/поставщик', 'Тестировщик', 'Отрасль', 'Дисперсионная среда', 'Методика', 'Результат', 'Финальное решение']:
                if s_2 == 'Производитель/поставщик':
                    type = {'КЗПМ':'КЗПМ', 'Китай':'Китай', 'Другой':'Другой'}
                    change = fr(type)
                elif s_2 == 'Тестировщик':
                    type = {'КЗПМ':'КЗПМ', 'ИГЕО':'ИГЕО', 'Другой':'Другой'}
                    change = fr(type)
                elif s_2 == 'Отрасль':
                    type = {'ЛКМ водоэмульсионные':'ЛКМ водоэмульсионные', 'ЛКМ масляные':'ЛКМ масляные', 'ЛКМ спирт':'ЛКМ спирт', 'ЛКМ алкидные':'ЛКМ алкидные', 'ЛКМ порошок':'ЛКМ порошок', 'Нефтегаз':'Нефтегаз', 'Литье':'Литье', 'Косметика':'Косметика', 'Клеи':'Клеи'}
                    change = fr(type)
                elif s_2 == 'Дисперсионная среда':
                    type = {'ДТ':'ДТ', 'Ксилол':'Ксилол', 'Керосин':'Керосин', 'ИПС':'ИПС', 'Мин масло':'Мин масло', 'Синтетика':'Синтетика'}
                    change = fr(type)
                elif s_2 == 'Методика':
                    type = {'ТУ Марка_1':'ТУ Марка_1', 'ТУ Марка_2':'ТУ Марка_2', 'Бурсервис':'Бурсервис', 'ХОС':'ХОС'}
                    change = fr(type)
                elif s_2 == 'Результат':
                    type = {'Соответствие':'Соответствие', 'Неполное соответствие':'Неполное соответствие', 'Несоответствие':'Несоответствие'}
                    change = fr(type)
                elif s_2 == 'Финальное решение':
                    type = {'Разрешен к поставке':'Разрешен к поставке', 'Повтор':'Повтор', 'Отвергнут':'Отвергнут'}
                    change = fr(type)
            elif s_2 == 'Дата теста/отправки':
                year = {}
                for i in range(2024,2030):
                    year[i] = i
                year_1 = st.selectbox('Выбрать год: ', list(year))
                month = {}
                for i in range(1,13):
                    month[i] = i
                month_1 = st.selectbox('Выбрать месяц: ', list(month))
                day = {}
                for i in range(1,32):
                    day[i] = i
                day_1 = st.selectbox('Выбрать день: ', list(day))
                with open("file_1.txt", "r") as file:  
                    text_11 = file.read()
                    file.close()
                st.warning(f'Содержание поля: {text_11}')
                change = [year_1, '-', month_1, '-', day_1]
                change = ' '.join([str(elem) for elem in change])
                
            else:
                with open("file_1.txt", "r") as file:  
                    text_1 = file.read()
                    file.close()
                rr_1 = st.checkbox('Глосовой ввод')
                if rr_1:
                    text_2 = voice()
                    if text_2:
                        with open("file_1.txt", "w") as file:  
                            file.write(text_2)
                            file.close()
                with open("file_1.txt", "r") as file:  
                    text_3 = file.read()
                    file.close()
                change = st.text_input('Содержание поля: ', value = text_3)
            with open("file_1.txt", "w") as file:  
                file.write(change)
                file.close()
            with open("file_1.txt", "r") as file:  
                text_3 = file.read()
                file.close()
            if st.button('Записать поле'):
                df_f = df_f.dropna(subset=['Компания', 'Имя образца'])
                df_f = df_f.drop_duplicates(subset=['Компания', 'Имя образца'])
                df_f.loc[(df_f['Компания'] == t_2) & (df_f['Имя образца'] == e_2), s_2] = text_3
                df_f.to_excel('Example.xlsx')
                print(df_f)
                
    else:
        st.title('Создание имени образца')
        a = {}
        for i in list(df_c['Компания'].unique()):
            a[i] = i
        t_1 = st.selectbox('Выбрать компанию: ', list(a))
        if t_1 in list(df_f['Компания']):
            df_e = df_f.loc[df_f['Компания'] == t_1]
            b = {}
            for i in list(df_e['Имя образца']):
                b[i] = i
            e_1 = st.selectbox('Проверить образцы: ', list(b))
        rr = st.checkbox('Голосовой ввод')
        text = 'new example'
        with open("file_1.txt", "w") as file:  
            file.write(text)
            file.close()
        if rr:
            text = voice()
            if text:
                with open("file_1.txt", "w") as file:  
                    file.write(text)
                    file.close()
        with open("file_1.txt", "r") as file:  
            text = file.read()
            file.close()
        print(text)
        example = st.text_input('Имя образца: ', value = text)
        with open("file_1.txt", "w") as file:  
            file.write(example)
            file.close()
        with open("file_1.txt", "r") as file:  
            text = file.read()
            file.close()
        print(text)
        ##wdf_c = st.checkbox('Записать название')
        if st.button('Создать новый образец'):
            now = datetime.now()
            g = 1
            if t_1 in list(df_f['Компания']): 
                g = len(df_e['Компания']) + 1
            new_row = {'Время':now, 'Компания': [t_1], 'Номер': [g], 'Имя образца': [text], 'Дата теста/отправки':[''], 'Производитель/поставщик':[''], 'Тестировщик':[''], 'Отрасль':[''], 'Дисперсионная среда':[''], 'Состав системы':[''], 'Методика':[''], 'Результат':[''], 'Примечания':[''], 'Финальное решение':['']}
            df_cc = pd.DataFrame(new_row)
            df_f = pd.concat([df_f, df_cc], axis=0, ignore_index = True)
            df_f = df_f.drop_duplicates(subset=['Компания','Имя образца'])
            df_f.dropna(subset=['Имя образца'])
            df_f.to_excel('Example.xlsx')
        
except:
    st.error("Ошибка ввода данных. Сделайте шаг назад или очистите кэш")

##       
