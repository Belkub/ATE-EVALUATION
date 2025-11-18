import streamlit as st
import numpy as np
import math
import pandas as pd

import pickle
#from pydub import AudioSegment
import cloudpickle
from datetime import datetime 

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


    st.header("chat_mobile/company")

    df_c = pd.read_excel('company_1.xlsx')
    if 'Unnamed: 0' in df_c.columns:
        df_c = df_c.drop('Unnamed: 0', axis = 1)
    df_c = df_c.dropna(subset=['Компания'])
    df_c = df_c.drop_duplicates()
    for i in df_c.columns:
        df_c[i] = df_c[i].astype(str)


    ##row_c = {'Компания': [''], 'Тип компании': [''], 'Менеджер': [''], 'Контакты':[''], 'Сфера деятельности':[''], 'Оптимальный бентонит':[''], 'Производитель ОБ':[''], 'Конкурентная цена':[''], 'Ожидания от продукта':[''], 'Методики контроля':['']}
    ##df_c = pd.DataFrame(row_c)

    ss = st.sidebar.checkbox('Внести коррективы/удалить')
    a = {}
    pp = st.sidebar.checkbox('Сформировать сводную таблицу компаний')
    if pp == True and ss == False:
        st.title('Создание сводной таблицы компаний')
        bn = st.checkbox('Вывести таблицу и ограничить число полей')
        if bn:
            number = df_c.columns
            num = st.multiselect('Выбрать значимые поля', list(df_c.columns))
            if num:
                number = num
            bnn = st.checkbox('Записать сводную таблицу в файл .xlsx')
            if bnn:
                st.warning('Создайте папку Data на диске С для записи данных')
                if st.button('Показать свод и записать в файл C:\Data\company.xlsx'):
                    st.dataframe(df_c[number]) 
                    df_c[number].to_excel("C:\\Data\\company.xlsx")
            else:
                if st.button('Показать таблицу'): 
                    st.dataframe(df_c[number]) 
        else:
            number = list(df_c.columns)

        spp = float(st.number_input('Установить число полей группировки: ', min_value = 1, max_value = 3, value = 1, step = 1))
        #new_row_1 = {'Компания':['Компания'],'Имя образца':['Имя образца'],'Дата теста/отправки':['Дата теста/отправки'], 'Производитель/поставщик':['Производитель/поставщик'], 'Тестировщик':['Тестировщик'], 'Отрасль':['Отрасль'], 'Дисперсионная среда':['Дисперсионная среда'], 'Состав системы':['Состав системы'], 'Методика':['Методика'], 'Результат':['Результат'], 'Примечания':['Примечания'], 'Финальное решение':['Финальное решение']}
        new_row_1 = {}
        for i in number:
            new_row_1[i] = i
        p_1 = st.selectbox('Выбрать поле_1: ', list(new_row_1))
        p_11 = st.multiselect(f'Выбрать позиции поля {p_1}', list(df_c[p_1].unique()))
        if p_11:
            df_cc = df_c.loc[df_c[p_1].isin(p_11)]
            if spp == 1:
                ta = st.checkbox('Записать таблицу в файл .xlsx')
                if ta:
                    st.warning('Создайте папку Data на диске С для записи данных')
                    if st.button('Показать и записать таблицу в файл C:\Data\company.xlsx'):
                        st.dataframe(df_cc[number]) 
                        df_cc[number].to_excel("C:\\Data\\company.xlsx")
                else:
                    if st.button('Показать сводную таблицу'): 
                        st.dataframe(df_cc[number]) 
            else:
                del new_row_1[p_1]
                p_2 = st.selectbox('Добавить поле_2: ', list(new_row_1))
                p_22 = st.multiselect(f'Выбрать позиции поля {p_2}', list(df_cc[p_2].unique()))
                if p_22:
                    df_cc = df_cc.loc[df_cc[p_2].isin(p_22)]
                    if spp == 2:
                        ta = st.checkbox('Записать таблицу в файл .xlsx')
                        if ta:
                            st.warning('Создайте папку Data на диске С для записи данных')
                            if st.button('Показать и записать таблицу в файл C:\Data\company.xlsx'):
                                st.dataframe(df_cc[number]) 
                                df_cc[number].to_excel("C:\\Data\\company.xlsx")
                        else:
                            if st.button('Показать сводную таблицу'): 
                                st.dataframe(df_cc[number]) 
                    else:
                        del new_row_1[p_2]
                        p_3 = st.selectbox('Добавить поле_3: ', list(new_row_1))
                        p_33 = st.multiselect(f'Выбрать позиции поля {p_3}', list(df_cc[p_3].unique()))
                        if p_33:
                            df_cc = df_cc.loc[df_cc[p_3].isin(p_33)]
                            if spp == 3:
                                ta = st.checkbox('Записать таблицу в файл .xlsx')
                                if ta:
                                    st.warning('Создайте папку Data на диске С для записи данных')
                                    if st.button('Показать и записать таблицу в файл C:\Data\company.xlsx'):
                                        st.dataframe(df_cc[number]) 
                                        df_cc[number].to_excel("C:\\Data\\company.xlsx")
                                else:
                                    if st.button('Показать сводную таблицу'): 
                                        st.dataframe(df_cc[number]) 
    elif pp == False and ss == True:
        
        for i in list(df_c['Компания']):
            a[i] = i
        t_1 = st.selectbox('Выбрать компанию: ', list(a))
        ss_1 = st.sidebar.checkbox('Изменить название компании/удалить')
        dd = st.sidebar.checkbox('Заполнить/изменить поле')
        print(df_c['Компания'])
        if ss_1 == True and dd == False:
            st.title('Изменение названия/удаление компании')
            with open("file.txt", "w") as file:  
                file.write(t_1)
                file.close()
            with open("file.txt", "r") as file:  
                company_1 = file.read()
                file.close()
            rr_1 = st.checkbox('Глосовой ввод')
            if rr_1: 
                text = voice()
                if text:
                    with open("file.txt", "w") as file:  
                        file.write(text)
                        file.close()
                    company_1 = text
            company = st.text_input('Название компании: ', value = company_1)
            with open("file.txt", "w") as file:  
                file.write(company)
                file.close()
            with open("file.txt", "r") as file:  
                text = file.read()
                file.close()
            if st.button('Записать новое имя'):
                df_c = df_c.dropna(subset=['Компания'])
                df_c = df_c.drop_duplicates()
                df_c.loc[df_c['Компания'] == t_1, 'Компания'] = text
                df_c.to_excel('company_1.xlsx')
            ss_2 = st.checkbox('Удалить компанию')
            if ss_2:     
                if st.button('Удалить информацию о компании'):
                    df_c = df_c.loc[df_c['Компания'] != t_1]
                    df_c.to_excel('company_1.xlsx')
                
        if dd == True and ss_1 == False:
            st.title('Заполнение/изменение полей')
            now = datetime.now()
            new_row_1 = {'Тип компании': ['Тип компании'], 'Контакт/менеджер': ['Контакт/менеджер'], 'Сфера деятельности':['Сфера деятельности'], 'Продукты':['Продукты']}
            t_2 = st.selectbox('Выбрать поле: ', list(new_row_1))
            comp = list(df_c['Компания'])
            if t_2 == 'Тип компании':
                type = {'Производитель':'Производитель', 'Потребитель':'Потребитель', 'ТД':'ТД'}
                
                with open("file.txt", "w") as file:  
                    file.write(str(df_c.loc[df_c['Компания'] == t_1, t_2][comp.index(t_1)]))
                    file.close()
                
                rr_11 = st.checkbox('Глосовой ввод')
                if rr_11:
                    text_22 = voice()
                    if text_22:
                        with open("file.txt", "w") as file:  
                            file.write(text_22)
                            file.close()
                with open("file.txt", "r") as file:  
                    text_11 = file.read()
                    file.close()
                change = st.text_input('Текущее содержание поля: ', value = text_11)
                ddd = st.checkbox('Выбрать тип компании')
                if ddd:
                    t_3 = st.selectbox('Тип компании: ', list(type))
                    change = st.text_input('Новое содержание поля: ', value = t_3)
            else:
                with open("file.txt", "w") as file:  
                    file.write(str(df_c.loc[df_c['Компания'] == t_1, t_2][comp.index(t_1)]))
                    file.close()
                with open("file.txt", "r") as file:  
                    text_1 = file.read()
                    file.close()
                rr_1 = st.checkbox('Глосовой ввод')
                if rr_1:
                    text_2 = voice()
                    if text_2:
                        with open("file.txt", "w") as file:  
                            file.write(text_2)
                            file.close()
                with open("file.txt", "r") as file:  
                    text_3 = file.read()
                    file.close()
            
                change = st.text_input('Содержание поля: ', value = text_3)
            with open("file.txt", "w") as file:  
                file.write(change)
                file.close()
            with open("file.txt", "r") as file:  
                text_3 = file.read()
                file.close()
            if st.button('Записать поле'):
                df_c = df_c.dropna(subset=['Компания'])
                df_c = df_c.drop_duplicates()
                df_c.loc[df_c['Компания'] == t_1, t_2] = text_3
                df_c.to_excel('company_1.xlsx')
                print(df_c)
            


    else:
        st.title('Создание компании')
        for i in list(df_c['Компания']):
            a[i] = i
        t_1 = st.selectbox('СОЗДАТЬ КОМПАНИЮ: ', list(a))
        rr = st.checkbox('Глосовой ввод')
        
        if rr:
            text = voice()
            if text:
                with open("file.txt", "w") as file:  
                    file.write(text)
                    file.close()
        with open("file.txt", "w") as file:  
            file.write(t_1)
            file.close()
        with open("file.txt", "r") as file:  
            text = file.read()
            file.close()
        company = st.text_input('Название компании: ', value = text)
        with open("file.txt", "w") as file:  
            file.write(company)
            file.close()
        with open("file.txt", "r") as file:  
            text = file.read()
            file.close()
        ##wdf_c = st.checkbox('Записать название')
        if st.button('Записать в df'):
            now = datetime.now()
            new_row = {'Время входа':now, 'Компания': [text], 'Тип компании': [''], 'Контакт/менеджер': [''], 'Сфера деятельности':[''], 'Продукты':['']}
            df_m = pd.DataFrame(new_row)
            df_c = pd.concat([df_c, df_m], axis=0, ignore_index = True)
            df_c = df_c.drop_duplicates(subset=['Компания'])
            df_c.dropna(subset=['Компания'])
            df_c.to_excel('company_1.xlsx')
except:
    st.error("Ошибка ввода данных. Сделайте шаг назад или очистите кэш")
       
