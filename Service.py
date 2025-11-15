import streamlit as st
import numpy as np
import math
import pandas as pd

import pickle
# from pydub import AudioSegment
import cloudpickle

import speech_recognition as sr

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


st.title("Сервис менеджера Оргбент")

df_c = pd.read_excel('company_1.xlsx')
df_c = df_c.drop('Unnamed: 0', axis = 1)
df_c = df_c.dropna(subset=['Компания'])
df_c = df_c.drop_duplicates()
for i in df_c.columns:
    df_c[i] = df_c[i].astype(str)


##row_c = {'Компания': [''], 'Тип компании': [''], 'Менеджер': [''], 'Контакты':[''], 'Сфера деятельности':[''], 'Оптимальный бентонит':[''], 'Производитель ОБ':[''], 'Конкурентная цена':[''], 'Ожидания от продукта':[''], 'Методики контроля':['']}
##df_c = pd.DataFrame(row_c)

ss = st.sidebar.checkbox('Внести коррективы/удалить')
a = {}
if ss:
    
    for i in list(df_c['Компания']):
        a[i] = i
    t_1 = st.selectbox('Выбрать компанию: ', list(a))
    ss_1 = st.sidebar.checkbox('Изменить название компании')
    dd = st.sidebar.checkbox('Заполнить/изменить поле')
    print(df_c['Компания'])
    if ss_1 == True and dd == False:
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
        new_row_1 = {'Тип компании': 'Тип компании', 'Менеджер': 'Менеджер', 'Контакты':'Контакты', 'Сфера деятельности':'Сфера деятельности', 'Оптимальный бентонит':'Оптимальный бентонит', 'Производитель ОБ':'Производитель ОБ', 'Конкурентная цена':'Конкурентная цена', 'Ожидания от продукта':'Ожидания от продукта', 'Методики контроля':'Методики контроля'}
        t_2 = st.selectbox('Выбрать поле: ', list(new_row_1))
        comp = list(df_c['Компания'])
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
        new_row = {'Компания': [text], 'Тип компании': [''], 'Менеджер': [''], 'Контакты':[''], 'Сфера деятельности':[''], 'Оптимальный бентонит':[''], 'Производитель ОБ':[''], 'Конкурентная цена':[''], 'Ожидания от продукта':[''], 'Методики контроля':['']}
        df_cc = pd.DataFrame(new_row)
        df_c = pd.concat([df_c, df_cc], axis=0, ignore_index = True)
        df_c = df_c.drop_duplicates()
        df_c.dropna(subset=['Компания'])
        df_c.to_excel('company_1.xlsx')
       


