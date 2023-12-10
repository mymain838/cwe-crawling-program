import tkinter as tk
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup
import pandas as pd
from deep_translator import GoogleTranslator

def extract_data(number, translator, is_translated):
    # 각 CWE 번호에 대한 데이터 추출
    url = f"https://cwe.mitre.org/data/definitions/{number}.html"
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # 데이터 초기화
    title = description = extended_description = bad_description = good_description = ""

    # title 데이터 추출
    anchor_tag = soup.find('a', {'name': str(number)})
    if anchor_tag:
        title = anchor_tag.find_next('h2').get_text(strip=True)
    else:
        title = "Blank(공란)"

    # description 데이터 추출
    description_tag = soup.find('div', {'id': f'oc_{number}_Description'})
    if description_tag:
        description = description_tag.find('div', {'class': 'indent'}).get_text(strip=True)
    else:
        description = "Blank(공란)"

    # Extends_description 데이터 추출
    extended_description_tag = soup.find('div', {'id': f'oc_{number}_Extended_Description'})
    if extended_description_tag:
        extended_description = extended_description_tag.find('div', {'class': 'indent'}).get_text(strip=True)
    else:
        extended_description = "Blank(공란)"

    # Bad Description 데이터 추출
    bad_description_tag = soup.find('div', {'class': 'indent Bad'})
    if bad_description_tag:
        bad_description = bad_description_tag.find_previous('p').get_text(strip=True)
    else:
        bad_description = "Blank(공란)"

    # Good Description 데이터 추출
    good_description_tag = soup.find('div', {'class': 'indent Good'})
    if good_description_tag:
        good_description = good_description_tag.find_previous('p').get_text(strip=True)
    else:
        good_description = "Blank(공란)"

    # 번역 여부에 따라 반환
    if is_translated == "true":
        translated_title = translator.translate(title)
        translated_description = translator.translate(description)
        translated_extended_description = translator.translate(extended_description)
        translated_bad_description = translator.translate(bad_description)
        translated_good_description = translator.translate(good_description)
        
        return [
            number, translated_title, translated_description,
            translated_extended_description, translated_bad_description,
            translated_good_description
        ]
    else:
        return [
            number, title, description, extended_description,
            bad_description, good_description
        ]

def create():
    global path_entry, cnum_entry
    path = path_entry.get()
    cnum_range = cnum_entry.get()

    # cnum_entry가 비어있는 경우
    if not cnum_range:
        messagebox.showwarning("경고", "CWE 번호를 입력하세요!")
        return
    
    # path_entry가 비어있는 경우
    if not path:
        messagebox.showwarning("경고", "저장하실 경로를 입력하세요!")
        return

    data_en = []
    data_ko = []
    translator = GoogleTranslator(source='auto', target='ko')

    # 콜론(:)을 기준으로 시작과 끝을 분리하여 리스트로 변환
    if ':' in cnum_range:
        start, end = map(int, cnum_range.split(':'))
        numbers = list(range(start, end + 1))

    # 쉼표(,)를 기준으로 분리하여 리스트로 변환    
    elif ',' in cnum_range:
        numbers = list(map(int, cnum_range.split(',')))

    # 숫자 하나만 입력되는 경우 처리
    else:
        numbers = [int(cnum_range)]

    for number in numbers:
        data_en.append(extract_data(number, translator, "false"))
        data_ko.append(extract_data(number, translator, "true"))

    # 데이터 프레임 생성
    df1 = pd.DataFrame(data_en, columns=["No", "Title", "Desc", "ExtendsDes", "BadDesc", "GoodDesc"])
    df2 = pd.DataFrame(data_ko, columns=["번호", "제목", "설명", "추가설명", "나쁜예제설명", "좋은예제설명"])

    # 언어별 데이터를 엑셀 파일로 저장
    df1.to_excel(f"{path}/result_en.xlsx", index=False)  
    df2.to_excel(f"{path}/result_ko.xlsx", index=False)  


root = tk.Tk()
root.title("CWE Example Creator")


root.geometry("600x200")

cnum_label = tk.Label(root, text="CWE 번호를 기입해주세요 ex )1 또는 1:20 또는 1,2,3")
cnum_label.pack()

cnum_entry = tk.Entry(root)
cnum_entry.pack()

path_label = tk.Label(root, text="폴더 생성 경로:")
path_label.pack()
path_entry = tk.Entry(root)
path_entry.pack()

create_button = tk.Button(root, text="생성", command=create)
create_button.pack()

root.mainloop()