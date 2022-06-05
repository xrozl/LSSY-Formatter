import os
import pyperclip
import yaml
import time
import PySimpleGUI as sg
import pandas as pd
import openpyxl as xl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager 

def main():
    # load config yaml
    config = None
    with open("config.yaml", "r") as f:
        config = yaml.safe_load(f)
    if config is None:
        print("config.yaml is not found")
        return
    url = config["url"]
    if url is None:
        print("url is not found")
        return

    option: Options = Options()
    option.add_experimental_option("prefs", {
        "download.default_directory": os.getcwd() + "/cache",
        "plugins.always_open_pdf_externally": True
    })
    # headless
    option.add_argument("--headless")
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=option)

    # delete cache
    if os.path.exists(os.getcwd() + "/cache"):
        os.system("rm -rf " + os.getcwd() + "/cache")
    os.system("mkdir " + os.getcwd() + "/cache")
    
    # open url
    driver.get(url)

    # wait for page load
    driver.implicitly_wait(10)
    print("ぺーじろーど")
    
    driver.find_element(By.XPATH, '/html/body/div[17]/div[2]/div[2]/div/label').click()
    time.sleep(2)
    print("確認画面クリック")
    driver.find_element(By.XPATH, '/html/body/div[16]/div[2]/div/div[4]/div[2]/div[2]/input').send_keys(config["username"])
    time.sleep(2)
    print("ユーザー名入力")
    driver.find_element(By.XPATH, '/html/body/div[16]/div[2]/div/div[4]/div[3]/span[1]').click()
    time.sleep(2)
    print("チェックボックスクリック")
    driver.find_element(By.XPATH, '/html/body/div[16]/div[2]/div/div[1]/div[2]/label').click()
    time.sleep(5)
    print("ログインボタンクリック")
    
    # reload
    driver.refresh()
    driver.implicitly_wait(10)
    print("ブラウザリロード")

    driver.find_element(By.XPATH, '/html/body/div[10]/div/div[1]/div/div[2]/div[5]/div/div[2]/div[2]').click()
    time.sleep(2)
    print("ダウンロードボタンクリック")
    driver.find_element(By.XPATH, '/html/body/div[17]/div[3]/div[2]/div/div[1]/a/span').click()
    time.sleep(2)
    print("確認ボタンクリック")
    for element in driver.find_elements_by_tag_name('span'):
        if element.text == 'Excel':
            try :
                element.click()
                break
            except Exception as e: print("Error起きたかも")
    print("xlsxダウンロード")

    print("圧縮してダウンロード中...")
    while True:
        t = False
        for file in os.listdir(os.getcwd() + "/cache"):
            if file.endswith(".xlsx"):
                t = True
                break
        if t:
            print("ダウンロード終わった")
            break
        time.sleep(1)
    # close browser
    driver.quit()
    print("ブラウザクローズ")

    os.system("mv " + os.getcwd() + "/cache/*.xlsx " + os.getcwd() + "/cache/compress.xlsx")

    # find xlsx file
    xlsx = None
    for file in os.listdir(os.getcwd() + "/cache"):
        if file.endswith(".xlsx"):
            xlsx = file
            break
    
    # read xlsx
    if xlsx is None:
        print("xlsx is not found")
        return
    wb = xl.load_workbook(os.getcwd() + "/cache/" + xlsx)

    # get sheets name
    sheets = wb.sheetnames
    if len(sheets) == 0:
        print("sheets is not found")
        return
    
    # check sheets if name is target
    target = config['sheet']
    if target is None:
        print("target is not found")
        return
    ws = None
    for sheet in sheets:
        if sheet[:-5] == target:
            ws = wb[sheet]
            break
    if ws is None:
        print("sheet is not found")
        return

    # get data
    data = []
    for row in ws.rows:
        data.append([cell.value for cell in row])
    if len(data) == 0:
        print("data is not found")
        return

    # print data
    print(data)



    """
    for name in files:
        if name.find("/") != -1:
            prefix = name.split("/")[0]
            suffix = name.split("/")[1]
            if suffix == config['csvname']:
                with open(os.getcwd() + "/cache/" + name, "r") as f:
                    csv = f.read()
                break
        else:
            continue
    
    if csv is None:
        print("csv file is not found")
        return
    
    print("データを整形します")

    # delete header
    csv = csv.split("\n")
    csv = csv[1:]
    csv = "\n".join(csv)

    # dict
    data = {}
    line1 = csv.split("\n")[0]
    offset = 2
    line1 = line1.split(",")

    cacheLine = csv.split("\n")[1+offset]
    line2 = []
    skipCount = 0
    next = True
    for i in range(len(line1)):
        if next:
            skipCount += 1
            next = False
            continue
        text = cacheLine.split(",")[i]
        if text.startswith('"¥'):
            i += 1
            text = text + cacheLine.split(",")[i]
            next = True
        if text.endswith('"'):
            text = text[:-1]
        if text.startswith('"'):
            text = text[1:]
        line2.append(text)
    selector = []

    for i in range(len(line2) - skipCount):
        data[line1[i]] = line2[i+skipCount]
    print("データ整形完了")
    # print data
    for key in data:
        print(key + " : " + data[key])
    """


    """
    sg.theme('Dark')
    layout = []
    window = sg.Window('LSSY-Formatter', layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
    """


if __name__ == "__main__":
    main()
