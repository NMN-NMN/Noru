import os
import re
import sys
import math
import time
import shutil
import subprocess
import requests
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.font as tkfont
import tkinter.messagebox as tkbox
import chromedriver_autoinstaller
from idlelib.tooltip import Hovertip
from threading import Thread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import pyodbc

class Crawling():
    def get(self, driver, mode, isThree):
        # KCC: ColorNm(3B), ColorNm(3P), ColorCode, MixDate, StdDesc, CarNm, ApplyYear, MixInfo
        # Noru: ColorNm(3B), ColorNm(3P), ColorCode, MixDate, StdDesc, CarNm, ApplyYear, MixInfo, MixCd, ColorMixStd
        data = {}
        detail_data_3B = []
        detail_data_3P = []

        if mode == "kcc":
            data["ColorNm(3B)"] = driver.find_element(By.XPATH, "//*[@id=\"contentArea\"]/div[2]/div[1]/div[2]/h3").text
            data["ColorCode"] = driver.find_element(By.XPATH, "//*[@id=\"contentArea\"]/div[2]/div[2]/div[2]/table/tbody/tr[1]/td").text
            data["MixDate"] = driver.find_element(By.XPATH, "//*[@id=\"contentArea\"]/div[2]/div[2]/div[1]/table/tbody/tr[2]/td").text
            data["StdDesc"] = "다" + driver.find_element(By.XPATH, "//*[@id=\"contentArea\"]/div[2]/div[2]/div[2]/table/tbody/tr[2]/td").text
            data["CarNm"] = driver.find_element(By.XPATH, "//*[@id=\"contentArea\"]/div[2]/div[3]/div[1]/table/tbody/tr[2]/td").text
            data["ApplyYear"] = driver.find_element(By.XPATH, "//*[@id=\"contentArea\"]/div[2]/div[3]/div[1]/table/tbody/tr[3]/td").text
            data["MixInfo"] = driver.find_element(By.XPATH, "//*[@id=\"contentArea\"]/div[2]/div[4]/div/div").text

            data["MixDate"] = data["MixDate"].replace(".", "")
            data["ApplyYear"] = data["ApplyYear"][:data["ApplyYear"].find(".")]

            index = 3
            while True:
                try:
                    detail_data_3B.append((
                        driver.find_element(By.XPATH, "//*[@id=\"mixColorArea\"]/div[1]/table/tbody/tr[" + str(index) + "]/td[1]").text,
                        driver.find_element(By.XPATH, "//*[@id=\"mixColorArea\"]/div[1]/table/tbody/tr[" + str(index) + "]/td[3]").text
                    ))
                    index += 1
                except:
                    break
            
            if isThree:
                data["ColorNm(3P)"] = data["ColorNm(3B)"]
                data["ColorNm(3B)"] = data["ColorNm(3B)"] + "-바탕"

                index = 2
                while True:
                    try:
                        detail_data_3P.append((
                            driver.find_element(By.XPATH, "//*[@id=\"threeCoatMixColorArea\"]/div[1]/table/tbody/tr[" + str(index) + "]/td[1]").text,
                            driver.find_element(By.XPATH, "//*[@id=\"threeCoatMixColorArea\"]/div[1]/table/tbody/tr[" + str(index) + "]/td[3]").text
                        ))
                        index += 1
                    except:
                        break
        else:
            driver.switch_to.window(driver.window_handles[0])
            data["ColorNm(3B)"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[1]/td[1]").text
            data["ColorCode"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[1]/td[2]").text
            data["MixDate"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[4]/td[1]").text
            data["StdDesc"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[6]/td[1]").text
            data["CarNm"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[2]/td[1]").text
            data["ApplyYear"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[3]/td[2]").text
            data["MixInfo"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[9]/td").text
            data["MixCd(3B)"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[2]/td[2]").text
            data["ColorMixStd"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[4]/td[2]").text

            data["MixInfo"] = data["CarNm"] + (", " if data["MixInfo"].replace(" ", "") != "" else "") + data["MixInfo"]

            driver.switch_to.window(driver.window_handles[1])
            
            for index in range(1, 16):
                TonerCd = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[3]/div[2]/table/tbody/tr[" + str(index) + "]/td[2]/input").get_attribute("value")
                if TonerCd == "":
                    continue

                detail_data_3B.append((
                        TonerCd.replace(" ", ""),
                        driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[3]/div[2]/table/tbody/tr[" + str(index) + "]/td[3]/input[3]").get_attribute("value")
                ))

            if isThree:
                driver.switch_to.window(driver.window_handles[2])
                data["ColorNm(3P)"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[1]/td[1]").text
                data["MixCd(3P)"] = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[6]/table/tbody/tr[2]/td[2]").text

                driver.switch_to.window(driver.window_handles[3])
                for index in range(1, 16):
                    TonerCd = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[3]/div[2]/table/tbody/tr[" + str(index) + "]/td[2]/input").get_attribute("value")
                    if TonerCd == "":
                        continue

                    detail_data_3P.append((
                            TonerCd.replace(" ", ""),
                            driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[3]/div[2]/table/tbody/tr[" + str(index) + "]/td[3]/input[3]").get_attribute("value")
                    ))
        
        return data, detail_data_3B, detail_data_3P, isThree
    
    def open(self, _ID, _PWD, code, mode, isThree, paintTy, isGet = False):
        options = webdriver.ChromeOptions()
        if isGet:
            options.add_argument("--headless")
        else:
            options.add_experimental_option(name="detach", value=True)
            options.add_experimental_option('excludeSwitches', ['enable-automation'])

        version = chromedriver_autoinstaller.get_chrome_version().split(".")[0]
        current_path = os.getcwd()
        driver_path = f"{current_path}/{version}/chromedriver.exe"

        driver = webdriver.Chrome(service=Service(driver_path), options=options)
        driver.implicitly_wait(2)

        if mode == "kcc":
            driver.get("https://refinish.co.kr/user/account/login_login")
            time.sleep(0.5)
            driver.find_element(By.ID, "id").send_keys(_ID)
            time.sleep(0.5)
            driver.find_element(By.ID, "pwd").send_keys(_PWD)
            time.sleep(0.5)
            driver.find_element(By.ID, "btnLogin").click()
            time.sleep(0.5)

            if driver.find_element(By.XPATH, "/html/body/div[5]/div").is_displayed():
                return False

            driver.get("https://refinish.co.kr/user/color/colorSearch_view/" + code)
            time.sleep(0.5)
            if isThree:
                driver.find_element(By.XPATH, "//*[@id=\"mixColorArea\"]/div[2]/div/div[1]/div[1]/button").click()
                time.sleep(0.1)
                driver.find_element(By.XPATH, "//*[@id=\"mixColorArea\"]/div[2]/div/div[1]/div[1]/div/ul/li[2]").click()
                time.sleep(0.1)
                driver.find_element(By.XPATH, "//*[@id=\"threeCoatMixColorArea\"]/div[2]/div/div[1]/div[1]/button").click()
                time.sleep(0.1)
                driver.find_element(By.XPATH, "//*[@id=\"threeCoatMixColorArea\"]/div[2]/div/div[1]/div[1]/div/ul/li[2]").click()
            else:
                driver.find_element(By.XPATH, "//*[@id=\"mixColorArea\"]/div[2]/div/div[1]/div[1]/button").click()
                time.sleep(0.1)
                driver.find_element(By.XPATH, "//*[@id=\"mixColorArea\"]/div[2]/div/div[1]/div[1]/div/ul/li[2]").click()
        else:
            driver.get("https://m.autorefinishes.co.kr/member/memberLogin.asp")
            time.sleep(0.5)
            driver.find_element(By.ID, "uid").send_keys(_ID)
            time.sleep(0.5)
            driver.find_element(By.ID, "upw").send_keys(_PWD)
            time.sleep(0.5)
            driver.find_element(By.ID, "loginbtn").click()

            try:
                driver.switch_to.alert
                return False
            except:
                pass

            driver.execute_script('window.open("about:blank", "_blank");')
            driver.switch_to.window(driver.window_handles[1])

            driver.get("https://www.autorefinishes.co.kr/member/login.asp")
            time.sleep(0.5)
            driver.find_element(By.ID, "uid").send_keys(_ID)
            time.sleep(0.5)
            driver.find_element(By.ID, "upw").send_keys(_PWD)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click()", driver.find_element(By.XPATH, "//*[@id=\"loginbtn\"]"))
            time.sleep(0.5)
            driver.switch_to.alert.accept()

            tabs = driver.window_handles
            while len(tabs) < 3:
                time.sleep(0.5)
                tabs = driver.window_handles
            
            driver.switch_to.window(tabs[2])
            driver.close()
            driver.switch_to.window(tabs[0])

            if len(code) > 1:
                driver.execute_script('window.open("about:blank", "_blank");')
                driver.execute_script('window.open("about:blank", "_blank");')

            tabs = driver.window_handles
            cott3 = ""

            for i in range(len(code)):
                driver.switch_to.window(tabs[0 if i == 0 else 2])
                driver.get("https://m.autorefinishes.co.kr/colorinformation/colormix_view.asp?MixCd=" + code[i] + "&PaintTy=" + paintTy)
                
                driver.switch_to.window(tabs[1 if i == 0 else 3])
                driver.get("https://www.autorefinishes.co.kr/colorinformation/colormix_user.asp?mixcd=" + code[i] + "&ptype=" + paintTy.lower() + "&Op=A")

                if not isThree and isGet:
                    try:
                        driver.switch_to.window(tabs[0])
                        cott3 = driver.find_element(By.XPATH, "//*[@id=\"mixlastForm\"]/div/div[4]/div[2]/div/p").text
                        isok = tkbox.askquestion("알람", "3코트 데이터입니다.\n이동하시겠습니까?")
                        if isok == "yes":
                            driver.execute_script('window.open("about:blank", "_blank");')
                            driver.execute_script('window.open("about:blank", "_blank");')
                            tabs = driver.window_handles

                            if cott3[16] == "B":
                                pearl = code[0][:]
                                code[0] = cott3[22:-1]
                                code.append(pearl)
                            else:
                                code.append(cott3.split("(")[1][:-1])

                            for p in range(len(code)):
                                driver.switch_to.window(tabs[0 if p == 0 else 2])
                                driver.get("https://m.autorefinishes.co.kr/colorinformation/colormix_view.asp?MixCd=" + code[p] + "&PaintTy=" + paintTy)
                                
                                driver.switch_to.window(tabs[1 if p == 0 else 3])
                                driver.get("https://www.autorefinishes.co.kr/colorinformation/colormix_user.asp?mixcd=" + code[p] + "&ptype=" + paintTy.lower() + "&Op=A")
                            isThree = True
                    except:
                        break

        if isGet:
            return self.get(driver, mode, isThree)

class SQL():
    cursors = ["", ""]
    tables = ["AUTO_CODE_TBL", "CARMAKER_TBL", "COLORCDGR_TBL", "COLORMIX_TBL", "COLORMIXDETAIL_TBL", "COLORMIXSTD_TBL", "MIXCOLORINFO_TBL"]
    types = {"VARCHAR":str, "LONGCHAR":str, "INTEGER":int, "DECIMAL":float}
    datas = {"code":[], "data":[]}
    datas_order = {"code":[], "data":[]}
    edit_function = []
    bind_function = None
    useble_widgets = []

    def Open(self, file, pwd=""):
        self.cursors[0 if file == "r_V2.accdb" else 1] = pyodbc.connect(
                r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
                r"DBQ=" + os.getcwd() + "/" + file + ";"
                r"PWD=" + pwd).cursor()

    def Close(self, cursor_index):
        self.cursors[cursor_index].close()
    
    def Save(self, data):
        for cursor_index in range(2):
            self.Open("r_V2.accdb" if cursor_index == 0 else "new_auto.mdb", "" if cursor_index == 0 else "EOGKSALSRNRSHFNVPDLS")
            cursor = self.cursors[cursor_index]
            input_index = 0

            if cursor_index == 0:
                cursor.execute("SELECT MAX([INDEX]) FROM RECORD")
                temp = cursor.fetchone()[0]
                if temp != None:
                    input_index = int(temp) + 1
                else:
                    input_index = 0

            if len(data) <= 3:
                if cursor_index == 0:
                    data[0].append(input_index)
                    data[1].append(input_index)
                    data[2].append(input_index)

                for index, table in enumerate(self.tables[:3]):
                    columns = [columns[3] for columns in cursor.columns(table=table)]
                    column_types = [columns[5] for columns in cursor.columns(table=table)]

                    input_columns = ""
                    for column in columns:
                        input_columns += f"[{column}], "
                    input_columns = f"({input_columns[:-2]})"

                    input_values = ""
                    for _type, input in zip(column_types, data[index]):
                        if self.types[_type] == int or self.types[_type] == float:
                            input_values += f"{str(input)}, "
                        else:
                            input_values += f"\'{str(input)}\', "
                    input_values = f"({input_values[:-2]})"

                    cursor.execute(f"INSERT INTO {table} {input_columns} VALUES {input_values}")
                    cursor.commit()
                
                if cursor_index == 0:
                    cursor.execute(f"INSERT INTO RECORD ([TYPE], [INDEX], [_DATE], [TRASH]) VALUES (\'CODE\', {input_index}, NOW(), FALSE)")
                    cursor.commit()
            else:
                if cursor_index == 0:
                    data[0][0].append(input_index)
                    for i in range(len(data[1][0])):
                        data[1][0][i].append(input_index)
                    data[2][0].append(input_index)
                    data[3][0].append(input_index)

                    try:
                        data[0][1].append(input_index)
                        for i in range(len(data[1][1])):
                            data[1][1][i].append(input_index)
                        data[2][1].append(input_index)
                        data[3][1].append(input_index)
                    except:
                        pass

                for index, table in enumerate(self.tables[3:]):
                    columns = [columns[3] for columns in cursor.columns(table=table)]
                    column_types = [columns[5] for columns in cursor.columns(table=table)]

                    if index == 1:
                        temp_columns = [columns[2], columns[4], columns[1], columns[3]]
                        temp_column_types = [column_types[2], column_types[4], column_types[1], column_types[3]]
                        del columns[1:5]
                        del column_types[1:5]
                        for column_index, column, column_type in zip(range(1, 5), temp_columns, temp_column_types):
                            columns.insert(column_index, column)
                            column_types.insert(column_index, column_type)

                    input_columns = ""
                    for column in columns:
                        if column != "order":
                            input_columns += f"[{column}], "
                    input_columns = f"({input_columns[:-2]})"

                    if index == 1:
                        for cott in data[index]:
                            for detail in cott:
                                input_values = ""

                                for _type, input in zip(column_types, detail):
                                    try:
                                        if self.types[_type] == int or self.types[_type] == float:
                                            input_values += f"{str(input)}, "
                                        else:
                                            input_values += f"\'{str(input)}\', "
                                    except:
                                        input_values += f"\'{str(input)}\', "

                                input_values = f"({input_values[:-2]})"

                                cursor.execute(f"INSERT INTO {table} {input_columns} VALUES {input_values}")
                                cursor.commit()
                    else:
                        for cott in data[index]:
                            input_values = ""

                            for _type, input in zip(column_types, cott):
                                if self.types[_type] == int or self.types[_type] == float:
                                    input_values += f"{str(input)}, "
                                else:
                                    input_values += f"\'{str(input)}\', "

                            input_values = f"({input_values[:-2]})"

                            cursor.execute(f"INSERT INTO {table} {input_columns} VALUES {input_values}")
                            cursor.commit()
                
                if cursor_index == 0:
                    pass
                    cursor.execute(f"INSERT INTO RECORD ([TYPE], [INDEX], [_DATE], [TRASH]) VALUES (\'DATA\', {input_index}, NOW(), FALSE)")
                    cursor.commit()
                        
            self.Close(cursor_index)
    
    def Update(self, _type, data, main_index, window, isdelete=False, istrash=False):
        if not isdelete:
            isok = tkbox.askquestion("경고", "데이터를 덮어씌우시겠습니까?\n이 행위는 되돌릴 수 없습니다.")
            if isok == "no":
                return False

        original = self.datas[_type][main_index]

        def get_update_columns(columns, column_types, data):
            update_column = ""
            for column, column_type, value in zip(columns, column_types, data):
                if column_type != "DATETIME" and (self.types[column_type] == int or self.types[column_type] == float):
                    update_column += f"[{column}] = {value}, "
                else:
                    update_column += f"[{column}] = \'{value}\', "
            update_column = f"{update_column[:-2]}"

            return update_column
        
        def get_condition(columns, column_types, original):
            if columns == None:
                if column_types == "code":
                    return f"[INDEX] = {original}"
                else:
                    if len(original) < 3:
                        return f"[MixCd] = \'{original[0]}\' AND [INDEX] = {original[1]}"
                    else:
                        return f"[MixCd] = \'{original[0]}\' AND [Nbr] = \'{original[1]}\' AND [INDEX] = {original[2]}"
            else:
                condition = ""
                for column, column_type, value in zip(columns, column_types, original):
                    if column_type != "DATETIME" and (self.types[column_type] == int or self.types[column_type] == float):
                        condition += f"[{column}] = {value} AND "
                    else:
                        condition += f"[{column}] = \'{value}\' AND "
                condition = f"{condition[:-5]}"

                return condition
        
        for cursor_index in ((1, 0) if original[3 if _type == "code" else 4] == 0 else (1,)):
            self.Open("r_V2.accdb" if cursor_index == 0 else "new_auto.mdb", "" if cursor_index == 0 else "EOGKSALSRNRSHFNVPDLS")
            cursor = self.cursors[cursor_index]

            for table_index, table in enumerate(self.tables[0 if _type == "code" else 3:3 if _type == "code" else 7]):
                columns = [columns[3] for columns in cursor.columns(table=table)]
                column_types = [columns[5] for columns in cursor.columns(table=table)]

                if len(original[table_index]) > 0:
                    if _type == "code":
                        if isdelete:
                            if cursor_index == 1 or (cursor_index == 0 and istrash):
                                command = f"DELETE FROM {table} WHERE ({get_condition(None if cursor_index == 0 else columns, _type if cursor_index == 0 else column_types, original[table_index][0][-1] if cursor_index == 0 else original[table_index][0])})"
                            else:
                                if cursor_index == 0 and not istrash:
                                    command = f"UPDATE RECORD SET [TRASH] = TRUE WHERE [INDEX] = {original[table_index][0][-1]}"
                                else:
                                    continue
                        else:
                            command = f"UPDATE {table} SET {get_update_columns(columns, column_types, data[table_index])} WHERE ({get_condition(None if cursor_index == 0 else columns, _type if cursor_index == 0 else column_types, original[table_index][0][-1] if cursor_index == 0 else original[table_index][0])})"
                        cursor.execute(command)
                        cursor.commit()
                    else:
                        _original = None
                        for cott in range(2 if len(original) > 7 else 1):
                            if cott == 1:
                                _original = original[7]
                            else:
                                _original = original

                            if table_index != 1:
                                if isdelete:
                                    if cursor_index == 1 or (cursor_index == 0 and istrash):
                                        command = f"DELETE FROM {table} WHERE ({get_condition(None if cursor_index == 0 else columns, _type if cursor_index == 0 else column_types, (_original[table_index][0][0], _original[table_index][0][-1]) if cursor_index == 0 else _original[table_index][0])})"
                                    else:
                                        if cursor_index == 0 and not istrash:
                                            command = f"UPDATE RECORD SET [TRASH] = TRUE WHERE [INDEX] = {_original[table_index][0][-1]}"
                                        else:
                                            continue
                                else:
                                    command = f"UPDATE {table} SET {get_update_columns(columns, column_types, data[table_index][cott])} WHERE ({get_condition(None if cursor_index == 0 else columns, _type if cursor_index == 0 else column_types, (_original[table_index][0][0], _original[table_index][0][-1]) if cursor_index == 0 else _original[table_index][0])})"
                                cursor.execute(command)
                                cursor.commit()
                            else:
                                copy_columns = columns[:]
                                copy_column_types = column_types[:]

                                temp_columns = copy_columns[1:5]
                                temp_column_types = copy_column_types[1:5]
                                
                                del copy_columns[1:5]
                                del copy_column_types[1:5]

                                for idx, replace_idx in enumerate([1, 3, 0, 2], 1):
                                    copy_columns.insert(idx, temp_columns[replace_idx])
                                    copy_column_types.insert(idx, temp_column_types[replace_idx])
                                
                                del columns[-1]
                                del column_types[-1]

                                if isdelete:
                                    for idx in range(len(_original[table_index])):
                                        if cursor_index == 1 or (cursor_index == 0 and istrash):
                                            command = f"DELETE FROM {table} WHERE ({get_condition(None if cursor_index == 0 else columns, _type if cursor_index == 0 else column_types, (_original[table_index][idx][0], _original[table_index][idx][1], _original[table_index][idx][-1]) if cursor_index == 0 else _original[table_index][idx])})"
                                            cursor.execute(command)
                                            cursor.commit()
                                        elif cursor_index == 0 and not istrash:
                                            command = f"UPDATE RECORD SET [TRASH] = TRUE WHERE [INDEX] = {_original[table_index][idx][-1]}"
                                            cursor.execute(command)
                                            cursor.commit()
                                            break
                                        else:
                                            break
                                else:
                                    for idx, p in enumerate(data[table_index][cott]):
                                        command = f"UPDATE {table} SET {get_update_columns(copy_columns, copy_column_types, p)} WHERE ({get_condition(None if cursor_index == 0 else columns, _type if cursor_index == 0 else column_types, (_original[table_index][idx][0], _original[table_index][idx][1], _original[table_index][idx][-1]) if cursor_index == 0 else _original[table_index][idx])})"
                                        cursor.execute(command)
                                        cursor.commit()
            
            self.Close(cursor_index)

        if _type == "code":
            if not isdelete:
                insert_index = original[table_index][0][-1]

                for idx in range(3):
                    data[idx].append(insert_index)
                    
                self.datas["code"][main_index][idx][0] = data[idx]
        else:
            if not isdelete:
                insert_index = original[table_index][0][-1]

                for idx in range(4):
                    for cott in range(2 if len(original) > 7 else 1):
                        if idx != 1:
                            data[idx][cott].append(insert_index)

                            if cott == 0:
                                self.datas["data"][main_index][idx][0] = data[idx][cott]
                            else:
                                self.datas["data"][main_index][7][idx][0] = data[idx][cott]
                        else:
                            for row_idx, row in enumerate(data[idx][cott]):
                                row.append(insert_index)

                                temp_data = row[:]
                                replace_data = temp_data[1:5]
                                del temp_data[1:5]

                                for data_idx, replace_idx in enumerate([2, 0, 3, 1], 1):
                                    temp_data.insert(data_idx, replace_data[replace_idx])
                                temp_data[-1] = temp_data[-1].replace(" 오전 9", " 09")

                                if cott == 0:
                                    self.datas["data"][main_index][idx][row_idx] = temp_data
                                else:
                                    self.datas["data"][main_index][7][idx][row_idx] = temp_data
        
        if isdelete:
            canvas_number = 0 if _type == "code" else 1
            self.window.nametowidget(f".record_frame.!frame.preview.outer{canvas_number}.canvas").delete(f"no_{str(main_index)}")
        else:
            canvas_number = 0 if _type == "code" else 1
            frame = self.window.nametowidget(f".record_frame.!frame.preview.outer{canvas_number}.canvas.no_{str(main_index)}")
            frame_color = frame["bg"]
            frame["bg"] = "#FF7F00"

            if canvas_number == 0:
                frame.children["code"]["text"] = self.datas["code"][main_index][2][0][1]
                frame.children["minorcd"]["text"] = self.datas["code"][main_index][0][0][1]
            else:
                frame.children["code"]["text"] = self.datas["data"][main_index][0][0][0].strip()
                frame.children["type"]["text"] = self.datas["data"][main_index][0][0][1].strip()
                if len(self.datas["data"][main_index][3]) > 0:
                    frame.children["color"]["text"] = self.datas["data"][main_index][3][0][4].strip()

            frame.after(1000, lambda: frame.config(bg=frame_color))

        if window != None:
            window.destroy()

    def Delete(self):
        for frame, _type in zip([self.window.nametowidget(".record_frame.!frame.preview.outer0.canvas"), self.window.nametowidget(".record_frame.!frame.preview.outer1.canvas")], ["code", "data"]):
            selected = []
            
            for row in frame.winfo_children():
                if "selected" in row.children["checkbutton"].state():
                    selected.append(row)

            if len(selected) > 0:
                isok = tkbox.askquestion(f"경고 ({_type})", "new_auto.mdb의 경우, 휴지통이 아닌 즉시 삭제됩니다.\n이 과정은 되돌릴 수 없습니다. 삭제하시겠습니까?")
                if isok == "no":
                    return False

                delete_data = []

                for row in selected:
                    row.children["checkbutton"].state(["!selected"])
                    main_index = int(row.winfo_name()[3:])

                    self.Update(_type, None, main_index, None, True, False if self.datas[_type][main_index][3 if _type == "code" else 4] == 1 else (self.datas[_type][main_index][5 if _type == "code" else 6]))
                    delete_data.append(self.datas[_type][main_index])

                    row.destroy()

                for row in delete_data:
                    del self.datas[_type][self.datas[_type].index(row)]

            for row in frame.winfo_children():
                row.children["checkbutton"].state(["!selected"])
        
        self.Apply_preview_result()

    def Trash(self):
        for frame in [self.window.nametowidget(".record_frame.!frame.preview.outer0.canvas"), self.window.nametowidget(".record_frame.!frame.preview.outer1.canvas")]:
            for row in frame.winfo_children():
                row.children["checkbutton"]["state"] = "disabled"
                row.children["edit"]["state"] = "disabled"

        for button in self.useble_widgets:
            button["state"] = "disabled"

        def INDEXING(cursor, isremove):
            if isremove:
                for table in self.tables:
                    try:
                        cursor.execute(f"DROP INDEX [INDEX] ON {table}")
                        cursor.commit()
                    except:
                        print(f"{table} Unindexing fail")

                try:
                    cursor.execute("DROP INDEX [TRASH] ON RECORD")
                    cursor.commit()
                except:
                    print("RECORD Trash Unindexing fail")

                try:
                    cursor.execute("DROP INDEX [INDEX] ON RECORD")
                    cursor.commit()
                except:
                    print("RECORD Index Unindexing fail")
            else:
                try:
                    cursor.execute("CREATE INDEX [TRASH] ON RECORD ([TRASH])")
                    cursor.commit()
                except:
                    print("RECORD Trash Indexing fail")

                try:
                    cursor.execute("CREATE INDEX [INDEX] ON RECORD ([INDEX])")
                    cursor.commit()
                except:
                    print("RECORD Index Indexing fail")

                for table in self.tables:
                    try:
                        cursor.execute(f"CREATE INDEX [INDEX] ON {table} ([INDEX])")
                        cursor.commit()
                    except:
                        print(f"{table} Indexing fail")

        def add_trash(root, data, i):
            frame = tk.Frame(root, name=f"no_{i}", bd=1, relief="raised", width=560, height=30)
        
            check_button = ttk.Checkbutton(frame, name="checkbutton", takefocus=False)
            check_button.state(["!alternate"])
            check_button.place(x=5, rely=0.5, anchor="w")

            tk.Label(frame, text=data[1]).place(x=30, rely=0.5, width=60, anchor="w")

            if data[1] == "CODE":
                tk.Label(frame, text=data[2]).place(x=95, rely=0.5, width=60, anchor="w")
                tk.Label(frame, text=data[3]).place(x=160, rely=0.5, width=67, anchor="w")
            else:
                tk.Label(frame, text=data[2]).place(x=232, rely=0.5, width=77, anchor="w")
                tk.Label(frame, text=data[3]).place(x=314, rely=0.5, width=35, anchor="w")
                tk.Label(frame, text=data[4]).place(x=354, rely=0.5, width=80, anchor="w")

            tk.Label(frame, text=data[-1][:10]).place(x=439, rely=0.5, width=110, anchor="w")

            root.create_window(0, 30 * i, anchor="nw", window=frame, tags=("result", f"no_{i}"))

        def select_all_data(root, isSelect):
            for i in root.winfo_children():
                i.children["checkbutton"].state(["selected" if isSelect else "!selected"])

        def delete(canvas):
            isok = tkbox.askquestion("경고", "데이터를 영구적으로 삭제하시겠습니까?")
            if isok == "no":
                canvas.nametowidget(canvas.winfo_toplevel()).lift()
                return False
            
            self.Open("r_V2.accdb", "")
            cursor = self.cursors[0]

            INDEXING(cursor, False)

            for i in canvas.winfo_children():
                if "selected" in i.children["checkbutton"].state():
                    index = self.trash_data[int(i.winfo_name()[3:])][0]

                    for table in self.tables:
                        cursor.execute(f"DELETE FROM {table} WHERE [INDEX] = {index}")
                        cursor.commit()
                    
                    cursor.execute(f"DELETE FROM RECORD WHERE [INDEX] = {index}")
                    cursor.commit()

                    i.destroy()
            
            INDEXING(cursor, True)

            self.Close(0)

            canvas.nametowidget(canvas.winfo_toplevel()).lift()
        
        def restore(canvas):
            isok = tkbox.askquestion("경고", "데이터를 복구하시겠습니까?")
            if isok == "no":
                canvas.nametowidget(canvas.winfo_toplevel()).lift()
                return False

            self.Open("r_V2.accdb", "")
            self.Open("new_auto.mdb", "EOGKSALSRNRSHFNVPDLS")

            INDEXING(self.cursors[0], False)

            for i in canvas.winfo_children():
                if "selected" in i.children["checkbutton"].state():
                    index = self.trash_data[int(i.winfo_name()[3:])][0]
                    _type = self.trash_data[int(i.winfo_name()[3:])][1]

                    for table in self.tables[0 if _type == "CODE" else 3:3 if _type == "CODE" else 7]:
                        self.cursors[0].execute(f"SELECT * FROM {table} WHERE [INDEX] = {index}")
                        for row in self.cursors[0].fetchall():
                            columns = [columns[3] for columns in self.cursors[1].columns(table=table)]
                            column_types = [columns[5] for columns in self.cursors[1].columns(table=table)]

                            column_command = ""
                            data_command = ""
                            for column, column_type, data in zip(columns, column_types, row):
                                column_command += f"[{column}], "
                                if column_type != "DATETIME" and (self.types[column_type] == int or self.types[column_type] == float):
                                    data_command += f"{data}, "
                                else:
                                    data_command += f"\'{data}\', "
                            column_command = column_command[:-2]
                            data_command = data_command[:-2]
                            
                            command = f"INSERT INTO {table} ({column_command}) VALUES ({data_command})"
                            self.cursors[1].execute(command)
                            self.cursors[1].commit()

                    self.cursors[0].execute(f"UPDATE RECORD SET [TRASH] = FALSE WHERE [INDEX] = {index}")
                    self.cursors[0].commit()

                    i.destroy()

            INDEXING(self.cursors[0], True)

            self.Close(0)
            self.Close(1)

            canvas.nametowidget(canvas.winfo_toplevel()).lift()

        def close(window):
            for frame in [self.window.nametowidget(".record_frame.!frame.preview.outer0.canvas"), self.window.nametowidget(".record_frame.!frame.preview.outer1.canvas")]:
                for row in frame.winfo_children():
                    row.children["checkbutton"]["state"] = "normal"
                    row.children["checkbutton"].state(["!alternate", "!selected"])
                    row.children["edit"]["state"] = "normal"

            for button in self.useble_widgets:
                button["state"] = "normal"
                if button.winfo_class() == "TCheckbutton":
                    button.state(["!alternate"])
            
            window.destroy()

        trash_sort = [False, False, False, False, False, False, False]

        trash_window = tk.Toplevel(self.window)
        trash_window.geometry("600x400")
        trash_window.title("휴지통")
        trash_window.protocol("WM_DELETE_WINDOW", lambda: close(trash_window))

        tk.Label(trash_window, text="휴지통").place(relx=0.5, y=20, anchor="n")
        
        frame = tk.Frame(trash_window, bd=1, relief="sunken")
        frame.place(relx=0.5, y=50, width=580, height=300, anchor="n")

        label_frame = tk.Frame(frame, bd=1, relief="raised")
        label_frame.place(x=0, y=0, width=578, height=30)

        def order_trash(num):
            if num == 0:
                self.trash_data.sort(key=lambda x: x[1], reverse=trash_sort[num])
            elif num == 1:
                self.trash_data.sort(key=lambda x: x[2] if x[1] == "CODE" else "", reverse=trash_sort[num])
            elif num == 2:
                self.trash_data.sort(key=lambda x: x[3] if x[1] == "CODE" else "", reverse=trash_sort[num])
            elif num == 3:
                self.trash_data.sort(key=lambda x: x[2] if x[1] == "DATA" else "", reverse=trash_sort[num])
            elif num == 4:
                self.trash_data.sort(key=lambda x: x[3] if x[1] == "DATA" else "", reverse=trash_sort[num])
            elif num == 5:
                self.trash_data.sort(key=lambda x: x[4] if x[1] == "DATA" else "", reverse=trash_sort[num])
            else:
                self.trash_data.sort(key=lambda x: x[-1], reverse=trash_sort[num])
            
            trash_sort[num] = not trash_sort[num]

            frame.children["canvas"].delete("result")

            for idx, data in enumerate(self.trash_data):
                add_trash(frame.children["canvas"], data, idx)
            
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))

        tk.Button(label_frame, text="데이터", command=lambda: order_trash(0)).place(x=30, rely=0.5, width=60, anchor="w")
        tk.Button(label_frame, text="그룹코드", command=lambda: order_trash(1)).place(x=95, rely=0.5, width=60, anchor="w")
        tk.Button(label_frame, text="Minorcd", command=lambda: order_trash(2)).place(x=160, rely=0.5, width=67, anchor="w")
        tk.Button(label_frame, text="배합코드", command=lambda: order_trash(3)).place(x=232, rely=0.5, width=77, anchor="w")
        tk.Button(label_frame, text="타입", command=lambda: order_trash(4)).place(x=314, rely=0.5, width=35, anchor="w")
        tk.Button(label_frame, text="컬러코드", command=lambda: order_trash(5)).place(x=354, rely=0.5, width=80, anchor="w")
        tk.Button(label_frame, text="시간", command=lambda: order_trash(6)).place(x=439, rely=0.5, width=110, anchor="w")

        canvas = tk.Canvas(frame, name="canvas", highlightthickness=0)
        canvas.place(x=0, y=30, width=560, height=270)

        scrollbar = tk.Scrollbar(frame, name="scrollbar")
        scrollbar.place(x=560, y=30, width=20, height=270)

        scrollbar.configure(command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        self.Open("r_V2.accdb", "")
        cursor = self.cursors[0]

        INDEXING(cursor, False)

        cursor.execute("SELECT [INDEX] FROM RECORD WHERE [TRASH] = TRUE")
        self.trash_data = []
        for i in cursor.fetchall():
            self.trash_data.append(list(i))

        for idx, data in enumerate(self.trash_data):
            # 데이터
            cursor.execute(f"SELECT [TYPE] FROM RECORD WHERE [INDEX] = {data[0]}")
            self.trash_data[idx].append(cursor.fetchall()[0][0])

            if self.trash_data[idx][-1] == "CODE":
                # 그룹코드
                cursor.execute(f"SELECT [Group_Code] FROM COLORCDGR_TBL WHERE [INDEX] = {data[0]}")
                self.trash_data[idx].append(cursor.fetchall()[0][0])

                # Minorcd
                for table, column in (("AUTO_CODE_TBL", "Minorcd"), ("CARMAKER_TBL", "MakerId"), ("COLORCDGR_TBL", "Maker_ENG")):
                    cursor.execute(f"SELECT [{column}] FROM {table} WHERE [INDEX] = {data[0]}")
                    temp = cursor.fetchall()
                    if len(temp) > 0:
                        self.trash_data[idx].append(temp[0][0])
                        break
            else:
                # 배합코드
                cursor.execute(f"SELECT [MixCd] FROM COLORMIX_TBL WHERE [INDEX] = {data[0]}")
                self.trash_data[idx].append(cursor.fetchall()[0][0])

                # 타입
                cursor.execute(f"SELECT [PaintTy] FROM COLORMIX_TBL WHERE [INDEX] = {data[0]}")
                self.trash_data[idx].append(cursor.fetchall()[0][0])

                # 컬러코드
                cursor.execute(f"SELECT [ColorCode] FROM MIXCOLORINFO_TBL WHERE [INDEX] = {data[0]}")
                self.trash_data[idx].append(cursor.fetchall()[0][0])

            # 시간
            cursor.execute(f"SELECT [_DATE] FROM RECORD WHERE [INDEX] = {data[0]}")
            self.trash_data[idx].append(cursor.fetchall()[0][0])

        for idx, data in enumerate(self.trash_data):
            add_trash(canvas, data, idx)

        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

        INDEXING(cursor, True)

        self.Close(0)

        select_all = ttk.Checkbutton(trash_window, text="전체 선택", takefocus=False)
        select_all.config(command=lambda: select_all_data(canvas, "selected" in select_all.state()))
        select_all.place(x=20, y=20)
        select_all.state(["!alternate"])

        tk.Button(trash_window, text="영구삭제", command=lambda: delete(canvas)).place(relx=0.25, rely=0.925, anchor="center")
        tk.Button(trash_window, text="복원하기", command=lambda: restore(canvas)).place(relx=0.5, rely=0.925, anchor="center")
        tk.Button(trash_window, text="뒤로가기", command=lambda: close(trash_window)).place(relx=0.75, rely=0.925, anchor="center")

        trash_window.mainloop()

    def Search_V2(self, cursor_indexs, search_texts, including_all_text, except_tables=None, except_columns=None, except_texts=None):
        for cursor_index in cursor_indexs:
            cursor = self.cursors[cursor_index]
            result = [[], [], [], [], [], [], []]

            for table_index, table in enumerate(self.tables):
                for column_index in [columns[3] if columns[3] != "INDEX" else "" for columns in cursor.columns(table=table)]:
                    if column_index != "":
                        try:
                            cursor.execute(f"CREATE INDEX {column_index} ON {table} ({column_index})")
                            cursor.commit()
                        except:
                            print(f"{table} {column_index} Indexing Fail")

                if except_tables != None and table in except_tables:
                    continue

                self.update_status_text(f"{table} 에서 검색중...")

                columns = [columns[3] if columns[3] not in except_columns[table] else "" for columns in cursor.columns(table=table)]
                column_types = [columns[5] if columns[3] not in except_columns[table] else "" for columns in cursor.columns(table=table)]
                while columns.count("") > 0:
                    columns.remove("")
                    column_types.remove("")
                try:
                    index = columns.index("INDEX")

                    del columns[index]
                    del column_types[index]
                except:
                    pass

                temp_results = []

                if search_texts == "":
                    cursor.execute(f"SELECT * FROM {table}")
                else:
                    search_command = f"SELECT * FROM {table} WHERE "

                    for column, column_type in zip(columns, column_types):
                        temp_search_command = f"[{column}] IS NOT NULL AND "

                        search_values = ""
                        for text in search_texts:
                            try:
                                if self.types[column_type] == str:
                                    search_values += f"REPLACE({column}, \' \', \'\') LIKE \'%{self.types[column_type](text)}%\' OR "
                                elif self.types[column_type] == int or self.types[column_type] == float:
                                    search_values += f"{column} = {str(self.types[column_type](text))} OR "
                            except:
                                continue
                        
                        if search_values != "":
                            search_command += f"({temp_search_command}({search_values[:-4]})) OR "
                    search_command = search_command[:-4]

                    cursor.execute(search_command)
                
                for i in cursor.fetchall():
                    temp_results.append(list(i))

                for index in range(len(temp_results)):
                    for data_index in range(len(temp_results[index])):
                        temp_results[index][data_index] = str(temp_results[index][data_index])
                
                def Find_match_all_search_text(lists):
                    for data in lists[:]:
                        match_list = [False for x in range(len(search_texts))]

                        for data_detail in data:
                            for index in range(len(search_texts)):
                                if search_texts[index] in str(data_detail):
                                    match_list[index] = True
                            
                        if False in match_list:
                            del lists[lists.index(data)]
                    
                    return lists
                
                def Find_match_except_text(lists):
                    flag = False

                    for data in lists[:]:
                        for data_detail in data:
                            for text in except_texts:
                                if text in str(data_detail):
                                    del lists[lists.index(data)]
                                    flag = True
                                    break
                            
                            if flag:
                                flag = False
                                break
                    
                    return lists

                if including_all_text:
                    result[table_index] = Find_match_except_text(Find_match_all_search_text(temp_results)) if except_texts != None else Find_match_all_search_text(temp_results)
                else:
                    result[table_index] = Find_match_except_text(temp_results) if except_texts != None else temp_results

            self.Finding_group(cursor_index, result)

        for i in cursor_indexs:
            self.Close(i)

        self.Apply_preview_result()

    def Finding_group(self, cursor_index, data):
        cursor = self.cursors[cursor_index]
        search_type_postition = [1, 2, 1, 12]
        code = data[:3]
        data = data[3:]

        if cursor_index == 0:
            def search_index(_type, index):
                result = []

                for table in self.tables[0 if _type == "code" else 3:3 if _type == "code" else 7]:
                    cursor.execute(f"SELECT TRASH FROM RECORD WHERE [INDEX] = {index}")
                    cursor.execute(f"SELECT * FROM {table} WHERE [INDEX] = {index}")
                    temp_result = []
                    for i in cursor.fetchall():
                        temp_result.append(list(i))

                    for idx1, row in enumerate(temp_result):
                        for idx2, value in enumerate(row):
                            temp_result[idx1][idx2] = str(value)
                    
                    result.append(temp_result)

                return result
            
            def search_time_and_trash(index):
                cursor.execute(f"SELECT [_DATE],[TRASH] FROM RECORD WHERE [INDEX] = {index}")
                return cursor.fetchall()[0]

            first_code = []
            duplicated_code = []

            first_data = []
            duplicated_data = []

            for i in code:
                for p in i:
                    if p[-1] not in first_code:
                        first_code.append(p[-1])
                    else:
                        if p[-1] not in duplicated_code:
                            duplicated_code.append(p[-1])
            
            for i in data:
                for p in i:
                    if p[-1] not in first_data:
                        first_data.append(p[-1])
                    else:
                        if p[-1] not in duplicated_data:
                            duplicated_data.append(p[-1])
            
            self.update_status_text("코드 0 / 0   |   데이터 0 / 0")
            self.update_progressbar(len(first_code) + len(first_data), None)

            # code
            for index in first_code:
                result = search_index("code", index)
                time_trash = search_time_and_trash(index)
                if time_trash[1]:
                    continue
                result.append(cursor_index)
                result.append(time_trash[0])
                result.append(time_trash[1])
                self.datas["code"].append(result)
                count = str(len(self.datas["code"]))
                self.update_status_text(f"코드 {count} / {str(len(first_code))}   |   데이터 0 / {str(len(first_data))}")
                self.update_progressbar(None, 1)
            
            # data
            for index in first_data:
                result = search_index("data", index)
                time_trash = search_time_and_trash(index)
                if time_trash[1]:
                    continue
                result.append(cursor_index)
                result.append(time_trash[0])
                result.append(time_trash[1])

                if len(result[0]) > 1:
                    temp_cott = [[result[0][1]], [row if row[0] == result[0][1][0] else "" for row in result[1]], [result[2][1]], [result[3][1]]]
                    while True:
                        if temp_cott[1][0] == "":
                            del temp_cott[1][0]
                        else:
                            break

                    for i in range(4):
                        if i != 1:
                            del result[i][1]
                        else:
                            for copy_cott in temp_cott[1]:
                                if copy_cott in result[1]:
                                    del result[1][result[1].index(copy_cott)]

                    result.append(temp_cott)
                
                self.datas["data"].append(result)
                count = str(len(self.datas["data"]))
                self.update_status_text(f"코드 {str(len(first_code))} / {str(len(first_code))}   |   데이터 {count} / {str(len(first_data))}")
                self.update_progressbar(None, 1)
            
            self.update_progressbar(len(first_code) + len(first_data), None, True)
        else:
            status_count = 0
            code_count = len(code[0]) + len(code[1]) + len(code[2])
            data_count = len(data[0]) + len(data[1]) + len(data[2]) + len(data[3])
            self.update_progressbar(code_count + data_count, None)

            # code
            for table_index in range(3):
                for code_list in code[table_index]:
                    result = [[], [], [], cursor_index]

                    result[table_index].append(code_list)
                    self.datas["code"].append(result)
                    
                    status_count += 1
                    self.update_status_text(f"코드 {status_count} / {code_count}   |   데이터 0 / {data_count}")
                    self.update_progressbar(None, 1)
            # data
            def search_data(table, mixcd, paintty):
                paintty_column_name = "PaintTy" if table != 2 else "paintTy"
                cursor.execute(f"SELECT * FROM {table} WHERE [MixCd] = \'{mixcd}\' AND [{paintty_column_name}] = \'{paintty}\'")

                result = []
                for i in cursor.fetchall():
                    result.append(list(i))

                for idx1, row in enumerate(result):
                    for idx2, value in enumerate(row):
                        result[idx1][idx2] = str(value)

                return result

            def remove_data(data_result):
                for original, result, idx in zip(data[:], data_result, range(4)):
                    for inner in result:
                        if inner in original:
                            del data[idx][data[idx].index(inner)]

            before_count = 0
            after_count = 0

            for table_index in range(4):
                while len(data[table_index]) > 0:
                    before_count = len(data[0]) + len(data[1]) + len(data[2]) + len(data[3])
                    MixCd = data[table_index][0][0]
                    PaintTy = data[table_index][0][search_type_postition[table_index]]
                    
                    data_result = [search_data(self.tables[x], MixCd, PaintTy) for x in range(3, 7)]
                    remove_data(data_result)
                    data_result.append(cursor_index)
                    self.datas["data"].append(data_result)

                    after_count = len(data[0]) + len(data[1]) + len(data[2]) + len(data[3])
                    self.update_progressbar(None, before_count - after_count)
                    self.update_status_text(f"코드 {code_count} / {code_count}   |   데이터 {data_count - (len(data[0]) + len(data[1]) + len(data[2]) + len(data[3]))} / {data_count}")
            
            packed_count = len(self.datas["data"])
            self.update_status_text(f"코드 {code_count} / {code_count}   |   데이터 {data_count} / {data_count} (그룹화된 데이터: {packed_count})")
        
            self.update_progressbar(code_count + data_count, None, True)
            
        for table in self.tables:
            for column_index in [columns[3] if columns[3] != "INDEX" else "" for columns in cursor.columns(table=table)]:
                if column_index != "":
                    try:
                        cursor.execute(f"DROP INDEX {column_index} ON {table}")
                        cursor.commit()
                    except:
                        print(f"{table} {column_index} Droping Index Fail")

    def Add_result(self, root, file_name, data_type, data, place_index, data_index):
        if data_type == "code":
            frame = tk.LabelFrame(root, name=f"no_{str(data_index)}", width=430, height=45, labelanchor="n")

            check_button = ttk.Checkbutton(frame, name="checkbutton", takefocus=False)
            check_button.state(["!alternate"])
            check_button.place(x=5, rely=0.5, anchor="w")

            _code = ""
            _minorcd = ""
            _time = ""
            _trash = ""

            if data[3] == 0:
                frame["text"] = "-"
                _code = data[2][0][1]
                _minorcd = data[0][0][1]
                _time = data[4][:10]
                _trash = str(data[5])
            elif len(data[0]) > 0:
                frame["text"] = "AUTO_CODE_TBL"
                _minorcd = data[0][0][1]
            elif len(data[1]) > 0:
                frame["text"] = "CARMAKER_TBL"
                _minorcd = data[1][0][0]
            else:
                frame["text"] = "COLORCDGR_TBL"
                _code = data[2][0][1]
                _minorcd = data[2][0][2]

            tk.Label(frame, name="name", text=file_name).place(x=28, rely=0.5, width=83, anchor="w")
            tk.Label(frame, name="code", text=_code).place(x=116, rely=0.5, width=60, anchor="w")
            tk.Label(frame, name="minorcd", text=_minorcd).place(x=181, rely=0.5, width=67, anchor="w")
            tk.Label(frame, name="time", text=_time).place(x=253, rely=0.5, width=70, anchor="w")
            tk.Label(frame, name="trash", text=_trash).place(x=328, rely=0.5, width=50, anchor="w")

            tk.Button(frame, name="edit", text="수정", command=lambda: self.Edit_row("code", data_index), takefocus=False).place(x=390, rely=0.5, width=30, anchor="w")

            root.create_window(0, 45 * place_index, window=frame, tags=("result", f"no_{str(data_index)}"))
        else:
            search_type_postition = [1, 2, 1 ,12]
            frame = tk.Frame(root, name=f"no_{str(data_index)}", bd=1, relief="raised", width=500, height=28)

            check_button = ttk.Checkbutton(frame, name="checkbutton", takefocus=False)
            check_button.place(x=5, rely=0.5, anchor="w")

            if not (len(data[0]) > 0 and len(data[1]) > 0 and len(data[2]) > 0 and len(data[3]) > 0):
                frame["bg"] = "#DF6464"
                style = ttk.Style()
                style.map(f"custom_{data_index}.TCheckbutton", background=[("", "#DF6464")])
                check_button["style"] = f"custom_{data_index}.TCheckbutton"
            elif len(data) > 7:
                frame["bg"] = "#3CB043"
                style = ttk.Style()
                style.map(f"custom_{data_index}.TCheckbutton", background=[("", "#3CB043")])
                check_button["style"] = f"custom_{data_index}.TCheckbutton"

            check_button.state(["!alternate"])

            _code = ""
            _type = ""
            _color = ""
            _time = ""
            _trash = ""

            for i in range(4):
                try:
                    _code = data[i][0][0].strip()
                    _type = data[i][0][search_type_postition[i]].strip()
                except:
                    continue
            
            if len(data[3]) > 0:
                _color = data[3][0][4].strip()

            try:
                _time = data[5][:10]
                _trash = str(data[6])
            except:
                pass

            tk.Label(frame, name="name", text=file_name).place(x=29, rely=0.5, width=83, anchor="w")
            tk.Label(frame, name="code", text=_code).place(x=117, rely=0.5, width=77, anchor="w")
            tk.Label(frame, name="type", text=_type).place(x=199, rely=0.5, width=35, anchor="w")
            tk.Label(frame, name="color", text=_color).place(x=239, rely=0.5, width=80, anchor="w")
            tk.Label(frame, name="time", text=_time).place(x=324, rely=0.5, width=70, anchor="w")
            tk.Label(frame, name="trash", text=_trash).place(x=399, rely=0.5, width=50, anchor="w")

            tk.Button(frame, name="edit", text="수정", command=lambda: self.Edit_row("data", data_index), takefocus=False).place(x=460, rely=0.5, width=30, anchor="w")

            root.create_window(0, 28 * place_index, window=frame, tags=("result", f"no_{str(data_index)}"))
    
    def Apply_preview_result(self):
        code_canvas = self.window.nametowidget(".record_frame.!frame.preview.outer0.canvas")
        data_canvas = self.window.nametowidget(".record_frame.!frame.preview.outer1.canvas")
        page_entry = self.window.nametowidget(".record_frame.!frame.!frame3.entry")
        page_now = self.window.nametowidget(".record_frame.!frame.!frame3.now")
        page_max = self.window.nametowidget(".record_frame.!frame.!frame3.max")
        view_count = int(self.window.nametowidget(".record_frame.!frame.!frame3.!combobox").get())

        code_canvas.delete("result")
        data_canvas.delete("result")

        page_entry["state"] = "normal"
        page_entry.delete(0, "end")
        page_entry.insert(0, "1")
        page_now["text"] = "1"
        page_max["text"] = str(math.ceil(len(self.datas["data"]) / view_count)) if len(self.datas["data"]) > 0 else "1"

        self.datas_order["code"] = [x for x in range(len(self.datas["code"]))]
        self.datas_order["data"] = [x for x in range(len(self.datas["data"]))]

        for i in self.datas_order["code"]:
            self.Add_result(code_canvas, "r_V2.accdb" if self.datas["code"][i][3] == 0 else "new_auto.mdb", "code", self.datas["code"][i], i, i)
        
        for i in code_canvas.winfo_children():
            i.children["checkbutton"].state(["!selected"])

        code_canvas.update_idletasks()
        code_canvas.configure(scrollregion=code_canvas.bbox("all"))

        for i in range(view_count) if len(self.datas["data"]) > view_count else self.datas_order["data"]:
            self.Add_result(data_canvas, "r_V2.accdb" if self.datas["data"][i][4] == 0 else "new_auto.mdb", "data", self.datas["data"][i], i, i)
        
        for i in data_canvas.winfo_children():
            i.children["checkbutton"].state(["!selected"])

        data_canvas.update_idletasks()
        data_canvas.configure(scrollregion=data_canvas.bbox("all"))

        for i in self.useble_widgets:
            i["state"] = "normal"
            if i.winfo_class() == "TCheckbutton":
                i.state(["!alternate"])
        
        self.update_status_text(f"코드: {str(len(self.datas['code']))}개   |   데이터: {str(len(self.datas['data']))}개")

        self.window.nametowidget(".record_frame.!frame.!frame3.left")["state"] = "disabled"
        self.window.nametowidget(".record_frame.!frame.!frame3.left_end")["state"] = "disabled"

        if page_now["text"] == page_max["text"]:
            self.window.nametowidget(".record_frame.!frame.!frame3.right")["state"] = "disabled"
            self.window.nametowidget(".record_frame.!frame.!frame3.right_end")["state"] = "disabled"

    def Edit_row(self, _type, row_index):
        window = tk.Toplevel(self.window)
        window.title(_type)
        window.geometry("750x380" if _type == "code" else "950x940")
        window.resizable(False, False)

        if _type == "code":
            code_frame = tk.Frame(window, name="code_frame", bd=1, relief="sunken")
            code_frame.place(x=0, y=0, width=750, height=380)
            self.edit_function[0](code_frame, window, row_index)
        else:
            data_frame = tk.Frame(window, name="data_frame", bd=1, relief="sunken")
            data_frame.place(x=0, y=0, width=950, height=940)
            self.edit_function[1](data_frame, window, row_index)

        window.bind("<MouseWheel>", lambda e: self.bind_function(e.widget, e.delta))

        window.mainloop()

    def update_status_text(self, text):
        status = self.window.nametowidget(".record_frame.!frame.!frame.state")
        status["text"] = "상태: " + text

    def update_progressbar(self, max, add, full=False):
        bar = self.window.nametowidget(".record_frame.!frame.!progressbar")
        
        if full:
            bar["value"] = max
            return

        if max != None:
            bar["maximum"] = max
            return
        
        bar.step(add)

    def search_for_color(self, keyword, canvas):
        canvas.delete("result")

        self.Open("r_V2.accdb", "")
        self.Open("new_auto.mdb", "EOGKSALSRNRSHFNVPDLS")
        
        for cursor_index in range(2):
            cursor = self.cursors[cursor_index]
            
            condition = f"\'%{keyword}%\'" if keyword.find("-") == -1 else f"\'{keyword}%\'"
            cursor.execute(f"SELECT [MixCd], [ColorCode] FROM MIXCOLORINFO_TBL WHERE [ColorCode] LIKE {condition}")
            for i in cursor.fetchall():
                cursor.execute(f"SELECT [StdDesc] FROM COLORMIXSTD_TBL WHERE [MixCd] = \'{i[0]}\'")
                data = cursor.fetchone()
                if data != None and len(data) > 0:
                    frame = tk.Frame(canvas, bd=1, relief="raised", width=258, height=25)
                    
                    one_text = i[1].strip() if i[1] != None else i[1]
                    two_text = i[0].strip() if i[0] != None else i[0]
                    three_text = data[0].strip() if data[0] != None else data[0]

                    one = tk.Label(frame, text=one_text)
                    one.place(x=5, rely=0.5, width=80, anchor="w")
                    two = tk.Label(frame, text=two_text)
                    two.place(x=95, rely=0.5, width=77, anchor="w")
                    three = tk.Label(frame, text=three_text)
                    three.place(x=182, rely=0.5, width=70, anchor="w")

                    Hovertip(one, one_text, hover_delay=0)
                    Hovertip(two, two_text, hover_delay=0)
                    Hovertip(three, three_text, hover_delay=0)

                    canvas.create_window(0, 25 * (len(canvas.winfo_children()) - 1), anchor="nw", window=frame, tags="result")

            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))

            self.Close(cursor_index)
    
    def get_TONERGRV_TBL(self):
        toners = {}
        
        self.Open("new_auto.mdb", "EOGKSALSRNRSHFNVPDLS")
        self.cursors[1].execute("SELECT [tonerCd], [gravity] FROM TONERGRV_TBL")
        for toner in self.cursors[1].fetchall():
            toners[toner[0].strip()] = float(toner[1])
        self.Close(1)

        return toners

class Calculation():
    def calculation_setting(self):
        self.weight = self.get_TONERGRV_TBL()
    
    def get_toner(self, main_index, mix_rates, color_codes):
        try:
            one = [r / c for r, c in zip(mix_rates, (self.weight[i] for i in color_codes))]
            two = [i / sum(one) for i in (500, 800, 900, 2000, 3600)]
            three = self.weight[color_codes[main_index]]
            return [one[main_index] * two[i] * three for i in range(5)]
        except:
            return False

class UI(SQL, Calculation):
    search_options = {"database":None, "including_text":None, "except_tables":None, "except_columns":None, "except_texts":None}

    def __init__(self):
        version = chromedriver_autoinstaller.get_chrome_version().split(".")[0]
        current_path = os.getcwd()
        driver_path = f"{current_path}/{version}/chromedriver.exe"
        
        if not os.path.exists(driver_path):
            chromedriver_autoinstaller.install(True)

        self.window = tk.Tk()
        self.window.title("노루 데이터베이스 편집 프로그램")
        self.window.geometry("400x400")
        self.window.resizable(False, False)

        self.setting()
        self.calculation_setting()
        self.window.attributes("-topmost", True)
        self.window.attributes("-topmost", False)
        self.window.mainloop()
    
    def generation_widget(self, length, names, visibles, shares, texts, widths, frame, tapables, passIndex = None):
        for index, name, isVisible, isShare, text, width, tap in zip(range(length), names, visibles, shares, texts, widths, tapables):
            label = tk.Label(frame, name="label" + str(index) + ("h" if not isVisible else ""), highlightthickness=2, text=name, width=width, anchor="center")
            entry = tk.Entry(frame, name="entry" + str(index) + ("h" if not isVisible else ""), highlightthickness=2, highlightcolor="SystemButtonFace", textvariable=isShare, width=width, takefocus=tap)

            label.grid(row=0, column=index, sticky="ew", padx=2)
            entry.grid(row=1, column=index, sticky="ew", padx=2)

            if not isVisible:
                label.grid_remove()
                entry.grid_remove()
            
            entry.insert(0, text)

            if passIndex == index:
                entry.destroy()

    def generation_combo(self, index, share, texts, width, frame):
        combobox = ttk.Combobox(frame, name="combobox" + str(index), textvariable=share, values=texts, width=width, takefocus=False)
        combobox.grid(row=1, column=index, sticky="ew", padx=2)
        combobox.set(texts[0])

    def code_setting(self, code_frame, main_frame, code_index = None):
        str_one = tk.StringVar()
        str_two = tk.StringVar()
        str_three = tk.StringVar(value="2023-02-28 오전 9:00:00")
        str_four = tk.StringVar(value="wgk9985")
        
        if code_index != None:
            trace_str1 = [tk.StringVar() for i in range(7)]
            trace_str2 = [tk.StringVar() for i in range(8)]
            trace_str3 = [tk.StringVar() for i in range(8)]

        one_frame = tk.Frame(code_frame)
        two_frame = tk.Frame(code_frame)
        three_frame = tk.Frame(code_frame)

        #   AUTO_CODE_TBL
        tk.Label(one_frame, text="AUTO_CODE_TBL").grid(row=0, column=0)
        inner_frame = tk.LabelFrame(one_frame, name="frame")
        inner_frame.grid(row=1, column=0)
        self.generation_widget(7,
                            ("Majorcd", "Minorcd", "Codenm", "Remark1", "Remark2", "Sortorder", "RegDate"),
                            (False, True, True, False, False, True, False),
                            (None, str_one, str_two, None, None, str_one, str_three),
                            ("60", "", "", "", "", "", ""),
                            (8, 8, 12, 8, 8, 8, 20),
                            inner_frame,
                            (False, True, True, False, False, True, False))
        
        #   CARMAKER_TBL
        tk.Label(two_frame, text="CARMAKER_TBL").grid(row=0, column=0)
        inner_frame = tk.LabelFrame(two_frame, name="frame")
        inner_frame.grid(row=1, column=0)
        self.generation_widget(8,
                            ("MakerId", "Maker_K", "Maker_E", "Maker_C", "UseYn", "RegId", "RegDate", "MakerNew_C"),
                            (True, True, True, False, False, False, False, False),
                            (str_one, None, str_two, None, None, str_four, str_three, None),
                            ("", "", "", "", "Y", "", "", ""),
                            (8, 12, 12, 8, 6, 8, 20, 12),
                            inner_frame,
                            (False, True, False, False, False, False, False, False))
        
        #   COLORCDGR_TBL
        tk.Label(three_frame, text="COLORCDGR_TBL").grid(row=0, column=0)
        inner_frame = tk.LabelFrame(three_frame, name="frame")
        inner_frame.grid(row=1, column=0)
        self.generation_widget(8,
                            ("idx", "Group_Code", "Maker_ENG", "Brand", "Site", "Origin", "RegId", "RegDate"),
                            (True, True, True, True, False, True, False, False),
                            (None, None, str_one, None, None, None, str_four, str_three),
                            ("", "", "", "", "KEC", "", "", ""),
                            (8, 10, 10, 6, 4, 6, 8, 20),
                            inner_frame,
                            (True, True, False, True, False, True, False, False))

        one_frame.place(relx=0.5, rely=0.15, anchor="center")
        two_frame.place(relx=0.5, rely=0.4, anchor="center")
        three_frame.place(relx=0.5, rely=0.65, anchor="center")

        def show():
            for i in (one_frame.children["frame"], two_frame.children["frame"], three_frame.children["frame"]):
                for o in i.winfo_children():
                    if "h" in o.winfo_name():
                        if o.winfo_ismapped():
                            o.grid_remove()
                        else:
                            o.grid()
        
        def clear(isAll):
            isok = tkbox.askquestion("경고", "입력하신 데이터를 지우시겠습니까?")
            if isok == "no":
                return False

            for i in (one_frame.children["frame"], two_frame.children["frame"], three_frame.children["frame"]):
                for o in i.winfo_children():
                    if o.winfo_class() == "Entry" and (isAll or (not isAll and not "h" in o.winfo_name())):
                        o.delete(0, "end")

        def reset():
            if clear(True) == False: return

            if code_index != None:
                result_set()
                return

            str_three.set("2023-02-28 오전 9:00:00")
            str_four.set("wgk9985")
            one_frame.children["frame"].children["entry0h"].insert(0, "60")
            two_frame.children["frame"].children["entry4h"].insert(0, "Y")
            three_frame.children["frame"].children["entry4h"].insert(0, "KEC")

        def ready_for_save():
            input = [[], [], []]

            for index, frame in enumerate((one_frame.children["frame"], two_frame.children["frame"], three_frame.children["frame"])):
                for entry in frame.winfo_children():
                    if entry.winfo_class() == "Entry":
                        input[index].append(entry.get())

            if code_index == None:
                self.Save(input)
            else:
                self.Update("code", input, code_index, main_frame)

        def result_set(is_bind = False):
            data = self.datas["code"][code_index]

            widget1 = [widget if widget.winfo_class() != "Label" else "deleteme" for widget in one_frame.children["frame"].winfo_children()]
            widget2 = [widget if widget.winfo_class() != "Label" else "deleteme" for widget in two_frame.children["frame"].winfo_children()]
            widget3 = [widget if widget.winfo_class() != "Label" else "deleteme" for widget in three_frame.children["frame"].winfo_children()]

            for i in (widget1, widget2, widget3):
                while True:
                    try:
                        del i[i.index("deleteme")]
                    except:
                        break

            for idx in range(3):
                if len(data[idx]) > 0:
                    for value, widget in zip(data[idx][0], widget1 if idx == 0 else (widget2 if idx == 1 else widget3)):
                        widget.delete(0, "end")
                        widget.insert(0, value)
                else:
                    one_frame.place_forget() if idx == 0 else (two_frame.place_forget() if idx == 1 else three_frame.place_forget())

            if is_bind:
                def change_alert(widget, original_data):
                    if widget.get() == original_data:
                        widget.config(highlightbackground = "SystemButtonFace", highlightcolor= "SystemButtonFace")
                    else:
                        widget.config(highlightbackground = "red", highlightcolor= "red")

                def ready_to_binding(variable, widget, value):
                    widget.configure(textvariable=variable)
                    variable.trace_add("write", lambda a, b, c: widget.after(1, lambda: change_alert(widget, value)))

                for pack in ((widget1, data[0][0] if len(data[0]) > 0 else None, trace_str1), (widget2, data[1][0] if len(data[1]) > 0 else None, trace_str2), (widget3, data[2][0] if len(data[2]) > 0 else None, trace_str3)):
                    if pack[1] == None:
                        continue

                    for widget, value, trace_var in zip(pack[0], pack[1], pack[2]):
                        if widget["textvariable"] != "":
                            for var in (str_one, str_two, str_three, str_four):
                                if str(var) == widget["textvariable"]:
                                    ready_to_binding(var, widget, value)
                                    break
                        else:
                            trace_var.set(value=value)
                            ready_to_binding(trace_var, widget, value)

        if code_index != None:
            result_set(True)

        tk.Button(code_frame, text="숨겨진 값 표시/숨기기", takefocus=False, command=show).place(relx=0.02, rely=0.9)
        tk.Button(code_frame, text="지우기", takefocus=False, command=lambda: clear(False)).place(relx=0.3, rely=0.9)
        tk.Button(code_frame, text="지우기 (고정값)", takefocus=False, command=lambda: clear(True)).place(relx=0.385, rely=0.9)
        tk.Button(code_frame, text="기본값으로 초기화", takefocus=False, command=reset).place(relx=0.54, rely=0.90)
        save_button = tk.Button(code_frame, text="저장하기", command=ready_for_save)
        save_button.place(relx=0.75, rely=0.9)
        save_button.bind("<Return>", lambda e: ready_for_save())
        tk.Button(code_frame, text="뒤로가기", takefocus=False, command=lambda: self.change_frames(main_frame, code_frame, 400, 400)).place(relx=0.88, rely=0.9)

    def data_setting(self, data_frame, main_frame, data_index = None):
        str_one = tk.StringVar()
        str_two = tk.StringVar() 
        str_three = tk.StringVar()
        isnormal = tk.IntVar()
        company = tk.IntVar()
        isshow = tk.IntVar(value=0)
        stop_spec = [tk.BooleanVar(value=False), tk.BooleanVar(value=False)]
        stop_filter = tk.BooleanVar(value=False)
        colormixstd_match = {
            "차량판넬":"000", "마스터판넬":"001", "라인판넬":"002", "실차":"003", 
            "글로벌칼라북(B)":"004", "글로벌칼라북(I)":"005", "글로벌칼라북(K)":"006", "글로벌칼라북(S)":"007", 
            "기타":"008", "현지제작판넬":"009", "글로벌칼라북(SD)":"010", "현장배합":"011", 
            "현장시편":"012"
        }

        if data_index != None:
            trace_str1_0 = [tk.StringVar() for i in range(18)]
            trace_str2_0 = [[tk.StringVar(), tk.StringVar()] for i in range(len(self.datas["data"][data_index][1]))]
            trace_str3_0 = [tk.StringVar() for i in range(4)]
            trace_str4_0 = [tk.StringVar() for i in range(13)]
            if len(self.datas["data"][data_index]) > 7:
                trace_str1_1 = [tk.StringVar() for i in range(18)]
                trace_str2_1 = [[tk.StringVar(), tk.StringVar()] for i in range(len(self.datas["data"][data_index][7][1]))]
                trace_str3_1 = [tk.StringVar() for i in range(4)]
                trace_str4_1 = [tk.StringVar() for i in range(13)]

        one_frame = tk.Frame(data_frame)
        two_frame = tk.Frame(data_frame)
        three_frame = tk.Frame(data_frame)
        four_frame = tk.Frame(data_frame)

        def input_filter(widget, text, edit_dir, function=None, function_arg=None):
            if stop_filter.get() == True: return True

            text = text.upper()
            if re.search("[^A-Z0-9-()]", text) != None:
                return False

            widget = self.window.nametowidget(widget)
            index = widget.index(tk.INSERT)
            if index == widget.index(tk.END):
                index = -1
            widget.delete(0, "end")
            widget.insert(0, text)
            if index != -1:
                widget.icursor(index + (1 if edit_dir == "1" else -1))
            widget.after_idle(lambda: widget.config(validate="all"))

            if function == "get_code":
                get_code(function_arg, text)
            elif function == "calculation":
                calculation(function_arg, text)

            return True

        def detail_gen(root, name, row):
            frame = tk.Frame(root, name=name, bd=1, relief="ridge", width=702, height=202)
            frame.grid(row=row, column=0, sticky="e")

            label_frame = tk.Frame(frame, name="label_frame", bd=1, relief="raised")
            label_frame.place(x=0, y=0, width=700, height=30)

            tk.Label(label_frame, text="MixRate", anchor="center").place(x=30, rely=0.5, anchor="w", width=70, height=22)
            tk.Label(label_frame, text="Nbr", anchor="center").place(x=100, rely=0.5, anchor="w", width=25, height=22)
            tk.Label(label_frame, text="TonerCd", anchor="center").place(x=125, rely=0.5, anchor="w", width=70, height=22)
            tk.Label(label_frame, text="Spec05", anchor="center").place(x=200, rely=0.5, anchor="w", width=70, height=22)
            tk.Label(label_frame, text="Spec08", anchor="center").place(x=275, rely=0.5, anchor="w", width=70, height=22)
            tk.Label(label_frame, text="Spec10", anchor="center").place(x=350, rely=0.5, anchor="w", width=70, height=22)
            tk.Label(label_frame, text="Spec20", anchor="center").place(x=425, rely=0.5, anchor="w", width=70, height=22)
            tk.Label(label_frame, text="Spec40", anchor="center").place(x=500, rely=0.5, anchor="w", width=70, height=22)
            tk.Label(label_frame, text=" 갯수: 0", bd=1, relief="sunken", anchor="w").place(x=580, rely=0.5, anchor="w", width=60, height=22)

            canvas = tk.Canvas(frame, name="canvas", highlightthickness=0)
            canvas.place(x=0, y=30, width=680, height=170)

            scrollbar = tk.Scrollbar(frame, name="scrollbar")
            scrollbar.place(x=680, y=30, width=20, height=170)

            scrollbar.configure(command=canvas.yview)
            canvas.configure(yscrollcommand=scrollbar.set)

            detail_frame = tk.Frame(canvas, name="detail")
            canvas.create_window((0, 0), window=detail_frame, anchor="nw")

        def calculation(widget, data, isUpdate = False):
            detail = widget
            if not isUpdate:
                detail = self.window.nametowidget(widget).master.master

            if stop_spec[int(detail.master.master.winfo_name()[-1])].get() == True:
                return

            def fail():
                for i in detail.winfo_children():
                    for id in range(5):
                        i.children[str(id)]["state"] = "normal"
                        i.children[str(id)].delete(0, "end")
                        i.children[str(id)].insert(0, "-")
                        i.children[str(id)]["state"] = "readonly"
            
            mix_rates = []
            color_codes = []
            try:
                for i in detail.winfo_children():
                    mix_rates.append(float(i.children["rate"].get()) if str(widget) != str(i.children["rate"]).replace(" ", "") else float(data))
                    color_codes.append(i.children["code"].get() if str(widget) != str(i.children["code"]).replace(" ", "") else data)
            except:
                fail()
                return

            if "" in mix_rates or "" in color_codes:
                fail()
                return

            for idx, i in enumerate(detail.winfo_children()):
                result = self.get_toner(idx, mix_rates, color_codes)
                if result == False:
                    fail()
                    return

                for id in range(5):
                    i.children[str(id)]["state"] = "normal"
                    i.children[str(id)].delete(0, "end")
                    i.children[str(id)].insert(0, str(round(result[id], 2)))
                    i.children[str(id)]["state"] = "readonly"

        def num_filter(text, input, widget):
            try:
                if stop_filter.get() == True: return True
                if text == "":
                    calculation(widget, text)
                    return True
                if input in ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."] and text.count(".") < 2 and text.find(".") != 0:
                    float(text)

                    calculation(widget, text)
                else:
                    return False
            except:
                return False

            return True

        def delete(me):
            outer = me.master.master.master
            canvas = me.master.master
            label = outer.children["label_frame"].children["!label9"]
            detail = me.master

            me.destroy()
            for index, frame in enumerate(detail.winfo_children()):
                frame.children["!label"]["text"] = str(index + 1).zfill(2)
                frame.grid(row=index, column=0)
            
            calculation(detail, "", True)

            label["text"] = " 갯수: " + str(len(detail.winfo_children()))
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))

        def add(root):
            canvas = root.children["canvas"]
            label_frame = root.children["label_frame"]
            detail = canvas.children["detail"]

            num = len(detail.winfo_children())
            if num > 14: return

            frame = tk.Frame(detail, width=680, height=27)
            frame.grid(row=num, column=0)

            tk.Button(frame, text="X", takefocus=False, command=lambda: delete(frame)).place(x=5, rely=0.5, anchor="w", width=20, height=25)
            tk.Entry(frame, name="rate", highlightthickness=2, highlightcolor="SystemButtonFace", validate="all", validatecommand=(frame.register(num_filter), "%P", "%S", "%W")).place(x=30, rely=0.5, anchor="w", width=70, height=22)
            tk.Label(frame, text=str(num + 1).zfill(2)).place(x=105, rely=0.5, anchor="w", width=15, height=22)
            tk.Entry(frame, name="code", highlightthickness=2, highlightcolor="SystemButtonFace", validate="all", validatecommand=(frame.register(input_filter), "%W", "%P", "%d", "calculation", "%W")).place(x=125, rely=0.5, anchor="w", width=70, height=22)

            tk.Entry(frame, name="0", state="readonly", takefocus=False, bd=1, relief="sunken").place(x=200, rely=0.5, anchor="w", width=70, height=22)
            tk.Entry(frame, name="1", state="readonly", takefocus=False, bd=1, relief="sunken").place(x=275, rely=0.5, anchor="w", width=70, height=22)
            tk.Entry(frame, name="2", state="readonly", takefocus=False, bd=1, relief="sunken").place(x=350, rely=0.5, anchor="w", width=70, height=22)
            tk.Entry(frame, name="3", state="readonly", takefocus=False, bd=1, relief="sunken").place(x=425, rely=0.5, anchor="w", width=70, height=22)
            tk.Entry(frame, name="4", state="readonly", takefocus=False, bd=1, relief="sunken").place(x=500, rely=0.5, anchor="w", width=70, height=22)

            calculation(detail, "", True)

            label_frame.children["!label9"]["text"] = " 갯수: " + str(num + 1)

            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))

        def get_code(widget, text):
            index = text.find("-")

            one_frame.nametowidget(widget).delete(0, "end")
            one_frame.nametowidget(widget).insert(0, text if index == -1 else text[:index])

        #   COLORMIX
        for i in range(2):
            tk.Label(one_frame, text="COLORMIX_TBL " + ("3B" if i == 0 else "3P")).grid(row=0 if i == 0 else 2, column=0)
            inner_frame = tk.LabelFrame(one_frame, name="frame" + str(i))
            inner_frame.grid(row=1 if i == 0 else 3, column=0)
            self.generation_widget(18,
                            ("MixCd", "PaintTy", "MixCdGr", "MixLocation", "MixMethod", "MixGubun", "ColorGubun", "StdMixCd", "PearlMixCd", "Desc1", "ChiDesc", "ModifyHist", "ChiModifyHist", "Site", "MixDate", "Designer", "RegDate", "ModifyDate"),
                            (True, True, True, False, False, False, True, True, True, False, False, False, False, False, True, False, False, False),
                            (str_one if i == 0 else str_two, str_three, None, None, None, None, None, str_one if i == 0 else str_two, str_two if i == 0 else str_one, None, None, None, None, None, None, None, None, None),
                            ("", "", "", "A", "A", "A", "", "", "", "", "", "", "", "KEC", "", "wgk9985", "2023-02-28", "2023-02-28"),
                            (12, 8, 8, 10, 10, 10, 10, 10, 10, 6, 8, 10, 12, 6, 10, 10, 14, 14),
                            inner_frame,
                            (True, False, True, False, False, False, True, False, False, False, False, False, False, False, True, False, False, False),
                            1 if i == 0 else None)
            
            inner_frame.children["entry0"].configure(validate="all", validatecommand=(inner_frame.register(input_filter), "%W", "%P", "%d", "get_code", inner_frame.children["entry2"]))
            inner_frame.children["entry7"]["state"] = "readonly"
            inner_frame.children["entry8"]["state"] = "readonly"

        def edible_spec(widget, stop_spec):
            stop_spec.set(True if "selected" in widget.state() else False)
            detail = two_frame.nametowidget(f"outer{'0' if widget.winfo_name()[-1] == '1' else '1'}.canvas.detail")
            for p in detail.winfo_children():
                for i in range(5):
                    p.children[str(i)]["state"] = "normal" if stop_spec.get() == True else "readonly"

        #   COLORMIXDETAIL
        for i in range(2):
            tk.Label(two_frame, text="COLORMIXDETAIL_TBL " + ("3B" if i == 0 else "3P")).grid(row=0 if i == 0 else 2, column=0, pady=2)
            if data_index == None:
                check = ttk.Checkbutton(two_frame, name=f"checkspec{str(i + 1)}", takefocus=False, text="Spec 수정")
                check.grid(row=0 if i == 0 else 2, column=0, padx=(250, 0), pady=2)

            detail_gen(two_frame, "outer" + str(i), 1 if i == 0 else 3)

        if data_index == None:
            two_frame.children["checkspec1"]["command"] = lambda: edible_spec(two_frame.children["checkspec1"], stop_spec[0])
            two_frame.children["checkspec2"]["command"] = lambda: edible_spec(two_frame.children["checkspec2"], stop_spec[1])
            two_frame.children["checkspec1"].state(["!alternate"])
            two_frame.children["checkspec2"].state(["!alternate"])

        for p in (two_frame.children["outer0"], two_frame.children["outer1"]):
            for i in range(5):
                add(p)

        tk.Button(two_frame.children["outer0"].children["label_frame"], text="추가", takefocus=False, command=lambda: add(two_frame.children["outer0"])).place(x=650, rely=0.5, anchor="w")
        tk.Button(two_frame.children["outer1"].children["label_frame"], text="추가", takefocus=False, command=lambda: add(two_frame.children["outer1"])).place(x=650, rely=0.5, anchor="w")

        #   COLORMIXSTD
        for i in range(2):
            tk.Label(three_frame, text="COLORMIXSTD_TBL " + ("3B" if i == 0 else "3P"), width=55).grid(row=0, column=i)
            inner_frame = tk.LabelFrame(three_frame, name="frame" + str(i), padx=2)
            inner_frame.grid(row=1, column=i)
            self.generation_widget(4,
                            ("MixCd", "paintTy", "ColorMixStd", "StdDesc"),
                            (True, True, True, True),
                            (str_one if i == 0 else str_two, str_three, None, None),
                            ("", "", "", ""),
                            (12, 8, 10, 20),
                            inner_frame,
                            (False, False, False, True),
                            2)
            
            self.generation_combo(2, None, ("000", "001", "002", "003", "004", "005", "006", "007", "008", "009", "010", "011", "012"), 7, inner_frame)

            inner_frame.children["entry0"]["state"] = "readonly"

        #   MIXCOLORINFO
        for i in range(2):
            tk.Label(four_frame, text="MIXCOLORINFO_TBL " + ("3B" if i == 0 else "3P")).grid(row=0 if i == 0 else 2, column=0)
            inner_frame = tk.LabelFrame(four_frame, name="frame" + str(i), padx=2)
            inner_frame.grid(row=1 if i == 0 else 3, column=0)
            self.generation_widget(13,
                            ("MixCd", "ColorNm", "EngColorNm", "ChiColorNm", "ColorCode", "EngColorCd", "CarNm", "EngCarNm", "ChiCarNm", "ApplyYear", "CBSFC", "MixInfo", "PaintTy"),
                            (True, True, False, False, True, False, True, False, False, True, False, True, True),
                            (str_one if i == 0 else str_two, None, None, None, None, None, None, None, None, None, None, None, str_three),
                            ("", "", "", "", "", "", "", "", "", "", "", "", ""),
                            (12, 20, 10, 10, 8, 10, 8, 8, 8, 8, 8, 8, 8),
                            inner_frame,
                            (False, True, False, False, True, False, True, False, False, True, False, True, False))
            
            inner_frame.children["entry4"].configure(validate="all", validatecommand=(inner_frame.register(input_filter), "%W", "%P", "%d"))
            inner_frame.children["entry0"]["state"] = "readonly"
        
        #   크롤링 부분
        inner_frame = tk.LabelFrame(two_frame, text="크롤링 [상태:       ]", labelanchor="n", width=200, height=228)
        inner_frame.grid(row=0, column=1, rowspan=4, padx=(10, 0), sticky="w")

        noru_frame = tk.Frame(inner_frame, width=190, height=60)
        noru_frame.place(relx=0.5, rely=0.3, anchor="center")
        tk.Label(noru_frame, text="3B", width=5).place(relx=0.05, rely=0.35, anchor="w")
        entry_3b = tk.Entry(noru_frame, name="3b", width=15, takefocus=False)
        entry_3b.place(relx=0.3, rely=0.3, anchor="w")
        tk.Label(noru_frame, text="3P", width=5).place(relx=0.05, rely=0.65, anchor="w")
        entry_3p = tk.Entry(noru_frame, name="3p", width=15, takefocus=False)
        entry_3p.place(relx=0.3, rely=0.7, anchor="w")

        noru_frame2 = tk.Frame(inner_frame, width=190, height=60)
        noru_frame2.place(relx=0.5, rely=0.3, anchor="center")
        tk.Label(noru_frame2, text="검색어", width=5).place(relx=0.05, rely=0.5, anchor="w")
        entry_noru = tk.Entry(noru_frame2, name="entry", width=15, takefocus=False)
        entry_noru.place(relx=0.3, rely=0.5, anchor="w")

        kcc_frame = tk.Frame(inner_frame, width=190, height=60)
        tk.Label(kcc_frame, text="검색어", width=5).place(relx=0.05, rely=0.5, anchor="w")
        entry_kcc = tk.Entry(kcc_frame, name="entry", width=15, takefocus=False)
        entry_kcc.place(relx=0.3, rely=0.5, anchor="w")

        login_frame = tk.Frame(inner_frame, width=190, height=75)
        login_frame.place(relx=0.5, rely=0.58, anchor="center")
        tk.Label(login_frame, text="페인트타입: ").place(relx=0.08, rely=0.28, anchor="w")
        type_entry = tk.Entry(login_frame, takefocus=False, width=10)
        type_entry.place(relx=0.46, rely=0.28, anchor="w")
        type_entry.insert(0, "BC")
        tk.Label(login_frame, text="아이디: ").place(relx=0.08, rely=0.58, anchor="w")
        id_entry = tk.Entry(login_frame, takefocus=False, width=10)
        id_entry.place(relx=0.46, rely=0.58, anchor="w")
        tk.Label(login_frame, text="비밀번호: ").place(relx=0.08, rely=0.88, anchor="w")
        pwd_entry = tk.Entry(login_frame, takefocus=False, width=10)
        pwd_entry.place(relx=0.46, rely=0.88, anchor="w")

        def crawling_entry_switch():
            if company.get() == 0:
                if isnormal.get() == 0:
                    noru_frame.place_forget()
                    noru_frame2.place(relx=0.5, rely=0.3, anchor="center")
                else:
                    noru_frame2.place_forget()
                    noru_frame.place(relx=0.5, rely=0.3, anchor="center")
                kcc_frame.place_forget()
            else:
                kcc_frame.place(relx=0.5, rely=0.3, anchor="center")
                noru_frame.place_forget()
                noru_frame2.place_forget()

        tk.Radiobutton(inner_frame, text="노루", value=0, variable=company, command=crawling_entry_switch, takefocus=False).place(relx=0.15, rely=0.075, anchor="center")
        tk.Radiobutton(inner_frame, text="KCC", value=1, variable=company, command=crawling_entry_switch, takefocus=False).place(relx=0.4, rely=0.075, anchor="center")
        tk.Label(inner_frame, text="유지").place(relx=0.615, rely=0.075, anchor="center")
        colorgubun_check = ttk.Checkbutton(inner_frame, takefocus=False)
        colorgubun_check.place(relx=0.75, rely=0.075, anchor="center")
        colorgubun_check.state(["!alternate"])
        colorgubun = tk.Entry(inner_frame, takefocus=False)
        colorgubun.place(relx=0.87, rely=0.075, width=30, anchor="center")

        def get_site_data():
            alert_frame = tk.Frame(data_frame, bd=4, relief="solid")
            alert_frame.place(relx=0.5, rely=0.5, anchor="center", width=900, height=300)

            alert_font = tkfont.nametofont("TkDefaultFont").copy()
            alert_font.config(size=30)

            tk.Label(alert_frame, text="잠시만 기다려주세요.\n(프로그램이 일시적으로 멈춥니다.)", font=alert_font).place(relx=0.5, rely=0.5, anchor="center")
            alert_frame.lift()
            self.window.update_idletasks()

            inner_frame["text"] = "크롤링 [상태: 성공]"

            try:
                _ID = id_entry.get()
                _PWD = pwd_entry.get()

                if _ID == "" or _PWD == "":
                    inner_frame["text"] = "크롤링 [상태: 실패]"
                    tkbox.showerror("실패", "아이디 혹은 패스워드를 확인하십시오.")
                    alert_frame.destroy()
                    return
                
                result = Crawling().open(_ID, _PWD, ([entry_3b.get(), entry_3p.get()] if isnormal.get() == 1 else [entry_noru.get()]) if company.get() == 0 else entry_kcc.get(),"noru" if company.get() == 0 else "kcc", bool(isnormal.get()), type_entry.get(), True)
            except:
                inner_frame["text"] = "크롤링 [상태: 실패]"
                tkbox.showerror("실패", "데이터를 가져오지 못했습니다.")
                alert_frame.destroy()
                return
            
            if result == False:
                inner_frame["text"] = "크롤링 [상태: 실패]"
                tkbox.showerror("실패", "아이디 혹은 패스워드를 확인하십시오.")
                alert_frame.destroy()
                return

            data = result[0]
            detail_data_3B = result[1]
            detail_data_3P = result[2]

            if result[3]:
                isnormal.set(1)
                show_normal_cott()

            # insert data
            clear(False)
            str_three.set(type_entry.get())

            stop_filter.set(True)

            for i in range(2 if isnormal.get() == 1 else 1):
                #ColorNm 3B, 3P
                four_frame.children["frame" + str(i)].children["entry1"].insert(0, data["ColorNm(3B)"] if i == 0 else data["ColorNm(3P)"])
                #ColorCode
                four_frame.children["frame" + str(i)].children["entry4"].insert(0, data["ColorCode"])
                #ColorGubun
                if "selected" in colorgubun_check.state():
                    one_frame.children["frame" + str(i)].children["entry6"].insert(0, colorgubun.get())
                #MixDate
                one_frame.children["frame" + str(i)].children["entry14"].insert(0, data["MixDate"])
                #StdDesc
                three_frame.children["frame" + str(i)].children["entry3"].insert(0, data["StdDesc"])
                #CarNm
                four_frame.children["frame" + str(i)].children["entry6"].insert(0, data["CarNm"])
                #ApplyYear
                four_frame.children["frame" + str(i)].children["entry9"].insert(0, data["ApplyYear"])
                #MixInfo
                four_frame.children["frame" + str(i)].children["entry11"].insert(0, data["MixInfo"])
                
                if company.get() == 0:
                    #MixCd
                    one_frame.children["frame" + str(i)].children["entry0"].insert(0, data["MixCd(3B)"] if i == 0 else data["MixCd(3P)"])
                    #MixCdGr
                    get_code(one_frame.children["frame" + str(i)].children["entry2"], data["MixCd(3B)"] if i == 0 else data["MixCd(3P)"])
                    #ColorMixStd
                    try:
                        three_frame.children["frame" + str(i)].children["combobox2"].set(colormixstd_match[data["ColorMixStd"].replace(" ", "")])
                    except:
                        inner_frame["text"] = "크롤링 [상테: StdDesc 없음.]"

                for p in range(len(detail_data_3B if i == 0 else detail_data_3P)):
                    add(two_frame.children["outer" + str(i)])

                for detail, frame in zip(detail_data_3B if i == 0 else detail_data_3P, two_frame.children["outer" + str(i)].children["canvas"].children["detail"].winfo_children()):
                    frame.children["rate"].insert(0, detail[1])
                    frame.children["code"].insert(0, detail[0])

                calculation(two_frame.children["outer0"].children["canvas"].children["detail"], "", True)
                if i == 1:
                    calculation(two_frame.children["outer1"].children["canvas"].children["detail"], "", True)
            
            stop_filter.set(False)
            alert_frame.destroy()

        tk.Button(inner_frame, text="가져오기", takefocus=False, command=get_site_data, width=10).place(relx=0.25, rely=0.9, anchor="center")
        tk.Button(inner_frame, text="열기", takefocus=False, command=lambda: Crawling().open(id_entry.get(), pwd_entry.get(), ([entry_3b.get(), entry_3p.get()] if isnormal.get() == 1 else [entry_noru.get()]) if company.get() == 0 else entry_kcc.get(),"noru" if company.get() == 0 else "kcc", bool(isnormal.get()), type_entry.get()), width=10).place(relx=0.75, rely=0.9, anchor="center")

        def show_normal_cott():
            if isnormal.get() == 0:
                one_frame.children["!label2"].grid_remove()
                one_frame.children["frame1"].grid_remove()

                one_frame.children["frame0"].children["entry6"]["takefocus"] = True
                one_frame.children["frame0"].children["entry6"]["state"] = "normal"
                one_frame.children["frame1"].children["entry6"]["state"] = "normal"
                one_frame.children["frame0"].children["entry6"].delete(0, "end")
                one_frame.children["frame1"].children["entry6"].delete(0, "end")

                two_frame.children["!label2"].grid_remove()
                if data_index == None:
                    two_frame.children["checkspec2"].grid_remove()
                two_frame.children["outer1"].grid_remove()

                three_frame.children["!label2"].grid_remove()
                three_frame.children["frame1"].grid_remove()

                four_frame.children["!label2"].grid_remove()
                four_frame.children["frame1"].grid_remove()

                noru_frame2.place(relx=0.5, rely=0.3, anchor="center")
                noru_frame.place_forget()

                three_frame.grid_columnconfigure(1, weight=0)
            else:
                one_frame.children["!label2"].grid()
                one_frame.children["frame1"].grid()

                one_frame.children["frame0"].children["entry6"]["takefocus"] = False
                one_frame.children["frame0"].children["entry6"].delete(0, "end")
                one_frame.children["frame0"].children["entry6"].insert(0, "3B")
                one_frame.children["frame0"].children["entry6"]["state"] = "readonly"

                one_frame.children["frame1"].children["entry6"]["takefocus"] = False
                one_frame.children["frame1"].children["entry6"].delete(0, "end")
                one_frame.children["frame1"].children["entry6"].insert(0, "3P")
                one_frame.children["frame1"].children["entry6"]["state"] = "readonly"

                two_frame.children["!label2"].grid()
                if data_index == None:
                    two_frame.children["checkspec2"].grid()
                two_frame.children["outer1"].grid()

                three_frame.children["!label2"].grid()
                three_frame.children["frame1"].grid()

                four_frame.children["!label2"].grid()
                four_frame.children["frame1"].grid()

                noru_frame.place(relx=0.5, rely=0.3, anchor="center")
                noru_frame2.place_forget()

                three_frame.grid_columnconfigure(1, weight=1)

        #   일반, 3코트
        tk.Radiobutton(data_frame, name="normal", text="일반", takefocus=False, variable=isnormal, value=0, command=show_normal_cott).place(x=5, y=5)
        tk.Radiobutton(data_frame, name="cott", text="3코트", takefocus=False, variable=isnormal, value=1, command=show_normal_cott).place(x=55, y=5)

        self.generation_combo(1, str_three, ("BC", "UU"), 5, one_frame.children["frame0"])

        def show():
            if main_frame.winfo_class() == "Frame":
                if self.window.winfo_width() == 950:
                    data_frame.place(width=1550)
                    self.window.geometry("1550x940")
                else:
                    data_frame.place(width=950)
                    self.window.geometry("950x940")
            else:
                if main_frame.winfo_width() == 950:
                    data_frame.place(width=1550)
                    main_frame.geometry("1550x940")
                else:
                    data_frame.place(width=950)
                    main_frame.geometry("950x940")

            for i in (one_frame, three_frame, four_frame):
                for o in i.winfo_children():
                    if o.winfo_class() == "Labelframe":
                        for p in o.winfo_children():
                            if "h" in p.winfo_name():
                                if isshow.get() == 1:
                                    p.grid_remove()
                                else:
                                    p.grid()
            
            isshow.set(0 if isshow.get() == 1 else 1)
        
        def clear(isall):
            isok = tkbox.askquestion("경고", "입력하신 데이터를 지우시겠습니까?")
            if isok == "no":
                return False

            if data_index != None:
                main_frame.lift()

            for p in (one_frame.children["frame0"], one_frame.children["frame1"]):
                for i in p.winfo_children():
                    if i.winfo_class() == "Entry":
                        if isall and "h" in i.winfo_name():
                            i.delete(0, "end")
                        elif "h" not in i.winfo_name():
                            i.delete(0, "end")

            for p in (two_frame.children["outer0"].children["canvas"].children["detail"], two_frame.children["outer1"].children["canvas"].children["detail"]):
                for i in p.winfo_children():
                    i.destroy()
                
            two_frame.children["outer0"].children["label_frame"].children["!label9"]["text"] = " 갯수: 0"
            two_frame.children["outer1"].children["label_frame"].children["!label9"]["text"] = " 갯수: 0"
            
            for p in (three_frame.children["frame0"], three_frame.children["frame1"]):
                for i in p.winfo_children():
                    if i.winfo_class() == "Entry":
                        i.delete(0, "end")
            
            for p in (four_frame.children["frame0"], four_frame.children["frame1"]):
                for i in p.winfo_children():
                    if i.winfo_class() == "Entry":
                        if isall and "h" in i.winfo_name():
                            i.delete(0, "end")
                        elif "h" not in i.winfo_name():
                            i.delete(0, "end")
        
        def reset():
            if clear(True) == False: return

            if data_index != None:
                result_set()
                return

            for p in (one_frame.children["frame0"], one_frame.children["frame1"]):
                for i in p.winfo_children():
                    if i.winfo_name() in "entry3h" or i.winfo_name() in "entry4h" or i.winfo_name() in "entry5h":
                        i.insert(0, "A")
                    elif i.winfo_name() == "entry13h":
                        i.insert(0, "KEC")
                    elif i.winfo_name() == "entry15h":
                        i.insert(0, "wgk9985")
                    elif i.winfo_name() == "entry16h" or i.winfo_name() == "entry17h":
                        i.insert(0, "2023-02-28")

            for p in (two_frame.children["outer0"], two_frame.children["outer1"]):
                for i in range(5):
                    add(p)
            
            if isnormal.get() == 1:
                one_frame.children["frame0"].children["entry6"].insert(0, "3B")
                one_frame.children["frame1"].children["entry6"].insert(0, "3P")

            str_three.set("BC")
        
        def force_edit():
            isok = tkbox.askquestion("경고", "일부 기능이 작동하지 않습니다.\n강제 수정을 끌 경우, 일부 데이터가 훼손될 수 있습니다.")
            if isok == "no":
                return False
            
            main_frame.lift()

            if stop_filter.get() == False:
                stop_filter.set(True)

                for i in (one_frame.children["frame0"].winfo_children(), one_frame.children["frame1"].winfo_children()):
                    for p in i:
                        p["state"] = "normal"

                for i in (two_frame.nametowidget("outer0.canvas.detail").winfo_children(), two_frame.nametowidget("outer1.canvas.detail").winfo_children()):
                    for p in i:
                        for o in p.winfo_children():
                            o["state"] = "normal"

                for i in (three_frame.children["frame0"].winfo_children(), three_frame.children["frame1"].winfo_children()):
                    for p in i:
                        p["state"] = "normal"

                for i in (four_frame.children["frame0"].winfo_children(), four_frame.children["frame1"].winfo_children()):
                    for p in i:
                        p["state"] = "normal"

                data_frame.children["default"]["state"] = "disabled"
                data_frame.children["force"]["text"] = "강제수정 (켜짐)"
                data_frame.children["force"]["bg"] = "#A7F8B0"
            else:
                stop_filter.set(False)

                one_frame.children["frame0"].children["entry7"]["state"] = "readonly"
                one_frame.children["frame0"].children["entry8"]["state"] = "readonly"
                one_frame.children["frame1"].children["entry7"]["state"] = "readonly"
                one_frame.children["frame1"].children["entry8"]["state"] = "readonly"

                for i in (two_frame.nametowidget("outer0.canvas.detail").winfo_children(), two_frame.nametowidget("outer1.canvas.detail").winfo_children()):
                    for p in i:
                        for o in p.winfo_children():
                            if o.winfo_class() == "Entry" and o.winfo_name().isdigit():
                                o["state"] = "readonly"

                three_frame.children["frame0"].children["entry0"]["state"] = "readonly"
                three_frame.children["frame1"].children["entry0"]["state"] = "readonly"
                four_frame.children["frame0"].children["entry0"]["state"] = "readonly"
                four_frame.children["frame1"].children["entry0"]["state"] = "readonly"

                data_frame.children["default"]["state"] = "normal"
                data_frame.children["force"]["text"] = "강제수정 (꺼짐)"
                data_frame.children["force"]["bg"] = "#F88787"

        def ready_for_save():
            input = [[], [], [], []]

            for frame in (one_frame.children["frame0"], one_frame.children["frame1"]):
                if not frame.winfo_ismapped(): continue
                temp_input = []

                for widget in frame.winfo_children():
                    if widget.winfo_class() != "Label":
                        if widget.winfo_class() == "TCombobox":
                            temp_input.insert(1, widget.get())
                        else:
                            temp_input.append(widget.get())
                
                input[0].append(temp_input)
            
            for index, frame in enumerate((two_frame.children["outer0"], two_frame.children["outer1"])):
                if not frame.winfo_ismapped(): continue
                detail = frame.children["canvas"].children["detail"]
                temp_input = []

                for inner_frame in detail.winfo_children():
                    temp_detail = [input[0][index][0], input[0][index][1]]

                    for widget in inner_frame.winfo_children():
                        if widget.winfo_class() == "Label":
                            temp_detail.append(widget["text"])
                        elif widget.winfo_class() == "Entry":
                            temp_detail.append(widget.get())

                    temp_detail.append("2023-02-28 오전 9:00:00")

                    temp_input.append(temp_detail)

                input[1].append(temp_input)

            for frame in (three_frame.children["frame0"], three_frame.children["frame1"]):
                if not frame.winfo_ismapped(): continue
                temp_input = []

                for widget in frame.winfo_children():
                    if widget.winfo_class() != "Label":
                        if widget.winfo_class() == "TCombobox":
                            temp_input.insert(2, widget.get())
                        else:
                            temp_input.append(widget.get())
                
                input[2].append(temp_input)

            for frame in (four_frame.children["frame0"], four_frame.children["frame1"]):
                if not frame.winfo_ismapped(): continue
                temp_input = []

                for widget in frame.winfo_children():
                    if widget.winfo_class() != "Label":
                        temp_input.append(widget.get())
                
                input[3].append(temp_input)
            
            if data_index == None:
                self.Save(input)
            else:
                self.Update("data", input, data_index, main_frame)

        def result_set(is_bind = False):
            data = self.datas["data"][data_index]
            
            for frame_idx in range(2 if len(data) > 7 else 1):
                widget1 = [widget for widget in one_frame.children[f"frame{str(frame_idx)}"].winfo_children() if widget.winfo_class() != "Label"]
                widget3 = [widget for widget in three_frame.children[f"frame{str(frame_idx)}"].winfo_children() if widget.winfo_class() != "Label"]
                widget4 = [widget for widget in four_frame.children[f"frame{str(frame_idx)}"].winfo_children() if widget.winfo_class() != "Label"]

                if frame_idx == 0:
                    temp_combo = widget1[-1]
                    del widget1[-1]
                    widget1.insert(1, temp_combo)

                temp_combo = widget3[-1]
                del widget3[-1]
                widget3.insert(2, temp_combo)

                # Insert value
                if (len(data[0]) if frame_idx == 0 else len(data[7][0])) > 0:
                    stop_filter.set(True)

                    for value, widget in zip(data[0][0] if frame_idx == 0 else data[7][0][0], widget1):
                        if widget.winfo_class() == "Entry":
                            widget.delete(0, "end")
                            widget.insert(0, value.strip())
                        else:
                            widget.set(value.strip())

                    stop_filter.set(False)
                else:
                    if len(data[0]) <= 0:
                        one_frame.grid_remove()
                
                for i in two_frame.nametowidget(f"outer{str(frame_idx)}.canvas.detail").winfo_children():
                    i.destroy()
                
                if (len(data[1]) if frame_idx == 0 else len(data[7][1])) > 0:
                    stop_filter.set(True)

                    for row in (data[1] if frame_idx == 0 else data[7][1]):
                        add(two_frame.children[f"outer{frame_idx}"])
                        two_frame.nametowidget(f"outer{str(frame_idx)}.canvas.detail").winfo_children()[-1].children["rate"].insert(0, row[4].strip())
                        two_frame.nametowidget(f"outer{str(frame_idx)}.canvas.detail").winfo_children()[-1].children["code"].insert(0, row[3].strip())

                    stop_filter.set(False)
                    calculation(two_frame.nametowidget(f"outer{str(frame_idx)}.canvas.detail"), "", True)
                else:
                    if len(data[1]) <= 0:
                        two_frame.grid_remove()

                if (len(data[2]) if frame_idx == 0 else len(data[7][2])) > 0:
                    stop_filter.set(True)

                    for value, widget in zip(data[2][0] if frame_idx == 0 else data[7][2][0], widget3):
                        if widget.winfo_class() == "Entry":
                            before_state = widget["state"]
                            widget["state"] = "normal"
                            widget.delete(0, "end")
                            widget.insert(0, value.strip())
                            widget["state"] = before_state
                        else:
                            widget.set(value.strip())
                    
                    stop_filter.set(False)
                else:
                    if len(data[2]) <= 0:
                        three_frame.grid_remove()

                if (len(data[3]) if frame_idx == 0 else len(data[7][3])) > 0:
                    stop_filter.set(True)

                    for value, widget in zip(data[3][0] if frame_idx == 0 else data[7][3][0], widget4):
                        if widget.winfo_class() == "Entry":
                            before_state = widget["state"]
                            widget["state"] = "normal"
                            widget.delete(0, "end")
                            widget.insert(0, value.strip())
                            widget["state"] = before_state
                        else:
                            widget.set(value.strip())
                    
                    stop_filter.set(False)
                else:
                    if len(data[3]) <= 0:
                        four_frame.grid_remove()

                if is_bind:
                    def change_alert(widget, original_data):
                        if widget.get() == original_data:
                            widget.config(highlightbackground = "SystemButtonFace", highlightcolor= "SystemButtonFace")
                        else:
                            widget.config(highlightbackground = "red", highlightcolor= "red")

                    def change_alert_label(widget, original_data, label):
                        if widget.get() == original_data:
                            label.config(highlightthickness=2, highlightbackground = "SystemButtonFace", highlightcolor= "SystemButtonFace")
                        else:
                            label.config(highlightthickness=2, highlightbackground = "red", highlightcolor= "red")

                    def ready_to_binding(variable, widget, value, label=None):
                        widget.configure(textvariable=variable)
                        if widget.winfo_class() == "Entry":
                            variable.trace_add("write", lambda a, b, c: widget.after(1, lambda: change_alert(widget, value)))
                        else:
                            variable.trace_add("write", lambda a, b, c: widget.after(1, lambda: change_alert_label(widget, value, label)))

                    trace_data1 = (data[0][0] if len(data[0]) > 0 else None) if frame_idx == 0 else (data[7][0][0] if len(data[7][0]) > 0 else None)
                    trace_data3 = (data[2][0] if len(data[2]) > 0 else None) if frame_idx == 0 else (data[7][2][0] if len(data[7][2]) > 0 else None)
                    trace_data4 = (data[3][0] if len(data[3]) > 0 else None) if frame_idx == 0 else (data[7][3][0] if len(data[7][3]) > 0 else None)

                    for pack in ([widget1, trace_data1, trace_str1_0 if frame_idx == 0 else trace_str1_1], [widget3, trace_data3, trace_str3_0 if frame_idx == 0 else trace_str3_1], [widget4, trace_data4, trace_str4_0 if frame_idx == 0 else trace_str4_1]):
                        if pack[1] == None:
                            pack[1] = [widget.get() for widget in pack[0]]
                        
                        for widget, value, trace_var in zip(pack[0], pack[1], pack[2]):
                            if widget.winfo_class() != "TCombobox":
                                if widget["textvariable"] != "":
                                    for var in (str_one, str_two, str_three):
                                        if str(var) == widget["textvariable"]:
                                            ready_to_binding(var, widget, value.strip())
                                            break
                                else:
                                    trace_var.set(value=value.strip())
                                    ready_to_binding(trace_var, widget, value.strip())
                            else:
                                if widget["textvariable"] == "":
                                    trace_var.set(value=value.strip())
                                ready_to_binding(trace_var if widget["textvariable"] == "" else str_three, widget, value.strip(), widget.master.children[f"label{str(widget)[-1]}"])
                    
                    for index, row, var in zip(range(len(trace_str2_0 if frame_idx == 0 else trace_str2_1)), two_frame.nametowidget(f"outer{str(frame_idx)}.canvas.detail").winfo_children(), trace_str2_0 if frame_idx == 0 else trace_str2_1):
                        var[0].set(data[1][index][4].strip() if frame_idx == 0 else data[7][1][index][4].strip())
                        var[1].set(data[1][index][3].strip() if frame_idx == 0 else data[7][1][index][3].strip())
                        ready_to_binding(var[0], row.children["rate"], data[1][index][4].strip() if frame_idx == 0 else data[7][1][index][4].strip())
                        ready_to_binding(var[1], row.children["code"], data[1][index][3].strip() if frame_idx == 0 else data[7][1][index][3].strip())

        def color_toplevel():
            window = tk.Toplevel(self.window)
            window.geometry("300x390")
            window.resizable(False, False)
            window.title("내역 검색")
            
            tk.Label(window, text="내역 검색 결과").place(relx=0.5, y=10, anchor="n")

            frame = tk.Frame(window, name="outer0", bd=1, relief="sunken")
            frame.place(relx=0.5, y=40, width=280, height=300, anchor="n")

            label_frame = tk.Frame(frame, bd=1, relief="raised")
            label_frame.place(x=0, y=0, width=278, height=30)

            canvas = tk.Canvas(frame, name="canvas", highlightthickness=0)
            canvas.place(x=0, y=30, width=258, height=270)

            scrollbar = tk.Scrollbar(frame, name="scrollbar")
            scrollbar.place(x=258, y=30, width=20, height=270)

            scrollbar.configure(command=canvas.yview)
            canvas.configure(yscrollcommand=scrollbar.set)

            tk.Button(label_frame, text="컬러코드").place(x=5, rely=0.5, width=80, anchor="w")
            tk.Button(label_frame, text="배합코드").place(x=95, rely=0.5, width=77, anchor="w")
            tk.Button(label_frame, text="다원화").place(x=182, rely=0.5, width=70, anchor="w")

            tk.Entry(window, name="entry").place(x=40, y=350, width=150, height=25)
            tk.Button(window, text="검색하기", command=lambda: self.search_for_color(window.children["entry"].get(), canvas)).place(x=200, y=350)

            window.children["entry"].bind("<Return>", lambda e: self.search_for_color(window.children["entry"].get(), canvas))
            window.bind("<MouseWheel>", lambda e: self.scroll_bind(e.widget, e.delta))

            window.mainloop()

        data_frame.grid_columnconfigure(0, weight=1)
        one_frame.grid_columnconfigure(0, weight=1)
        two_frame.grid_columnconfigure(0, weight=1)
        two_frame.grid_columnconfigure(1, weight=1)
        three_frame.grid_columnconfigure(0, weight=1)
        three_frame.grid_columnconfigure(1, weight=1)
        four_frame.grid_columnconfigure(0, weight=1)

        one_frame.grid(row=0, column=0, pady=10, sticky="ew")
        two_frame.grid(row=1, column=0, pady=10, sticky="ew")
        three_frame.grid(row=2, column=0, pady=10, sticky="ew")
        four_frame.grid(row=3, column=0, pady=10, sticky="ew")

        show_normal_cott()
        if data_index != None:
            result_set(True)

            if len(self.datas["data"][data_index]) > 7:
                isnormal.set(1)
                show_normal_cott()
            
            data_frame.children["normal"].place_forget()
            data_frame.children["cott"].place_forget()

        tk.Button(data_frame, text="숨겨진 값 표시/숨기기", takefocus=False, command=show).place(relx=0.05, rely=0.96)
        tk.Button(data_frame, text="지우기", takefocus=False, command=lambda: clear(False)).place(relx=0.25, rely=0.96)
        tk.Button(data_frame, text="지우기 (고정값)", takefocus=False, command=lambda: clear(True)).place(relx=0.3125, rely=0.96)
        tk.Button(data_frame, name="default", text="기본값으로 초기화", takefocus=False, command=reset).place(relx=0.425, rely=0.96)
        save_button = tk.Button(data_frame, text="저장하기", command=ready_for_save)
        save_button.place(relx=0.75, rely=0.96)
        save_button.bind("<Return>", lambda e: ready_for_save())
        tk.Button(data_frame, name="force", text="강제수정 (꺼짐)", bg="#F88787", takefocus=False, command=force_edit).place(relx=0.635, rely=0.96)
        if data_index == None:
            tk.Button(data_frame, text="내역보기", takefocus=False, command=color_toplevel).place(relx=0.565, rely=0.96)
        tk.Button(data_frame, text="뒤로가기", takefocus=False, command=lambda: self.change_frames(main_frame, data_frame, 400, 400)).place(relx=0.85, rely=0.96)

    def setting(self):
        self.edit_function = [self.code_setting, self.data_setting]
        self.bind_function = self.scroll_bind
        main_frame = tk.Frame(self.window, name="main_frame")
        code_frame = tk.Frame(self.window, name="code_frame", bd=1, relief="sunken")
        data_frame = tk.Frame(self.window, name="data_frame", bd=1, relief="sunken")
        record_frame = tk.Frame(self.window, name="record_frame")
        patch_frame = tk.Frame(self.window, name="patch_frame")

        with open(f"{os.getcwd()}\\version.txt", "r") as file:
            MAIN_VERSION = file.read()

        tk.Label(main_frame, font=("TkDefaultFont", 12), text=f"현재 버전: {MAIN_VERSION}").place(relx=0.5, rely=0.9, anchor="n")

        def record_setting():
            code_sort = [False, False, False, False, False]
            data_sort = [False, False, False, False, False, False]

            one_frame = tk.Frame(record_frame)
            one_frame.place(x=0, y=0, width=1000, height=470)

            preview_frame = tk.Frame(one_frame, name="preview")
            preview_frame.place(x=10, y=30, width=980, height=260, anchor="nw")

            state_frame = tk.Frame(one_frame, bd=1, relief="raised")
            state_frame.place(x=10, y=325, width=980, height=30, anchor="nw")

            tk.Label(state_frame, name="state", text="상태: ").place(relx=0.5, rely=0.5, anchor="center")

            loading_bar = ttk.Progressbar(one_frame)
            loading_bar.place(x=10, y=360, width=980, height=15, anchor="nw")

            menu_frame = tk.Frame(one_frame, bd=1, relief="sunken")
            menu_frame.place(x=10, y=390, width=980, height=70, anchor="nw")
            
            menu_one = tk.Frame(menu_frame)
            menu_one.place(relx=0.5, rely=0.5, width=978, height=68, anchor="center")

            except_frame = tk.LabelFrame(record_frame, text="예외 처리")
            except_frame.place(x=10, y=480, width=980, height=510, anchor="nw")

            page_frame = tk.Frame(one_frame, bd=1, relief="sunken")
            page_frame.place(x=470, y=295, width=520, height=25)

            def check_all(canvas, isall):
                for row in canvas.winfo_children():
                    row.children["checkbutton"].state(["selected" if isall else "!selected"])

            # Code Preview
            tk.Label(one_frame, text="코드 테이블 검색 결과", anchor="center").place(x=10, y=10, width=450, height=20)

            code_preview_frame = tk.Frame(preview_frame, name="outer0", bd=1, relief="sunken")
            code_preview_frame.place(x=0, y=0, width=450, height=260)

            label_frame = tk.Frame(code_preview_frame, bd=1, relief="raised")
            label_frame.place(x=0, y=0, width=448, height=30)

            def order_code(num):
                if num == 0:
                    self.datas["code"].sort(key=lambda x: x[3], reverse=code_sort[num])
                elif num == 1:
                    self.datas["code"].sort(key=lambda x: x[2][0][1] if len(x[2]) > 0 else "", reverse=code_sort[num])
                elif num == 2:
                    self.datas["code"].sort(key=lambda x: x[0][0][1] if len(x[0]) > 0 else "", reverse=code_sort[num])
                elif num == 3:
                    self.datas["code"].sort(key=lambda x: x[4] if len(x) > 4 else "", reverse=code_sort[num])
                else:
                    self.datas["code"].sort(key=lambda x: bool(x[5]) if len(x) > 4 else False, reverse=code_sort[num])
                
                code_sort[num] = False if code_sort[num] else True
                canvas = code_preview_frame.children["canvas"]

                canvas.delete("result")

                for i in range(len(self.datas["code"])):
                    self.Add_result(canvas, "r_V2.accdb" if self.datas["code"][i][3] == 0 else "new_auto.mdb", "code", self.datas["code"][i], i, i)

                canvas.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))

            tk.Button(label_frame, text="파일명", command=lambda: order_code(0)).place(x=28, rely=0.5, width=83, anchor="w")
            tk.Button(label_frame, text="그룹코드", command=lambda: order_code(1)).place(x=116, rely=0.5, width=60, anchor="w")
            tk.Button(label_frame, text="Minorcd", command=lambda: order_code(2)).place(x=181, rely=0.5, width=67, anchor="w")
            tk.Button(label_frame, text="작성시간", command=lambda: order_code(3)).place(x=253, rely=0.5, width=70, anchor="w")
            tk.Button(label_frame, text="휴지통", command=lambda: order_code(4)).place(x=328, rely=0.5, width=50, anchor="w")

            for i in label_frame.winfo_children():
                self.useble_widgets.append(i)

            canvas = tk.Canvas(code_preview_frame, name="canvas", highlightthickness=0)
            canvas.place(x=0, y=30, width=430, height=230)

            scrollbar = tk.Scrollbar(code_preview_frame, name="scrollbar")
            scrollbar.place(x=430, y=30, width=20, height=230)

            scrollbar.configure(command=canvas.yview)
            canvas.configure(yscrollcommand=scrollbar.set)

            code_check_all = ttk.Checkbutton(one_frame, text="전체 선택", takefocus=False)
            code_check_all.config(command=lambda: check_all(code_preview_frame.children["canvas"], "selected" in code_check_all.state()))
            code_check_all.place(x=10, y=10, height=20)
            code_check_all.state(["!alternate"])

            # Data Preview
            tk.Label(one_frame, text="데이터 테이블 검색 결과", anchor="center").place(x=470, y=10, width=520, height=20)

            data_preview_frame = tk.Frame(preview_frame, name="outer1", bd=1, relief="sunken")
            data_preview_frame.place(x=460, y=0, width=520, height=260)

            label_frame = tk.Frame(data_preview_frame, bd=1, relief="raised")
            label_frame.place(x=0, y=0, width=518, height=30)

            def order_data(num):
                if num == 0:
                    self.datas["data"].sort(key=lambda x: x[4], reverse=data_sort[num])
                elif num == 1:
                    self.datas["data"].sort(key=lambda x: x[0][0][0] if len(x[0]) > 0 else (x[1][0][0] if len(x[1]) > 0 else (x[2][0][0] if len(x[2]) > 0 else (x[3][0][0] if len(x[3]) > 0 else ""))), reverse=data_sort[num])
                elif num == 2:
                    self.datas["data"].sort(key=lambda x: x[0][0][1] if len(x[0]) > 0 else (x[1][0][2] if len(x[1]) > 0 else (x[2][0][1] if len(x[2]) > 0 else (x[3][0][12] if len(x[3]) > 0 else ""))), reverse=data_sort[num])
                elif num == 3:
                    self.datas["data"].sort(key=lambda x: x[3][0][4] if len(x[3]) > 0 else "", reverse=data_sort[num])
                elif num == 4:
                    self.datas["data"].sort(key=lambda x: x[5] if len(x) > 5 else "", reverse=data_sort[num])
                else:
                    self.datas["data"].sort(key=lambda x: bool(x[6]) if len(x) > 5 else False, reverse=data_sort[num])

                data_sort[num] = False if data_sort[num] else True

                change_page(int(self.window.nametowidget(".record_frame.!frame.!frame3.now")["text"]))

            tk.Button(label_frame, text="파일명", command=lambda: order_data(0)).place(x=29, rely=0.5, width=83, anchor="w")
            tk.Button(label_frame, text="배합코드", command=lambda: order_data(1)).place(x=117, rely=0.5, width=77, anchor="w")
            tk.Button(label_frame, text="타입", command=lambda: order_data(2)).place(x=199, rely=0.5, width=35, anchor="w")
            tk.Button(label_frame, text="컬러코드", command=lambda: order_data(3)).place(x=239, rely=0.5, width=80, anchor="w")
            tk.Button(label_frame, text="작성시간", command=lambda: order_data(4)).place(x=324, rely=0.5, width=70, anchor="w")
            tk.Button(label_frame, text="휴지통", command=lambda: order_data(5)).place(x=399, rely=0.5, width=50, anchor="w")

            for i in label_frame.winfo_children():
                self.useble_widgets.append(i)

            canvas = tk.Canvas(data_preview_frame, name="canvas", highlightthickness=0)
            canvas.place(x=0, y=30, width=500, height=230)

            scrollbar = tk.Scrollbar(data_preview_frame, name="scrollbar")
            scrollbar.place(x=500, y=30, width=20, height=230)

            scrollbar.configure(command=canvas.yview)
            canvas.configure(yscrollcommand=scrollbar.set)

            data_check_all = ttk.Checkbutton(one_frame, text="전체 선택", takefocus=False)
            data_check_all.config(command=lambda: check_all(data_preview_frame.children["canvas"], "selected" in data_check_all.state()))
            data_check_all.place(x=470, y=10, height=20)
            data_check_all.state(["!alternate"])

            # Data Preview page
            def change_page(page):
                if page == None:
                    self.window.nametowidget(".record_frame.!frame.!frame3.entry").insert(0, self.window.nametowidget(".record_frame.!frame.!frame3.now")["text"])
                    return

                max = self.window.nametowidget(".record_frame.!frame.!frame3.max")["text"]
                view_count = int(self.window.nametowidget(".record_frame.!frame.!frame3.!combobox").get())
                datas = self.datas["data"][view_count * (page - 1):view_count * page if page != int(max) else len(self.datas["data"])]

                data_preview_frame.children["canvas"].delete("result")
                if page == int(max):
                    self.window.nametowidget(".record_frame.!frame.!frame3.right")["state"] = "disabled"
                    self.window.nametowidget(".record_frame.!frame.!frame3.right_end")["state"] = "disabled"
                elif page == 1:
                    self.window.nametowidget(".record_frame.!frame.!frame3.left")["state"] = "disabled"
                    self.window.nametowidget(".record_frame.!frame.!frame3.left_end")["state"] = "disabled"

                for i in range(len(datas)):
                    self.Add_result(data_preview_frame.children["canvas"], "r_V2.accdb" if datas[i][4] == 0 else "new_auto.mdb", "data", datas[i], i, view_count * (page - 1) + i)
                
                data_canvas = self.window.nametowidget(".record_frame.!frame.preview.outer1.canvas")
                data_canvas.update_idletasks()
                data_canvas.configure(scrollregion=data_canvas.bbox("all"))

            def left(isend):
                now = self.window.nametowidget(".record_frame.!frame.!frame3.now")
                self.window.nametowidget(".record_frame.!frame.!frame3.right")["state"] = "normal"
                self.window.nametowidget(".record_frame.!frame.!frame3.right_end")["state"] = "normal"
                if isend:
                    now["text"] = "1"
                else:
                    now["text"] = str(int(now["text"]) - 1)
                self.window.nametowidget(".record_frame.!frame.!frame3.entry").delete(0, "end")
                self.window.nametowidget(".record_frame.!frame.!frame3.entry").insert(0, now["text"])
                change_page(int(now["text"]))

            def right(isend):
                now = self.window.nametowidget(".record_frame.!frame.!frame3.now")
                max = self.window.nametowidget(".record_frame.!frame.!frame3.max")["text"]
                self.window.nametowidget(".record_frame.!frame.!frame3.left")["state"] = "normal"
                self.window.nametowidget(".record_frame.!frame.!frame3.left_end")["state"] = "normal"
                if isend:
                    now["text"] = max
                else:
                    now["text"] = str(int(now["text"]) + 1)
                self.window.nametowidget(".record_frame.!frame.!frame3.entry").delete(0, "end")
                self.window.nametowidget(".record_frame.!frame.!frame3.entry").insert(0, now["text"])
                change_page(int(now["text"]))

            def change_combo(selected):
                if len(self.datas["data"]) <= 0:
                    return
                
                data_preview_frame.children["canvas"].delete("result")
                self.window.nametowidget(".record_frame.!frame.!frame3.now")["text"] = "1"
                self.window.nametowidget(".record_frame.!frame.!frame3.entry").delete(0, "end")
                self.window.nametowidget(".record_frame.!frame.!frame3.entry").insert(0, "1")
                self.window.nametowidget(".record_frame.!frame.!frame3.max")["text"] = str(math.ceil(len(self.datas["data"]) / int(selected)))
                self.window.nametowidget(".record_frame.!frame.!frame3.left")["state"] = "disabled"
                self.window.nametowidget(".record_frame.!frame.!frame3.left_end")["state"] = "disabled"
                self.window.nametowidget(".record_frame.!frame.!frame3.right")["state"] = "normal"
                self.window.nametowidget(".record_frame.!frame.!frame3.right_end")["state"] = "normal"
                change_page(1)

            def num_filter(text, input):
                if text == "" or input in ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"] or text.isdecimal():
                    return True
                return False

            tk.Button(page_frame, name="left_end", state="disabled", text="<<", command=lambda: left(True), takefocus=False, width=3).place(relx=0.1, y=0, height=24, anchor="n")
            tk.Button(page_frame, name="left", state="disabled", text="<", command=lambda: left(False), takefocus=False, width=3).place(relx=0.2, y=0, height=24, anchor="n")
            tk.Entry(page_frame, name="entry", validate="all", validatecommand=(page_frame.register(num_filter), "%P", "%S"), state="disabled", takefocus=False).place(relx=0.325, y=0, width=35, height=22, anchor="n")
            tk.Button(page_frame, text="이동", state="disabled", takefocus=False, command=lambda: change_page(int(self.window.nametowidget(".record_frame.!frame.!frame3.entry").get()) if self.window.nametowidget(".record_frame.!frame.!frame3.entry").get() != "" else None)).place(relx=0.4, y=0, width=35, height=22, anchor="n")
            tk.Label(page_frame, name="now", text="1", anchor="e").place(relx=0.47, y=0, width=25, height=22, anchor="n")
            tk.Label(page_frame, text="/").place(relx=0.5, y=0, width=5, height=22, anchor="n")
            tk.Label(page_frame, name="max", text="1", anchor="w").place(relx=0.5325, y=0, width=25, height=22, anchor="n")
            tk.Button(page_frame, name="right", state="disabled", text=">", command=lambda: right(False), takefocus=False, width=3).place(relx=0.6, y=0, height=24, anchor="n")
            tk.Button(page_frame, name="right_end", state="disabled", text=">>", command=lambda: right(True), takefocus=False, width=3).place(relx=0.7, y=0, height=24, anchor="n")
            tk.Label(page_frame, text="표시 갯수: ").place(relx=0.825, y=0, height=24, anchor="n")

            combo = ttk.Combobox(page_frame, values=("5", "10", "25", "30", "50", "100", "150", "200"), state="readonly", takefocus=False, width=3)
            combo.place(relx=0.93, y=0, height=24, anchor="n")
            combo.set("50")
            combo.bind("<<ComboboxSelected>>", lambda e: change_combo(combo.get()))

            for i in page_frame.winfo_children():
                self.useble_widgets.append(i)

            # Options Frame
            self.search_options["database"] = [tk.IntVar(value=1), tk.IntVar(value=1)]
            tk.Checkbutton(menu_one, text="r_V2.accdb", variable=self.search_options["database"][0]).place(x=10, rely=0.75, anchor="w")
            tk.Checkbutton(menu_one, text="new_auto.mdb", variable=self.search_options["database"][1]).place(x=110, rely=0.75, anchor="w")

            self.search_options["including_text"] = tk.IntVar(value=0)
            tk.Checkbutton(menu_one, text="검색시 입력한 단어 무조건 포함 (AND)", variable=self.search_options["including_text"]).place(x=230, rely=0.75, anchor="w")

            # except
            def checkbutton_select(type, all):
                if type == "code":
                    one = code_except_frame.children["checkbutton " + self.tables[0]]
                    two = code_except_frame.children["checkbutton " + self.tables[1]]
                    three = code_except_frame.children["checkbutton " + self.tables[2]]

                    if all:
                        if is_except_code.get() == 0:
                            one.state(["!selected"])
                            two.state(["!selected"])
                            three.state(["!selected"])
                        else:
                            one.state(["selected"])
                            two.state(["selected"])
                            three.state(["selected"])
                        return
                    
                    if "selected" in one.state() and "selected" in two.state() and "selected" in three.state():
                        is_except_code.set(1)
                    else:
                        is_except_code.set(0)
                else:
                    one = data_except_frame.children["checkbutton " + self.tables[3]]
                    two = data_except_frame.children["checkbutton " + self.tables[4]]
                    three = data_except_frame.children["checkbutton " + self.tables[5]]
                    four = data_except_frame.children["checkbutton " + self.tables[6]]

                    if all:
                        if is_except_data.get() == 0:
                            one.state(["!selected"])
                            two.state(["!selected"])
                            three.state(["!selected"])
                            four.state(["!selected"])
                        else:
                            one.state(["selected"])
                            two.state(["selected"])
                            three.state(["selected"])
                            four.state(["selected"])
                        return
                    
                    if "selected" in one.state() and "selected" in two.state() and "selected" in three.state() and "selected" in four.state():
                        is_except_data.set(1)
                    else:
                        is_except_data.set(0)

            self.search_options["except_texts"] = tk.StringVar()
            tk.Label(except_frame, text="제외단어").place(relx=0.01, y=20, anchor="w")
            tk.Entry(except_frame, name="except_text", bd=2, textvariable=self.search_options["except_texts"]).place(relx=0.07, y=20, width=300, height=25, anchor="w")

            self.useble_widgets.append(except_frame.children["except_text"])

            code_except_frame = tk.LabelFrame(except_frame, text=" Code ( 전체제외:      )")
            code_except_frame.place(x=10, y=40, width=440, height=390)
            
            is_except_code = tk.IntVar(value=0)
            all_check = tk.Checkbutton(except_frame, variable=is_except_code, command=lambda: checkbutton_select("code", True))
            all_check.place(x=118, y=37, width=17)

            self.useble_widgets.append(all_check)

            data_except_frame = tk.LabelFrame(except_frame, text=" Data ( 전체제외:      )")
            data_except_frame.place(x=460, y=0, width=510, height=480)
            
            is_except_data = tk.IntVar(value=0)
            all_check = tk.Checkbutton(except_frame, variable=is_except_data, command=lambda: checkbutton_select("data", True))
            all_check.place(x=564, y=-3, width=17)

            self.useble_widgets.append(all_check)

            column_names = [("Majorcd", "Minorcd", "Codenm", "Remark1", "Remark2", "Sortorder", "RegDate"),
                            ("MakerId", "Maker_K", "Maker_E", "Maker_C", "UseYn", "RegId", "RegDate", "MakerNew_C"),
                            ("idx", "Group_Code", "Maker_ENG", "Brand", "Site", "Origin", "RegId", "RegDate"),
                            ("MixCd", "PaintTy", "MixCdGr", "MixLocation", "MixMethod", "MixGubun", "ColorGubun", "StdMixCd", "PearlMixCd", "Desc1", "ChiDesc", "ModifyHist", "ChiModifyHist", "Site", "MixDate", "Designer", "RegDate", "ModifyDate"),
                            ("MixCd", "Nbr", "PaintTy", "TonerCd", "MixRate", "Spec05", "Spec08", "Spec10", "Spec20", "Spec40", "RegDate"),
                            ("MixCd", "paintTy", "ColorMixStd", "StdDesc"),
                            ("MixCd", "ColorNm", "EngColorNm", "ChiColorNm", "ColorCode", "EngColorCd", "CarNm", "EngCarNm", "ChiCarNm", "ApplyYear", "CBSFC", "MixInfo", "PaintTy")]

            check_places = [(120, 18),
                            (120, 133),
                            (120, 248),
                            (150, 14),
                            (150, 161),
                            (150, 271),
                            (150, 343),]

            # except: code columns
            for index in range(3):
                inner_frame = tk.LabelFrame(code_except_frame, name=str(index), text=self.tables[index] + "       ")
                inner_frame.grid(row=index, column=0, padx=(10, 0), pady=(20, 0), sticky="w")

                table_check = ttk.Checkbutton(code_except_frame, name="checkbutton " + self.tables[index], takefocus=False, command=lambda: checkbutton_select("code", False))
                table_check.place(x=check_places[index][0], y=check_places[index][1])
                code_except_frame.children["checkbutton " + self.tables[index]].state(["!alternate"])

                self.useble_widgets.append(table_check)

                for idx in range(len(column_names[index])):
                    tk.Label(inner_frame, name="l"+str(idx), text=column_names[index][idx], font=("TkDefaultFont", 8), width=9).grid(row=0 + int(idx / 6) * 2, column=idx % 6)
                    check = ttk.Checkbutton(inner_frame, name="c"+str(idx), takefocus=False)
                    check.grid(row=1 + int(idx / 6) * 2, column=idx % 6)
                    check.state(["!alternate"])
                    
                    self.useble_widgets.append(check)

            # except: data columns
            for index in range(3, 7):
                inner_frame = tk.LabelFrame(data_except_frame, name=str(index), text=self.tables[index] + ("               " if index == 3 else "       "))
                inner_frame.grid(row=index, column=0, padx=(10, 0), pady=(15, 0), sticky="w")

                table_check = ttk.Checkbutton(data_except_frame, name="checkbutton " + self.tables[index], takefocus=False, command=lambda: checkbutton_select("data", False))
                table_check.place(x=check_places[index][0], y=check_places[index][1])
                data_except_frame.children["checkbutton " + self.tables[index]].state(["!alternate"])

                self.useble_widgets.append(table_check)

                for idx in range(len(column_names[index])):
                    tk.Label(inner_frame, name="l"+str(idx), text=column_names[index][idx], font=("TkDefaultFont", 8), width=9).grid(row=0 + int(idx / 7) * 2, column=idx % 7)
                    check = ttk.Checkbutton(inner_frame, name="c"+str(idx), takefocus=False)
                    check.grid(row=1 + int(idx / 7) * 2, column=idx % 7)
                    check.state(["!alternate"])

                    self.useble_widgets.append(check)

            def switch_except_frame():
                if self.window.winfo_height() == 470:
                    record_frame.place(x=0, y=0, width=1000, height=1000)
                    self.window.geometry("1000x1000")
                else:
                    record_frame.place(x=0, y=0, width=1000, height=470)
                    self.window.geometry("1000x470")

            def ready_for_search():
                # Lock button, entry
                for i in self.useble_widgets:
                    i["state"] = "disabled"

                self.window.nametowidget(".record_frame.!frame.!progressbar")["value"] = 0

                code_preview_frame.children["canvas"].delete("result")
                data_preview_frame.children["canvas"].delete("result")
                self.datas = {"code":[], "data":[]}

                # Keyword
                keyword = menu_one.children["search_entry"].get()
                if keyword != "":
                    keywords = []
                    while keyword.find(",") > -1:
                        keywords.append(keyword[:keyword.find(",")].strip())
                        keyword = keyword[keyword.find(",") + 1:]
                    keywords.append(keyword.strip())

                    while keywords.count("") > 0:
                        del keywords[keywords.index("")]
                else:
                    keywords = ""

                # Select database
                select_databases = []

                if self.search_options["database"][0].get() == 1:
                    self.Open("r_V2.accdb", "")
                    select_databases.append(0)
                
                if self.search_options["database"][1].get() == 1:
                    self.Open("new_auto.mdb", "EOGKSALSRNRSHFNVPDLS")
                    select_databases.append(1)

                # Get except tables
                self.search_options["except_tables"] = []

                for index, table in enumerate(self.tables):
                    if index < 3:
                        if "selected" in code_except_frame.children["checkbutton " + table].state():
                            self.search_options["except_tables"].append(table)
                    else:
                        if "selected" in data_except_frame.children["checkbutton " + table].state():
                            self.search_options["except_tables"].append(table)
                
                if len(self.search_options["except_tables"]) <= 0:
                    self.search_options["except_tables"] = None
                
                # Get except columns
                self.search_options["except_columns"] = {}

                for index in range(7):
                    if index < 3:
                        self.search_options["except_columns"][self.tables[index]] = []
                        for check_idx, check_widget in enumerate(code_except_frame.children[str(index)].winfo_children()):
                            if check_widget.winfo_name() == "c"+str(int(check_idx / 2)) and "selected" in check_widget.state():
                                self.search_options["except_columns"][self.tables[index]].append(code_except_frame.children[str(index)].children["l"+str(int(check_idx / 2))]["text"])
                    else:
                        self.search_options["except_columns"][self.tables[index]] = []
                        for check_idx, check_widget in enumerate(data_except_frame.children[str(index)].winfo_children()):
                            if check_widget.winfo_name() == "c"+str(int(check_idx / 2)) and "selected" in check_widget.state():
                                self.search_options["except_columns"][self.tables[index]].append(data_except_frame.children[str(index)].children["l"+str(int(check_idx / 2))]["text"])

                # Get except texts
                except_keyword = except_frame.children["except_text"].get()
                if except_keyword != "":
                    except_keywords = []
                    while except_keyword.find(",") > -1:
                        except_keywords.append(except_keyword[:except_keyword.find(",")].strip())
                        except_keyword = except_keyword[except_keyword.find(",") + 1:]
                    except_keywords.append(except_keyword.strip())

                    while except_keywords.count("") > 0:
                        del except_keywords[except_keywords.index("")]

                    self.search_options["except_texts"] = except_keywords
                else:
                    self.search_options["except_texts"] = None

                Thread(target=self.Search_V2, daemon=True, args=(select_databases, keywords, bool(self.search_options["including_text"].get()), self.search_options["except_tables"], self.search_options["except_columns"], self.search_options["except_texts"])).start()

            tk.Label(menu_one, text="통합검색").place(relx=0.01, rely=0.3, anchor="w")
            tk.Entry(menu_one, name="search_entry", bd=2, takefocus=False).place(relx=0.07, rely=0.3, width=300, height=25, anchor="w")
            tk.Button(menu_one, text="검색하기", takefocus=False, command=lambda: ready_for_search()).place(relx=0.38, rely=0.3, anchor="w")
            tk.Button(menu_one, text="검색 예외", takefocus=False, command=switch_except_frame).place(relx=0.58, rely=0.3, anchor="w")
            tk.Button(menu_one, text="삭제하기", command=self.Delete, takefocus=False).place(relx=0.74, rely=0.3, anchor="w")
            tk.Button(menu_one, text="휴지통", command=self.Trash, takefocus=False).place(relx=0.82, rely=0.3, anchor="w")
            tk.Button(menu_one, text="뒤로가기", takefocus=False, command=lambda: self.change_frames(main_frame, record_frame, 400, 400)).place(relx=0.925, rely=0.3, anchor="w")

            menu_one.children["search_entry"].bind("<Return>", lambda e: ready_for_search())

            for i in menu_one.winfo_children():
                self.useble_widgets.append(i)

        def patch_setting():
            canvas = tk.Canvas(patch_frame, name="canvas", highlightthickness=0, bd=1, relief="sunken")
            canvas.place(x=10, y=10, width=360, height=340)

            scrollbar = tk.Scrollbar(patch_frame, name="scrollbar")
            scrollbar.place(x=370, y=10, width=20, height=340)

            scrollbar.configure(command=canvas.yview)
            canvas.configure(yscrollcommand=scrollbar.set)

            context = """V0.51(2024.06.20)
        + 기록조회시 3코드를 불러올때 멈추는 버그 수정

V0.5(2024.06.10)
        + 가져오는 데이터가 3코트인지 체크하는 기능추가
        + 이제 비중이 new_auto.mdb를 기준으로 함.
        + 크롤링 상태 표시 일부 수정.
        + ColorGubun 입력칸 추가.

V0.43(2024.04.29)
        + 모든 입력칸에 직접 입력이 가능하게 수정.

V0.42(2024.04.25)
        + Spec칸에 직접 입력이 가능하게 수정.
        + 비중에 'CLEAR','HARDER'를 추가함.
        + PaintTy에 따라 데이터를 가져오는 형식이
        달라짐.

V0.41(2024.04.11)
        + 데이터 가져올 시 ColorMixStd 매치가 되지않아
        발생한 프리징 현상을 해결함.
        + 업데이트 코드를 좀 더 좋게 변경함.

V0.4(2024.03.15):
        + 수정사항 해결됐음.

V0.31 (2024.03.09):
        + 문제가 있던 업데이트 런처를 없애고
        단일 실행 파일로 합쳐졌음.
        
V0.3 (2024.03.08):
        + 기록 검색 결과 정렬 기능 추가.
        + 기타 버그 및 에러 수정.

V0.2 (2024.03.07): 
        + 자동 업데이트 기능 추가.
        + new_auto.mdb 데이터 삭제가 안되던 오류 픽스.

V0.1 (2024.03.03):
        최초배포버전."""
            text = tk.Label(canvas, font=("TkDefaultFont", 10), text=context, width=43, justify="left", anchor="nw", wraplength=345)
            canvas.create_window(5, 5, anchor="nw", window=text)

            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))

            tk.Button(patch_frame, text="뒤로가기", takefocus=False, command=lambda: self.change_frames(main_frame, patch_frame, 400, 400)).place(relx=0.825, rely=0.91)

        self.code_setting(code_frame, main_frame)
        self.data_setting(data_frame, main_frame)
        record_setting()
        patch_setting()

        # Setting main screen

        main_font = tkfont.nametofont("TkDefaultFont").copy()
        main_font.config(size=18)

        tk.Button(main_frame, text="코드 입력", font=main_font, bd=3, width=12, command=lambda: self.change_frames(code_frame, main_frame, 750, 380)).place(relx=0.5, rely=0.25, anchor="center")
        tk.Button(main_frame, text="데이터 입력", font=main_font, bd=3, width=12, command=lambda: self.change_frames(data_frame, main_frame, 950, 940)).place(relx=0.5, rely=0.5, anchor="center")
        tk.Button(main_frame, text="기록 조회", font=main_font, bd=3, width=12, command=lambda: self.change_frames(record_frame, main_frame, 1000, 470)).place(relx=0.5, rely=0.75, anchor="center")
        tk.Button(main_frame, font=("TkDefaultFont", 10), text="패치 내역", command=lambda: self.change_frames(patch_frame, main_frame, 400, 400)).place(relx=0.875, rely=0.9, anchor="n")

        main_frame.place(x=0, y=0, width=400, height=400)

        self.window.bind("<MouseWheel>", lambda e: self.scroll_bind(e.widget, e.delta))

    def loading_bar_processing(self):
        bar = self.window.nametowidget(".record_frame.!frame.!progressbar")
        bar.step(1)

    def change_frames(self, show_frame, hide_frame, width, height):
        if show_frame.winfo_class() == "Frame":
            self.window.geometry(str(width) + "x" + str(height))
            show_frame.place(x=0, y=0, width=width, height=height)
            hide_frame.place_forget()
        else:
            show_frame.destroy()

    def scroll_bind(self, widget, delta):
        if "outer" in str(widget):
            scrollbar = self.window.nametowidget(str(widget)[:str(widget).find("outer") + 6] + ".scrollbar")
            canvas = self.window.nametowidget(str(widget)[:str(widget).find("outer") + 6] + ".canvas")

        try:
            if scrollbar.get()[0] == 0.0 and scrollbar.get()[1] == 1.0: return
            else: canvas.yview_scroll(-1 * int(delta/120), "units")
        except: return

if __name__ == "__main__":
    IS_SOURCES = True if os.path.splitext(__file__)[1] == ".py" else False

    if IS_SOURCES:
        UI()
    else:
        if len(sys.argv) < 2:
            VERSION = ""
            if os.path.exists(f"{os.getcwd()}\\version.txt"):
                with open(f"{os.getcwd()}\\version.txt", "r") as file:
                    VERSION = file.read()
            else:
                VERSION = "first_run"

            OWNER = "NMN-NMN"
            REPO = "Noru"
            API_URL = f"https://api.github.com/repos/{OWNER}/{REPO}"
            API_KEY = "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"

            result = requests.get(f"{API_URL}/releases/latest")
            if result.status_code != 200:
                tkbox.showerror("실패", "깃허브 서버 접속에 실패했습니다.")
                sys.exit(0)

            result = result.json()
            NEW_VERSION = result["tag_name"]

            if NEW_VERSION != VERSION:
                shutil.copyfile("NoruV2.exe", "Updater.exe")

                with open(f"{os.getcwd()}\\version.txt", "w") as file:
                    file.write(result["tag_name"])

                download_url = result["assets"][0]["url"]
                subprocess.Popen([f"{os.getcwd()}\\Updater.exe", "update", f"{download_url}"])
            else:
                UI()
        elif sys.argv[1] == "update":
            download = requests.get(sys.argv[2], headers={'Accept': 'application/octet-stream'}, stream=True)

            if download.status_code == 200:
                file_name = f"{os.getcwd()}\\NoruV2.exe"

                with open(file_name, "wb") as file:
                    for chunk in download.iter_content(chunk_size=8192*1024):
                        file.write(chunk)
            else:
                tkbox.showerror("실패", "다운로드에 실패했습니다.")
                sys.exit(0)
            
            subprocess.Popen([f"{os.getcwd()}\\NoruV2.exe", "finished"])
        elif sys.argv[1] == "finished":
            os.remove(f"{os.getcwd()}\\Updater.exe")
            UI()