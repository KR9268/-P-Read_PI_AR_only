import tabula
import re
import numpy as np
import pandas as pd
import xlwings as xw
import os
import time
import json
from tkinter import messagebox
import logging
import traceback
import os
from pathlib import Path

#로그저장 세팅 DEBUG, INFO, WARNING, ERROR, WARNING
logging.basicConfig(filename='./Read_PI_AR_only_log' + str(time.time()) + '.log', level=logging.ERROR)

# PI글자수 조정용 함수
def shorten_pi(text):
    if len(text)>30:
        return text.replace('-','')
    else:
        return text
    
try:
    #기본사항 세팅
    with open('./Read_PI_AR_only_defaultOption.json', 'r', encoding='UTF-8')as f:
        path_all = json.load(f, strict=False)
    with open('./Read_PI_AR_only_changePort.json', 'r', encoding='UTF-8')as f:
        change_port = json.load(f, strict=False)

    #파일리스트 읽기
    file_path = str(Path(path_all['file_path'])) + '\\'
    excel_path = str(Path(path_all['excel_path'])) + '\\'
    output_path = str(Path(path_all['output_path'])) + '\\'
    #pdf파일리스트
    file_list = [file for file in os.listdir(file_path) if file.endswith('pdf')]
    #엑셀파일리스트
    excel_list = [file for file in os.listdir(excel_path) if file.endswith('xlsx')]
    excel_list = [file for file in excel_list if not file.startswith('~$')]
    excel_list = sorted(excel_list,reverse=True)

    #엑셀열기
    wb = xw.Book(excel_path + excel_list[0])
    sht = wb.sheets[0]
    range_df = sht.range('A2:BL'+str(sht.used_range.last_cell.row))

    excel_table = range_df.options(pd.DataFrame).value

    #pdf읽기
    base_df = pd.DataFrame()
    error_pi = []
    for file in file_list:
        try:
            #임시저장용 dict 초기화
            each_pi = {}
            #pdf 읽고 페이지의 파트 분배
            pdf_text = tabula.read_pdf(file_path + file, pages="all", encoding='CP949')
            pdf_page_number = len(pdf_text)
            pdf_text[0] = pdf_text[0].replace(np.nan,'')
            part1 = pdf_text[0]['BUYER']
            part2 = pdf_text[0]['P/I NUMBER']

            #pdf내용 정리
            crit_POL = part2[part2.isin(['PORT OF LOADING'])].index+1
            crit_POD = part2[part2.isin(['PORT OF DESTINATION'])].index+1
            crit_opendate = part2[part2.str.contains('P/I DATE')].index+1
            crit_Partial = part1.str.contains('PARTIAL|Partial')
            crit_Tranship = part1.str.contains('TRANS-SHIPMENT|Trans Shipment')

            each_pi['PI_NO'] = part2[0]
            each_pi['POL1'] = change_port[part2[crit_POL].item().upper()][:2]
            each_pi['POL2'] = change_port[part2[crit_POL].item().upper()][2:]
            each_pi['POD1'] = change_port[part2[crit_POD].item().upper()][:2]
            each_pi['POD2'] = change_port[part2[crit_POD].item().upper()][2:]

            each_pi['OPEN_DATE'] = (re.findall('\d{4}-\d{2}-\d{2}',part2[crit_opendate].item()))[0]
            each_pi['SHIPMENT_DATE'] = ''
            each_pi['EXPIRY_DATE'] = ''
            each_pi['Partial'] = 'O' if ('NOT ALLOWED') in part1[crit_Partial].item().upper() else ''
            each_pi['Tranship'] = 'O' if ('NOT ALLOWED') in part1[crit_Tranship].item().upper() else ''

            #엑셀내용 정리
            each_pi['CUR'] = excel_table[excel_table['P/I No.']==each_pi['PI_NO']]['Unit'].item()
            each_pi['AMOUNT'] = excel_table[excel_table['P/I No.']==each_pi['PI_NO']]['Amount(CIP)'].item()
            each_pi['TERM'] = excel_table[excel_table['P/I No.']==each_pi['PI_NO']]['Payment(GERP)'].item()
            each_pi['INCO'] = excel_table[excel_table['P/I No.']==each_pi['PI_NO']]['Incoterms'].item()[:3]
            each_pi['INCO_TEXT'] = excel_table[excel_table['P/I No.']==each_pi['PI_NO']]['Incoterms'].item()[3:]

            base_df = pd.concat([base_df, pd.DataFrame(each_pi, index=[0])])
        except:
            error_pi.append(file)            

    # 레이아웃 조절 후 엑셀저장
    base_df2 = pd.DataFrame()
    if len(base_df)>0:
        base_df2 = base_df[['PI_NO','POL1','POL2','POD1','POD2','CUR','AMOUNT','TERM','INCO','INCO_TEXT','OPEN_DATE','SHIPMENT_DATE','EXPIRY_DATE','Partial','Tranship']]
        base_df2['OPEN_DATE'] = base_df2['OPEN_DATE'].str.replace('-','.')
        base_df2['PI_NO_Shorten'] = base_df2['PI_NO_Shorten'] = base_df2['PI_NO'].apply(shorten_pi)
        base_df2 = base_df2[['PI_NO','PI_NO_Shorten','POL1','POL2','POD1','POD2','CUR','AMOUNT','TERM','INCO','INCO_TEXT','OPEN_DATE','SHIPMENT_DATE','EXPIRY_DATE','Partial','Tranship']]

        base_df2.to_excel(output_path + 'list_pi' + str(time.time()) + '.xlsx')
    else:
        #with open(output_path + 'list_pi_추출된 건이 없습니다' + str(time.time()) + '.xlsx', 'w') as f:
        pass
    # 에러건 txt저장
    if len(error_pi)>0:
        with open(output_path + 'list_error' + str(time.time()) + '.txt', 'w') as f:
            for pi_name in error_pi:
                f.write(pi_name+'\n')
    else:
        #with open(output_path + '에러건이 없습니다' + str(time.time()) + '.txt', 'w') as f:
        pass

    msg_result = '{0}건 추출, {1}건 오류로 제외되었습니다'.format(len(base_df2),len(error_pi))
    messagebox.showinfo('알림', msg_result)

except: #실행오류시 오류내역 저장
    logging.error(traceback.format_exc())
