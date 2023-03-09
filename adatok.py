import streamlit as st
import pandas as pd
import numpy as np
import xlsxwriter
import requests
import json
import requests
import urllib3
from io import BytesIO
import copy

def select_city(city_url, city):

    import requests
    import json
    import pandas as pd

    #print("Írd be a várost!")
    #city = input().strip()
    code = ""

    print("Várj....")

    try:
      req = requests.get(city_url)
      #print(req.text)
      city_dict = json.loads(req.text)["list"]
      df = pd.DataFrame(city_dict, columns=['megnev', 'maz', 'taz'])
      nev = df.query('megnev==@city')
      maz = nev["maz"].tolist()[0]
      taz = nev["taz"].tolist()[0]

      link = "https://vtr.valasztas.hu/ogy2022/data/04161400/szavossz/"+maz+"/SzavkorJkv-"+maz+"-"+taz+".json"
      print(link)
    
    
      return link, maz, taz

    except Exception as e:
      #print(e)
      print("Rossz város!")  
      return 0, 0, 0

@st.cache
def ogy(eredmeny_url, maz, taz):

    print("OGY...")

    import requests
    import json
    import pandas as pd

    city_url = "https://vtr.valasztas.hu/ogy2022/data/04022333/ver/Telepulesek.json"
    jeloltek_url = "https://vtr.valasztas.hu/ogy2022/data/04022333/ver/EgyeniJeloltek.json"
    partok_url = "https://vtr.valasztas.hu/ogy2022/data/04022333/ver/Szervezetek.json"

    partok = {
        6: 'Fidesz',
        1: 'ellenzék',
        5: 'Mi Hazánk',
        3: 'MKKP',
        4: 'MEMO',
        2: 'Normális Párt'
    }

    req = requests.get(eredmeny_url)
    eredmeny_dict = json.loads(req.text)

    req = requests.get(jeloltek_url)
    jeloltek_dict = json.loads(req.text)["list"]

    df = pd.DataFrame(jeloltek_dict, columns=['neve', 'jlcs_nev', 'ej_id'])

    jeloltek = {}

    eredmenyek = {}
    for ered in range(len(eredmeny_dict["list"])):
        szk_no = int(eredmeny_dict["list"][ered]["sorsz"])
        lista_ered = eredmeny_dict["list"][ered]["listas_jkv"]["tetelek"]
        eredmenyek[szk_no] = {} 
        eredmenyek[szk_no]["Szavazókör"] = szk_no
        egyeni_ered = eredmeny_dict["list"][ered]["egyeni_jkv"]["tetelek"]
        for szk in egyeni_ered:
            id = int(szk['ej_id'])
            szavazat = szk['szavazat']
            if id not in jeloltek:
                jeloltek = build_jeloltek(jeloltek_dict, jeloltek, szk)  
            jelolt = jeloltek[int(id)]       
            eredmenyek[szk_no][jelolt] = szavazat
        for szk in lista_ered:
            if 'sorsz' in szk:
                no = szk['sorsz']
                szavazat = szk['szavazat']
                part = partok[int(no)]          
                eredmenyek[szk_no][part] = szavazat
                

    pd.DataFrame(eredmenyek).transpose()

    return(eredmenyek)

def build_jeloltek(jeloltek_dict, jeloltek, ered):
    
    import pandas as pd

    df = pd.DataFrame(jeloltek_dict, columns=['neve', 'jlcs_nev', 'ej_id'])

    id = ered['ej_id']
    nev = df.query('ej_id==@id')
    jeloltek[id] = nev["neve"].tolist()[0] + " - " + nev["jlcs_nev"].tolist()[0]
    
    return jeloltek

def get_winner():
    evk_dict = data_polgi[1]
    for evk in evk_dict:
        osszes = {}
        if data_polgi[2] != {}:
          for szk in evk_dict[evk]['szavazokorok']:
            for nev in data_polgi[2][szk].keys():
              if nev not in osszes:
                  osszes[nev] = 0
              osszes[nev] += int(data_polgi[2][szk][nev])
          max = 0
          max_nev = ''
          for elem in osszes:
              if max < osszes[elem]:
                max = osszes[elem]
                max_nev = elem
          evk_dict[evk]["győztes"] = max_nev
        else:
          evk_dict[evk]["győztes"] = "nincsenek egyéni körzetek"
    return evk_dict

@st.cache
def ep(eredmeny_url, maz, taz):
    from bs4 import BeautifulSoup
    import pandas as pd
    import requests 
    import time
    
    print("EP...")

    data = {}

    eu_url = "https://www.valasztas.hu/szavazokorok_ep2019?_epszavazokorok_WAR_nvinvrportlet_formDate=32503680000000&p_p_id=epszavazokorok_WAR_nvinvrportlet&p_p_lifecycle=1&p_p_state=maximized&p_p_mode=view&_epszavazokorok_WAR_nvinvrportlet_vlId=291&_epszavazokorok_WAR_nvinvrportlet_vltId=684&_epszavazokorok_WAR_nvinvrportlet_megyeKod="+maz+"&_epszavazokorok_WAR_nvinvrportlet_telepulesKod="+taz+"&_epszavazokorok_WAR_nvinvrportlet_valasztasTipusKod=E&_epszavazokorok_WAR_nvinvrportlet_searchSortColumn=SORSZAM&_epszavazokorok_WAR_nvinvrportlet_searchSortType=asc&_epszavazokorok_WAR_nvinvrportlet_megyeKod2="+maz+"&_epszavazokorok_WAR_nvinvrportlet_telepulesKod2="+taz+"&_epszavazokorok_WAR_nvinvrportlet_szavkorTypes=&_epszavazokorok_WAR_nvinvrportlet_settlement=Sz%C3%A1zhalombatta&_epszavazokorok_WAR_nvinvrportlet_searchWard=&p_auth=&p_auth=#_epszavazokorok_WAR_nvinvrportlet_paginator"
    req = requests.get(eu_url + '1')
    soup = BeautifulSoup(req.content)
    pages = int(soup.find('small', attrs = {'class':'search-results'}).text.split(' / ')[1]) // 20 + 1
    for page in range(1, pages+1):
      eu_url = "https://www.valasztas.hu/szavazokorok_ep2019?p_p_id=epszavazokorok_WAR_nvinvrportlet&p_p_lifecycle=1&p_p_state=maximized&p_p_mode=view&_epszavazokorok_WAR_nvinvrportlet_searchSortColumn=SORSZAM&_epszavazokorok_WAR_nvinvrportlet_megyeKod2="+maz+"&_epszavazokorok_WAR_nvinvrportlet_vlId=291&_epszavazokorok_WAR_nvinvrportlet_searchSortType=asc&_epszavazokorok_WAR_nvinvrportlet_vltId=684&_epszavazokorok_WAR_nvinvrportlet_telepulesKod2="+taz+"&_epszavazokorok_WAR_nvinvrportlet_valasztasTipusKod=E&_epszavazokorok_WAR_nvinvrportlet_delta=20&_epszavazokorok_WAR_nvinvrportlet_resetCur=false&_epszavazokorok_WAR_nvinvrportlet_cur=%s#_epszavazokorok_WAR_nvinvrportlet_paginator"
      req = requests.get(eu_url %str(page))
      soup = BeautifulSoup(req.content)
      link_divs = soup.find_all('div', attrs = {"class" : "nvi-search-container-row"})
      links = []
      for div in link_divs:
         links.append(div.find('a', href=True)["href"]) 
      for link in links:
         szk_no = ""
         while szk_no == "":
          try:
            req = requests.get(link)
            time.sleep(0.5)
            soup = BeautifulSoup(req.content)
            szk_no = soup.find('h1', attrs = {"class": "pb-1"}).text
            szk_no = szk_no.split(". számú")[0].split(" ")[-1]
            if szk_no not in data:
                data[szk_no] = {}

            conts = soup.find_all('div', attrs = {"class":  "nvi-search-container-row"})
            for cont in conts:
                text = cont.text.split(" ")
                text = list(filter(None, text))
                data[szk_no][text[1]] = int(text[-2])
                data[szk_no]["szk"] = int(szk_no)
          except:
            continue
      
    return(data)

@st.cache
def onk(eredmeny_url, maz, taz, keys):
    import requests
    from bs4 import BeautifulSoup

    print("ONK...")

    #14 033

    szk_ind = 0
    evk_dict = {}
    polgi_dict = {}
    kepv_dict = {}
    megyei_dict = {}

    szk_exists = True

    for szk in keys:
      url = "https://www.valasztas.hu/telepules-adatlap_onk2019?p_p_id=onkszavazokorieredmenyek_WAR_nvinvrportlet&p_p_lifecycle=1&p_p_state=maximized&p_p_mode=view&_onkszavazokorieredmenyek_WAR_nvinvrportlet_telepulesKod="+taz+"&_onkszavazokorieredmenyek_WAR_nvinvrportlet_megyeKod2="+maz+"&_onkszavazokorieredmenyek_WAR_nvinvrportlet_megyeKod="+maz+"&_onkszavazokorieredmenyek_WAR_nvinvrportlet_vlId=294&_onkszavazokorieredmenyek_WAR_nvinvrportlet_vltId=687&_onkszavazokorieredmenyek_WAR_nvinvrportlet_telepulesKod2="+taz+"&_onkszavazokorieredmenyek_WAR_nvinvrportlet_szavkorSorszam="+str(szk)+"&_onkszavazokorieredmenyek_WAR_nvinvrportlet_valaltipKod=H"
      req = requests.get(url)
      soup = BeautifulSoup(req.content)
      divs = soup.find_all('div', attrs = {"class" : "nvi-search-container-row"})
      if divs == []:
        szk_exists = False
      else:       
        evk = soup.find_all('h2', attrs = {"class": "f-size-20"})
        if evk != []:
          evk = evk[0].text.split(". egyéni")[0].split(" ")[2]
          if evk not in evk_dict:
            evk_dict[evk] = {}
            evk_dict[evk]["szavazokorok"] = []
          evk_dict[evk]["szavazokorok"].append(szk)
        else:
          evk_dict[str(szk)] = {'szavazokorok' : [szk]}

        if szk not in polgi_dict:
          polgi_dict[szk] = {}
        for div in divs:
          arr = div.text.split("Érvényes szavazatok:")
          nev = arr[0].strip()
          arr = arr[1].split("Jelölő szervezet:")
          part = arr[1].strip().split("     ")
          part = [p.split("   ")[0].strip("-").strip() for p in part] 
          st.write(part)
          nev = nev + " - " + part
          szavazatok = arr[0].split(" (")[0].strip()
          polgi_dict[szk][nev] = szavazatok
        
        
        filter = soup.find('div', attrs = {"class" : "nvi-electoral-district-filter"}).text
        if "EVK-választás" in filter:
           new_url = url + "&_onkszavazokorieredmenyek_WAR_nvinvrportlet_tabId2=EVK_KEPVISELO_VALASZTASA"
           req = requests.get(new_url)
           soup = BeautifulSoup(req.content)
           if szk not in kepv_dict:
              kepv_dict[szk] = {}
           divs = soup.find_all('div', attrs = {"class" : "nvi-search-container-row"})
           for div in divs:
              arr = div.text.split("Érvényes szavazatok:")
              nev = arr[0].strip()
              arr = arr[1].split("Jelölő szervezet:")
              part = arr[1].strip()
              nev = nev + " - " + part
              szavazatok = arr[0].split(" (")[0].strip()
              kepv_dict[szk][nev] = szavazatok
        if "Megyei" in filter:
            new_url = url + "&_onkszavazokorieredmenyek_WAR_nvinvrportlet_tabId2=MEGYEI_KOZGYULES_VALASZTASA"
            req = requests.get(new_url)
            soup = BeautifulSoup(req.content)
            if szk not in megyei_dict:
                megyei_dict[str(szk)] = {}
            divs = soup.find_all('div', attrs = {"class" : "nvi-search-container-row"})
            for div in divs:
              arr = div.text.split("Érvényes szavazatok:")
              part = arr[0].replace("Lista:", "").strip().split(" ")[0].strip()
              szavazatok = arr[1].strip()
              megyei_dict[str(szk)][part] = int(szavazatok)
              megyei_dict[str(szk)]['szk'] = int(szk)
        szk_ind += 1


    
    return(polgi_dict, evk_dict, kepv_dict, megyei_dict)

def writeExcel(data_ogy, data_ep, data_polgi):
    import xlsxwriter
    import math

    output = BytesIO()
    
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet_ogy = workbook.add_worksheet('OGY')
    worksheet_ep = workbook.add_worksheet('EP') 
    worksheet_onk = workbook.add_worksheet('ONK') 
    worksheet_megyei = workbook.add_worksheet('MEGYEI') 
    
    bold = workbook.add_format({'bold': 1})
    orange_format = workbook.add_format({'bold': True, 'bg_color': 'orange'})
    blue_format = workbook.add_format({'bold': True, 'bg_color': '#66CCFF'})
    magenta_format = workbook.add_format({'bold': True, 'bg_color': 'magenta'})
    red_format = workbook.add_format({'bold': True, 'bg_color': 'red'})
    green_format = workbook.add_format({'bold': True, 'bg_color': '#33FF33'})
    grey_format = workbook.add_format({'bold': True, 'bg_color': '#909090'})
    silver_format = workbook.add_format({'bold': True, 'bg_color': 'silver'})
    yellow_format = workbook.add_format({'bold': True, 'bg_color': 'yellow'})
    darkred_format = workbook.add_format({'bold': True, 'bg_color': '#CC3333'})

    data_ep = format_data_ep(data_ep, data_polgi[1])

    worksheet_ep.write(0, 0, "Szavazókör", bold)
    i = 1
    for col in data_ep.columns:
      if col != "szk":
        if col == "FIDESZ":
            worksheet_ep.write(0, i, col, orange_format)
        elif col == "DK":
            worksheet_ep.write(0, i, col, blue_format)
        elif col == "MOMENTUM":
            worksheet_ep.write(0, i, col, magenta_format)
        elif col == "MKKP":
            worksheet_ep.write(0, i, col, yellow_format)
        elif col == "MSZP":
            worksheet_ep.write(0, i, col, red_format)
        elif col == "LMP":
            worksheet_ep.write(0, i, col, green_format)
        elif col == "MUNKÁSPÁRT":
            worksheet_ep.write(0, i, col, darkred_format)
        elif col == "MI HAZÁNK":
            worksheet_ep.write(0, i, col, grey_format)
        elif col == "JOBBIK":
            worksheet_ep.write(0, i, col, silver_format)
        else:
            worksheet_ep.write(0, i, col, bold)
        i = i + 1
    
    cols = list(data_ep.columns)
    cols.remove("szk")
    
    row_no = 1
    for index, row in data_ep.iterrows():
       worksheet_ep.write(row_no, 0, row["szk"])
       worksheet_ep.write(row_no, 1, row[cols[0]])
       worksheet_ep.write(row_no, 2, row[cols[1]])
       worksheet_ep.write(row_no, 3, row[cols[2]])
       worksheet_ep.write(row_no, 4, row[cols[3]])
       worksheet_ep.write(row_no, 5, row[cols[4]])
       worksheet_ep.write(row_no, 6, row[cols[5]])
       worksheet_ep.write(row_no, 7, row[cols[6]])
       worksheet_ep.write(row_no, 8, row[cols[7]])
       worksheet_ep.write(row_no, 9, row[cols[8]])
       row_no = row_no + 1

    center_format = workbook.add_format()

    center_format.set_align('center')
    center_format.set_align('vcenter')
    
    row_no = 1
    for d in data_polgi[1]:
      length = len(data_polgi[1][d]['szavazokorok'])
      if length > 1:
          data_to_write = data_ep.loc[str(data_polgi[1][d]['szavazokorok'][0])]["EP 2019 - DK / Momentum % (Az összes szavazat arányában a DK és a Momentum százaléka"]
          worksheet_ep.merge_range(row_no, 10, row_no+length-1, 10, data_to_write, center_format)
          data_to_write = data_ep.loc[str(data_polgi[1][d]['szavazokorok'][0])]["EP 2019 - MI (DK, Momentum, MSZP, LMP, Jobbik) %"] 
          worksheet_ep.merge_range(row_no, 11, row_no+length-1, 11, data_to_write, center_format)
          data_to_write = data_ep.loc[str(data_polgi[1][d]['szavazokorok'][0])]["EP 2019 - FIDESZ %"]
          worksheet_ep.merge_range(row_no, 12, row_no+length-1, 12, data_to_write, center_format)
          data_to_write = data_ep.loc[str(data_polgi[1][d]['szavazokorok'][0])]["EP 2019 - EGYÉB %"]
          worksheet_ep.merge_range(row_no, 13, row_no+length-1, 13, data_to_write, center_format)
          data_to_write = data_ep.loc[str(data_polgi[1][d]['szavazokorok'][0])]['2019 ÖNK eredmény / nyertes'].strip('-')
          worksheet_ep.merge_range(row_no, 14, row_no+length-1, 14, data_to_write, center_format)
      else:
          data_to_write = data_ep.loc[str(data_polgi[1][d]['szavazokorok'][0])]["EP 2019 - DK / Momentum % (Az összes szavazat arányában a DK és a Momentum százaléka"]
          worksheet_ep.write(row_no, 10, data_to_write)
          data_to_write = data_ep.loc[str(data_polgi[1][d]['szavazokorok'][0])]["EP 2019 - MI (DK, Momentum, MSZP, LMP, Jobbik) %"] 
          worksheet_ep.write(row_no, 11, data_to_write)
          data_to_write = data_ep.loc[str(data_polgi[1][d]['szavazokorok'][0])]["EP 2019 - FIDESZ %"]
          worksheet_ep.write(row_no, 12, data_to_write)
          data_to_write = data_ep.loc[str(data_polgi[1][d]['szavazokorok'][0])]["EP 2019 - EGYÉB %"]
          worksheet_ep.write(row_no, 13, data_to_write)
          data_to_write = data_ep.loc[str(data_polgi[1][d]['szavazokorok'][0])]['2019 ÖNK eredmény / nyertes'].strip('-')
          worksheet_ep.write(row_no, 14, data_to_write)
      row_no +=length
       

    relevant_rows = int(data_ep.shape[0])-1
    worksheet_ep.write(int(relevant_rows), 0, "Összesen", bold)
    worksheet_ep.write(int(relevant_rows)+1, 0, "Százalék", bold)
   
    
    
    data_ogy = format_data_ogy(data_ogy, data_polgi[1])
    
    worksheet_ogy.write(0, 0, "Szavazókör", bold)
    i = 1
    for col in data_ogy.columns:
      if col != "Szavazókör":
        if "FIDESZ" in col or "Fidesz" in col:
            worksheet_ogy.write(0, i, col, orange_format)
        elif "DK" in col or "ellenzék" in col:
            worksheet_ogy.write(0, i, col, blue_format)        
        elif "MI HAZÁNK" in col or "Mi Hazánk" in col:
            worksheet_ogy.write(0, i, col, grey_format)
        elif "MKKP" in col:
            worksheet_ogy.write(0, i, col, yellow_format)
        else:
            worksheet_ogy.write(0, i, col, bold)
        i = i + 1
  
    columns = data_ogy.columns.to_list()[:-3]
    
    row_no = 1
    for index, row in data_ogy.iterrows():
       col_no = 0
       for col in columns:
         worksheet_ogy.write(row_no, col_no, row[col])
         col_no = col_no + 1
       row_no = row_no + 1
    col_index = len(columns)
    row_no = 1
    for d in data_polgi[1]:

      length = len(data_polgi[1][d]['szavazokorok'])
      if length > 1:         
          index= int(data_polgi[1][d]['szavazokorok'][0])        
          data_to_write = data_ogy.loc[index]['Nyertes egyéni %']
          worksheet_ogy.merge_range(row_no, col_index, row_no+length-1, col_index, data_to_write, center_format)
          data_to_write = data_ogy.loc[index]['Nyertes lista %']
          worksheet_ogy.merge_range(row_no, col_index + 1, row_no+length-1, col_index + 1, data_to_write, center_format)
          data_to_write = data_ogy.loc[index]['2019 ÖNK eredmény / nyertes'].strip('-')
          worksheet_ogy.merge_range(row_no, col_index + 2, row_no+length-1, col_index + 2, data_to_write, center_format)
      else:
          index = int(data_polgi[1][d]['szavazokorok'][0])
          data_to_write = data_ogy.loc[index]['Nyertes egyéni %']
          worksheet_ogy.write(row_no, col_index, data_to_write)
          data_to_write = data_ogy.loc[index]['Nyertes lista %']
          worksheet_ogy.write(row_no, col_index + 1, data_to_write)
          data_to_write = data_ogy.loc[index]['2019 ÖNK eredmény / nyertes'].strip('-')
          worksheet_ogy.write(row_no, col_index + 2, data_to_write)

      #(print(data_to_write))
      row_no +=length
    relevant_rows = int(data_ogy.shape[0])-1
    worksheet_ogy.write(int(relevant_rows), 0, "Összesen", bold)
    worksheet_ogy.write(int(relevant_rows)+1, 0, "Százalék", bold)


    row_no = 0
    
    columns= list(data_polgi[0][list(data_polgi[0].keys())[0]].keys())
    
    polgi_columns_dict = {}
    kepv_columns_dict = {}
    k = 1
    for column in columns:
          polgi_columns_dict[column] = k
          k = k + 1       
    
    polgi_list_len = len(columns) + 3
   

    for d in data_polgi[1]:
          worksheet_onk.write(row_no, 0, "Szavazókör", bold)
          col_no = 1
          
          for column in columns:
            if 'FIDESZ' in column.strip('-'):
                worksheet_onk.write(row_no, col_no, column.strip('-'), orange_format)
            elif 'DK' in column.strip('-') or 'MOMENTUM' in column.strip('-') or 'LMP' in column.strip('-') or 'JOBBIK' in column.strip('-') or 'MSZP' in column.strip('-') or 'PÁRBESZÉD' in column.strip('-'):
                worksheet_onk.write(row_no, col_no, column.strip('-'), blue_format)
            else:
                worksheet_onk.write(row_no, col_no, column.strip('-'), bold)
            col_no  += 1
          
          if data_polgi[2] != {}:
            worksheet_onk.write(row_no, polgi_list_len, "Szavazókör", bold)
            kepv_columns = list(data_polgi[2][data_polgi[1][str(d)]["szavazokorok"][0]].keys())
            col_no = polgi_list_len + 1
            for kepv_column in kepv_columns:
              if 'FIDESZ' in kepv_column.strip('-'):
                worksheet_onk.write(row_no, col_no, kepv_column.strip('-'), orange_format)
                kepv_columns_dict[kepv_column] = col_no
              elif 'DK' in kepv_column.strip('-') or 'MOMENTUM' in kepv_column.strip('-') or 'LMP' in kepv_column.strip('-') or 'JOBBIK' in kepv_column.strip('-') or 'MSZP' in kepv_column.strip('-') or 'PÁRBESZÉD' in kepv_column.strip('-'):
                worksheet_onk.write(row_no, col_no, kepv_column.strip('-'), blue_format)
                kepv_columns_dict[kepv_column] = col_no
              else:
                worksheet_onk.write(row_no, col_no, kepv_column.strip('-'), bold)
                kepv_columns_dict[kepv_column] = col_no
              col_no  += 1
          row_no += 1
          
          ossz = {}
          for szk in data_polgi[1][d]['szavazokorok']:
                       
            
            kepv_columns = []
           
            col_no = 1
            worksheet_onk.write(row_no, 0, szk)
            if data_polgi[2] != {}:
              kepv_columns = list(data_polgi[2][str(szk)].keys())
              worksheet_onk.write(row_no, polgi_list_len, szk)
            
            for column in columns:
              worksheet_onk.write(row_no, polgi_columns_dict[column], int(data_polgi[0][str(szk)][column]))
              #worksheet_onk.write(row_no, col_no + polgi_list_len, int(data_polgi[2][int(szk)][column]))
              if (column+"polgi") not in ossz:
                ossz[column + "polgi"] = 0
              ossz[column + "polgi"] += int(data_polgi[0][str(szk)][column])
              col_no += 1
            if data_polgi[2] != {}:
              
              col_no = polgi_list_len + 1
                          
              
              for kepv_column in kepv_columns:
                #worksheet_onk.write(row_no, col_no, int(data_polgi[0][int(szk)][column]))
                worksheet_onk.write(row_no, kepv_columns_dict[kepv_column], int(data_polgi[2][str(szk)][kepv_column]))
                if (kepv_column+"kepv") not in ossz:
                  ossz[kepv_column + "kepv"] = 0
                ossz[kepv_column + "kepv"] += int(data_polgi[2][str(szk)][kepv_column])
                col_no += 1
            row_no += 1
            if data_polgi[2] != {}:
              col_no = polgi_list_len + 1
              worksheet_onk.write(row_no, polgi_list_len, "Összesen")
              for kepv_column in kepv_columns:
                worksheet_onk.write(row_no, kepv_columns_dict[kepv_column], ossz[kepv_column + "kepv"])
                col_no += 1
          col_no = 1         
          worksheet_onk.write(row_no, 0, "Összesen")           
          for column in columns:
             worksheet_onk.write(row_no, polgi_columns_dict[column], ossz[column + "polgi"])
             col_no += 1
          row_no += 2

    if data_polgi[3] != {}:
        data_megyei = format_data_megyei(data_polgi[3], data_polgi[1]) 
        worksheet_ep.write(0, 0, "Szavazókör", bold)
        i = 1
        for col in data_megyei.columns:
          if col != "szk":
            if col == "FIDESZ":
                worksheet_megyei.write(0, i, col, orange_format)
            elif col == "DK":
                worksheet_megyei.write(0, i, col, blue_format)
            elif col == "MOMENTUM":
                worksheet_megyei.write(0, i, col, magenta_format)
            elif col == "MKKP":
                worksheet_megyei.write(0, i, col, yellow_format)
            elif col == "MSZP":
                worksheet_megyei.write(0, i, col, red_format)
            elif col == "LMP":
                worksheet_megyei.write(0, i, col, green_format)
            elif col == "MUNKÁSPÁRT":
                worksheet_megyei.write(0, i, col, darkred_format)
            elif col == "MI HAZÁNK":
                worksheet_megyei.write(0, i, col, grey_format)
            elif col == "JOBBIK":
                worksheet_megyei.write(0, i, col, silver_format)
            else:
                worksheet_megyei.write(0, i, col, bold)
            i = i + 1
        row_no = 1
        cols = list(data_megyei.columns)
        cols.remove("szk")
        print(cols)
       
        for index, row in data_megyei.iterrows():
            try:
              worksheet_megyei.write(row_no, 0, row["szk"])
              worksheet_megyei.write(row_no, 1, row[cols[0]])
              worksheet_megyei.write(row_no, 2, row[cols[1]])
              worksheet_megyei.write(row_no, 3, row[cols[2]])
              worksheet_megyei.write(row_no, 4, row[cols[3]])
              worksheet_megyei.write(row_no, 5, row[cols[4]])
            except:
                continue
            row_no = row_no + 1  

        row_no = 1
        for d in data_polgi[1]:
          length = len(data_polgi[1][d]['szavazokorok'])
          try:
              if length > 1:
                  data_to_write = data_megyei.loc[str(data_polgi[1][d]['szavazokorok'][0])]["MEGYEI ÖNK 2019 - DK / Momentum %"]
                  worksheet_megyei.merge_range(row_no, 6, row_no+length-1, 6, data_to_write, center_format)
                  data_to_write = data_megyei.loc[str(data_polgi[1][d]['szavazokorok'][0])]["MEGYEI ÖNK 2019 - MI (DK, Momentum, MSZP, Jobbik) %"] 
                  worksheet_megyei.merge_range(row_no, 7, row_no+length-1, 7, data_to_write, center_format)
                  data_to_write = data_megyei.loc[str(data_polgi[1][d]['szavazokorok'][0])]["MEGYEI ÖNK 2019 - FIDESZ %"]
                  worksheet_megyei.merge_range(row_no, 8, row_no+length-1, 8, data_to_write, center_format)
                  data_to_write = data_megyei.loc[str(data_polgi[1][d]['szavazokorok'][0])]["2019 ÖNK eredmény / nyertes"]
                  worksheet_megyei.merge_range(row_no, 9, row_no+length-1, 9, data_to_write, center_format)
              else:
                  data_to_write = data_megyei.loc[str(data_polgi[1][d]['szavazokorok'][0])]["MEGYEI ÖNK 2019 - DK / Momentum %"]
                  worksheet_megyei.write(row_no, 6, data_to_write)
                  data_to_write = data_megyei.loc[str(data_polgi[1][d]['szavazokorok'][0])]["MEGYEI ÖNK 2019 - MI (DK, Momentum, MSZP, Jobbik) %"] 
                  worksheet_megyei.write(row_no, 7, data_to_write)
                  data_to_write = data_megyei.loc[str(data_polgi[1][d]['szavazokorok'][0])]["MEGYEI ÖNK 2019 - FIDESZ %"]
                  worksheet_megyei.write(row_no, 8, data_to_write)
                  data_to_write = data_megyei.loc[str(data_polgi[1][d]['szavazokorok'][0])]["2019 ÖNK eredmény / nyertes"]
                  worksheet_megyei.write(row_no, 9, data_to_write)
          except:
            continue
          row_no +=length

          relevant_rows = int(data_megyei.shape[0])-1
          worksheet_megyei.write(int(relevant_rows), 0, "Összesen", bold)
          worksheet_megyei.write(int(relevant_rows)+1, 0, "Százalék", bold)
    
    workbook.close()
    return output.getvalue()

def format_data_megyei(data, evk_dict):
    import numpy as np
    import pandas as pd
    import math

    df = pd.DataFrame(data).T
    df.loc['Összesen']= df.sum()
    mylist = list(df.loc["Összesen"])
    arr = [0 if math.isnan(x) else x for x in mylist]
    szk_sum = df.loc["Összesen"].apply(np.sum)["szk"]
    total_sum = sum(arr)-szk_sum
    df["összeg"] = df.sum(axis = 1) - df["szk"]

    evk_list = []
    for d in evk_dict:
      for szam in evk_dict[d]['szavazokorok']:
        evk_list.append(szam)
    evk_list = [str(x) for x in evk_list]
    df= df.loc[evk_list]

    columns = df.columns.to_list()
    
    try:
        df['MEGYEI ÖNK 2019 - DK / Momentum %'] = df.apply(lambda row : custom_row('ep1', row["szk"], evk_dict, df), axis = 1)
        df['MEGYEI ÖNK 2019 - MI (DK, Momentum, MSZP, Jobbik) %'] = df.apply(lambda row : custom_row('megyei2', row["szk"], evk_dict, df), axis = 1)   
        df["MEGYEI ÖNK 2019 - FIDESZ %"] = df.apply(lambda row : custom_row('ep3', row["szk"], evk_dict, df), axis = 1)    

        df["2019 ÖNK eredmény / nyertes"] = df.apply(lambda row : custom_row('ep5', row["szk"], evk_dict, df), axis = 1)
    
    except:
        pass
    
    df.loc['Összesen']= df[columns].sum()
    df.loc['Százalék'] = (df.loc["Összesen"]/total_sum)*100

    df = df.drop('összeg', axis=1)

    df = df.fillna(0)

    return df

def format_data_ep(data, evk_dict):
    import numpy as np
    import pandas as pd
    import math

    
    df = pd.DataFrame(data).T
    df.loc['Összesen']= df.sum()
    mylist = list(df.loc["Összesen"])
    arr = [0 if math.isnan(x) else x for x in mylist]
    szk_sum = df.loc["Összesen"].apply(np.sum)["szk"]
    total_sum = sum(arr)-szk_sum
    df["összeg"] = df.sum(axis = 1) - df["szk"]
    df = df.rename(columns={"MI": "MI HAZÁNK"})
    
    evk_list = []
    for d in evk_dict:
      for szam in evk_dict[d]['szavazokorok']:
        evk_list.append(szam)
    evk_list = [str(x) for x in evk_list]
    df= df.loc[evk_list]

    df['EP 2019 - DK / Momentum % (Az összes szavazat arányában a DK és a Momentum százaléka'] = df.apply(lambda row : custom_row('ep1', row["szk"], evk_dict, df), axis = 1)
    df['EP 2019 - MI (DK, Momentum, MSZP, LMP, Jobbik) %'] = df.apply(lambda row : custom_row('ep2', row["szk"], evk_dict, df), axis = 1)   
    df["EP 2019 - FIDESZ %"] = df.apply(lambda row : custom_row('ep3', row["szk"], evk_dict, df), axis = 1)    
    df["EP 2019 - EGYÉB %"] = df.apply(lambda row : custom_row('ep4', row["szk"], evk_dict, df), axis = 1)
    
    columns = df.columns.to_list()
    
    df["2019 ÖNK eredmény / nyertes"] = df.apply(lambda row : custom_row('ep5', row["szk"], evk_dict, df), axis = 1)

    df.loc['Összesen']= df[columns].sum()
    df.loc['Százalék'] = (df.loc["Összesen"]/total_sum)*100

    df = df.drop('összeg', axis=1)

    df = df.fillna(0)

    return df

def format_data_ogy(data, evk_dict):
    import numpy as np
    import pandas as pd
    import math
    
    df = pd.DataFrame(data).T
    df.loc['Összesen']= df.sum()
    mylist = list(df.loc["Összesen"])
    arr = [0 if math.isnan(x) else x for x in mylist]
    szk_sum = df.loc["Összesen"].apply(np.sum)["Szavazókör"]
    total_sum = sum(arr)-szk_sum
    df["összeg_list"] = df[["Fidesz", "ellenzék", "MKKP", 'MEMO', 'Normális Párt', 'Mi Hazánk']].sum(axis = 1)

    df["max_list"] = df[['Fidesz', 'ellenzék']].max(axis=1)
    
    evk_list = []
    for d in evk_dict:
      for szam in evk_dict[d]['szavazokorok']:
        if int(szam) in list(df.index):
            evk_list.append(szam)
    evk_list = [int(x) for x in evk_list]
    df= df.loc[evk_list]
    
    columns = df.columns.to_list()
    columns.remove('Szavazókör')
    columns.remove('Fidesz')
    columns.remove('ellenzék')
    columns.remove('MKKP')
    columns.remove('MEMO')
    columns.remove('Normális Párt')
    columns.remove('Mi Hazánk')
    columns.remove('összeg_list')
    columns.remove('max_list')

    df["max_egyeni"] = df[columns].max(axis=1)
    df["összeg_egyeni"] = df[columns].sum(axis = 1)

    df['Nyertes egyéni %'] = df.apply(lambda row : custom_row('ogy1', row["Szavazókör"], evk_dict, df), axis = 1)
    df['Nyertes lista %'] = df.apply(lambda row : custom_row('ogy2', row["Szavazókör"], evk_dict, df), axis = 1)

    columns = df.columns.to_list()
    
    df["2019 ÖNK eredmény / nyertes"] = df.apply(lambda row : custom_row('ep5', row["Szavazókör"], evk_dict, df), axis = 1)

    df.loc['Összesen']= df[columns].sum()
    df.loc['Százalék'] = (df.loc["Összesen"]/total_sum)*100

    df = df.drop('összeg_egyeni', axis=1)
    df = df.drop('összeg_list', axis=1)
    df = df.drop('max_list', axis=1)
    df = df.drop('max_egyeni', axis=1)

    df = df.fillna(0)

    return df

def custom_row(param, szk, evk_dict, df):
    keys = list(evk_dict.keys())
    i = 0
    evk = keys[i]
    found = True
    evk_dict = get_winner()

    while found:
        if int(szk) in evk_dict[evk]['szavazokorok'] or str(int(szk)) in evk_dict[evk]['szavazokorok']:
            found = False
        else:
            evk = keys[i]
            i += 1

    if param == "ep1":
        dkmomentum = 0
        osszeg = 0
        for sz in evk_dict[evk]['szavazokorok']:
            dkmomentum += df.loc[str(sz)]["DK"] + df.loc[str(sz)]["MOMENTUM"]
            osszeg += df.loc[str(sz)]["összeg"]
        return ((dkmomentum / osszeg)*100)
    elif param == "ep2":
        ellenzek = 0
        osszeg = 0
        for sz in evk_dict[evk]['szavazokorok']:
            ellenzek += df.loc[str(sz)]["DK"] + df.loc[str(sz)]["MOMENTUM"] + df.loc[str(sz)]["JOBBIK"] + df.loc[str(sz)]["MSZP"] + df.loc[str(sz)]["LMP"]
            osszeg += df.loc[str(sz)]["összeg"]
        return ((ellenzek / osszeg)*100)
    elif param == "ep3":
        fidesz = 0
        osszeg = 0
        for sz in evk_dict[evk]['szavazokorok']:
            fidesz += df.loc[str(sz)]["FIDESZ"]
            osszeg += df.loc[str(sz)]["összeg"]
        return ((fidesz / osszeg)*100)
    elif param == "ep4":
        kicsi_osszeg = 0
        osszeg = 0
        for sz in evk_dict[str(evk)]['szavazokorok']:
            kicsi_osszeg += df.loc[str(sz)]["MKKP"] + df.loc[str(sz)]["MI HAZÁNK"] + df.loc[str(sz)]["MUNKÁSPÁRT"]
            osszeg += df.loc[str(sz)]["összeg"]
        return ((kicsi_osszeg / osszeg)*100)
    elif param == "ep5":
        return (evk_dict[str(evk)]['győztes'])
    elif param == "ogy1":
        max = 0
        osszeg = 0
        for sz in evk_dict[str(evk)]['szavazokorok']:
            if int(sz) in list(df.index):
                max += df.loc[int(sz)]["max_egyeni"]
                osszeg += df.loc[int(sz)]["összeg_egyeni"]
        return ((max / osszeg)*100)
    elif param == "ogy2":
        max = 0
        osszeg = 0
        for sz in evk_dict[evk]['szavazokorok']:
            if int(sz) in list(df.index):
                max += df.loc[int(sz)]["max_list"]
                osszeg += df.loc[int(sz)]["összeg_list"]
        return ((max / osszeg)*100)
    elif param == "megyei2":
        ellenzek = 0
        osszeg = 0
        for sz in evk_dict[evk]['szavazokorok']:            
            ellenzek += df.loc[str(sz)]["DK"] + df.loc[str(sz)]["MOMENTUM"] + df.loc[str(sz)]["JOBBIK"] + df.loc[str(sz)]["MSZP"]
            osszeg += df.loc[str(sz)]["összeg"]
        return ((ellenzek / osszeg)*100)

requests.packages.urllib3.disable_warnings()
requests.packages.urllib3.util.ssl_.DEFAULT_CIPHERS += ':HIGH:!DH:!aNULL'

try:
    requests.packages.urllib3.contrib.pyopenssl.util.ssl_.DEFAULT_CIPHERS += ':HIGH:!DH:!aNULL'
except AttributeError:
    # no pyopenssl support used / needed / available
    pass

st.title('Választási adatok')

st.write("A város nevét úgy írd be, ahogy az a választás.hu-n található. Ez elsősorban Budapesten érdekes, ahol a \"Budapest XIV. kerület\" formátum a jó." )

st.write("Az adatletöltés gyorsasága a település méretétől függ." )

st.write("Készítette, hibabejelentés: Tóth Gy. Bori, +36 30 648 0643, bori.tothgy@gmail.com")

st.write("Tudott hibák: Budapesten önkormányzati, Heves megye megyei listás választás")

city_url = "https://vtr.valasztas.hu/ogy2022/data/04022333/ver/Telepulesek.json"

city = st.text_input('Írd be a várost!')

if city:

    eredmeny_url, maz, taz = select_city(city_url, city)

    if eredmeny_url == 0:
        st.write("Rossz város!")
    else:
        data_load_state = st.text('OGY...')
        data_ogy=ogy(eredmeny_url, maz, taz)

        data_load_state = st.text('EP...')
        data_ep=ep(eredmeny_url, maz, taz)

        data_load_state = st.text('ONK...')
        polgi_data = onk(eredmeny_url, maz, taz, list(data_ep.keys()))
        data_polgi = copy.deepcopy(polgi_data)

        workbook = writeExcel(data_ogy, data_ep, data_polgi)  

        st.write("Kész!")
        label = "Letöltés " + city + " adatok"

        st.download_button(
            label=label,
            data=workbook,
            file_name=city+".xlsx",
            mime="application/vnd.ms-excel"
        )    
       
   
