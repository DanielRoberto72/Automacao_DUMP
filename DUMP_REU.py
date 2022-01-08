#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import selenium, os, time, pandas as pd, csv, warnings, shutil, sys, lxml, re, itertools, openpyxl, glob, mysql.connector, smtplib
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup as BS
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import win32com.client
import os
import glob
warnings.filterwarnings("ignore")
log = ''

tempo = datetime.now() - timedelta()
timestamp = tempo.strftime('%Y-%m-%d')
timestamp_envio = tempo.strftime('%d-%m')
#-----------------------------------------------------------------------------------------------
#SETANDO INFORMACOES FIXAS
login = "login"
senha = "senha"
dirRaiz = 'C:/Prod/Python/BuscaReu/'
diretorio = dirRaiz + 'arquivos/'

#-----------------------------------------------------------------------------------------------
#INICIANDO O CHROMEDRIVER
chrome_options = webdriver.ChromeOptions()
chromedriver = dirRaiz+"Driver/chromedriver.exe"
prefs = {"download.default_directory": r"C:\Prod\Python\BuscaReu\arquivos"}
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument('ignore-certificate-errors')
driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=chromedriver)
#-----------------------------------------------------------------------------------------------
def wait_xpath_click(y):
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, y))).click()
#-----------------------------------------------------------------------------------------------
#BAIXANDO RELATÓRIO DA SMS_BLOCKING

try:
    driver.get("link")
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '//*[@id="LoginForm"]/div/div[1]/input[1]'))).send_keys(login_REU)
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '//*[@id="LoginForm"]/div/div[2]/input[1]'))).send_keys(senha_REU)
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '//*[@id="btnLogin"]'))).click()
    print('lOGADO!')
    print('Entrando na pagina do SMS blocking')
    driver.get('link')
    print('Buscando pagina do SMS blocking')
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '//*[@id="unload_button"]'))).click()
    print('Entrando na pagina de download do SMS blocking')
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '//*[@id="radio_button_2_id"]'))).click()
    print('Selecionado local disk')
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '//*[@id="textfontid2"]/input'))).click()
    print('realizando download')
    time.sleep(20)
    print('realizado download do sms blocking')
    
    file_oldname = os.path.join(diretorio, "unload.txt")
    file_newname_newfile = os.path.join(diretorio, "SMS_blocking_"+timestamp_envio+".txt")
    os.rename(file_oldname, file_newname_newfile)
    file_SMS = diretorio+"SMS_blocking_"+timestamp_envio+".txt"
    
    
#-----------------------------------------------------------------------------------------------
#BAIXANDO RELATÓRIO DA SMS_BLOCKING


    print('----------------------------------------')
    print('Entrando na pagina do CAMEL blocking')
    driver.get('link')
    print('Buscando na pagina do CAMEL blocking')
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '//*[@id="unload_button"]'))).click()
    print('Entrando na pagina de download do CAMEL blocking')
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '//*[@id="radio_button_2_id"]'))).click()
    print('Selecionado local disk')
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '//*[@id="textfontid2"]/input'))).click()
    print('realizando download')
    time.sleep(20)
    print('realizado download do camel blocking')
    file_oldname = os.path.join(diretorio, "unload.txt")
    file_newname_newfile = os.path.join(diretorio, "camel_blocking_"+timestamp_envio+".txt")
    os.rename(file_oldname, file_newname_newfile)
    file_Camel = diretorio+"camel_blocking_"+timestamp_envio+".txt"
    driver.close()
    print('----------------------------------------')
    
    #-----------------------------------------------------------------------------------------------
    #ENVIO DO EMAIL PARA OS DESTINATÁRIOS FIXOS
    print ('Enviando e-mail aos destinatários')
    tempo = datetime.now() - timedelta()
    timestamp_envio = tempo.strftime('%d-%m')
    try:
        email = 'email@email.com'
        password = 'senha'
        send_to_email = ['destinatarios','destinatarios']
        subject = 'DUMP SMS blocking e Camel blocking '+timestamp_envio
        message ='''
    Bom dia!

    Segue em anexo o DUMP do SMS blocking e do Camel blocking


    Atenciosamente. '''

        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = ", ".join(send_to_email)
        msg['Subject'] = subject

        msg.attach(MIMEText(message, 'plain'))
        
        # Setup the attachment
        filename = os.path.basename(file_SMS)
        attachment = open(file_SMS, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        
        filename = os.path.basename(file_Camel)
        attachment = open(file_Camel, "rb")
        part1 = MIMEBase('application', 'octet-stream')
        part1.set_payload(attachment.read())
        encoders.encode_base64(part1)
        part1.add_header('Content-Disposition', "attachment; filename= %s" % filename)


        # Attach the attachment to the MIMEMultipart object
        
        msg.attach(part)
        msg.attach(part1)


        server = smtplib.SMTP('SMTP.office365.com',587)
        server.starttls()
        server.login(email, password)
        text = msg.as_string()
        server.sendmail(email, send_to_email, text)
        server.quit()
        
        attachment.close()

        print('Email enviado COM SUCESSO PARA OS DESTINATÁRIOS')
    except:
        print('Falha ao enviar o Email!')
        the_type, the_value, the_traceback = sys.exc_info()
        print(the_type, ',' ,the_value,',', the_traceback)
        pass
    try:
        print('----------------------------------------')
        print('Deletando os arquivos')
        os.remove(diretorio+'SMS_blocking_'+timestamp_envio+'.txt')
        os.remove(diretorio+'camel_blocking_'+timestamp_envio+'.txt')
        print('Processo de remover os arquivos finalizado!!!')
    except:
        print('falha ao remover arquivo')
        the_type, the_value, the_traceback = sys.exc_info()
        print(the_type, ',' ,the_value,',', the_traceback)


except:
    print('Falha ao executar o script, script finalizado!')
    sys.exit()


# In[ ]:




