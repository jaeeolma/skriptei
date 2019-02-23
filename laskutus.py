#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import print_function
import smtplib
import pandas 
import datetime
import time
from email.mime.text import MIMEText
from email.header import Header
from email.charset import Charset
import sys

# NOW WITH PYTHON 3.6 WOOOO

#Tähän omat gmailtiedot
APP_PWD = LAITA
GLB_USER = OMASI


def send_email(recipient, subject, body):
    """Lähettää laskut gmailin kautta
    """
    gmail_user = GLB_USER
    gmail_pwd = APP_PWD
    FROM = 'OMA NIMI'
    TO = recipient 
    SUBJECT = subject
    TEXT = body
    
    #valmistele viesti
    
    msg = MIMEText(body, 'plain', 'UTF-8')
    msg['Subject'] = Header(SUBJECT.encode('utf-8'), 'UTF-8').encode()
    msg['From'] = Header(FROM.encode('utf-8'), 'UTF-8').encode()
    msg['to'] = Header(TO.encode('utf-8'), 'UTF-8').encode()
   
    try:
        server_ssl = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server_ssl.ehlo()
        server_ssl.login(gmail_user, gmail_pwd)
        server_ssl.sendmail(FROM, TO, msg.as_string())
        server_ssl.close()
    except:
        print('failed to send mail to:', recipient)

# laskee halutun määrän viitenumeroita annetun numeron pohjalta
# voi laskea max 1000 numeroa helposti
def count_nbr(nbr, amount):
    nbrs = []
    i = 0
    add = '00'
    for i in range(amount):
        if i == 10:
            add = '0'
        if i == 100:
            add = ''
        nbr_string = str(nbr) + add + str(i)
        kertoimet = (7,3,1)
        summa = sum(kertoimet[i % 3] * int(x) for i, x in enumerate(reversed(nbr_string)))
        check = (10 - (summa % 10)) % 10
        if check == 10:
            check = 0
        viite = nbr_string + str(check)
        nbrs.append(viite)
    return nbrs


def read_xlsx(file):
    """Lukee excel-taulukosta sarakkeet 'Nimi', 'Sähköposti' ja 'Hinta'
    """
    df = pandas.read_excel(file,options={'encoding':'utf-8'}, keep_default_na=False)
    FORMAT = [u'Nimi', u'Sähköposti', u'Hinta']
    df_selected = df[FORMAT]
    return df_selected

    
def combine_nbrs(nbrs, df, outputfile):
    """Tallentaa exceltaulukon, jossa laskutettaville on lisätty henkilökohtainen
    viite
    """
    df[u'Viite'] = nbrs
    writer = pandas.ExcelWriter(outputfile, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Taul1')
    writer.save()

def add_invoice_to_text(bodytext, df, i, due):
    """Luo sähköpostiviestin ja liittää laskun siihen
    """
    maksaja = u'Maksaja: ' + df.loc[i, 'Nimi']
    saaja = u'Saaja: TÄYTÄ OMA'
    tili = u'Saajan tilinumero: TÄYTÄ OMA'
    viite = u'Viitenumero: ' + df.loc[i, 'Viite']
    expdate = datetime.datetime.now() + datetime.timedelta(days=int(due))
    date = u'Eräpäivä: ' + expdate.strftime('%d.%m.%Y')
    summa = u'Maksettava summa: ' + str(df.loc[i, 'Hinta']) + u' \u20ac'
    sign = u'Terveisin' + '\n' + u'NIMI'+ '\n' +u'TITTELI'+ '\n' + u'YHTEISÖ'
    code = u'\nVoit maksaa myös käyttämällä alla olevaa virtuaaliviivakoodia:\n' + virtuaaliviivakoodi(df.loc[i, 'Hinta'], df.loc[i,'Viite'], expdate)
    text = """\n%s\n%s\n%s\n%s\n%s\n%s\n%s\n\n%s""" % (maksaja, saaja, tili, viite, date, summa, code, sign)
    viesti = bodytext + text
    return viesti
    
def virtuaaliviivakoodi(sum, ref, date):
    """Luo virtuaaliviivakoodin laskulle
    """
    ver = '4'
    tili = 'täytä oma' # Tähän oma IBAN ilman FI-alkua 
    eurot = str(sum)
    sentit = '00'
    vara = '000'
    viite = ref
    year = str(date.year - 2000)
    mon = str(date.month)
    day = str(date.day)
    if len(mon) == 1:
        mon = '0' + mon
    if len(day) == 1:
        day = '0' + day
    for _ in range(20 - len(viite)):
        viite = '0' + viite
    for _ in range(6 - len(eurot)):
        eurot = '0' + eurot
    code = ver + tili + eurot + sentit + vara + viite + year + mon + day
    return code
    
def main():
    """Otetaan komennot vastaan ja lähetetään viestit
    """
    ref_start = input('Anna viitenumeron alkuosa: ')
    original_data = input('Anna excelin tiedostonimi: ')
    df = read_xlsx(original_data)
    nbrs = count_nbr(ref_start, len(df))
    filename = input('Anna tallennettavan tiedoston nimi: ')
    combine_nbrs(nbrs, df, filename)
    duedate = input('Anna maksuajan pituus: ')
    subject = input('Anna viestin otsikko: ')
    bodytext = input('Anna viestirungon nimi: ')
    textfile = open(bodytext, 'r', encoding='utf-8')
    teksti = textfile.read()
    for i in range(len(df)):
        text = add_invoice_to_text(teksti, df, i, duedate)
        #Testailua varten korvaa 'maili' omalla sähköpostilla
        send_email('maili', subject, df.loc[i, u'Sähköposti'] + '\n\n' + text)
        #varsinainen lähetyskomento
        #send_email(df.loc[i, u'Sähköposti'], subject, text)
        print('>>> {}/{} laskua lahetetty'.format(i + 1, len(df)), end='\r')
        sys.stdout.flush()
        if i % 20 == 0:
            time.sleep(2)
    print('')
    
if __name__ == '__main__':
    main()
