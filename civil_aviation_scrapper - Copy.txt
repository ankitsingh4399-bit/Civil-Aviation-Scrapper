from bs4 import BeautifulSoup
import requests
import re
import warnings

import win32com
import win32com.client
import sys
import time
from datetime import datetime

warnings.filterwarnings('ignore')



def error_email(date_time,error):
    # Create an Outlook application instance
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Email details
    recipients = 'ankit.singh8@goindigo.in; divyanshu.upadhyay@goindigo.in'
    mail.To = recipients
    mail.Subject = f'Civil Aviation Update Error: {date_time}'
    mail.HTMLBody = f'''
                        <html>
                            <body>
                                <p>Dear Team,<p>
                                <p>Error {error} occured while running the script.<p>
                                <br>
                                <br>
                                <br>    
                                <br>
                                <p>Thanks & Regards,<br><br>Publication Team<br>(Comm,ISC)<p> 
                    '''
    
    mail.Send()
    
    print("\nError Email sent!\n")



def notification_email(date_time, domestic, international, otp, plf):
    # Create an Outlook application instance
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Email details
    recipients = 'skdteam@goindigo.in'
    mail.To = recipients
    mail.Subject = f'Civil Aviation Website Update: {date_time}'
    mail.HTMLBody = f'''
                        <html>
                            <body>
                                <p>Dear All,<p>
                                <p>Please find the below data as of today from Civil Aviation website:<p>
                                <br>
                                <p><b>{domestic}</b><p>
                                <br>
                                <p><b>{international}</b><p>
                                <br>
                                <p><b>{otp}</b><p>
                                <br>
                                <p><b>{plf}</b><p>
                                <br>
                                <br>    
                                <br>
                                <p>Thanks & Regards,<br><br>Publication Team<br>(Comm,ISC)<p> 
                    '''
    
    mail.Send()



    print("\nNotification Email sent successfully!\n")

def run_scrapper():
    flag = 0
    
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36", "Accept-Encoding":"gzip, deflate", "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8", "DNT":"1","Connection":"close", "Upgrade-Insecure-Requests":"1"}

        try:
            page = requests.get(url = 'https://www.civilaviation.gov.in', headers=headers)
        except:
            page = requests.get(url = 'http://www.civilaviation.gov.in', headers=headers)
        
        soup = BeautifulSoup(page.content, 'html.parser')
        soup = BeautifulSoup(soup.prettify(), 'html.parser')

        DEVANAGARI_START = re.compile(r'^\s*[\u0900-\u097F]')
        domestic_traffic  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 domestic-traffic"}).get_text()
        domestic_traffic = [i for i in domestic_traffic.split('\n') if i !=' ' or i != '']
        domestic_traffic = [i.strip() for i in domestic_traffic]
        domestic_traffic = [i for i in domestic_traffic if i!='']
        domestic_traffic = [s for s in domestic_traffic if not DEVANAGARI_START.match(s or "")]

        domestic_traffic = [f"{domestic_traffic[i]}: {domestic_traffic[i+1]}"
                for i in range(0, len(domestic_traffic)-1, 2)]

        domestic = "<br>".join(domestic_traffic)

        international_traffic  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 international-traffic"}).get_text()
        international_traffic = [i for i in international_traffic.split('\n') if i !=' ' or i != '']
        international_traffic = [i.strip() for i in international_traffic]
        international_traffic = [i for i in international_traffic if i!='']
        international_traffic = [s for s in international_traffic if not DEVANAGARI_START.match(s or "")]

        international_traffic = [f"{international_traffic[i]}: {international_traffic[i+1]}"
                for i in range(0, len(international_traffic)-1, 2)]

        international = "<br>".join(international_traffic)

        on_time_performance  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 on-time-performance"}).get_text()
        on_time_performance = [i for i in on_time_performance.split('\n') if i !=' ' or i != '']
        on_time_performance = [i.strip() for i in on_time_performance]
        on_time_performance = [i for i in on_time_performance if i!='']
        on_time_performance = [s for s in on_time_performance if not DEVANAGARI_START.match(s or "")]

        on_time_performance = [f"{on_time_performance[i]}: {on_time_performance[i+1]}"
                for i in range(0, len(on_time_performance)-1, 2)]

        otp = "<br>".join(on_time_performance)

        passenger_load_factor  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 passenger-load-factor"}).get_text()
        passenger_load_factor = [i for i in passenger_load_factor.split('\n') if i !=' ' or i != '']
        passenger_load_factor = [i.strip() for i in passenger_load_factor]
        passenger_load_factor = [i for i in passenger_load_factor if i!='']
        passenger_load_factor = [s for s in passenger_load_factor if not DEVANAGARI_START.match(s or "")]

        passenger_load_factor = [f"{passenger_load_factor[i]}: {passenger_load_factor[i+1]}"
                for i in range(0, len(passenger_load_factor)-1, 2)]

        plf = "<br>".join(passenger_load_factor)

        now = datetime.now().strftime('%d-%b-%y %H:%M')

        notification_email(now, domestic, international, otp, plf)
        flag = 1
        return flag
        
    except Exception as e:
        now = datetime.now().strftime('%d-%b-%y %H:%M')
        print(f'Error {e} Ocurred!')
        
        flag = 0
        return flag
        error_email(now,e)

    
if __name__ == "__main__":
    x = 0
    i = 0
    limit = 10
    while x!=1:
        if i < limit:
            i +=1
            # print(f'Running {i} time(s)')
            # time.sleep(1)
            x = run_scrapper()
        else:
            now = datetime.now().strftime('%d-%b-%y %H:%M')
            error_email(now, 'Limit (10) reached for number of iterations')
            sys.exit()     
    # print('Successful.')
    # sys.exit()
    
    # try:
    #     headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36", "Accept-Encoding":"gzip, deflate", "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8", "DNT":"1","Connection":"close", "Upgrade-Insecure-Requests":"1"}

    #     try:
    #         page = requests.get(url = 'https://www.civilaviation.gov.in', headers=headers)
    #     except:
    #         page = requests.get(url = 'http://www.civilaviation.gov.in', headers=headers)
        
    #     soup = BeautifulSoup(page.content, 'html.parser')
    #     soup = BeautifulSoup(soup.prettify(), 'html.parser')

    #     DEVANAGARI_START = re.compile(r'^\s*[\u0900-\u097F]')
    #     domestic_traffic  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 domestic-traffic"}).get_text()
    #     domestic_traffic = [i for i in domestic_traffic.split('\n') if i !=' ' or i != '']
    #     domestic_traffic = [i.strip() for i in domestic_traffic]
    #     domestic_traffic = [i for i in domestic_traffic if i!='']
    #     domestic_traffic = [s for s in domestic_traffic if not DEVANAGARI_START.match(s or "")]

    #     domestic_traffic = [f"{domestic_traffic[i]}: {domestic_traffic[i+1]}"
    #             for i in range(0, len(domestic_traffic)-1, 2)]

    #     domestic = "<br>".join(domestic_traffic)

    #     international_traffic  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 international-traffic"}).get_text()
    #     international_traffic = [i for i in international_traffic.split('\n') if i !=' ' or i != '']
    #     international_traffic = [i.strip() for i in international_traffic]
    #     international_traffic = [i for i in international_traffic if i!='']
    #     international_traffic = [s for s in international_traffic if not DEVANAGARI_START.match(s or "")]

    #     international_traffic = [f"{international_traffic[i]}: {international_traffic[i+1]}"
    #             for i in range(0, len(international_traffic)-1, 2)]

    #     international = "<br>".join(international_traffic)

    #     on_time_performance  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 on-time-performance"}).get_text()
    #     on_time_performance = [i for i in on_time_performance.split('\n') if i !=' ' or i != '']
    #     on_time_performance = [i.strip() for i in on_time_performance]
    #     on_time_performance = [i for i in on_time_performance if i!='']
    #     on_time_performance = [s for s in on_time_performance if not DEVANAGARI_START.match(s or "")]

    #     on_time_performance = [f"{on_time_performance[i]}: {on_time_performance[i+1]}"
    #             for i in range(0, len(on_time_performance)-1, 2)]

    #     otp = "<br>".join(on_time_performance)

    #     passenger_load_factor  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 passenger-load-factor"}).get_text()
    #     passenger_load_factor = [i for i in passenger_load_factor.split('\n') if i !=' ' or i != '']
    #     passenger_load_factor = [i.strip() for i in passenger_load_factor]
    #     passenger_load_factor = [i for i in passenger_load_factor if i!='']
    #     passenger_load_factor = [s for s in passenger_load_factor if not DEVANAGARI_START.match(s or "")]

    #     passenger_load_factor = [f"{passenger_load_factor[i]}: {passenger_load_factor[i+1]}"
    #             for i in range(0, len(passenger_load_factor)-1, 2)]

    #     plf = "<br>".join(passenger_load_factor)

    #     now = datetime.now().strftime('%d-%b-%y %H:%M')

    #     notification_email(now)
        
    # except Exception as e:
    #     now = datetime.now().strftime('%d-%b-%y %H:%M')
        
    #     error_email(now,e)