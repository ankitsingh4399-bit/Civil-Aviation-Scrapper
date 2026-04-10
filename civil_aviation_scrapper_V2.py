from bs4 import BeautifulSoup
import requests
import re
import warnings
import win32com
import win32com.client
import sys
import time
from datetime import datetime
import unicodedata

warnings.filterwarnings('ignore')



def sublists_between_hindi_word_markers(
    lst,
    include_pre_first=False,
    include_post_last=False,
    skip_empty=False,
    normalize_unicode=True,
):

    def norm(s):
        return unicodedata.normalize('NFC', s) if normalize_unicode else s

    devanagari_start_re = re.compile(r'^\s*[\u0900-\u097F]')

    def starts_with_hindi_word(item) -> bool:
        s = norm(str(item))
        return bool(devanagari_start_re.match(s))

    result = []
    last_marker_idx = None

    for i, item in enumerate(lst):
        if starts_with_hindi_word(item):
            if last_marker_idx is None:
                if include_pre_first and i > 0:
                    seg = lst[:i]
                    if not (skip_empty and len(seg) == 0):
                        result.append(seg)
            else:
                seg = lst[last_marker_idx + 1 : i]
                if not (skip_empty and len(seg) == 0):
                    result.append(seg)
            last_marker_idx = i

    if include_post_last:
        seg = lst[:] if last_marker_idx is None else lst[last_marker_idx + 1 :]
        if not (skip_empty and len(seg) == 0):
            result.append(seg)

    return result



def join_elements_pattern(list_of_lists):

    out = []

    for sub in list_of_lists:
        if not sub:
            continue

        sub_strs = [str(x).strip() for x in sub]

        if len(sub_strs) > 2:
            base = f"{sub_strs[0]}: {sub_strs[1]}"
            for elem in sub_strs[2:]:
                out.append(f"{base} {elem}")
                
        elif len(sub_strs) == 2:
            base = f"{sub_strs[0]}: {sub_strs[1]}"
            out.append(base)

    return out



def error_email(date_time,error):
    # Create an Outlook application instance
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Email details
    recipients = 'ankit.singh8@goindigo.in'
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

        domestic_traffic  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 domestic-traffic"}).get_text()
        domestic_traffic = [i for i in domestic_traffic.split('\n') if i !=' ' or i != '']
        domestic_traffic = [i.strip() for i in domestic_traffic]
        domestic_traffic = [i for i in domestic_traffic if i!='']
        domestic_traffic = join_elements_pattern(sublists_between_hindi_word_markers(domestic_traffic))

        domestic = "<br>".join(domestic_traffic)

        international_traffic  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 international-traffic"}).get_text()
        international_traffic = [i for i in international_traffic.split('\n') if i !=' ' or i != '']
        international_traffic = [i.strip() for i in international_traffic]
        international_traffic = [i for i in international_traffic if i!='']
        international_traffic = join_elements_pattern(sublists_between_hindi_word_markers(international_traffic))

        international = "<br>".join(international_traffic)

        on_time_performance  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 on-time-performance"}).get_text()
        on_time_performance = [i for i in on_time_performance.split('\n') if i !=' ' or i != '']
        on_time_performance = [i.strip() for i in on_time_performance]
        on_time_performance = [i for i in on_time_performance if i!='']
        on_time_performance = join_elements_pattern(sublists_between_hindi_word_markers(on_time_performance))

        otp = "<br>".join(on_time_performance)

        passenger_load_factor  = soup.find("div", attrs={'class':"views-element-container col-lg-4 col-md-6 col-sm-12 passenger-load-factor"}).get_text()
        passenger_load_factor = [i for i in passenger_load_factor.split('\n') if i !=' ' or i != '']
        passenger_load_factor = [i.strip() for i in passenger_load_factor]
        passenger_load_factor = [i for i in passenger_load_factor if i!='']
        passenger_load_factor = join_elements_pattern(sublists_between_hindi_word_markers(passenger_load_factor))

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
            x = run_scrapper()
        else:
            now = datetime.now().strftime('%d-%b-%y %H:%M')
            error_email(now, 'Limit (10) reached for number of iterations')
            sys.exit()     