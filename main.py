import datetime
from distutils.command.build import build
from bs4 import BeautifulSoup
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import re
import time
from colorama import init
from colorama import Fore
import xlsxwriter
init()

def main():
    while True:
        keys = input(Fore.WHITE +
            'Выберите файл с лицензиями:\nДля продолжения нажмите любую клавишу'
        )
        Tk().withdraw()
        html_license = askopenfilename()
        #os.system(os.path.realpath(html_license))
        expired_licese = []
        excel_license = [] # для форматирования excel
        try:
            with open(html_license, 'r') as f:
                tmp = f.read()
                soap = BeautifulSoup(tmp, 'html.parser')
                table = soap.find('table')
                rows = table.find_all('tr')
                data = []
                for row in rows:
                    cols = row.find_all('td')
                    cols = [val.text.strip() for val in cols]
                    data.append(cols)
                print(data[0])
                def complete_date(value: str) -> datetime:
                    match_week_mounth = re.findall(r'[A-Za-z]{3}', value)
                    match_day_time = re.findall(r'\d{1,2}', value) 
                    match_year = re.findall(r'\d{4}', value)
                    #print(match_week_mounth, match_day_time, match_year)
                    build_date = datetime.datetime(
                        day=int(match_day_time[0]),
                        month=int(time.strptime(match_week_mounth[1], '%b').tm_mon),
                        year=int(match_year[0]),
                        hour=int(match_day_time[1]),
                        minute=int(match_day_time[2]),
                        second=int(match_day_time[3])
                    )
                    return build_date

                for line in data:
                    domain_name = line[1]
                    match_license_date = re.findall(r' [A-za-z]{3} [A-za-z]{3} {1,2}\d{0,2} \d{1,2}:\d{1,2}:\d{1,2} \d{4}', line[2])
                    #print(domain_name, match_license_date)
                    
                    try:
                        #print(complete_date(match_license_date[0].strip()))
                        result = complete_date(match_license_date[1].strip()) -datetime.datetime.now()
                        if result.days <=30:
                            print(Fore.RED + 'Срок лицензии истекает: {}\nДаты: {}\nОсталось:{} дней\n'.format(
                                domain_name, match_license_date, result.days
                            ))
                            expired_licese.append('Срок лицензии истекает: {}\nДаты: {}\nОсталось:{} дней\n'.format(
                                domain_name, match_license_date, result.days
                            ).encode('utf-8'))
                            excel_license.append([domain_name, match_license_date, result.days])
                    except IndexError as e:
                        print('Нет лицензии!', e)
                with open('result.txt', 'wb') as file:
                    for val in expired_licese:
                        file.write(val)
                        file.write('{}\n'.format('-'*20).encode('utf-8'))
                    file.write('Количество истекающих лицензий: {}'.format(len(expired_licese)).encode('utf-8'))
                try:
                    workbook = xlsxwriter.Workbook('result.xlsx')
                    worksheet = workbook.add_worksheet()
                    for i in range(len(expired_licese)):
                        worksheet.write('A{}'.format(i+1), excel_license[i][0])
                        worksheet.write('B{}'.format(i+1), f'{excel_license[i][1]}')
                        worksheet.write('C{}'.format(i+1), excel_license[i][2])
                    workbook.close()
                except Exception:
                    print('Файл result.xlsx занят! Закройте файл и перезапустите программу.')
            print(Fore.YELLOW + 'Количество истекающих лицензий: {}'.format(Fore.WHITE + str(len(expired_licese))))
            print(Fore.GREEN + 'Результаты выполнения храняться в файле result.txt рядом с программой\n')            
        except FileNotFoundError:
            print('Файл не найден или не существует! Проверьте путь к файлу.')

main()