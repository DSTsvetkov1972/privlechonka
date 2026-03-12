from processors.source_checker import source_checker
from processors.prepared_maker import prepared_maker
from processors.sent_maker import sent_maker
from processors.fns import init_project
from logo import logo_colored


from colorama import Fore, Style, init

init()
print(Style.BRIGHT)
print(logo_colored)

while True:
    if not init_project():
        print(Fore.RED,
            'Название папки проекта не должно содеражать подчёркиваний!\n'
            ' Закройте программу и перименуйте папку проекта!',
            Fore.RESET)
        input()
    else:
        break
    
while True:
    try:
        print()
        print(Fore.BLUE, '1 - получить сводку по файлу "свод по каждому заказу.xlsx"', Fore.RESET)
        print(Fore.BLUE, '2 - подготовить файл передачи на актирование', Fore.RESET)
        print(Fore.BLUE, '3 - передать файл на актирование', Fore.RESET)
        choise = input("Ваш выбор: ")

        if choise == '1':
            source_checker()
        elif choise == '2':
            prepared_maker()
        elif choise == '3':
            sent_maker()
    except Exception as e:
        print(Fore.RED, str(e), Fore.RESET)