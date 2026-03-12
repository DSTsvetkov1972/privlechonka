# Сгенерировано с помощью https://www.asciiart.eu/text-to-ascii-art
# ANSI Shadow

from colorama import Fore


logo = Fore.BLACK + """
********************************************************************************
                    ПОДГОТОВКА К ОТПРАВКЕ НА АКТИРОВАНИЕ
                         ПРИВЛЕЧЕННЫХ КОНТЕЙНЕРОВ
********************************************************************************
""" + Fore.RESET



version = " v.2026-03-11\n"

advertisement = """
 Нужна быстрая автоматизация или аналитика
 без ТЗ, совещаний, бюрократии и миллионного бюджета?
 Обращайтесь:
 Цветков Дмитрий Сергеевич (ЦКП, Москва)
 TsvetkovDS@trcont.ru
 +7-903-617-77-55
"""

advertisement_colored = Fore.YELLOW + advertisement + Fore.RESET

logo_colored = ''
for ch in list(logo):
    if ch in ['╝', '═', '║', '╔', '╚', '╗']:
        ch_colored = Fore.MAGENTA+ ch + Fore.RESET
    elif ch in ('█', '▄', '▀'):    
        ch_colored = Fore.CYAN + chr(9619) + Fore.RESET
    else:
        ch_colored = ch
    logo_colored += ch_colored

version_colored = Fore.MAGENTA + version + Fore.RESET    

logo_colored = logo_colored + version_colored + advertisement_colored