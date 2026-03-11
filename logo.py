# Сгенерировано с помощью https://www.asciiart.eu/text-to-ascii-art
# ANSI Shadow

from colorama import Fore


logo = """
 ██████╗  █████╗ ██████╗  ██████╗ ███████╗  
 ██╔══██╗██╔══██╗██╔══██╗██╔════╝ ██╔════╝  
 ██████╔╝███████║██║  ██║██║  ███╗█████╗    
 ██╔══██╗██╔══██║██║  ██║██║   ██║██╔══╝    
 ██████╔╝██║  ██║██████╔╝╚██████╔╝███████╗  
 ╚═════╝ ╚═╝  ╚═╝╚═════╝  ╚═════╝ ╚══════╝  
                                            
 ███╗   ███╗ █████╗ ██╗  ██╗███████╗██████╗ 
 ████╗ ████║██╔══██╗██║ ██╔╝██╔════╝██╔══██╗
 ██╔████╔██║███████║█████╔╝ █████╗  ██████╔╝
 ██║╚██╔╝██║██╔══██║██╔═██╗ ██╔══╝  ██╔══██╗
 ██║ ╚═╝ ██║██║  ██║██║  ██╗███████╗██║  ██║
 ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝
"""   



version = " v.2025-11-17"

advertisement = """
 Нужна быстрая автоматизация или аналитика?
 Всегда рад помочь коллегам - обращайтесь!
 Сделаю без ТЗ, совещаний и миллионного бюджета!
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