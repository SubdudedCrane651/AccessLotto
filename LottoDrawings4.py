import json
import random
import os
import sys
import requests
from concurrent.futures import ThreadPoolExecutor

# Best works with pypy3.10 for more performance including threads

os.system('cls')

def choose(i):
    switcher = {
        1: 'Lotto649.json',
        2: 'LottoMax.json',
        3: 'Grande_Vie.json',
        4: 'ToutouRien.json'
    }
    return switcher.get(i, 'Invalid Number')

loto649 = """
  __     ___  _   ___  
 / /_   / / || | / _ \\ 
| '_ \\ / /| || || (_) |
| (_) / / |__   _\\__, |
 \\___/_/     |_|   /_/ 
 
"""

LottoMax = """
 _          _   _        __  __            
| |    ___ | |_| |_ ___ |  \\/  | __ ___  __
| |   / _ \\| __| __/ _ \\| |\\/| |/ _` \\ \\/ /
| |__| (_) | |_| || (_) | |  | | (_| |>  < 
|_____\\___/ \\__|\\__\\___/|_|  |_|\\__,_/_/\\_\\
                                           
"""

GrandeVie = """
  ____                     _       __     ___      
 / ___|_ __ __ _ _ __   __| | ___  \\ \\   / (_) ___ 
| |  _| '__/ _` | '_ \\ / _` |/ _ \\  \\ \\ / /| |/ _ \\
| |_| | | | (_| | | | | (_| |  __/   \\ V / | |  __/
 \\____|_|  \\__,_|_| |_|\\__,_|\\___|    \\_/  |_|\\___| 

"""

ToutouRien = """
 _____           _                   ____  _            
|_   _|__  _   _| |_    ___  _   _  |  _ \\(_) ___ _ __  
  | |/ _ \\| | | | __|  / _ \\| | | | | |_) | |/ _ \\ '_ \\ 
  | | (_) | |_| | |_  | (_) | |_| | |  _ <| |  __/ | | |
  |_\\___/ \\__,_|\\__|  \\___/ \\__,_| |_| \\_\\_|\___|_| |_|
                          
"""                   

try:
    if sys.argv[1] != "":
        lotto = int(sys.argv[1])

except:
    print("""       1 = Lotto 6/49
       2 = LottoMax
       3 = Grande Vie
       4 = Tout Ou Rien""")
    print()
    lotto = int(input("What Lotto numbers do you wish to pick: "))

jsonfile = choose(lotto)

url = "https://richard-perreault.com/Documents/" + jsonfile

response = requests.get(url)
data = json.loads(response.text)

count = len(data)
drawnumbers = []

def PrintStatus():
    sys.stdout.write("|\r"); sys.stdout.flush()
    sys.stdout.write("/\r"); sys.stdout.flush()
    sys.stdout.write("\\\r"); sys.stdout.flush()
    sys.stdout.write("|\r"); sys.stdout.flush()
    sys.stdout.write("/\r"); sys.stdout.flush()

def PickLottoNumbers(samenumber, total, numbers):
    seen = set(numbers)
    while samenumber:
        rnd = random.randint(1, total)
        if rnd not in seen:
            seen.add(rnd)
            numbers.append(rnd)
            samenumber -= 1
    numbers.sort()
    return numbers

def PickNum(data, numbers, lotto):
    hits = 0
    PickNumbers = True
    numbers_set = set(numbers)
    
    for pan in data:
        PrintStatus()

        # Common set operations for various lotto types
        drawn_numbers = {
            1: {int(pan.get("P1", 0)), int(pan.get("P2", 0)), int(pan.get("P3", 0)), int(pan.get("P4", 0)), int(pan.get("P5", 0)), int(pan.get("P6", 0))},
            2: {int(pan.get("P1", 0)), int(pan.get("P2", 0)), int(pan.get("P3", 0)), int(pan.get("P4", 0)), int(pan.get("P5", 0)), int(pan.get("P6", 0)), int(pan.get("P7", 0))},
            3: {int(pan.get("p1", 0)), int(pan.get("p2", 0)), int(pan.get("p3", 0)), int(pan.get("p4", 0)), int(pan.get("p5", 0)), int(pan.get("gn", 0))},
            4: {int(pan.get("p1", 0)), int(pan.get("p2", 0)), int(pan.get("p3", 0)), int(pan.get("p4", 0)), int(pan.get("p5", 0)), int(pan.get("p6", 0)), int(pan.get("p7", 0)), int(pan.get("p8", 0)), int(pan.get("p9", 0)), int(pan.get("p10", 0)), int(pan.get("p11", 0)), int(pan.get("p12", 0))}
        }[lotto]
        
        hit = len(numbers_set & drawn_numbers)
        
        if (lotto == 1 and hit >= 4) or (lotto == 2 and hit >= 4) or (lotto == 3 and (hit == 5 or hit == 6)) or (lotto == 4 and (hit == 12 or hit >= 8 or hit <= 4)):
            PickNumbers = True
            hits += 1
        else:
            PickNumbers = False
    return PickNumbers, hits

def lotto_drawings(rangenum, drawingnum, same, drawnumbers):
    while True:
        numbers = []
        numbers = PickLottoNumbers(same, drawingnum, numbers)
        PickNumbers, hits = PickNum(data, numbers, lotto)
        if PickNumbers:
            break
    return numbers

# Use ThreadPoolExecutor for better thread management
with ThreadPoolExecutor(max_workers=4) as executor:
    if lotto == 1:
        print(loto649)
        lottonumbers = executor.submit(lotto_drawings, 7, 49, 6, drawnumbers).result()
        os.system('C:\\Users\\rchrd\\AppData\\Local\\Microsoft\\WindowsApps\\python3.9.exe C:\\Users\\rchrd\\Documents\\Python\\text2speech.py "--lang=fr" "Voici les numéros gagnants de Lotto 6/49"')
        print(f"The winning 6/49 numbers are {lottonumbers} in a total of {count} drawings")

    if lotto == 2:
        print(LottoMax)
        lottonumbers = executor.submit(lotto_drawings, 8, 50, 7, drawnumbers).result()
        os.system('C:\\Users\\rchrd\\AppData\\Local\\Microsoft\\WindowsApps\\python3.9.exe C:\\Users\\rchrd\\Documents\\Python\\text2speech.py "--lang=fr" "Voici les numéros gagnants de Lotto Max"')
        print(f"The LottoMax winning numbers are {lottonumbers} in a total of {count} drawings")

    if lotto == 3:
        print(GrandeVie)
        lottonumbers = executor.submit(lotto_drawings, 6, 49, 5, drawnumbers).result()
        os.system('C:\\Users\\rchrd\\AppData\\Local\\Microsoft\\WindowsApps\\python3.9.exe C:\\Users\\rchrd\\Documents\\Python\\text2speech.py "--lang=fr" "Voici les numéros gagnants de la Grande Vie"')
        print(f"The winning Grande Vie numbers are {lottonumbers} in a total of {count} drawings")

    if lotto == 4:
        print(ToutouRien)
        lottonumbers = executor.submit(lotto_drawings, 13, 24, 12, drawnumbers).result()
        os.system('C:\\Users\\rchrd\\AppData\\Local\\Microsoft\\WindowsApps\\python3.9.exe C:\\Users\\rchrd\\Documents\\Python\\text2speech.py "--lang=fr" "Voici les numéros gagnants de Tout ou rien"')
        print(f"The winning Tout ou Rien numbers are {lottonumbers} in a total of {count} drawings")

input("<PRESS ENTER>")
