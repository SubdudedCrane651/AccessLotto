import json
import numbers
import random
import os
import sys
import requests
import threading

#Best works with pypy3.10 for more performance including threads


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

global count
count = len(data)

global drawnumbers
drawnumbers = []

def PrintStatus():
    sys.stdout.write("|\r");sys.stdout.flush()
    sys.stdout.write("/\r");sys.stdout.flush()
    sys.stdout.write("\\\r");sys.stdout.flush()
    sys.stdout.write("|\r");sys.stdout.flush()
    sys.stdout.write("/\r");sys.stdout.flush()

def PickLottoNumbers(samenumber, total,numbers):

    count2 = 0
    rnd=0
    samenumber=1

    for number in numbers:

        if numbers[count2 - 1] == number and count2 != 0:

         rnd = random.randint(1,total)

         numbers[count2] = rnd

         rnd = random.randint(1,total)

         samenumber += 1

        else:

           samenumber -= 1

        count2 += 1

        numbers.sort()

    return numbers, samenumber

def PickNum(data,numbers,lotto,hits):
         num=0
         hit=0
         PickNumbers=True
         for pan in data:
         
                    PrintStatus()
         
                    hit=0

                    #Lotto 6/49 Drawings
                    if lotto == 1:
                        
                        for num in range(0,6):
                            if numbers[num] == int(pan["P1"]) or numbers[num] == int(pan["P2"]) \
                            or numbers[num] == int(pan["P3"]) or numbers[num] == int(pan["P4"]) \
                            or numbers[num] == int(pan["P5"]) or numbers[num] == int(pan["P6"]) \
                            or numbers[num] == int(pan["P7"]):
                             hit+=1
                        if hit >= 4:
                            PickNumbers=True
                            hits+=1
                        else:
                            PickNumbers=False

                    #LottoMax Drawings        
                    if lotto == 2:

                        for num in range(0,7):
                            if numbers[num] == int(pan["P1"]) or numbers[num] == int(pan["P2"]) \
                            or numbers[num] == int(pan["P3"]) or numbers[num] == int(pan["P4"]) \
                            or numbers[num] == int(pan["P5"]) or numbers[num] == int(pan["P6"] \
                            or numbers[num] == int(pan["P7"])):
                              hit+=1
                            if hit >3:
                               PickNumbers=True
                               hits+=1
                            else:
                               PickNumbers=False 

                    #Grande Vie Drawings
                    elif lotto == 3:

                        if numbers[0] == int(pan["p1"]) and numbers[1] == int(pan["p2"]) \
                        and numbers[2] == int(pan["p3"]) and numbers[3] == int(pan["p4"])\
                        and numbers[4] == int(pan["p5"]):
                            PickNumbers=True
                            hits=+1
                        else:
                            PickNumbers=False

                        if numbers[0] == int(pan["p1"]) and numbers[1] == int(pan["p2"]) \
                        and numbers[2] == int(pan["p3"]) and numbers[3] == int(pan["p4"]) \
                        and numbers[4] == int(pan["p5"]) and numbers[5] == int(pan["gn"]):
                            PickNumbers=True
                            hits=+1
                        else:
                            PickNumbers=False 

                        for num in range(0,5):
                            if numbers[num] == int(pan["p1"]) or numbers[num] == int(pan["p2"]) \
                            or numbers[num] == int(pan["p3"]) or numbers[num] == int(pan["p4"]):       
                                hit+=1
                            if hit >= 5:
                                PickNumbers=True
                                hits+=1
                            else:
                                PickNumbers=False
            
                    #Tout ou Rien Drawings
                    if lotto == 4:
                        for num in range(0,12):
                            if numbers[num] == int(pan["p1"]) or numbers[num] == int(pan["p2"]) \
                            or numbers[num] == int(pan["p3"]) or numbers[num] == int(pan["p4"]) \
                            or numbers[num] == int(pan["p5"]) or numbers[num] == int(pan["p6"]) \
                            or numbers[num] == int(pan["p7"]) or numbers[num] == int(pan["p8"]) \
                            or numbers[num] == int(pan["p9"]) or numbers[num] == int(pan["p10"]) \
                            or numbers[num] == int(pan["p11"]) or numbers[num] == int(pan["p12"]):          
                                hit+=1
                            if hit==12:
                                PickNumbers=True
                                hits+=1
                            else:
                                PickNumbers=False
         return PickNumbers,hits


class LottoDrawings(threading.Thread):

    def __init__(self, rangenum, drawingnum, same, drawnumbers):
        threading.Thread.__init__(self)
        self.rangenum = rangenum
        self.drawingnum = drawingnum
        self.same = same
        self.drawnumbers = drawnumbers

    def run(self):
        PickNumbers = True
        hits = 0

        while PickNumbers or hits > 0:
            numbers = []
            hits = 0

            for count in range(1, self.rangenum):
                rnd = random.randint(1, self.drawingnum)
                samenumber = 0
                numbers2 = PickLottoNumbers(samenumber, self.drawingnum, numbers)
                numbers2[0].append(rnd)
                numbers2[0].sort()
                samenumber = 0

            while samenumber != self.same:
                numbers2 = PickLottoNumbers(samenumber, self.drawingnum, numbers)
                samenumber = numbers2[1]
            numbers = numbers2[0]

            if self.same == -4:
                rnd = random.randint(1, 6)
                numbers.append(rnd)

            self.drawnumbers = numbers
            PickNumbers, hits = PickNum(data, numbers, lotto, hits)


# Lotto 6/49 Drawings
if lotto == 1:
    print(loto649)
    lottonumbers = LottoDrawings(7, 49, -5, drawnumbers)
    lottonumbers.start()
    lottonumbers.join()
    print("The winning 6/49 numbers are " + str(lottonumbers.drawnumbers) + " in a total of " + str(count) + " drawings")
    os.system('C:\\Users\\rchrd\\AppData\\Local\\Microsoft\\WindowsApps\\python3.9.exe C:\\Users\\rchrd\\Documents\\Python\\text2speech.py "--lang=fr" "Voici les numéros gagnants de Lotto 6/49"')
    input("<PRESS ENTER>")

# LottoMax Drawings
if lotto == 2:
    print(LottoMax)
    lottonumbers = LottoDrawings(8, 50, -6, drawnumbers)
    lottonumbers.start()
    lottonumbers.join()
    print("The LottoMax winning numbers are " + str(lottonumbers.drawnumbers) + " in a total of " + str(count) + " drawings")
    os.system('C:\\Users\\rchrd\\AppData\\Local\\Microsoft\\WindowsApps\\python3.9.exe C:\\Users\\rchrd\\Documents\\Python\\text2speech.py "--lang=fr" "Voici les numéros gagnants de Lotto Max"')
    input("<PRESS ENTER>")

# Grande Vie Drawings
if lotto == 3:
    print(GrandeVie)
    lottonumbers = LottoDrawings(6, 49, -4, drawnumbers)
    lottonumbers.start()
    lottonumbers.join()
    print("The winning Grande Vie numbers are " + str(lottonumbers.drawnumbers) + " in a total of " + str(count) + " drawings")
    os.system('C:\\Users\\rchrd\\AppData\\Local\\Microsoft\\WindowsApps\\python3.9.exe C:\\Users\\rchrd\\Documents\\Python\\text2speech.py "--lang=fr" "Voici les numéros gagnants de la Grande Vie"')
    input("<PRESS ENTER>")

# Tout ou Rien Drawings
if lotto == 4:
    print(ToutouRien)
    lottonumbers = LottoDrawings(13, 24, -11, drawnumbers)
    lottonumbers.start()
    lottonumbers.join()
    print("The winning Tout ou Rien numbers are " + str(lottonumbers.drawnumbers) + " in a total of " + str(count) + " drawings")
    os.system('C:\\Users\\rchrd\\AppData\\Local\\Microsoft\\WindowsApps\\python3.9.exe C:\\Users\\rchrd\\Documents\\Python\\text2speech.py "--lang=fr" "Voici les numéros gagnants de Tout ou rien"')
    input("<PRESS ENTER>")
