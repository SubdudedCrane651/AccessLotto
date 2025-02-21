import threading
import random
import time
import pyodbc

# Function to generate a random combination
def generate_combination(task_id, results):
    pick12 = []
    while len(pick12) < 12:
        rnd_num = random.randint(1, 24)
        if rnd_num not in pick12:
            pick12.append(rnd_num)
    results[task_id] = sorted(pick12)

# Function to check if at least 8 or at most 4 numbers are different
def check_combination(pick12, P):
    count = sum(1 for i in range(len(pick12)) if pick12[i] != P[i])
    return count >= 8 or count <= 4

# Main function to handle the multithreading and processing
def tout_ou_rien_special():
    # Database connection
    conn_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        r"DBQ=C:/Users/rchrd/Documents/Richard/LottoDrawings.mdb;"
    )
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    cursor.execute("SELECT DrawDate, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12 FROM ToutouRien ORDER BY DrawDate ASC")
    rows = cursor.fetchall()

    P = [0] * 12
    ChangeCaption = 1

    results = [None] * 12
    threads = []

    # Create and start threads
    for i in range(12):
        thread = threading.Thread(target=generate_combination, args=(i, results))
        threads.append(thread)
        thread.start()

    # Wait for all threads to complete
    for thread in threads:
        thread.join()

    # Process results
    pick12 = results[0]  # Use the first result (or choose another logic)
    while not check_combination(pick12, P):
        threads = []
        for i in range(12):
            thread = threading.Thread(target=generate_combination, args=(i, results))
            threads.append(thread)
            thread.start()
        for thread in threads:
            thread.join()
        pick12 = results[0]

    with open("C:/Users/rchrd/Documents/Richard/test.txt", "w", encoding="utf-8") as fileout:
        for row in rows:
            DrawDate, *P = row

            if all(P[i] == pick12[i] for i in range(12)):
                fileout.write(f"{DrawDate}\n")
            else:
                count = sum(1 for i in range(12) if str(pick12[i]) in map(str, P))
                if count > 7:
                    fileout.write(f"{DrawDate}\n")
                    identical = True
                else:
                    identical = False

    # Call the text-to-speech Python script
    import subprocess
    subprocess.run(["python3", "C:/Users/rchrd/Documents/Python/text2speech.py", "--lang=fr", "Voici les numéros gagnants de Tout ou rien"])

    # Display results
    result_str = " ".join(map(str, pick12))
    print(result_str)

if __name__ == "__main__":
    tout_ou_rien_special()
