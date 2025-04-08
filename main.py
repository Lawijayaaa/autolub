import time, socket, json
import pandas as pd
from toolbox import generate_domain
from fuzzylogic.classes import Rule
from datetime import datetime

host = '192.168.1.192'
port = 6000

INVENTORY = '04 FF 0F'

PRESET_Value = 0xFFFF
POLYNOMIAL = 0x8408

timeout = 30

listNull = ['05', '00', '0F', 'FB', 'E2', 'A7']

# Membership Function
durasiOven = generate_domain("Durasi Oven (Menit)", 120, 1440, ["sangat_sebentar", "sebentar", "sedang", "lama", "sangat_lama"])
pastLub = generate_domain("Rentang Lubrikasi Terakhir (Jam)", 12, 672, ["sangat_sebentar", "sebentar", "sedang", "lama", "sangat_lama"])
lastLub = generate_domain("Durasi Lubrikasi Terakhir (Milisecond)", 50, 6000, ["tidak_spray", "sebentar", "sedang", "lama", "sangat_lama"])
lubDur = generate_domain("Durasi Lubrikasi (Milisecond)", 20, 6000, ["tidak_spray", "sebentar", "sedang", "lama", "sangat_lama"])

# Aturan fuzzy dari file CSV
df = pd.read_csv("rulebases.csv")
rule_dict = {(row["Durasi_Oven"], row["Lubrikasi_Terakhir"], row["Durasi_Terakhir"]): row["Durasi_Lubrikasi"] for _, row in df.iterrows()}

rules = []
for (dur_oven, lub_terakhir, dur_terakhir), dur_lubrikasi in rule_dict.items():
    rules.append(Rule({(getattr(durasiOven, dur_oven), 
                        getattr(pastLub, lub_terakhir),
                        getattr(lastLub, dur_terakhir)): 
                        getattr(lubDur, dur_lubrikasi)}))

final_rule = sum(rules)

def crc(cmd):
    cmd = bytes.fromhex(cmd)
    viCrcValue = PRESET_Value
    for x in cmd:
        viCrcValue ^= x
        for _ in range(8):
            if viCrcValue & 0x0001:
                viCrcValue = (viCrcValue >> 1) ^ POLYNOMIAL
            else:
                viCrcValue >>= 1
    crc_H = (viCrcValue >> 8) & 0xFF
    crc_L = viCrcValue & 0xFF
    return cmd + bytes([crc_L, crc_H])

def send_cmd(cmd):
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(2)
            s.connect((host, port))
            message = crc(cmd)
            s.sendall(message)
            data = s.recv(64)
            if not data:
                return ""
            response_hex = data.hex().upper()
            if len(response_hex)%2 == 0 :
                output = []
                for i in range(0, int(len(response_hex)/2)):
                    output.append(response_hex[i*2: i*2 + 2])
            else:
                output = ['FF']
            return output
    except (socket.timeout, socket.error) as e:
        return "ERROR"
    
def scan_rfid():
    start_time = time.time()
    last_state = ""
    while True:
        result = send_cmd(INVENTORY)
        if result != last_state:
            if result == listNull:
                print("Menunggu tag...")
            elif result == "ERROR":
                print("Kesalahan koneksi ke RFID reader.")
            else:                
                resultInt = int((result[6] + result[7]), 16)
                return resultInt
            last_state = result
        time.sleep(0.04)
        if time.time() - start_time > timeout:
            print("Timeout terjadi. Mereset program...")
            return 999
        
def calc_lub(idTag):
    strJsonName = "cart" + str(idTag) + ".json"
    with open(strJsonName, "r") as file:
        data = json.load(file)   
    #INPUT 1 RENTANG WAKTU SEJAK DETEKSI TERAKHIR (DALAM MENIT)
    data["lastTS"] = datetime.fromisoformat(data["lastLubTS"])
    lastScan = (round(((datetime.now() - data["lastTS"]).total_seconds())*100))/100
    lastScan = (round(lastScan / 0.06))/1000
    print(f"Input 1 : {lastScan} Menit")
    #INPUT 2 RENTANG WAKTU SEJAK LUBRIKASI TERAKHIR (DALAM JAM)
    data["lastLubTS"] = datetime.fromisoformat(data["lastLubTS"])
    lastLubSpan = (round((((datetime.now() - data["lastLubTS"]).total_seconds())*100)))/100
    lastLubSpan = (round(lastLubSpan / 3.6))/1000
    print(f"Input 2 : {lastLubSpan} Jam")
    #INPUT 3 DURASI LUBRIKASI TERAKHIR (DALAM MILIDETIK)
    lastLubDur = data["lastLubDur"]
    print(f"Input 3 : {lastLubDur} MiliDetik")
    #OUTPUT  DURASI LUBRIKASI AKTUATOR (DALAM MILIDETIK)
    values = {durasiOven: lastScan, pastLub: lastLubSpan, lastLub: lastLubDur}
    result = round(final_rule(values))
    print(f"Output untuk kereta {idTag}: {result} Milidetik")
    if result < 500:
        result = 0
        data["lastLubTS"] = data["lastLubTS"].isoformat()
        data["lastTS"] = datetime.now().isoformat()
        data["lastLubDur"] = 0          
    else:
        data["lastLubTS"] = datetime.now().isoformat()
        data["lastTS"] = datetime.now().isoformat()            
        data["lastLubDur"] = result
    with open(strJsonName, "w") as file:
        json.dump(data, file, indent=4)
    return result

def main():
    while True:
        prox_stat = input("Tekan 1 untuk melanjutkan")
        if prox_stat == "1":
            idTag = scan_rfid()
            lubDur = calc_lub(idTag)
            print(lubDur)

main()