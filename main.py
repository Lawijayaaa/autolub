import time, json, os
from datetime import datetime, date
import pandas as pd
import tkinter as tk
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from pymodbus.client import ModbusSerialClient
from fuzzylogic.classes import Rule
from toolbox import (
    generate_domain, send_cmd, log_power_to_excel, read_power_meter
)

# ===================== Konfigurasi =====================
client = ModbusSerialClient(port='/dev/ttyACM0', baudrate=9600)
INVENTORY = '04 FF 0F'
listNull = ['05', '00', '0F', 'FB', 'E2', 'A7']
timeout = 30

# Membership Functions
durasiOven = generate_domain("Durasi Oven (Menit)", 120, 1440, ["sangat_sebentar", "sebentar", "sedang", "lama", "sangat_lama"])
pastLub = generate_domain("Rentang Lubrikasi Terakhir (Jam)", 12, 672, ["sangat_sebentar", "sebentar", "sedang", "lama", "sangat_lama"])
lastLub = generate_domain("Durasi Lubrikasi Terakhir (Milisecond)", 50, 6000, ["tidak_spray", "sebentar", "sedang", "lama", "sangat_lama"])
lubDur = generate_domain("Durasi Lubrikasi (Milisecond)", 20, 6000, ["tidak_spray", "sebentar", "sedang", "lama", "sangat_lama"])

# Fuzzy Rules
df = pd.read_csv("rulebases.csv")
rule_dict = {(row["Durasi_Oven"], row["Lubrikasi_Terakhir"], row["Durasi_Terakhir"]): row["Durasi_Lubrikasi"] for _, row in df.iterrows()}
rules = [Rule({(getattr(durasiOven, a), getattr(pastLub, b), getattr(lastLub, c)): getattr(lubDur, d)})
         for (a, b, c), d in rule_dict.items()]
final_rule = sum(rules)

# ===================== Fungsi Utama =====================
def scan_rfid():
    start_time = time.time()
    last_state = ""
    while True:
        result = send_cmd(INVENTORY)
        if result != last_state:
            if result == listNull:
                print("Menunggu tag...")
            elif result == "ERROR":
                print("Koneksi ke RFID gagal.")
            else:
                resultInt = int((result[6] + result[7]), 16)
                return resultInt
            last_state = result
        time.sleep(0.04)
        if time.time() - start_time > timeout:
            print("Timeout RFID. Reset.")
            return 999

def calc_lub(idTag):
    strJsonName = f"cart{idTag}.json"
    with open(strJsonName, "r") as file:
        data = json.load(file)

    data["lastTS"] = datetime.fromisoformat(data["lastLubTS"])
    lastScan = round((datetime.now() - data["lastTS"]).total_seconds() / 60, 3)
    lastLubSpan = round((datetime.now() - datetime.fromisoformat(data["lastLubTS"])).total_seconds() / 3600, 3)
    lastLubDur = data["lastLubDur"]

    print(f"Input 1 : {lastScan} Menit")
    print(f"Input 2 : {lastLubSpan} Jam")
    print(f"Input 3 : {lastLubDur} ms")

    values = {durasiOven: lastScan, pastLub: lastLubSpan, lastLub: lastLubDur}
    result = round(final_rule(values))
    print(f"Output Lubrikasi: {result} ms")

    data["lastTS"] = datetime.now().isoformat()
    if result < 500:
        result = 0
        data["lastLubDur"] = 0
    else:
        data["lastLubDur"] = result
        data["lastLubTS"] = datetime.now().isoformat()

    with open(strJsonName, "w") as file:
        json.dump(data, file, indent=4)

    log_cart_to_excel(idTag, data)
    return result

def log_cart_to_excel(idTag, data):
    filename = f"cart_log_{date.today()}.xlsx"
    sheet = "CartEvent"
    log_data = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        idTag,
        data["lastTS"],
        data["lastLubTS"],
        data["lastLubDur"]
    ]

    try:
        import openpyxl
        wb = openpyxl.load_workbook(filename)
        if sheet in wb.sheetnames:
            ws = wb[sheet]
        else:
            ws = wb.create_sheet(sheet)
            ws.append(["Waktu", "CartID", "lastTS", "lastLubTS", "lastLubDur"])
    except FileNotFoundError:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet
        ws.append(["Waktu", "CartID", "lastTS", "lastLubTS", "lastLubDur"])

    ws.append(log_data)
    wb.save(filename)

# ===================== GUI =====================
class RealtimePlot:
    def __init__(self, root):
        self.root = root
        self.root.title("Monitoring Tegangan dan Arus")
        self.root.geometry("800x600")

        self.fig, self.axs = plt.subplots(2, 1, figsize=(6, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.root)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.timestamps, self.v_ab, self.i_a = [], [], []

        self.update_plot()

    def update_plot(self):
        data = read_power_meter(client, ct_ratio=150)
        if data:
            ts, vab, vbc, vca, ia, ib, ic = data
            self.timestamps.append(ts)
            self.v_ab.append(vab)
            self.i_a.append(ia)

            # Max data points
            if len(self.timestamps) > 60:
                self.timestamps = self.timestamps[-60:]
                self.v_ab = self.v_ab[-60:]
                self.i_a = self.i_a[-60:]

            self.axs[0].clear()
            self.axs[1].clear()
            self.axs[0].plot(self.timestamps, self.v_ab, label='V_AB (Volt)', color='blue')
            self.axs[1].plot(self.timestamps, self.i_a, label='I_A (Ampere)', color='red')
            self.axs[0].legend()
            self.axs[1].legend()
            self.axs[0].set_ylabel("Volt")
            self.axs[1].set_ylabel("Ampere")
            self.axs[1].set_xlabel("Waktu")

            self.fig.tight_layout()
            self.canvas.draw()

        log_power_to_excel(client, ct_ratio=150)
        self.root.after(1000, self.update_plot)

# ===================== Loop Utama =====================
def main_loop():
    prev_stat = 0
    while True:
        try:
            getPLC = client.read_holding_registers(address=502, count=1, slave=1)
            stat = getPLC.registers[0]
        except:
            stat = 0

        print(f"Status PLC: {stat}")
        if prev_stat == 0 and stat == 1:
            print("Status berubah ke 1")
            idTag = scan_rfid()
            if idTag != 999:
                durasi = calc_lub(idTag)
                client.write_register(address=501, value=durasi, slave=1)
        prev_stat = stat
        time.sleep(1)

# ===================== Start Program =====================
if __name__ == "__main__":
    import threading
    root = tk.Tk()
    app = RealtimePlot(root)
    thread = threading.Thread(target=main_loop, daemon=True)
    thread.start()
    root.mainloop()