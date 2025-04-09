import time, socket, json, random, threading
import pandas as pd
from toolbox import generate_domain
from fuzzylogic.classes import Rule
from datetime import datetime, date
from pymodbus.client import ModbusSerialClient
from openpyxl import Workbook, load_workbook
from tkinter import *
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# === Global Config ===
client = ModbusSerialClient(port='/dev/ttyACM0', baudrate=9600)
host = '192.168.1.192'
port = 6000
INVENTORY = '04 FF 0F'
PRESET_Value = 0xFFFF
POLYNOMIAL = 0x8408
timeout = 30
stat = 0
listNull = ['05', '00', '0F', 'FB', 'E2', 'A7']
randomify = False  # MODE TEST

# === Domain Fuzzy ===
durasiOven = generate_domain("Durasi Oven (Menit)", 120, 1440, ["sangat_sebentar", "sebentar", "sedang", "lama", "sangat_lama"])
pastLub = generate_domain("Rentang Lubrikasi Terakhir (Jam)", 12, 672, ["sangat_sebentar", "sebentar", "sedang", "lama", "sangat_lama"])
lastLub = generate_domain("Durasi Lubrikasi Terakhir (Milisecond)", 50, 6000, ["tidak_spray", "sebentar", "sedang", "lama", "sangat_lama"])
lubDur = generate_domain("Durasi Lubrikasi (Milisecond)", 20, 6000, ["tidak_spray", "sebentar", "sedang", "lama", "sangat_lama"])

# === Aturan Fuzzy ===
df = pd.read_csv("rulebases.csv")
rule_dict = {(row["Durasi_Oven"], row["Lubrikasi_Terakhir"], row["Durasi_Terakhir"]): row["Durasi_Lubrikasi"] for _, row in df.iterrows()}
rules = [Rule({(getattr(durasiOven, do), getattr(pastLub, lt), getattr(lastLub, dt)): getattr(lubDur, dl)})
         for (do, lt, dt), dl in rule_dict.items()]
final_rule = sum(rules)

# === CRC untuk RFID ===
def crc(cmd):
    cmd = bytes.fromhex(cmd)
    crc_val = PRESET_Value
    for b in cmd:
        crc_val ^= b
        for _ in range(8):
            crc_val = (crc_val >> 1) ^ POLYNOMIAL if crc_val & 0x0001 else crc_val >> 1
    return cmd + bytes([(crc_val & 0xFF), (crc_val >> 8) & 0xFF])

# === Komunikasi RFID ===
def send_cmd(cmd):
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(2)
            s.connect((host, port))
            s.sendall(crc(cmd))
            data = s.recv(64)
            return [data.hex().upper()[i*2:i*2+2] for i in range(len(data))] if data else ['FF']
    except:
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
                print("RFID error")
            else:
                return int(result[6] + result[7], 16)
            last_state = result
        if time.time() - start_time > timeout:
            print("Timeout")
            return 999
        time.sleep(0.05)

# === Hitung Durasi Lubrikasi ===
last_data_json = {}

def calc_lub(idTag):
    global last_data_json
    filename = f"cart{idTag}.json"
    with open(filename, "r") as file:
        data = json.load(file)
    data["lastTS"] = datetime.fromisoformat(data["lastLubTS"])
    lastScan = round((datetime.now() - data["lastTS"]).total_seconds() / 60, 3)
    lastLubSpan = round((datetime.now() - datetime.fromisoformat(data["lastLubTS"])).total_seconds() / 3600, 3)
    lastLubDur = data["lastLubDur"]

    result = round(final_rule({durasiOven: lastScan, pastLub: lastLubSpan, lastLub: lastLubDur}))
    data["lastTS"] = datetime.now().isoformat()
    data["lastLubTS"] = data["lastTS"]
    data["lastLubDur"] = result if result >= 500 else 0

    with open(filename, "w") as file:
        json.dump(data, file, indent=4)

    last_data_json = {
        "Cart ID": idTag,
        "lastTS": data["lastTS"],
        "lastLubTS": data["lastLubTS"],
        "lastLubDur": data["lastLubDur"]
    }
    log_rfid(idTag, last_data_json)
    return result

# === Logging Excel ===
def init_excel():
    today = date.today().strftime("%Y-%m-%d")
    try:
        wb = load_workbook("log_data.xlsx")
    except FileNotFoundError:
        wb = Workbook()
    if today not in wb.sheetnames:
        ws1 = wb.create_sheet(title=today)
        ws1.append(["CartID", "lastTS", "lastLubTS", "lastLubDur"])
        ws2 = wb.create_sheet(title=f"{today}_power")
        ws2.append(["Timestamp", "Voltage", "Current"])
    wb.save("log_data.xlsx")

def log_rfid(cart_id, data):
    wb = load_workbook("log_data.xlsx")
    ws = wb[date.today().strftime("%Y-%m-%d")]
    ws.append([cart_id, data["lastTS"], data["lastLubTS"], data["lastLubDur"]])
    wb.save("log_data.xlsx")

def log_power(voltage, current):
    wb = load_workbook("log_data.xlsx")
    ws = wb[f"{date.today().strftime('%Y-%m-%d')}_power"]
    ws.append([datetime.now().isoformat(), voltage, current])
    wb.save("log_data.xlsx")

# === GUI & Realtime Chart ===
root = Tk()
root.title("Monitoring Tegangan & Arus")

status_label = Label(root, text="Status: ", font=("Arial", 14))
status_label.pack()

json_label = Label(root, text="Data JSON: ", font=("Courier", 10), justify=LEFT)
json_label.pack()

fig = Figure(figsize=(6, 3), dpi=100)
ax = fig.add_subplot(111)
line_v, = ax.plot([], [], label="Tegangan")
line_i, = ax.plot([], [], label="Arus")
ax.legend()
ax.set_ylim(0, 500)
ax.set_xlim(0, 100)
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack()

voltage_data, current_data = [], []

def update_graph():
    global stat
    try:
        getPLC = client.read_holding_registers(address=502, count=1, slave=1)
        stat = getPLC.registers[0]
    except:
        stat = 0
    status_label.config(text=f"Status Kereta: {'ADA' if stat == 1 else 'TIDAK'} | Mode: {'Random' if randomify else 'Real'}")

    if stat == 1:
        idTag = scan_rfid()
        if idTag != 999:
            dur = calc_lub(idTag)
            client.write_register(address=501, value=dur, slave=1)

    # Baca tegangan & arus
    if randomify:
        voltage = random.randint(200, 240)
        current = round(random.uniform(0.5, 2.5), 2)
    else:
        try:
            reg = client.read_holding_registers(address=0, count=2, slave=2)
            voltage = reg.registers[0] / 10.0
            current = reg.registers[1] / 100.0
        except:
            voltage, current = 0, 0

    voltage_data.append(voltage)
    current_data.append(current)
    if len(voltage_data) > 100:
        voltage_data.pop(0)
        current_data.pop(0)

    log_power(voltage, current)

    line_v.set_ydata(voltage_data)
    line_i.set_ydata(current_data)
    line_v.set_xdata(range(len(voltage_data)))
    line_i.set_xdata(range(len(current_data)))
    ax.relim()
    ax.autoscale_view()
    canvas.draw()

    json_label.config(text=f"Data Lubrikasi: {json.dumps(last_data_json, indent=2)}")
    root.after(1000, update_graph)

# === Main ===
init_excel()
update_graph()
root.mainloop()