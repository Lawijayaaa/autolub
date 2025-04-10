from pymodbus.client import ModbusSerialClient
import time

#init modbus device
client = ModbusSerialClient(port='/dev/ttyACM0', baudrate=9600)
joni = False

def testBatch():
    #getElect = client.read_holding_registers(0, 29,
    getPLC = client.read_holding_registers(address=502, count=1, slave=1)
    #writePLC = client.write_register(address=501, value=6000, slave=1)
    #print(getElect.registers)
    print(getPLC.registers)
    #print(writePLC)
    print("~~~")

#Loop
if joni:
    while True:
        testBatch()
        time.sleep(2)
        
else:   
    testBatch()
