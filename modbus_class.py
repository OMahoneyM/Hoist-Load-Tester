from pymodbus.client import ModbusTcpClient

class ModbusConnect:
    def __init__(self, ip):
        # set ip address of networked device
        self.ip = ip

    def client_connect(self):
        try:
            # attempt to connect to networked device @ ip address
            ModbusTcpClient(self.ip).connect()
        except Exception as e:
            print(f"An error occured attempting to connect to ${self.ip}: {e}")
    
    def client_close(self):
        # close connection to networked device
        ModbusTcpClient(self.ip).close()
