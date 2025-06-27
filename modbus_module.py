import struct
import time
from PySide6.QtCore import Signal, QObject
from pymodbus.client import ModbusTcpClient


class ModbusConnect(ModbusTcpClient):
    def __init__(self, host):
        super().__init__(host)
        # set ip address of networked device
        #self.ip = ip

    def client_connect(self):
        try:
            # attempt to connect to networked device @ ip address
            ModbusTcpClient(self.host).connect()
        except Exception as e:
            print(f"An error occured attempting to connect to ${self.host}: {e}")
    
    def client_close(self):
        # close connection to networked device
        ModbusTcpClient(self.host).close()


class ReadRegisterWorker(QObject):
    progress_updated = Signal(int)
    amp_readings = Signal(dict)
    error_catch = Signal(bool, Exception)
    finished = Signal()
    # set list of parameters to measure
    names = ["volts_1", "volts_2", "volts_3", "current_1", "current_2", "current_3"]

    def __init__(self, client, params_raw={}, params_max={}, parent=None):
        self.client = client
        self.params_raw = params_raw
        self.params_max = params_max
        super().__init__(parent)
        self._is_running = True

    def read_registers(self):
        try:
            total_iterations = 15

            for i in range(total_iterations):
                # store registers from Modbus
                result = self.client.read_input_registers(0, count=12)
                
                # check if result retrieved valid response
                if not result or not hasattr(result, "registers") or result.registers is None or len(result.registers) != 12:
                    print(f"Read failed: {result}")
                    return {}
                
                # iterate over names list and push results to params_raw dictionary
                for j, name in enumerate(self.names):
                    high = result.registers[j * 2]
                    low = result.registers[j * 2 + 1]
                    raw = (high << 16) | low
                    bytes_ = raw.to_bytes(4, byteorder='big')
                    value = struct.unpack('>f', bytes_)[0]

                    # check if params_raw[name] exists. If not append value
                    self.params_raw.setdefault(name, []).append(value)
            
                # signal an update to progress
                self.progress_updated.emit(i)
                # put program to sleep for 200ms
                time.sleep(0.2)

            # populate params_max dictionary with largest element in key list
            for key, value in self.params_raw.items():
                self.params_max.update({ key : max(value) })
            
            print("Successfully read registers!")
            # signal the output
            self.amp_readings.emit(self.params_max)
            # signal the job is done
            self.finished.emit()

        except Exception as e:
            self.error_catch.emit(True, e)
            print(f"Exception during Modbus read: {e}")

    def stop(self):
        self._is_running = False
