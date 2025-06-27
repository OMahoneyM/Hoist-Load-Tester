import struct
import time
from pymodbus.client import ModbusTcpClient

def read_first_six_3000_parameters(ip, unit=1):
    """
    Reads the first six 3X (input) parameters, where each parameter is stored in two consecutive 16-bit registers.
    Returns a dictionary with parameter names and their 32-bit values.
    """

    # attempt to connect to Waveshare
    try: 
        client = ModbusTcpClient(ip)
        client.connect()

    except:
        print(f"Connection to {ip} failed.")
        return {}

    # try to read registers
    try:      
        # set list of parameters to measure
        names = ["volts_1", "volts_2", "volts_3", "current_1", "current_2", "current_3"]

        # declare empty dictionary to house register read results
        params_raw = {}
        
        # declare empty dictionary with max values
        params_max = {}

        # create loop to read Modbus 15 times in 3 secs
        for i in range(15):

            # store registers from Modbus
            result = client.read_input_registers(0, count=12)

            # check if result retrieved valid response
            if not result or not hasattr(result, "registers") or result.registers is None or len(result.registers) != 12:
                print(f"Read failed: {result}")
                return {}

            # iterate over names list and push results to params_raw dictionary
            for j, name in enumerate(names):
                high = result.registers[j * 2]
                low = result.registers[j * 2 + 1]
                raw = (high << 16) | low
                bytes_ = raw.to_bytes(4, byteorder='big')
                value = struct.unpack('>f', bytes_)[0]

                # check if params_raw[name] exists. If not create it or append value
                params_raw.setdefault(name, []).append(value)
                
            # put program to sleep for 200ms
            time.sleep(0.2)

        # populate params_max dictionary with largest element in key list
        for key, value in params_raw.items():
            params_max.update({ key : max(value) })

        return params_max
    
    except Exception as e:
        print(f"Exception during Modbus read: {e}")
        return {}
    
    finally:
        client.close()