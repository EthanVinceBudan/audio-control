import serial
from pycaw.pycaw import AudioUtilities, AudioDeviceState, IAudioEndpointVolume
from comtypes import CLSCTX_ALL

DIAL_MAPPINGS = (
    "Yeti Speakers",
    "Realtek Speakers"
)

ARDUINO_PORT = 'COM3'
BAUD_RATE = "9600"
NUM_DIALS = len(DIAL_MAPPINGS)


# find device based on given search term
# inputs: list of devices, search string
def find_device(deviceList, searchStr):
    searchTerms = searchStr.split()
    for d in deviceList:
        if all(term in d.FriendlyName for term in searchTerms):
            return d
    raise IndexError(f"Could not find device with search terms {searchTerms}")


def extract_pot_values(rawBytes):
    extractedVals = list(map(int, rawBytes.split(b"|")))[:NUM_DIALS]
    if len(extractedVals) != NUM_DIALS or any(i == None for i in extractedVals):
        raise RuntimeError()
    return extractedVals


def set_device_volume(interface, value):
    try:
        interface.SetMasterVolumeLevelScalar(round(value,3),None)

        if value < 6/1024:
            interface.SetMute(True, None)
        else:
            interface.SetMute(False, None)
    except:
        # device was most likely unplugged; skip
        return


if __name__ == "__main__":
    # find devices based on 'friendly name'
    devices = []
    interfaces = []
    activeDevices = [device for device in AudioUtilities.GetAllDevices()
                if device.state == AudioDeviceState.Active]

    for s in DIAL_MAPPINGS:
        foundDevice = find_device(activeDevices, s)
        devices.append(foundDevice)
        interface = foundDevice._dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        interfaces.append(interface.QueryInterface(IAudioEndpointVolume))

    # begin reading from arduino
    with serial.Serial(port=ARDUINO_PORT, baudrate=BAUD_RATE, timeout=1) as ser:
        ser.readline()
        while True:
            try:
                line = ser.readline()
                print(f"Received: {line}")
                if line == b'':
                    continue
                vals = extract_pot_values(line)
                for i in range(NUM_DIALS):
                    set_device_volume(interfaces[i],vals[i]/1024)
            except KeyboardInterrupt:
                print("\nCtl+C detected, stopping...")
                break
            except RuntimeError:
                print("\nIncomplete serial message detected, stopping...")
                break
