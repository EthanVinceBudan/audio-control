from tkinter import ttk, Tk, StringVar, IntVar
from pystray import Icon, Menu, MenuItem
from PIL import Image, ImageDraw
from serial import Serial
from serial.tools.list_ports import comports as list_comports
from pycaw.pycaw import AudioUtilities, AudioDeviceState, IAudioEndpointVolume
from comtypes import CLSCTX_ALL

from threading import Thread, Event

class TrayApplication(Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.comport_var = StringVar()
        self.baud_var = IntVar()
        AudioDeviceUpdater.find_all_devices()
        self.serial_port = ThreadedPortReader(timeout=0.25)
        self.all_comports = {}

        self.protocol("WM_DELETE_WINDOW", self.withdraw)

        self.scan_comports()
        self.icon = self.create_tray_icon()
        self.create_window_content()

    def graceful_exit(self):
        self.serial_port.stop()
        self.serial_port.close()
        self.deiconify()
        self.quit()
        self.icon.stop()

    def scan_comports(self):
        self.all_comports = {}
        for c in list_comports():
            self.all_comports[str(c)] = c

    def start_comport(self):
        print(f"Opening COM port {self.comport_var.get()}")
        self.serial_port.port = self.all_comports[self.comport_var.get()].device
        self.serial_port.baudrate = self.baud_var.get()
        self.serial_port.open()
        self.serial_port.start()
        self.startButton.state(['disabled'])
        self.portSelect.state(['disabled'])
        self.baudSelect.state(['disabled'])
        self.stopButton.state(['!disabled'])

    def stop_comport(self):
        print("Closing COM port.")
        self.serial_port.stop()
        self.serial_port.close()
        self.stopButton.state(['disabled'])
        self.portSelect.state(['!disabled'])
        self.baudSelect.state(['!disabled'])
        self.startButton.state(['!disabled'])

    def create_image(self, width, height, color1, color2):
        image = Image.new('RGB', (width, height), color1)
        dc = ImageDraw.Draw(image)
        dc.rectangle(
            (width // 2, 0, width, height // 2),
            fill=color2
        )
        dc.rectangle(
            (0, height // 2, width // 2, height),
            fill=color2
        )
        return image

    def create_menu(self):
        exitItem = MenuItem("Exit", self.graceful_exit)
        showItem = MenuItem("Open", self.deiconify, default="True")
        return Menu(showItem, exitItem)

    def create_tray_icon(self):
        icon = self.create_image(64, 64, 'black', 'white')
        menu = self.create_menu()
        return Icon(self, title="AudioControl", icon=icon, menu=menu)

    def create_window_content(self):
        mainFrame = ttk.Frame(self)

        portLabel = ttk.Label(mainFrame, text="COM Port:")
        portLabel.grid(row=0, column=0)

        self.portSelect = ttk.Combobox(mainFrame, textvariable=self.comport_var,
                values=list(self.all_comports.keys()), width=50)
        self.portSelect.state(['readonly'])
        self.portSelect.current(0)
        self.portSelect.grid(row=0, column=1)

        baudLabel = ttk.Label(mainFrame, text="Baud Rate:")
        baudLabel.grid(row=1, column=0)

        self.baudSelect = ttk.Combobox(mainFrame, textvariable=self.baud_var,
                values=[1200,1800,2400,4800,9600,19200])
        self.baudSelect.state(['readonly'])
        self.portSelect.current(0)
        self.baudSelect.grid(row=1, column=1)

        self.startButton = ttk.Button(mainFrame, text="Start",
                command=self.start_comport)
        self.startButton.grid(row=2, column=0)

        self.stopButton = ttk.Button(mainFrame, text="Stop",
                command=self.stop_comport)
        self.stopButton.state(['disabled'])
        self.stopButton.grid(row=2, column=1)

        mainFrame.pack()

    def mainloop(self):
        self.icon.run_detached()
        super().mainloop()


class ThreadedPortReader(Serial):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.readEvent = Event()
        self.readThread = None

    def start(self):
        self.readEvent.set()
        self.readThread = Thread(target=self.continual_read, args=[self.readEvent])
        self.readThread.start()

    def stop(self):
        self.readEvent.clear()
        self.readThread.join()

    def continual_read(self, event):
        while event.is_set():
            line = self.readline()
            if line:
                print(line)
                AudioDeviceUpdater.set_device_volumes(AudioDeviceUpdater.extract_pot_values(line))


class AudioDeviceUpdater:
    devices = []
    interfaces = []
    dialMappings = (
        "Yeti Speakers",
        "Realtek Speakers"
    )

    def find_all_devices():
        activeDevices = [device for device in AudioUtilities.GetAllDevices()
                if device.state == AudioDeviceState.Active]
        AudioDeviceUpdater.devices = [AudioDeviceUpdater.find_device(activeDevices, dm) for dm in AudioDeviceUpdater.dialMappings]
        AudioDeviceUpdater.interfaces = [d._dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None).QueryInterface(IAudioEndpointVolume) for d in AudioDeviceUpdater.devices]
        print(f"Found devices: {[i.FriendlyName for i in AudioDeviceUpdater.devices]}")
        
    def find_device(deviceList, searchStr):
        searchTerms = searchStr.split()
        for d in deviceList:
            if all(term in d.FriendlyName for term in searchTerms):
                return d
        raise IndexError(f"Could not find device with search terms {searchTerms}")
        
    def extract_pot_values(rawBytes):
        extractedVals = list(map(int, rawBytes.split(b"|")))[:2]
        if len(extractedVals) != 2 or any(i == None for i in extractedVals):
            raise RuntimeError()
        return extractedVals

    def set_device_volumes(valueArray):
        for i, value in enumerate(valueArray):
            value = value/1024
            try:
                interface = AudioDeviceUpdater.interfaces[i]
                interface.SetMasterVolumeLevelScalar(round(value,3),None)

                if value < 6/1024:
                    interface.SetMute(True, None)
                else:
                    interface.SetMute(False, None)
            except:
                # device was most likely unplugged; skip
                return


if __name__ == "__main__":
    trayApp = TrayApplication()
    trayApp.title("AudioControl")
    trayApp.mainloop()
