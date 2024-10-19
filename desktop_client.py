from tkinter import ttk, Tk, StringVar, IntVar, HORIZONTAL
from pystray import Icon, Menu, MenuItem
from PIL import Image
from serial import Serial
from serial.tools.list_ports import comports as list_comports
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from pycaw.constants import EDataFlow, DEVICE_STATE
from comtypes import CLSCTX_ALL, COMError
from threading import Thread, Event

class TrayApplication(Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.iconbitmap('icon.ico')
        self.comport_var = StringVar()
        self.baud_var = IntVar()
        self.device_vars = [StringVar() for i in range(4)]
        self.all_comports = self.scan_comports()
        self.all_audio_devices = self.scan_audio_devices()

        self.protocol("WM_DELETE_WINDOW", self.withdraw)

        self.deviceUpdater = DeviceUpdater(self.all_audio_devices.values())
        self.serial_port = ThreadedPortReader(self.deviceUpdater, timeout=1)

        self.icon = self.create_tray_icon()
        self.create_window_content()

    def graceful_exit(self):
        self.serial_port.stop()
        self.serial_port.close()
        self.deiconify()
        self.quit()
        self.icon.stop()

    def scan_comports(self):
        return {str(c): c for c in list_comports()}

    def scan_audio_devices(self):
        deviceEnumerator = AudioUtilities.GetDeviceEnumerator()

        deviceCollection = deviceEnumerator.EnumAudioEndpoints(
            EDataFlow.eRender.value, DEVICE_STATE.ACTIVE.value
        )
        activeDevices = []
        for i in range(deviceCollection.GetCount()):
            dev = deviceCollection.Item(i)
            if dev is not None:
                activeDevices.append(AudioUtilities.CreateDevice(dev))

        return {d.FriendlyName: d for d in activeDevices}

    def start_controlling(self):
        print(f"Opening COM port {self.comport_var.get()}")
        self.serial_port.port = self.all_comports[self.comport_var.get()].device
        self.serial_port.baudrate = self.baud_var.get()
        self.deviceUpdater.select_devices([v.get() for v in self.device_vars])
        self.serial_port.open()
        self.serial_port.start()

        for cb in self.deviceCbList:
            cb.state(['disabled'])
        self.startButton.state(['disabled'])
        self.refreshButton.state(['disabled'])
        self.portSelect.state(['disabled'])
        self.baudSelect.state(['disabled'])
        self.stopButton.state(['!disabled'])

    def stop_comport(self):
        print("Closing COM port.")
        self.serial_port.stop()
        self.serial_port.close()

        for cb in self.deviceCbList:
            cb.state(['!disabled'])
        self.stopButton.state(['disabled'])
        self.refreshButton.state(['!disabled'])
        self.portSelect.state(['!disabled'])
        self.baudSelect.state(['!disabled'])
        self.startButton.state(['!disabled'])

    def refresh_device_list(self):
        self.all_audio_devices = self.scan_audio_devices()
        for cb in self.deviceCbList:
            cb['values'] = ["None"] + list(self.all_audio_devices.keys())
            if cb.get() not in cb['values']:
                cb.current(0)
        self.deviceUpdater = DeviceUpdater(self.all_audio_devices.values())
        self.serial_port = ThreadedPortReader(self.deviceUpdater, timeout=1)

    def create_menu(self):
        exitItem = MenuItem("Exit", self.graceful_exit)
        showItem = MenuItem("Open", self.deiconify, default="True")
        return Menu(showItem, exitItem)

    def create_tray_icon(self):
        with Image.open('icon.ico') as im:
            menu = self.create_menu()
            return Icon(self, title="AudioControl", icon=im, menu=menu)

    def create_window_content(self):
        mappingFrame = ttk.Labelframe(self, text="Device Mapping")
        mappingFrame.grid(row=0,column=0, sticky="NSEW", padx=5, pady=5)

        self.deviceCbList = []
        for i in range(4):
            cbLabel = ttk.Label(mappingFrame, text=f"D{i}:")
            cbLabel.grid(row=i,column=0, padx=3)
            cb = ttk.Combobox(mappingFrame, textvariable=self.device_vars[i],
                    values = ["None"] + list(self.all_audio_devices.keys()), width=50)
            cb.state(['readonly'])
            cb.grid(row=i,column=1, pady=0 if i != 3 else (0,5))
            cb.current(0)
            self.deviceCbList.append(cb)

        optionsFrame = ttk.Labelframe(self, text="Options")
        optionsFrame.grid(row=2,column=0, sticky="NSEW", padx=5, pady=5)
        optionsFrame.columnconfigure(1, weight=1)

        portLabel = ttk.Label(optionsFrame, text="COM Port:")
        portLabel.grid(row=0, column=0, padx=3)

        self.portSelect = ttk.Combobox(optionsFrame, textvariable=self.comport_var,
                values=list(self.all_comports.keys()), width=40)
        self.portSelect.state(['readonly'])
        self.portSelect.current(0)
        self.portSelect.grid(row=0, column=1, sticky="W")

        baudLabel = ttk.Label(optionsFrame, text="Baud Rate:")
        baudLabel.grid(row=1, column=0, padx=3, pady=(0,5))

        self.baudSelect = ttk.Combobox(optionsFrame, textvariable=self.baud_var,
                values=[1200,1800,2400,4800,9600,19200], width=15)
        self.baudSelect.state(['readonly'])
        self.baudSelect.current(0)
        self.baudSelect.grid(row=1, column=1, sticky="W", pady=(0,5))

        buttonFrame = ttk.Frame(self)
        buttonFrame.grid(row=4,column=0, sticky="NSEW", padx=5, pady=5)
        buttonFrame.columnconfigure(3, weight=1)

        self.startButton = ttk.Button(buttonFrame, text="Start",
                command=self.start_controlling)
        self.startButton.grid(row=0, column=0)

        self.stopButton = ttk.Button(buttonFrame, text="Stop",
                command=self.stop_comport)
        self.stopButton.state(['disabled'])
        self.stopButton.grid(row=0, column=1)

        self.refreshButton = ttk.Button(buttonFrame, text="Refresh Devices",
                command=self.refresh_device_list)
        self.refreshButton.grid(row=0,column=2)

        quitButton = ttk.Button(buttonFrame, text="Quit",
                command=self.graceful_exit)
        quitButton.grid(row=0, column=3, sticky="E")

    def mainloop(self):
        self.icon.run_detached()
        super().mainloop()


class ThreadedPortReader(Serial):
    def __init__(self, updater, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.readEvent = Event()
        self.updater = updater
        self.readThread = Thread(target=None)

    def start(self):
        self.readEvent.set()
        self.readThread = Thread(target=self.continual_read,
                args=[self.updater.set_device_volumes, self.readEvent])
        self.readThread.start()

    def stop(self):
        self.readEvent.clear()
        if self.readThread.is_alive():
            self.readThread.join()

    def continual_read(self, updateFunc, event):
        while event.is_set():
            line = self.readline()
            if line:
                updateFunc(line)


class DeviceUpdater:
    def __init__(self, deviceList):
        self.all_devices = None
        self.selected_interfaces = []
        self.set_deviceList(deviceList)

    def set_deviceList(self, deviceList):
        newInterfaces = list(map(self.open_device, deviceList))
        self.all_devices = {d.FriendlyName: i for d,i in zip(deviceList, newInterfaces)}

    def select_devices(self, deviceNameList):
        self.selected_interfaces = []
        for d in deviceNameList:
            if d in self.all_devices:
                self.selected_interfaces.append(self.all_devices[d])

    def open_device(self, device):
        return device._dev.Activate(
            IAudioEndpointVolume._iid_,
            CLSCTX_ALL,
            None
        ).QueryInterface(IAudioEndpointVolume)

    def set_device_volumes(self, rawBytes):
        def extract_pot_values(rawBytes):
            return list(map(int, rawBytes.split(b"|")))

        potValues = [i/1024 for i in extract_pot_values(rawBytes)]
        for i, interface in enumerate(self.selected_interfaces):
            try:
                if i < len(potValues):
                    interface.SetMasterVolumeLevelScalar(round(potValues[i],3),None)

                if potValues[i] < 6/1024:
                    interface.SetMute(True, None)
                else:
                    interface.SetMute(False, None)
            except COMError as ce:
                print(f"Device #{i}: {ce.text}")

if __name__ == "__main__":
    trayApp = TrayApplication()
    trayApp.title("AudioControl")
    trayApp.mainloop()
