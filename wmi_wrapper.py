from wmi import WMI, x_wmi_timed_out
import pythoncom
import threading
from .helpers import convert_to_human_readable


class WmiHandler():


    @staticmethod
    def get_serial_number():
        """Thread Safe WMI Query for Serial Number in Win32_Bios Class"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI()
        try:
            for sn in w.Win32_Bios(["SerialNumber"]):
                return sn.SerialNumber
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

    @staticmethod
    def get_operating_system_info():
        """Thread Safe WMI Query for Caption, build, arch, systemDrive in Win32_OperatingSystem Class"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI()
        try:
            for os in w.Win32_OperatingSystem(["Caption", "OSArchitecture", "MUILanguages", "Version", "SystemDrive"]):
                return {
                    "name": os.Caption,
                    "arch": os.OSArchitecture,
                    "languages": os.MUILanguages,
                    'build': os.Version,
                    'system_drive': os.SystemDrive
                }
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

    @staticmethod
    def get_model():
        """Thread Safe WMI Query for Serial Number in Win32_Bios Class"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI()
        try:
            for m in w.Win32_ComputerSystem(["Model"]):
                return m.Model
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

    @staticmethod
    def get_manufacturer():
        """Thread Safe WMI Query for Serial Number in Win32_Bios Class"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI()
        try:
            for m in w.Win32_ComputerSystem(["Manufacturer"]):
                return m.Manufacturer
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()


    @staticmethod
    def get_processors():
        """Thread Safe WMI Query for Processors info in Win32_Processor Class"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI()
        processors=[]
        try:
            for p in w.Win32_Processor(["Name"]):
                processors.append(p.Name)
            return processors
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

    @staticmethod
    def is_computer_on_AC_power():
        """Thread Safe WMI Query return True if AC adapter is connected"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI()
        try:
            for bat in w.Win32_Battery(["BatteryStatus"]):
                return bat.BatteryStatus in ['2', '3', '6', '7', '8', '9', '11']
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()


    @staticmethod
    def get_estimated_battery_level():
        """Thread Safe WMI Query for remaining battery level"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI()
        try:
            for bat in w.Win32_Battery(["EstimatedChargeRemaining"]):
                return bat.EstimatedChargeRemaining
        except Exception as e:
            print(e)
            return -1
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()


    @staticmethod
    def set_display_brightness(brightness):
        """Thread Safe WMI Query for setting display brightness"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI(namespace="WMI")
        try:
            for mon in w.WmiMonitorBrightnessMethods():
                mon.WmiSetBrightness(brightness, 0)
                return True
        except Exception as e:
            print(e)
            return False
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

    @staticmethod
    def get_total_memory_amount(get_human_readable_value=False):
        """Thread Safe WMI Query for setting display brightness"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI()
        try:
            total_amount = 0
            for mem in w.Win32_PhysicalMemory(['Capacity']):
                total_amount += int(mem.Capacity)
            if get_human_readable_value:
                return convert_to_human_readable(total_amount, use_exact_size=True)
            return total_amount
        except Exception as e:
            print(e)
            return 0
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()


    @staticmethod
    def get_pnp_devices():
        """Thread Safe WMI Query for list of system devices"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI()
        try:
            devices = []
            for dev in w.Win32_PnPEntity():
                try:
                    devices.append({'caption': dev.Caption,
                     'hw_id': dev.HardwareID[0],
                     'name': dev.Name,
                     'status': dev.Status
                     })
                except Exception as ex:
                    pass
            return devices
        except Exception as e:
            print(e)
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

    @staticmethod
    def get_disks_drives():
        """Thread Safe WMI Query for Disk drives in Win32_DiskDrive Class
        Returns: array of Dictionary: {index, model, serial_number, capacity, human_capacity
        """
        disks=[]
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()
        w = WMI()
        wql = 'SELECT * FROM Win32_DiskDrive WHERE InterfaceType <> "USB"'
        try:
            for dd in w.query(wql):
                d ={}
                d['index'] = dd.Index
                d['model'] = dd.Caption
                d['serial_number'] = dd.SerialNumber
                d['capacity'] = dd.Size
                d['human_capacity'] = convert_to_human_readable(dd.Size)
                disks.append(d)
            return disks
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

    @staticmethod
    def get_system_sku(manufacturer):
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()

        try:
            if 'DELL' in manufacturer.upper() or 'ALIENWARE' in manufacturer.upper():
                w = WMI(namespace='WMI')
                for sku in w.MS_SystemInformation(["SystemSKU"]):
                    return sku.SystemSKU
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

    @staticmethod
    def is_ssd_drive(disk_id):
        """Thread Safe WMI Query for checking if disk drive is SSD  in MSFT_PhysicalDisk class"""
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()

        try:
            w = WMI(namespace='Microsoft\Windows\Storage')
            for m in w.MSFT_PhysicalDisk(["MediaType"], DeviceId=disk_id ):
                return int(m.MediaType)==4
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

class UsbInsertWatcher:
    def __init__(self):
        self.stop_wanted=False
    def stop(self):
        self.stop_wanted = True
    def watch_for_events(self):
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()

        try:
            w = WMI()
            # WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'")
            watcher = w.Win32_USBHub.watch_for('creation')
            while not self.stop_wanted:
                try:
                    event = watcher(timeout_ms=1000)
                except x_wmi_timed_out:
                    pass
                else:
                    print(event)
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

class VolumeCreationWatcher:
    def __init__(self, callback=None):
        self.stop_wanted=False
        self.callback=callback
    def stop(self):
        self.stop_wanted = True
    def watch_for_events(self):
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()

        try:
            w = WMI()
            # WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'")
            watcher = w.Win32_VolumeChangeEvent.watch_for(EventType=2)
            while not self.stop_wanted:
                try:
                    event = watcher(timeout_ms=1000)
                except x_wmi_timed_out:
                    pass
                else:
                    print(event.DriveName)
                    if self.callback is not None:
                        self.callback(event.DriveName)
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()


class VolumeRemovalWatcher:
    def __init__(self, callback=None):
        self.stop_wanted=False
        self.callback=callback
    def stop(self):
        self.stop_wanted = True
    def watch_for_events(self):
        if not threading.current_thread() is threading.main_thread():
            pythoncom.CoInitialize()

        try:
            w = WMI()
            watcher = w.Win32_VolumeChangeEvent.watch_for(EventType=3)
            while not self.stop_wanted:
                try:
                    event = watcher(timeout_ms=1000)
                except x_wmi_timed_out:
                    pass
                else:
                    print(event.DriveName)
                    if self.callback is not None:
                        self.callback(event.DriveName)
        except Exception as e:
            print(e)
            return None
        finally:
            if not threading.current_thread() is threading.main_thread():
                pythoncom.CoUninitialize()

if __name__=="__main__":
    print(WmiHandler.get_system_sku('DELL'))
    #print(WmiHandler.get_disks_drives())


    # t =threading.Thread(target=lambda: print(WmiHandler.is_computer_on_AC_power()))
    # t.start()




