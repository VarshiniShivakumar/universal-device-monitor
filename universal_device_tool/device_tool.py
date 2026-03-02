import wmi
import psutil
import time
import socket
import serial.tools.list_ports
from datetime import datetime


# ==========================================
# PRINT HEADER
# ==========================================

def print_header():

    print("\n======================================================")
    print("        COMPLETE DEVICE + NETWORK MONITOR")
    print("======================================================")
    print("Monitoring hardware and network changes...")
    print("======================================================\n")


# ==========================================
# NETWORK INTERFACES
# ==========================================

def get_network_interfaces():

    interfaces = {}

    stats = psutil.net_if_stats()
    addrs = psutil.net_if_addrs()

    for name, stat in stats.items():

        if stat.isup:

            ip_address = None
            mac_address = None

            if name in addrs:

                for addr in addrs[name]:

                    if addr.family == socket.AF_INET:
                        ip_address = addr.address

                    if addr.family == psutil.AF_LINK:
                        mac_address = addr.address


            interfaces[name] = {

                "ip": ip_address,
                "mac": mac_address
            }

    return interfaces


# ==========================================
# USB DEVICES (WMI)
# ==========================================

def get_usb_devices():

    c = wmi.WMI()

    devices=set()

    for usb in c.Win32_USBHub():

        devices.add(usb.DeviceID)

    return devices


# ==========================================
# SERIAL / COM PORTS
# ==========================================

def get_com_ports():

    ports = serial.tools.list_ports.comports()

    devices=set()

    for port in ports:

        devices.add(port.device)

    return devices


# ==========================================
# STORAGE DEVICES
# ==========================================

def get_storage_devices():

    partitions = psutil.disk_partitions()

    drives=set()

    for p in partitions:

        drives.add(p.device)

    return drives


# ==========================================
# OPEN PORTS
# ==========================================

def get_open_ports():

    connections = psutil.net_connections(kind='inet')

    open_ports=set()

    for conn in connections:

        if conn.status == "LISTEN":

            open_ports.add(conn.laddr.port)

    return sorted(open_ports)


# ==========================================
# PRINT CURRENT STATUS
# ==========================================

def print_current_status():

    print("CURRENT SYSTEM STATUS\n")


    # Network
    interfaces=get_network_interfaces()

    print("ACTIVE NETWORK INTERFACES\n")

    for name,details in interfaces.items():

        print("Interface   :",name)
        print("IP Address  :",details["ip"])
        print("MAC Address :",details["mac"])
        print("-"*40)


    # USB
    print("\nUSB DEVICES DETECTED:")

    for dev in get_usb_devices():

        print(dev)


    # COM Ports
    print("\nCOM PORTS:")

    for port in get_com_ports():

        print(port)


    # Storage
    print("\nSTORAGE DEVICES:")

    for drive in get_storage_devices():

        print(drive)


    print("\nOPEN LISTENING PORTS:")

    print(get_open_ports())

    print("="*60)



# ==========================================
# MONITOR CHANGES
# ==========================================

def monitor_all():

    prev_net=get_network_interfaces()
    prev_usb=get_usb_devices()
    prev_com=get_com_ports()
    prev_storage=get_storage_devices()


    while True:

        time.sleep(2)


        # NETWORK

        curr_net=get_network_interfaces()

        for n in curr_net:

            if n not in prev_net:

                print("\nNETWORK CONNECTED")
                print("Time:",datetime.now())
                print("Interface:",n)

        for n in prev_net:

            if n not in curr_net:

                print("\nNETWORK DISCONNECTED")
                print("Time:",datetime.now())
                print("Interface:",n)


        prev_net=curr_net


        # USB

        curr_usb=get_usb_devices()

        for d in curr_usb:

            if d not in prev_usb:

                print("\nUSB CONNECTED")
                print("Time:",datetime.now())
                print("Device:",d)


        for d in prev_usb:

            if d not in curr_usb:

                print("\nUSB DISCONNECTED")
                print("Time:",datetime.now())
                print("Device:",d)


        prev_usb=curr_usb


        # COM PORTS

        curr_com=get_com_ports()

        for c in curr_com:

            if c not in prev_com:

                print("\nCOM PORT CONNECTED")
                print("Time:",datetime.now())
                print("Port:",c)


        for c in prev_com:

            if c not in curr_com:

                print("\nCOM PORT DISCONNECTED")
                print("Time:",datetime.now())
                print("Port:",c)


        prev_com=curr_com


        # STORAGE

        curr_storage=get_storage_devices()

        for s in curr_storage:

            if s not in prev_storage:

                print("\nSTORAGE CONNECTED")
                print("Time:",datetime.now())
                print("Drive:",s)


        for s in prev_storage:

            if s not in curr_storage:

                print("\nSTORAGE REMOVED")
                print("Time:",datetime.now())
                print("Drive:",s)


        prev_storage=curr_storage



# ==========================================
# MAIN
# ==========================================

if __name__=="__main__":

    print_header()

    print_current_status()

    monitor_all()