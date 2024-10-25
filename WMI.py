import wmi

# Função para visualização
def wmi_print(object_list):
    for obj in object_list:
        properties = list(obj._properties) # Lista os atributos

        # Verifica qual atributo tem a maior quantidade de caracteres
        max_length = 0
        for property in properties:
            length = len(property)
            if max_length < length:
                max_length = length

        # Imprime na tela
        print(f"\ninstância do {obj.Path_.Class}\n{{")
        for property in properties:
            print(f"{property.rjust(max_length + 8)} : {getattr(obj, property)}")
        print(f"}};\n")

if __name__ == "__main__":
    c = wmi.WMI()
    d = wmi.WMI(namespace="root\SecurityCenter2")

    cpu = c.query("SELECT Name FROM Win32_Processor")
    memory = c.query("SELECT Capacity, Manufacturer, PartNumber, SerialNumber, SMBIOSMemoryType, Speed FROM Win32_PhysicalMemory")
    cd_rom = c.query("SELECT Name FROM Win32_CDROMDrive")
    disk_drive = c.query("SELECT MediaType, Model, SerialNumber, Size FROM Win32_DiskDrive")
    gpu = c.query("SELECT AdapterRAM, Name, VideoProcessor FROM Win32_VideoController")
    net = c.query("SELECT MACAddress, Manufacturer, Name, NetConnectionID FROM Win32_NetworkAdapter WHERE PhysicalAdapter = '1' AND PNPDeviceID LIKE 'PCI\\%'")
    bios = c.query("SELECT Manufacturer, Name, SerialNumber, Version FROM Win32_BIOS")
    pc = c.query("SELECT Domain, Manufacturer, Model, Name, SystemFamily, TotalPhysicalMemory, UserName, Workgroup FROM Win32_ComputerSystem")
    os = c.query("SELECT BuildNumber, Caption, Manufacturer, OSLanguage, SerialNumber FROM Win32_OperatingSystem")
    key = c.query("SELECT OA3xOriginalProductKey FROM SoftwareLicensingService")
    product = c.query("SELECT InstallDate, Name, Vendor, Version FROM Win32_Product")
    partial_key = c.query("SELECT Name, PartialProductKey FROM SoftwareLicensingProduct WHERE LicenseStatus = '1' AND Name LIKE '%Office%'")
    security = d.query("SELECT displayName FROM AntiVirusProduct")

    wmi_print(cpu)
    wmi_print(memory)
    wmi_print(cd_rom)
    wmi_print(disk_drive)
    wmi_print(gpu)
    wmi_print(net)
    wmi_print(bios)
    wmi_print(pc)
    wmi_print(os)
    wmi_print(key)
    wmi_print(product)
    wmi_print(partial_key)
    wmi_print(security)
