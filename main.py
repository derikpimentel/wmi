from wmi import WMI

# Carregando os dados do equipamento.

Processor = WMI().query("""
    SELECT Name
    FROM Win32_Processor
""")

PhysicalMemory = WMI().query("""
    SELECT Capacity, Manufacturer, PartNumber, SerialNumber, SMBIOSMemoryType, Speed 
    FROM Win32_PhysicalMemory
""")

CDROMDrive = WMI().query("""
    SELECT Name
    FROM Win32_CDROMDrive
""")

DiskDrive = WMI().query("""
    SELECT MediaType, Model, SerialNumber, Size
    FROM Win32_DiskDrive
""")

VideoController = WMI().query("""
    SELECT AdapterRAM, Name, VideoProcessor
    FROM Win32_VideoController
""")

NetworkAdapter = WMI().query(""" 
    SELECT MACAddress, Manufacturer, Name, NetConnectionID
    FROM Win32_NetworkAdapter
    WHERE PhysicalAdapter = '1' AND PNPDeviceID LIKE 'PCI\\%'
""")

BIOS = WMI().query("""
    SELECT Manufacturer, Name, SerialNumber, Version
    FROM Win32_BIOS
""")

ComputerSystem = WMI().query("""
    SELECT Domain, Manufacturer, Model, Name, SystemFamily, TotalPhysicalMemory, UserName, Workgroup
    FROM Win32_ComputerSystem
""")

OperatingSystem = WMI().query("""
    SELECT BuildNumber, Caption, Manufacturer, OSLanguage, SerialNumber
    FROM Win32_OperatingSystem
""")

SoftwareLicensingService = WMI().query("""
    SELECT OA3xOriginalProductKey
    FROM SoftwareLicensingService
""")

AntiVirusProduct = WMI(namespace="root\SecurityCenter2").query("""
    SELECT displayName
    FROM AntiVirusProduct
""")

Product = WMI().query("""
    SELECT InstallDate, Name, Vendor, Version
    FROM Win32_Product
""")

SoftwareLicensingProduct = WMI().query("""
    SELECT Name, PartialProductKey
    FROM SoftwareLicensingProduct
    WHERE LicenseStatus = '1' AND Name LIKE '%Office%'
""")

# Depois de carregar os dados, exibimos os resultados.

print('\n','Processador(es):')
for cpu in Processor:
    print(' + ', str(cpu.Name))

print('\n','Memória(s):')
for memory in PhysicalMemory:
    print(' + ', str(memory.Capacity).ljust(10), ' | ', str(memory.Manufacturer).ljust(15), ' | ', str(memory.PartNumber).ljust(20), ' | ', str(memory.SerialNumber).ljust(15), ' | ', str(memory.SMBIOSMemoryType).ljust(3), ' | ', str(memory.Speed).ljust(0))

print('\n','Unidade(s) de Leitor Óptico:')
for cdrom in CDROMDrive:
    print(' + ', str(cdrom.Name))

print('\n','Unidade(s) de Disco Rígido:')
for disk in DiskDrive:
    print(' + ', str(disk.MediaType).ljust(25), ' | ', str(disk.Model).ljust(40), ' | ', str(disk.SerialNumber).ljust(40), ' | ', str(disk.Size).ljust(0))

print('\n','Controlador(es) de Vídeo:')
for gpu in VideoController:
    print(' + ', str(gpu.AdapterRAM).ljust(10), ' | ', str(gpu.Name).ljust(20), ' | ', str(gpu.VideoProcessor).ljust(0))

print('\n','Controlador(es) de Rede:')
for net in NetworkAdapter:
    print(' + ', str(net.MACAddress).ljust(20), ' | ', str(net.Manufacturer).ljust(30), ' | ', str(net.Name).ljust(40), ' | ', str(net.NetConnectionID).ljust(0))

print('\n','BIOS:')
for bios in BIOS:
    print(' + ', 'Fabricante:'.rjust(12), str(bios.Manufacturer))
    print(' + ', 'Nome:'.rjust(12), str(bios.Name))
    print(' + ', 'Nº de Série:'.rjust(12), str(bios.SerialNumber))
    print(' + ', 'Versão:'.rjust(12), str(bios.Version))

print('\n','Computador:')
for pc in ComputerSystem:
    print(' + ', 'Domínio:'.rjust(18), str(pc.Domain))
    print(' + ', 'Fabricante:'.rjust(18), str(pc.Manufacturer))
    print(' + ', 'Modelo:'.rjust(18), str(pc.Model))
    print(' + ', 'Nome:'.rjust(18), str(pc.Name))
    print(' + ', 'Família:'.rjust(18), str(pc.SystemFamily))
    print(' + ', 'Total de Memória:'.rjust(18), str(pc.TotalPhysicalMemory))
    print(' + ', 'Usuário:'.rjust(18), str(pc.UserName))
    print(' + ', 'Grupo de Trabalho:'.rjust(18), str(pc.Workgroup))

print('\n','Sistema Operacional:')
for os in OperatingSystem:
    print(' + ', 'Build:'.rjust(18), str(os.BuildNumber))
    print(' + ', 'Sistema:'.rjust(18), str(os.Caption))
    print(' + ', 'Fabricante:'.rjust(18), str(os.Manufacturer))
    print(' + ', 'Linguagem Nativa:'.rjust(18), str(os.OSLanguage))
    print(' + ', 'Chave:'.rjust(18), str(os.SerialNumber))

for key in SoftwareLicensingService:
    print('\n','Chave do Windows: ', str(key.OA3xOriginalProductKey))

print('\n','Sistema(s) de Segurança:')
for security in AntiVirusProduct:
    print(' + ', str(security.displayName))

print('\n','Software(s) Instalado(s):')
for soft in Product:
    print(' + ', str(soft.InstallDate).ljust(10), ' | ', str(soft.Name).ljust(65), ' | ', str(soft.Vendor).ljust(30), ' | ', str(soft.Version).ljust(0))

print('\n','Chave(s) do MSOffice:')
for keys in SoftwareLicensingProduct:
    print(' + ', str(keys.Name).ljust(50), ' | ', str(keys.PartialProductKey).ljust(0))