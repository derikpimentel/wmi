from typing import List
import wmi

wmi_classes = [
    {
        'name': "Win32_Processor",
        'properties': ["Name"]
    },
    {
        'name': "Win32_PhysicalMemory",
        'properties': ["Capacity", "Manufacturer", "PartNumber", "SerialNumber", "SMBIOSMemoryType", "Speed"]
    },
    {
        'name': "Win32_CDROMDrive",
        'properties': ["Name"]
    },
    {
        'name': "Win32_DiskDrive",
        'properties': ["MediaType", "Model", "SerialNumber", "Size"]
    },
    {
        'name': "Win32_VideoController",
        'properties': ["AdapterRAM", "Name", "VideoProcessor"]
    },
    {
        'filter': "PhysicalAdapter = '1' AND PNPDeviceID LIKE 'PCI\\%'",
        'name': "Win32_NetworkAdapter",
        'properties': ["MACAddress", "Manufacturer", "Name", "NetConnectionID"]
    },
    {
        'name': "Win32_BIOS",
        'properties': ["Manufacturer", "Name", "SerialNumber", "Version"]
    },
    {
        'name': "Win32_ComputerSystem",
        'properties': ["Domain", "Manufacturer", "Model", "Name", "SystemFamily", "TotalPhysicalMemory", "UserName", "Workgroup"]
    },
    {
        'name': "Win32_OperatingSystem",
        'properties': ["BuildNumber", "Caption", "Manufacturer", "OSLanguage", "SerialNumber"]
    },
    {
        'name': "SoftwareLicensingService",
        'properties': ["OA3xOriginalProductKey"]
    },
    {
        'name': "Win32_Product",
        'properties': ["InstallDate", "Name", "Vendor", "Version"]
    },
    {
        'filter': "LicenseStatus = '1' AND Name LIKE '%Office%'",
        'name': "SoftwareLicensingProduct",
        'properties': ["Name", "PartialProductKey"]
    },
    {
        'name': "AntiVirusProduct",
        'namespace': "root\SecurityCenter2",
        'properties': ["displayName"]
    }
]

# Função para interromper
def pause() -> None:
    input("[ENTER]")

# Função que estrutura uma string do comando SELECT para SQL
def select_query_str(table: str, columns: List[str] = ["*"], where_clause: str = None) -> str:
    if table:
        columns_str = ", ".join(columns)
        where_clause_str = ""
        if where_clause:
            where_clause_str = f" WHERE {where_clause}"
        return f"SELECT {columns_str} FROM {table}{where_clause_str}"

# Função que inicializa o WMI e realiza uma busca otimizada
def wmi_object(select_query: str, namespace: str = None) -> wmi:
    if select_query:
        return wmi.WMI(namespace=namespace).query(select_query)

# Função para visualização
def wmi_print(object_list: wmi) -> None:
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
    for wmi_class in wmi_classes:
        table = wmi_class.get('name')
        columns = wmi_class.get('properties')
        where_clause = wmi_class.get('filter', None)
        select_query = select_query_str(table, columns, where_clause)
        namespace = wmi_class.get('namespace', None)
        object_list = wmi_object(select_query, namespace)
        wmi_print(object_list)
        pause()