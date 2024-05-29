import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException
from com.sun.star.uno import Exception as UnoException
import sys
import time
import subprocess

def launch_libreoffice():
    try:
        process = subprocess.Popen([
            'soffice', '--accept=socket,host=localhost,port=2002;urp;',
            '--headless', '--norestore', '--invisible'
        ])
        return process
    except Exception as e:
        print(f"Error launching LibreOffice: {e}")
        sys.exit(1)

def connect_to_libreoffice(retries=5, delay=3):
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context)

    attempt = 0
    while attempt < retries:
        try:
            context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            return context
        except NoConnectException:
            print(f"Connection attempt {attempt + 1} failed. Retrying in {delay} seconds...")
            time.sleep(delay)
            attempt += 1
        except UnoException as e:
            print(f"UNO Exception: {e}")
            sys.exit(1)
    print("Failed to connect to LibreOffice.")
    sys.exit(1)

def open_document(desktop, file_path):
    try:
        properties = PropertyValue()
        properties.Name = "Hidden"
        properties.Value = False
        doc = desktop.loadComponentFromURL(file_path, "_blank", 0, (properties,))
        return doc
    except UnoException as e:
        print(f"Error opening document: {e}")
        sys.exit(1)

def get_sheet(doc, sheet_name):
    try:
        sheet = doc.Sheets.getByName(sheet_name)
        return sheet
    except UnoException as e:
        print(f"Error accessing sheet: {e}")
        sys.exit(1)

def set_cell_value(sheet, cell_address, value):
    try:
        cell = sheet.getCellRangeByName(cell_address)
        cell.setValue(value)
    except UnoException as e:
        print(f"Error setting cell value: {e}")
        sys.exit(1)

def get_cell_value(sheet, cell_address):
    try:
        cell = sheet.getCellRangeByName(cell_address)
        cell_type = cell.getType()
        if cell_type == uno.Enum("com.sun.star.table.CellContentType", "TEXT"):
            return cell.getString()
        else:
            return cell.getValue()
    except UnoException as e:
        print(f"Error getting cell value: {e}")
        sys.exit(1)

def close_document(doc):
    try:
        if not doc.isModified():
            doc.store()
        doc.close(True)
    except UnoException as e:
        print(f"Error closing document: {e}")

def terminate_libreoffice(process):
    try:
        process.terminate()
        process.wait(timeout=10)
    except Exception as e:
        print(f"Error terminating LibreOffice process: {e}")

def main():
    libreoffice_process = launch_libreoffice()
    context = connect_to_libreoffice()
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

    file_path = "file:///inputs/Excel_Automation/test_AXA.xlsx"
    doc = open_document(desktop, file_path)

    sheet_name = "AXA 1"
    sheet = get_sheet(doc, sheet_name)

    set_cell_value(sheet, "B10", 79844)
    cell_value = get_cell_value(sheet, "C10")
    print(cell_value)

    doc.calculateAll()

    close_document(doc)
    terminate_libreoffice(libreoffice_process)

if _name_ == "_main_":
    main()