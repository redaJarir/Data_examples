import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException
from com.sun.star.uno import Exception as UnoException
import sys
import time
import subprocess
import pandas as pd

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
        properties.Value = True
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

def load_extract(df):
    libreoffice_process = launch_libreoffice()
    context = connect_to_libreoffice()
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

    file_path = "file:///0_inputs/Data_examples/test_AXA.xlsx"
    doc = open_document(desktop, file_path)

    sheet_name = "data"
    sheet = get_sheet(doc, sheet_name)

    # Convert DataFrame to a list of lists
    data = [df.columns.tolist()] + df.values.tolist()

    # Determine the range to be filled
    start_row = 0
    start_col = 0
    end_row = start_row + len(data) - 1
    end_col = start_col + len(data[0]) - 1

    # Create a range to insert data
    cell_range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)

    # Set the data in one shot
    cell_range.setDataArray(tuple(tuple(row) for row in data))
   

    doc.calculateAll()

    sheet_name = "Statistics"
    sheet = get_sheet(doc, sheet_name)

    # Determine the used area in the sheet
    used_range = sheet.getCellRangeByPosition(0, 0, 4,25)

    # Get data as a nested tuple
    data_tuple = used_range.getDataArray()

    # Convert to DataFrame
    data = [list(row) for row in data_tuple]
    df_out = pd.DataFrame(data[1:], columns=data[0])  # Assuming the first row is the header

    # Save the DataFrame to a CSV file
    #df_out.to_csv("test.csv", index=False)
    #doc.store()

    close_document(doc)
    terminate_libreoffice(libreoffice_process)
    return df_out

#if __name__=="__main__":
    #df = pd.read_csv('/0_inputs/Data_examples/ouput1.csv')
    #load_extract(df)