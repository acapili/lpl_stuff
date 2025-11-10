"""
Purpose:  Efficiently audit the ASPEN model by reading or writing information to a input/output template
"""
__author__    = "Ayrthon Capili"
__version__   = "0.0.0"



import json

with open('Capstone/storage.JSON', 'r') as f:
    data = json.load(f)


import sqlite3

#logs to both file and console, with rotation
import logging
from logging.handlers import RotatingFileHandler
import os
from datetime import datetime
from ast import literal_eval

# OlxAPI and OlxObj imports
import sys
import os

# Add the parent directory to the system path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import OlxAPI
import OlxAPILib
import OlxObj
from OlxAPIConst import *

from ctypes import *
import cmath


# Log file path (in Capstone folder)
log_file = os.path.join(os.path.dirname(__file__), "capstone.log")

# Create handlers
file_handler = RotatingFileHandler(log_file, maxBytes=1_000_000, backupCount=3)
file_handler.setLevel(logging.INFO)


# Formatter
formatter = logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s")
file_handler.setFormatter(formatter)


# Root logger setup
logging.basicConfig(level=logging.INFO, handlers=[file_handler])


#import from LPL files
from AspenDataExtraction import BusEquipmentExtractor, olr_file_utils, AquireNetworkInfo
from AspenDataImporting import UpdateModel
from ExcelUtility import ExcelUtility
from CSVUtility import CSVUtility
from Fault_analysis import faults
from AddToModel import import_to_aspen
from DatabaseUtility import DatabaseUtility, AddtoDatabase
# from something import DatabaseUtility


# db = DatabaseUtility("network.db")

# db.create_table("buses", "name TEXT, location TEXT, kv REAL, angle REAL, x REAL, y REAL")
# db.create_table("lines", "name TEXT, id TEXT, length REAL, type TEXT, r REAL, x REAL, r0 REAL, x0 REAL, bus1 TEXT, bus2 TEXT")
# db.create_table("xfmrs2", "name TEXT, id TEXT, baseMVA REAL, r REAL, x REAL, r0 REAL, x0 REAL, bus1 TEXT, bus2 TEXT")

def write_aspen_data_to_excel() -> None:
    '''
    Writes information about equipment connected to buses at specified locations into an Excel
    file using openpyxl.
    Information comes directly from Aspen OLR file using OlxAPI.
    '''

    # Load the Excel file once (instead of reopening for each location)
    excel_util = ExcelUtility(data['fileLocations']['aspen_excel_path'])

    #erase previous data
    excel_util.clear_rows_in_range(2, 1000, sheet_names=data["SheetsForReading"])

    # Read locations from the "Substation List" sheet, column A, starting from row 3
    locations_to_search = excel_util.read_column(data["SheetsForWriting"][0],start_row= 3)
    cleaned_list = [x for x in locations_to_search if x is not None]

    sheets = list(data["SheetsForReading"].keys())

    # Track current row positions for each sheet/section
    row_positions = {
        "buses": {"sheet": sheets[0], "row": 2,"col": 1},
        "generator": {"sheet": sheets[1], "row": 2, "col": 1},
        "gen_unit": {"sheet": sheets[1], "row": 2, "col": 12},
        "lines": {"sheet": sheets[2], "row": 2, "col": 1},
        "xfmr2": {"sheet": sheets[3], "row": 2, "col": 1},
        "xfmr3": {"sheet": sheets[4], "row": 2, "col": 1},
        "switches": { "sheet": sheets[9], "row": 2, "col": 2},
        "fuse": {"sheet": sheets[9], "row": 2, "col": 10},
        "recloser": {"sheet": sheets[9], "row": 2, "col": 17},
        "oc_phase": {"sheet": sheets[5], "row": 2, "col": 1},
        "oc_ground": {"sheet": sheets[6], "row": 2, "col": 1},
        "dist_phase": {"sheet": sheets[7], "row": 2, "col": 1},
        "dist_ground": {"sheet": sheets[8], "row": 2, "col": 1}
        
    }

    #loop through each location
    for loc in cleaned_list:
        if loc is None:
            continue
        #find all buses at location
        buses = AquireNetworkInfo().get_buses_at_location(loc)


        #extracting infromation from each bus
        for i in buses:
            # create bus entry
            bus = BusEquipmentExtractor(i)
            #start with relays beacuase it is 3D array and needs to be extracted outside of the data map
            relays = bus.relay_information()
            data_map = {
                #equipment
                "buses": bus.bus_information(),
                "generator": bus.gen_information(),
                "gen_unit": bus.genunit_information(),
                "lines": bus.line_information(),
                "xfmr2": bus.xfmr2_information(),
                "xfmr3": bus.xfmr3_information(),
                "switches": bus.switch_info(),
                #devices
                "oc_phase": relays[0],
                "oc_ground": relays[1],
                "dist_phase": relays[2],
                "dist_ground": relays[3],
                "fuse": relays[4],
                "recloser": relays[5]
                }

            #write to excel using the data map and update row positions
            for key, data_array in data_map.items():

                if not data_array:
                    continue  # skip if no data to write

                pos = row_positions[key]
                excel_util.write_2d_array(
                    data_array,
                    sheet_name=pos["sheet"],
                    start_row=pos["row"],
                    start_col=pos.get("col", 1)
                )

    # Open workbook once after writing everything
    excel_util.open_workbook()


import sqlite3, csv
from pathlib import Path

def export_sqlite_to_csvs(db_path, out_dir):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(str(db_path))
    cur = conn.cursor()

    cur.execute("""
        SELECT name
        FROM sqlite_master
        WHERE type='table' AND name NOT LIKE 'sqlite_%'
        ORDER BY name
    """)
    tables = [r[0] for r in cur.fetchall()]
    if not tables:
        raise RuntimeError("No tables found. Check if file path is correct")

    for t in tables:
        cur.execute(f"SELECT * FROM '{t}'")
        rows = cur.fetchall()
        headers = [d[0] for d in cur.description]

        out_file = out_dir / f"{t}.csv" #output file path
        with out_file.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(headers)
            for r in rows:
                w.writerow([v.hex() if isinstance(v, (bytes, bytearray)) else v for v in r])

    conn.close()
    return [str(out_dir / f"{t}.csv") for t in tables]

# Example:
# files = export_sqlite_to_csvs("C:/path/to/your_database.db", "C:/path/to/csv_out")
# print(files)



def bus_faults() -> None:

    """
    Used to create a spreadsheet with fault current information

    Can be used for Requiremnt 2 of PRC-027 or for general model auditing
    """

    def results() -> tuple[float, float]:
        """
        Returns the magnitude and angle of the fault current in polar form.
        Originates from Aspen OlxObj FltSimResult.
        """
        r, theta = 0.0, 0.0
        for r1 in OlxObj.FltSimResult:
            # set scope for compute solution results
            r1.setScope(tiers=1)
            #returns complex rectangular form
            phaseABC_currents = r1.current()
            #converts to polar form
            r, theta = cmath.polar(phaseABC_currents[0])
            #converts radians to degrees
            theta = theta * (180/cmath.pi)
        return r, theta

    # Load the Excel file once (instead of reopening for each location)
    excel_util = ExcelUtility(data['fileLocations']['PRC_027_excel_path'], False)
    excel_util_backup = ExcelUtility(data['fileLocations']['PRC_027_Backup_excel_path'], False)

    # Read locations from the "Substation List" sheet, column A, starting from row 3
    locations_to_search = excel_util.read_column(data["PRC027Sheets"][0],start_row= 3)
    cleaned_list = [x for x in locations_to_search if x is not None] #removes empty cells


    #loop through each location
    for loc in cleaned_list:
        if loc is None:
            continue

        #find all buses at location
        buses = AquireNetworkInfo().get_buses_at_location(loc)

        for b1 in buses:
            current_sheet = data["PRC027Sheets"][2]     
            # Determine the latest entry for a bus
            refrence = excel_util.read_column(current_sheet,col_num= 7,start_row=8, end_row=300)
            cleaned_refrence = [x for x in refrence if x is not None]
            # Define the fault specification
            fs1 = OlxObj.SPEC_FLT.Classical(obj=b1, fltApp='BUS', fltConn='3LG')
            # Run fault simulation
            OlxObj.OLCase.simulateFault(fs1, 1)
            #gains results
            threelgr, threelgtheta = results()
            # Define the fault specification
            fs2 = OlxObj.SPEC_FLT.Classical(obj=b1, fltApp='BUS', fltConn='1LG:A')
            # Run fault simulation
            OlxObj.OLCase.simulateFault(fs2, 1)
            #gains results
            onelgr, onelgtheta = results()
            
            #checks if transmission or distribution based on kv
            if b1.KV > 60:
                location = loc + " -  Transmission"
            else:
                location = loc + " -  Distribution"
            

            #format going into excel
            analysis = [datetime.now().strftime("%m/%d/%Y"), location, b1.NAME, b1.KV, f"{round(threelgr)}", f"{round(onelgr)}"]
            #backup with angles (without the brackets they point to the same memory)
            backup_analysis = analysis[:]
            backup_analysis[4] = f"{round(threelgr)} @  {round(threelgtheta)}"
            backup_analysis[5] = f"{round(onelgr)}@ {round(onelgtheta)}"


            def find_existing_bus_entry(bus_name: str, reference_list: list[str] = cleaned_refrence) -> int:
                """Return the index of the last occurrence of a bus name in reference_list.
                Returns -1 if the bus is not found.
                """
                for idx in range(len(reference_list) - 1, -1, -1):  # iterate backwards
                    if reference_list[idx] == bus_name:
                        return idx
                return -1
            #call created function
            existing_index = find_existing_bus_entry(b1.NAME)

            #excel location details
            row_location_after_sample = 8
            column_start = 4

            #if the bus already exists, update the entry else create a new entry
            if existing_index != -1:
                previous_faults = excel_util.read_row(current_sheet, row_num= existing_index + row_location_after_sample, start_col=9, end_col=10)
                previous_faults_backup = excel_util_backup.read_row(current_sheet, row_num= existing_index + row_location_after_sample, start_col=9, end_col=10)
                for x in previous_faults:
                    analysis.insert(-2, x)
                for x in previous_faults_backup:
                    backup_analysis.insert(-2, x)

                analysis.insert(1, "Short Circuit Analysis")
                backup_analysis.insert(1, "Short Circuit Analysis")
                #insert into a new spot
                excel_util.insert_row(sheet_name=current_sheet, starting_row= existing_index + row_location_after_sample + 1, starting_column=column_start, value=analysis)
                excel_util_backup.insert_row(sheet_name=current_sheet, starting_row= existing_index + row_location_after_sample + 1, starting_column=column_start, value=backup_analysis)
            else:
                #append to the end of the list, making a new entry
                analysis.insert(1, "Establishing a baseline")
                excel_util.insert_row(sheet_name=current_sheet, starting_row= row_location_after_sample + len(cleaned_refrence), starting_column=column_start, value=analysis)
                excel_util_backup.insert_row(sheet_name=current_sheet, starting_row= row_location_after_sample + len(cleaned_refrence), starting_column=column_start, value=backup_analysis)


    #open workbook once after writing everything
    excel_util.open_workbook()
                
def relay_operation_study() -> None:
    """
    To conduct studies of relay operations for buses that have been shown to have faults that devate more than 15%

    this will involve conducting multiple faults
    """

    # Load the Excel file once (instead of reopening for each location)
    excel_util = ExcelUtility(data['fileLocations']['PRC_027_excel_path'])
    # Read locations from the "Substation List" sheet, column A, starting from row 3
    locations_to_search = excel_util.read_column(data["PRC027Sheets"][0], col_num=2,start_row= 3)
    cleaned_list = [x for x in locations_to_search if x is not None]#removes empty cells

    #the buses that will have thier faults run through them
    buses_over_acceptable_deviation = []

    network = AquireNetworkInfo()

    #If locations are not provided it will check for all buses with more than 15% deviation
    if cleaned_list is None or len(cleaned_list) == 0:
        #obtain lists of buses that have more than 15% deviation
        threeLG_deviation = excel_util.read_column(data["PRC027Sheets"][2], col_num=13, start_row= 8)
        oneLG_deviation = excel_util.read_column(data["PRC027Sheets"][2], col_num=14, start_row= 8)

        #checks each deviation value
        for idx in range(len(threeLG_deviation)):
            #checks if the cell is empty/NA
            if isinstance(threeLG_deviation[idx], str) or isinstance(oneLG_deviation[idx], str):
                continue
            #if deviation is equal to or greater than 15%
            elif int(threeLG_deviation[idx]) >=15 or int(oneLG_deviation[idx]) >=15:
                x = excel_util.read_row(data["PRC027Sheets"][2], row_num= idx + 8, start_col=7, end_col=8)
                buses_over_acceptable_deviation.append(x)

        if not buses_over_acceptable_deviation:
            print("No buses found with more than 15% deviation in fault current for  locations provided. Exiting study.")
            return
        else:
            #using the bus names to aquire the bus objects and fill the list with bus objects
            for idx, x in enumerate(buses_over_acceptable_deviation):
                buses_over_acceptable_deviation[idx] = network.get_bus_with_name(x)
    
            
    #If locations are provided it will use those instead
    else:            
        #loop through each location
        for loc in cleaned_list:
            #find all buses at location
            for x in network.get_buses_at_location(loc):
                buses_over_acceptable_deviation.append(x)
    
    #path for CTI data
    folder_path = data["fileLocations"]["CTI_data_path"]


    

    #clears out the CTI information folder
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}: {e}")


#TODO wait for them to confirm whether or not equipment outside of tranmission lines are usable.
#if they are then use this


    #loop through each bus that needs relay operation study
    for b1 in buses_over_acceptable_deviation:
        #obtain important information about each relay group connected to the bus
        for relayGRP in b1.RLYGROUP:
            b2 = relayGRP.BUS[1]
            Cir_ID = relayGRP.EQUIPMENT.CID
            equip_type = 1 
            #run the relay operation study command        
            checkParams = '<CHECKRELAYOPERATIONSEA' \
                            f' REPORTPATHNAME="{folder_path}"' \
                            f' SELECTEDOBJ="{b1.NO}; \'{b1.NAME}\'; 13.8; {b2.NO};\'{b2.NAME}\'; 345; \'{Cir_ID}\'; {equip_type};" ' \
                            ' FAULTTYPE="3LG 1LF"' \
                            ' DEVICETYPE="OCG, OCP, DSG, DSP, LOGIC, VOLTAGE, DIFF" ' \
                            ' OUTAGELINE="1" OUTAGEPILOT="1"/>'
            print(checkParams)
            if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(checkParams):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            print("Success")


    #TODO wait for them to confirm whether or not equipment outside of tranmission lines are usable.
    #if they are not use this

    # #run the relay operation study command        
    # checkParams = '<CHECKRELAYOPERATIONSEA' \
    #                 f' REPORTPATHNAME="{folder_path}"' \
    #                 ' FAULTTYPE="3LG 1LF"' \
    #                 ' DEVICETYPE="OCG, OCP, DSG, DSP, LOGIC, VOLTAGE, DIFF" ' \
    #                 ' OUTAGELINE="1" OUTAGEPILOT="1"/>'
    # print(checkParams)
    # if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(checkParams):
    #     raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    # print("Success")
    
    CSV_files = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.endswith(('.csv', '.CSV'))
    ]

    data_array_relay_coor = {}

 
    # for b1 in buses_over_acceptable_deviation:
    for csv_file in CSV_files:
        ACTIVATE_dict = False
        if "Summary" in csv_file:
            continue
        else:
            CSV_array = CSVUtility(csv_file).read_csv()
            for i in CSV_array:
                #checking whether or not there 
                if len(i) == 1 and "Stepped Event Simulation" in str(i[0]):
                    #write to excel
                    analysis = data_array_relay_coor.values()
                    
                    if ACTIVATE_dict:
                        print(data_array_relay_coor.keys())
                        print(analysis)
                        excel_util.write_2d_array([analysis], sheet_name=data["PRC027Sheets"][3], start_row=8 + len(excel_util.read_column(sheet_name=data["PRC027Sheets"][3])), start_col=7)
                        data_array_relay_coor.clear()
                    else:
                        ACTIVATE_dict = True
                    
                    #resume clearing data array for next use
                    data_array_relay_coor["scenario"] = i[0][50:67]
                #contingency    
                elif len(i) == 1 and "outage" in str(i[0]):
                    data_array_relay_coor["contingency"] = i[0][9:]
                #inputs the trip times
                elif len(i) == 8 and i[1].replace('.', '', 1).isdigit():
                    data_array_relay_coor["trip_time"] = float(i[1])
                    data_array_relay_coor["prim_rel"] = i[2]
                    data_array_relay_coor["back_rel"] = i[5]
                    data_array_relay_coor["CTI"] = i[7]
                #if there is no back up relay or an error accured
                elif len(i) == 5 and i[1].replace('.', '', 1).isdigit():
                     data_array_relay_coor["trip_time"] = float(i[1])
                     data_array_relay_coor["loc"] = i[2]
                     data_array_relay_coor["rel"] = i[3]
                     data_array_relay_coor["error"] = i[4]
                else:
                    continue
    # open workbook once after writing everything
    excel_util.open_workbook()


def updateDatabase():
    """
    Extract network elements from the active Aspen OneLiner case (OlxObj.OLCase)
    and write them into an SQLite database.

    Workflow Summary:
        1. Initialize the AddtoDatabase object (handles SQLite operations).
        2. Filter the OLCase.BUS objects to only include those located in defined substations.
        3. Load and insert bus data into the database.
        4. Loop through all other equipment types (lines, transformers, relays, etc.),
           extract their objects, and perform UPSERT operations into their corresponding tables.
    """
    # 1) Initialize the AddtoDatabase utility.
    up = AddtoDatabase("test_of_code.db", data["substation list"], data["equipment list"])
    
    # 2) Identify all bus objects within substations listed in storage.JSON.
    list_of_buses = []
    for bus in OlxObj.OLCase.BUS:
        if getattr(bus, 'LOCATION', None) in data["substation list"]:
            list_of_buses.append(bus)

    # 3) Insert the bus data into SQL first, as it will be faster to look through the buses
    #then to look at all equipment
    up.bus(list_of_buses)

    # 4) Begin equipment extraction and database population.
    device_flag = False
    for equi in data["equipment list"][1:]:
        if equi == 'rlygroup':
            device_flag = True
            continue
        # 4.1) Retrieve all Aspen objects corresponding to the current equipment type.
        list_of_equi = up.equipment_find(equi, device_flag)

        # 4.2) Validate that equipment was found before attempting a database insert.
        if len(list_of_equi)==0:
            raise RuntimeError(f"No {equi}s found.")
        
        # 4.3) UPSERT (insert or update) the equipment records into SQLite
        up.upsert_from_objects(
            table_name = equi,
            objects=list_of_equi,
            drop_and_recreate=True, #TODO remove the need for this specification
        )




def main():
    user_input = 'update_database'
    # user_input =  'read_excel_import_aspen'
    # user_input =  'export_database_write_excel'
    # user_input =  'test_faults'
    # user_input =  'run_study'

    #user input
    while True:
        match user_input:
            case 'read_excel_import_aspen':
                open_olr_file_with_retry(True)
                import_to_aspen().read_from_excel()
                break
            case 'export_database_write_excel':  
                # open_olr_file_with_retry(False)
                # write_aspen_data_to_excel()
                export_sqlite_to_csvs("U:/Python/test_of_code.db", "U:/Python/SQL_Exports")
                break
            case 'test_faults':
                open_olr_file_with_retry(False)
                bus_faults()
                break
            case 'run_study':
                open_olr_file_with_retry(False)
                relay_operation_study()
                break
            case 'update_database':
                open_olr_file_with_retry(False)
                updateDatabase()
                break
    olr_file_utils.close_olr_file()

def open_olr_file_with_retry(Read_Excel: bool) -> None:
    """Opens a specific OLR file depending on bool
        with retry mechanism if the file is already open."""
    
    if Read_Excel:
        x = 0
        olr_file_path = data['fileLocations']['import_olr_path']
    else:
        x = 1
        olr_file_path = data['fileLocations']['extract_olr_path']

    while True:
        try:
            olr_file_utils.open_olr_file(olr_file_path, x)
            break
        except Exception as e:
            print(f"Error opening OLR file: {e}")
            print("Please close the OLR file if it is open and press Enter to try again.")
            input("")

if __name__ == "__main__":
    main()
    # olr_file_utils.open_olr_file("C:/Users/184876/Documents/Fall25/Ayrthon Model for updating the network.OLR", 1)
    # olr_file_utils.close_olr_file()
