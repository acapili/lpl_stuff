import sys
import os

# Add the parent directory to the system path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))


import OlxAPI
import OlxObj
from OlxAPIConst import *

import ast


# Takes in and saves user input
class UserInput:
    """
    Class to handle user input and control the program flow.
    """
    def __init__(self):
        self.user_input = 0  # Default value for user input
        self.clarification = "" # Clarification message from the user
        self.programming_running = True # Flag to control the program loop
        self.file_location = ""# file location for the OLR file

    def str_input(self):
        """
        Takes in a string input from the user. Most basic input method.
        """
        try:
            self.clarification = str(input("Input: "))
        except ValueError:
            print("Invalid input. Please enter a String.")
            self.user_input = 0  # Default to 0 or any other fallback value 

    def int_input(self):
        """
        Takes in a string input from the user. Most basic input method.
        """
        try:
            self.user_input = int(input("Input: "))
        except ValueError:
            print("Invalid input. Please enter a number.")
            self.user_input = 0  # Default to 0 or any other fallback value 


    def command_list(self):
        """
        Displays a list of commands for the user to choose from.
        """
        print("\nEnter a number for which operation to perform:")
        print("1:Exit Program\n2:Print all areas\n3:Print all zones\n4:Print all relays")
        print("      41: Print all relays of a certain type")
        print("      42: Retrive information for relays inbetween two buses")
        print("        421: Print out numbers for buses within LPL\n")
        print("5:Change file to be used")
        print("=======================================================================")
        try:
            self.user_input = int(input("Input: "))
        except ValueError:
            print("\n=======================================")
            print("||Invalid input. Please enter a number||")
            print("=======================================")
            self.user_input = 0  # Default to 0 or any other fallback value

    def get_file_location(self):
        """
        Gets the file location of the OLR file from the user.
        """
        global olr_file_path  # Declare the global variable
        self.file_location = input("Enter the file location of the OLR file: ")
        olr_file_path = self.file_location  # Modify the global variable



class locations:
    """
    Class to represent retrive information regarding areas and zones in the OLR file.
    
    Utilizes the OlxAPI to retrieve area and zone information.
    """

    def __init__(self):
        self.areas = []
        self.zones = []

    def get_all_areas(self):
        """Retrieve all area numbers and their names within the OLR file."""
        for area_no in range(1, 1201):  # Iterate from 1 to 2000
            area_name = OlxAPI.GetAreaName(area_no)
            if not area_name.startswith("GetAreaName failure:"):
                self.areas.append({"number": area_no, "name": area_name})
        return self.areas
    
    def get_all_zones(self):
        """Retrieve all zones numbers and their names within the OLR file."""
        for zones_no in range(1, 1201):  # Iterate from 1 to 2000
            zones_name = OlxAPI.GetZoneName(zones_no)
            if not zones_name.startswith("GetZoneName failure:"):
                self.zones.append({"number": zones_no, "name": zones_name})
        return self.zones
    
    def print_all_areas(self):
        self.get_all_areas()
        print("Total Number of Areas:" + str(len(self.areas)) + "\n")
        print("=======================")
        for area in self.areas:
            print(f"Area Number: {area['number']}, Area Name: {area['name']}")
    
    def print_all_zones(self):
        self.get_all_zones()
        print("Total Number of Zones:" + str(len(self.zones)) + "\n")
        print("=======================")
        for area in self.zones:
            print(f"Zone Number: {area['number']}, Zone Name: {area['name']}")



class network_info:
    """
    Class to represent and retrieve information regarding buses in the OLR file.
    """
    
    def __init__(self):
        self.relay_grps = OlxObj.OLCase.RLYGROUP
        self.buses = OlxObj.OLCase.BUS
        self.lines = OlxObj.OLCase.LINE
        self.XFMR2= OlxObj.OLCase.XFMR
        self.XFMR3= OlxObj.OLCase.XFMR3

        self.LPL_buses: list[OlxObj.BUS] = []
        self.ONCOR_buses: list[OlxObj.BUS] = []

        self.lpl_lines: list[OlxObj.LINE] = []
        self.lpl_XFMR2: list[OlxObj.XFMR] = []
        self.lpl_XFMR3: list[OlxObj.XFMR3] = []


    def print_all_buses(self):
        for i in range(len(self.buses)):
            testing_buses = self.buses[i]
            print(f"Bus name: {testing_buses.NAME}")

    def get_LPL_buses(self):
        if len(self.LPL_buses) == 0:
            for i in range(len(self.buses)):
                testing_buses = self.buses[i]#buses are in alphabetical order
                if testing_buses.AREANO == 31:
                    self.LPL_buses.append(testing_buses.NAME)
        else:
            return

    def get_lpl_lines(self):
        """
        Utilizing the get LPL buses it takes the list and checks for all lines that have any of the bus lines and puts that into a class list
        """
        temp_lines = []

        if(len(self.LPL_buses) == 0):
            self.get_LPL_buses()

        for i in range(len(self.lines)):
            lines = self.lines[i]
            temp = [lines.BUS1.NAME, lines.BUS2.NAME] # if speed is important, removing checking the second bus name will increase the speed
            if any(item in temp for item in self.LPL_buses):
                temp_lines.append(lines)

        self.lpl_lines = temp_lines

        # # Debugging statment to print all of the items in the lpl_lines list
        # for i in range(len(self.lpl_lines)):
        #     print(f"{i}. Name of line {self.lpl_lines[i].NAME} BUS 1: {self.lpl_lines[i].BUS1.NAME}  BUS2:{self.lpl_lines[i].BUS2.NAME}\n")


    def get_lpl_XFMR2(self):
        """
        Utilizing the get LPL buses it takes the list and checks for all 2-Winding XFMRs that are connected to any of the bus lines in LPL and puts that into a class list
        """
        temp_lines = []

        if(len(self.LPL_buses) == 0):
            self.get_LPL_buses()

        for i in range(len(self.XFMR2)):
            lines = self.XFMR2[i]
            temp = [lines.BUS1.NAME, lines.BUS2.NAME] # if speed is important, removing checking the second bus name will increase the speed
            if any(item in temp for item in self.LPL_buses):
                temp_lines.append(lines)

        self.lpl_XFMR2 = temp_lines

        # # Debugging statment to print all of the items in the lpl_lines list
        # for i in range(len(self.lpl_XFMR2)):
        #     print(f"{i}. Name of line {self.lpl_XFMR2[i].NAME} BUS 1: {self.lpl_XFMR2[i].BUS1.NAME}  BUS2:{self.lpl_XFMR2[i].BUS2.NAME}\n")

    def get_lpl_XFMR3(self):
        """
        Utilizing the get LPL buses it takes the list and checks for all 3-Winding XFMRs that are connected to any of the bus lines in LPL and puts that into a class list
        """
        temp_lines = []

        if(len(self.LPL_buses) == 0):
            self.get_LPL_buses()

        for i in range(len(self.XFMR3)):
            lines = self.XFMR3[i]
            temp = [lines.BUS1.NAME, lines.BUS2.NAME, lines.BUS3.NAME] # if speed is important, removing checking the second bus name will increase the speed
            if any(item in temp for item in self.LPL_buses):
                temp_lines.append(lines)

        self.lpl_XFMR3 = temp_lines

        # # Debugging statment to print all of the items in the lpl_lines list
        # for i in range(len(self.lpl_XFMR3)):
        #     print(f"{i}. Name of line {self.lpl_XFMR3[i].NAME} BUS 1: {self.lpl_XFMR3[i].BUS1.NAME}  BUS2:{self.lpl_XFMR3[i].BUS2.NAME}\n")

    def get_all_relays(self):
        """
        Prints details of all relays within the OLR file.

        Args:
            relays (List[OlxObj.RLYGROUP]): A list of relay group objects.
        """
        for i in range(len(self.relays)):
            testing_relays = self.relays[i]
            print(f"Relay group index: {i+1}")
            print(f"Relay group location: {testing_relays.KEYSTR}") #Outputs: [bus 1,bus 2,Circuit ID, BRCODE]
            print(f"Relay group location: {testing_relays.KEYSTR[0]}") #Outputs: [bus 1,bus 2,Circuit ID, BRCODE]
            print(f"Relay group handle: {testing_relays.HANDLE}")
            print(f"Relays #: {testing_relays.RELAY}")
            print("=======================================================")
            for i in range(len(testing_relays.RELAY)):
                relay_in_grp = testing_relays.RELAY[i]
                if isinstance(relay_in_grp, OlxObj.RLYOCG) or isinstance(relay_in_grp, OlxObj.RLYOCP) or isinstance(relay_in_grp, OlxObj.FUSE) or isinstance(relay_in_grp, OlxObj.RLYDSP):
                    print(f"Relay type: {relay_in_grp.TYPE} Relay ID: {relay_in_grp.ID} Relay Handle: {relay_in_grp.HANDLE}")
                elif isinstance(relay_in_grp, OlxObj.RECLSR):
                    print(f"Recloser: {relay_in_grp.PARAMSTR}")
                else:
                    print("Unassigned relay type")
            print("=======================================================\n\n\n")


    def identify_connected_buses(self, first_bus :str) -> list:
        """
        Returns a list of relay objects that are connected to the first bus.
        """
        x = first_bus.upper()# Capitalizing each letter from input and adding a space to the end of the string

        connected_lines = [] #stores the lines as objects in a list
        connected_XFMR2 = [] #stores the 2-winding XFMR as objects in a list
        connected_XFMR3 = [] #stores the 3-windering XFMR as objects in a list

        if(len(self.lpl_lines) == 0):
            self.get_lpl_lines()
            print(len(self.lpl_lines))
        if(len(self.lpl_XFMR2) == 0):
            self.get_lpl_XFMR2()
            print(len(self.lpl_XFMR2))
        if(len(self.lpl_XFMR3) == 0):
            self.get_lpl_XFMR3()
            print(len(self.lpl_XFMR3))


        for i in range(len(self.lpl_lines)):
            if (x == self.lpl_lines[i].BUS1.NAME or x == self.lpl_lines[i].BUS2.NAME):
                connected_lines.append(self.lpl_lines[i])
            # debugging statement too show why certain relays are not being added to the list
            # else:
            #     print(f"{self.lpl_lines[i].BUS1.NAME} nor {self.lpl_lines[i].BUS2.NAME} contain {x}")

        for i in range(len(self.lpl_XFMR2)):
            if (x == self.lpl_XFMR2[i].BUS1.NAME or x == self.lpl_XFMR2[i].BUS2.NAME):
                connected_XFMR2.append(self.lpl_XFMR2[i])
            # else:
            #     print(f"{self.lpl_XFMR2[i].BUS1.NAME} nor {self.lpl_XFMR2[i].BUS2.NAME} contain {x}")

        for i in range(len(self.lpl_XFMR3)):
            if (x == self.lpl_XFMR3[i].BUS1.NAME or x == self.lpl_XFMR3[i].BUS2.NAME or x == self.lpl_XFMR3[i].BUS3.NAME):
                connected_XFMR3.append(self.lpl_XFMR3[i])
            # else:
            #     print(f"{self.lpl_XFMR2[i].BUS1.NAME} nor {self.lpl_XFMR3[i].BUS2.NAME} contain {x}")


        # #debugging statement
        if len(connected_lines) == 0 and len(connected_XFMR2) == 0 and len(connected_XFMR3) == 0:
            print(f"Nothing is connected to {x}.")
            return

        # debugging statment shows the connected lines (buses) and their locations
        print("list of connected lines")
        for i in range(len(connected_lines)):
            print(f"Location 1: {connected_lines[i].BUS1.NAME} Location 2: {connected_lines[i].BUS2.NAME}") #Outputs: [bus 1,bus 2,Circuit ID, BRCODE]
        print("list of connected 2-winding XFMRS")
        for i in range(len(connected_XFMR2)):
            print(f"Location 1: {connected_XFMR2[i].BUS1.NAME} Location 2: {connected_XFMR2[i].BUS2.NAME}") #Outputs: [bus 1,bus 2,Circuit ID, BRCODE]
        print("list of connected 3-winding XFMRS")
        for i in range(len(connected_XFMR3)):
            print(f"Location 1: {connected_XFMR3[i].BUS1.NAME} Location 2: {connected_XFMR3[i].BUS2.NAME} Location 3: {connected_XFMR3[i].BUS3.NAME}") #Outputs: [bus 1,bus 2,Circuit ID, BRCODE]

        connected_relays = set()
        connected_lines
        
        return(connected_relays) # Return the list of connected buses
    

    
    def check_for_int(self, x):
        """
        Checks if the input is an integer.
        """
        try:
            int(x)
            return True
        except ValueError:
            return False


    def between_buses(self, first_bus):
        """
        Returns a list of relays between two buses. This function needs to be refactored
        """
        rlygrp_list = [] # List to store the relay group object between the two buses
        string_list = [] # List to hold the two bus names

        self.get_LPL_buses() # Get the list of LPL buses

        if first_bus == "b" or first_bus == "B":
            return
        elif self.check_for_int(first_bus) and int(first_bus) < len(self.LPL_buses):
            string_list.append(self.LPL_buses[int(first_bus)][0])
        else:
            string_list.append(first_bus.upper()) # Capitilizing each letter from input and adding the first bus name to the list

        print(f"Now looking for the second bus connected to {string_list[0]}")

        list_connected_buses = self.identify_connected_buses(string_list[0])

        if len(list_connected_buses) != 0:
            #prints the list of buses connected to the first bus
            for i in range(len(list_connected_buses)):
                parsed_keystr = ast.literal_eval(list_connected_buses[i].KEYSTR)
                if string_list[0] == parsed_keystr[0][0]:
                    print(f"{i+1}: {parsed_keystr[1][0]}")
                else:
                    print(f"{i+1}: {parsed_keystr[0][0]}")
        else:
            print(f"No buses connected to {string_list[0]}.")
            print("Please check the bus name and try again")
            self.between_buses(input("Enter the first bus name: "))# Recursively call the function to get a new bus name
            return

        #simple input validation to make sure the user enters a number between 1 and the number of connected buses
        x = int(input(f"Please enter a number between 1 and {len(list_connected_buses)} to select second bus:    "))
        while True:
            if x > len(list_connected_buses):
                print(f"Invalid input. Please enter a number between 1 and {len(list_connected_buses)}")
            else:
                break

        parsed_keystr = ast.literal_eval(list_connected_buses[x-1].KEYSTR)

        #adding the second bus name to the list
        if(string_list[0] in parsed_keystr[0][0]):
            string_list.append(parsed_keystr[1][0])
        else:
            string_list.append(parsed_keystr[0][0])

        #looks through self.relays too find every instance that contains some combination of the two strings within the string_list
        #then puts each instance into the rlygrp_list
        #I want to change this but am indeed scared of breaking something (⋟﹏⋞)
        for item in range(len(self.relays)):
            # Parse the ketstr into a list
            parsed_keystr = ast.literal_eval(self.relays[item].KEYSTR)
            if (string_list[0] in parsed_keystr[0][0] or string_list[1] in parsed_keystr[0][0]) and (string_list[0] in parsed_keystr[1][0] or string_list[1] in parsed_keystr[1][0]):
                rlygrp_list.append(self.relays[item])
        
        #when there are 2 objects within the list then it will operate as normally

            print(rlygrp_list)
            for item in rlygrp_list:
                if isinstance(item, OlxObj.RLYGROUP):
                    print(f"Relay Group connected too {item.KEYSTR}: {item.RELAY}")
            return(rlygrp_list)


        #if there is only 1 object that either means that only one relay group is associated with that
        #bus or there is a transformer within the middle and the names of the relay need to be identified
        # in order too link the groups together
    
            rly_IDs_to_be_compared = []
            needed_rly_IDs = []
            needed_rly_IDs.append(rlygrp_list[0].RELAY[0].KEYSTR)
            needed_rly_IDs.append(rlygrp_list[0].RELAY[1].KEYSTR)


            parsed_keystr_rly = ast.literal_eval(needed_rly_IDs[0])
            needed_rly_IDs.insert(0, [parsed_keystr_rly[2], parsed_keystr_rly[3]])
            parsed_keystr_rly = ast.literal_eval(needed_rly_IDs[1])
            needed_rly_IDs.insert(1, [parsed_keystr_rly[2], parsed_keystr_rly[3]])
            print(f"{needed_rly_IDs}\n\n")


            for i in range(len(list_connected_buses)):
                temp = []
                for j in range(len(list_connected_buses[i].RELAY)):
                    parsed_keystr = ast.literal_eval(list_connected_buses[i].RELAY[j].KEYSTR)
                    temp.append([parsed_keystr[2], parsed_keystr[3]])
                rly_IDs_to_be_compared.append(temp)
            for i in range(len(list_connected_buses)):
              print(f"{list_connected_buses[i].KEYSTR}\n\n")

            print(f"{rly_IDs_to_be_compared}\n\n")

            current_index = 0
            for lst in rly_IDs_to_be_compared:
                if any(item in lst for item in needed_rly_IDs):
                    rly_IDs_to_be_compared.clear
                    rly_IDs_to_be_compared.append(lst)
                    break
                current_index += 1


            rlygrp_list.append(list_connected_buses[current_index])
            print(rlygrp_list)
            for item in rlygrp_list:
                if isinstance(item, OlxObj.RLYGROUP):
                    print(f"Relay Group connected too {item.KEYSTR}: {item.RELAY}")

        else:
            print(f"Could not identify relays between those {string_list[0]} and {string_list[1]}.\n")
            print("Please check the bus names and try again.")
            print("=========================================")
            self.get_LPL_buses()
            self.print_list_of_buses(1)
        



def main():
    #Initializing user inputs
    User = UserInput()

    # Connect to OneLiner
    OlxAPI.InitOlxAPI()
    olr_file_path = "S:/lpl_t&d_eng/NEW LP&L ENGINEERING/SUBSTATION/Modeling&Automation/ASPEN/ASPEN Models/2 - Archived/AyrthonTesting/AyrthonsTestModel.OLR"
    while True:    
        try:
            OlxObj.OLCase.open(olr_file_path, 1)
            OlxAPI.LoadDataFile(olr_file_path, 1)
            break  # Exit the loop if the file is loaded successfully
        except Exception as e:
            print(f"Error loading OLR file: {e}")
            User.get_file_location()
            olr_file_path = User.file_location

    #initializing the information classes
    area = locations()
    relay = network_info()

    while User.programming_running:
        match User.user_input:
            case 0:
                User.command_list()
            case 1:
                print("Exiting program.")
                User.user_input = 0
                User.programming_running = False
            case 11:
                print(f"Changing the file location to {User.file_location}")

            case 2:
                print("Printing all areas.")
                area.print_all_areas()
                User.user_input = 0  # Reset user input to allow for new input
            case 3:
                print("Printing all zones.")
                area.print_all_zones()
                User.user_input = 0
            case 4:
                print("Printing all relays.")
                relay.get_all_relays()
                User.user_input = 0
            case 41:
                print("Printing all relays of a certain type:")
                User.str_input()
                relay.get_specific_relays(User.clarification)
                User.user_input = 0
            case 42:
                print("Type a bus name or its number to find relays between two buses. type 'b' to go back.")
                print("Example: \'LP_southest3\' or \'40\'")
                User.str_input()
                relay.between_buses(User.clarification)
                User.user_input = 0
            case 421:
                print("Printing all buses in LPL area.")
                relay.get_LPL_buses()
                relay.print_list_of_buses(1)
                User.user_input = 0
            case default:
                print("\n=====================================")
                print("||Invalid input. Please try again.||")
                print("=====================================")
                User.command_list()

    # Cleanup
    OlxAPI.CloseDataFile()



if __name__ == '__main__':
    # main()

    # Connect to OneLiner
    OlxAPI.InitOlxAPI()
    olr_file_path = "S:/lpl_t&d_eng/NEW LP&L ENGINEERING/SUBSTATION/Modeling&Automation/ASPEN/ASPEN Models/2 - Archived/AyrthonTesting/AyrthonsTestModel.OLR"
    OlxObj.OLCase.open(olr_file_path, 1)
    OlxAPI.LoadDataFile(olr_file_path, 1)

    bus = network_info()

    # bus.get_lpl_lines()
    # bus.get_all_relays()
    # bus.between_buses("lp_vicksbrg3")
    bus.identify_connected_buses("lp_vicksbrg3")
    # bus.get_lpl_XFMR3()

    # Cleanup
    OlxAPI.CloseDataFile()