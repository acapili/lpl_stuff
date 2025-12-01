import logging
import ast

import sys
import os

# Add the parent directory to the system path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import OlxAPI
from OlxAPIConst import *
import os,sys
from ctypes import * # pyright: ignore[reportWildcardImportFromLibrary]


class UpdateModel:
    OlxAPI.InitOlxAPI()
    intArray = c_int*MXOBJPARAMS
    tokens = intArray()
    voidpArray = c_void_p*MXOBJPARAMS
    params = voidpArray()
    doubleArray = c_double*20
    OlxAPI.CreateNetwork(100)

    def check_length(self, list1: list, list2: list) -> None:
        if len(list1) != len(list2):
            logging.error("Keys and values you do not match length")
            raise ValueError(f"Values are not equal: {len(list1)} != {len(list2)}")
        else:
            return



    def parse(self, idx, data, literal):
            """
            
            """
            # Parameter tokens
            # 100s: string
            # 200s: double
            # 300s: integers
            # 400s: arrays of strings (tab delimited)
            # 500s: arrays of doubles
            # 600s: arrays of integers
            if literal < 200:
                str_var = c_char_p(OlxAPI.encode3(data)) # pyright: ignore[reportArgumentType]
                self.params[idx] = cast(pointer(str_var),c_void_p)
                self.tokens[idx] = literal
            elif literal < 300:
                self.params[idx] = cast(pointer(c_double(data)),c_void_p)
                self.tokens[idx] = literal
            elif literal < 400:
                self.params[idx] = cast(pointer(c_int(data)),c_void_p)
                self.tokens[idx] = literal
            elif literal < 600:
                ratings_formatted = self.doubleArray()
                for index, i in enumerate(data):
                    ratings_formatted[index] = i
                self.params[idx] = cast(pointer(ratings_formatted),c_void_p)
                self.tokens[idx] = literal
            else:
                logging.error("Unknown type")


    def finally_add_to_ASPEN(self, constant: int, hndLoc = 0, BusSide = 0) -> int | None:
        """
        Adds in the equipment or device and then checks to see if the 
        """
        if hndLoc:
            hnd = OlxAPI.AddDevice(c_int(constant),c_int(hndLoc), BusSide,self.tokens,self.params)
        else:
            hnd = OlxAPI.AddEquipment(c_int(constant),self.tokens,self.params)

        if hnd == 0:
            logging.error("AddEquipment failed with tokens: %s", list(self.tokens))
            logging.error("Params: %s", [p for p in self.params])
            logging.error("OlxAPI error: %s", OlxAPI.ErrorString())
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        logging.info( OlxAPI.PrintObj1LPF(hnd) + " had been added successfully")
        intArray = c_int*MXOBJPARAMS
        self.tokens = intArray()
        voidpArray = c_void_p*MXOBJPARAMS
        self.params = voidpArray()
        return hnd
    

    def use_dict(self, input:dict, constant: int, equipment: bool ) -> int | None:
        """
        
        """

        if equipment:
            for idx, (key, value) in enumerate(input.items()):
                self.parse(idx, value, key)
            self.tokens[len(input)] = 0  #zero terminated list
            return self.finally_add_to_ASPEN(constant)
        else:
            hndLoc = input["hndLoc"]
            BusSide = input["BusSide"]
            del input["BusSide"]
            del input["hndLoc"]
            for idx, (key, value) in enumerate(input.items()):
                self.parse(idx, value, key)
            self.tokens[len(input)] = 0  #zero terminated list

            return self.finally_add_to_ASPEN(constant, hndLoc, BusSide)
    

    def addBus(self, values_from_excel: list) -> int | None:
        """
        Add a bus to the model

        parameters:
        values_from_excel = {bus_name: str, location: str ,kv: float, spcx: c_double, spcy: c_double}
        """

        keys = [BUS_sName,BUS_sLocation, BUS_dKVnominal, BUS_dSPCx, BUS_dSPCy]
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        return self.use_dict(my_dict, TC_BUS, True)
    
    def addLine( self, values_from_excel: list) -> int | None:
        """
        Add a line between two buses


        """
        keys = [LN_sName, LN_sID, LN_dLength, LN_sType, LN_dR, LN_dR0, LN_dX, LN_dX0, LN_vdRating, LN_nMeteredEnd, LN_nBus1Hnd, LN_nBus2Hnd]
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        return self.use_dict(my_dict, TC_LINE, True)

    def addXFMR2(self, values_from_excel: list) -> int | None:
        """
        Add a XFMR between two buses
        """

        keys = [XR_sName, XR_sID, XR_dPriTap, XR_dSecTap, XR_dBaseMVA, XR_dMVA1, XR_dMVA2, XR_dMVA3,
                XR_sCfgP, XR_sCfgS, XR_sCfgST,XR_dR, XR_dR0, XR_dX, XR_dX0, XR_nMeteredEnd, XR_nBus1Hnd,XR_nBus2Hnd ]
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        return self.use_dict(my_dict, TC_XFMR, True)

    def addXFMR3(self, values_from_excel: list) -> int | None:
        """
        Add a XFMR between three buses
        """
        keys = [X3_sName, X3_sID, X3_dPriTap, X3_dSecTap, X3_dTerTap, 
                X3_dBaseMVA, X3_dMVA1, X3_dMVA2, X3_dMVA3,
                X3_sCfgP, X3_sCfgS, X3_sCfgST, X3_sCfgT, X3_sCfgTT,
                X3_dRps,X3_dXps, X3_dRpt, X3_dXpt, X3_dRst,X3_dXst, 
                X3_nZ0Method, X3_nBus1Hnd,X3_nBus2Hnd, X3_nBus3Hnd
                ] #AAAHHHHH I just spent half an hour looking for a bug and I had missed that X3_nBus3Hnd was written like X3_nBus2Hnd
        
        #adds in the proper variables depending on if it is utilizing Classical T model impedance
        # or Short circuit impedance
        beginning_of_impedance_vales = 20

        #checks if the model is classic T or short circuit impedance
        if values_from_excel[-4]:
            temp_array = [X3_dR0pm, X3_dX0pm, X3_dR0sm, X3_dX0sm, X3_dR0mg, X3_dX0mg]
            for i, x in enumerate(temp_array):
                keys.insert(beginning_of_impedance_vales + i, x)
        else:
            temp_array = [X3_dR0ps, X3_dX0ps, X3_dR0pt, X3_dX0pt, X3_dR0st, X3_dX0st]
            for i, x in enumerate(temp_array):
                keys.insert(beginning_of_impedance_vales + i, x)
        
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        return self.use_dict(my_dict, TC_XFMR3, True)
    
    def addSWITCH(self, values_from_excel) -> int | None:
        """
        
        """
        keys = [SW_sName,SW_sID, SW_dRating, SW_nDefault, SW_nBus1Hnd, SW_nBus2Hnd]
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        return self.use_dict(my_dict, TC_SWITCH, True)
    
    def addGEN(self, values_from_excel) -> int | None:
        """
        
        """ 
        # GE_dPgen, GE_dQgen
        keys = [OBJ_sGUID, GE_dRefAngle, GE_dCurrLimit1, GE_dCurrLimit2, GE_dVSourcePU, GE_nCtrlBusHnd, GE_nBusHnd]
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        return self.use_dict(my_dict, TC_GEN, True)
    
    def addGENUNIT(self, values_from_excel) -> None:
        """
        
        """
        keys = [GU_sID, GU_dMVArating, GU_vdR, GU_vdX, GU_dRz, GU_dXz,
                GU_dSchedP, GU_dSchedQ, GU_dPmax, GU_dPmin,GU_dQmax, GU_dQmin,GU_nGenHnd]
        
        list_GU_vdR = list(map(float, values_from_excel[2].split(",")))
        list_GU_vdX = list(map(float, values_from_excel[3].split(",")))
        values_from_excel[2] = list_GU_vdR
        values_from_excel[3] = list_GU_vdX
        
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        self.use_dict(my_dict, TC_GENUNIT, True)


    def addFUSE(self, values_from_excel) -> None:
        """
        
        """
        keys = [FS_sID, FS_dRating, FS_sType, "hndLoc", "BusSide"]
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        self.use_dict(my_dict, TC_FUSE, True)


    def addReclose(self, values_from_excel) -> None:
        """
        Reclosers have an interesting problem to deal with
        """
        keys = [CP_sID, CP_dRecIntvl1, CP_dRecIntvl2, CP_dRecIntvl3, CP_nTotalOps, CP_nFastOps, CP_nCurveInUse,
                CP_nCurveInUse, CP_nCurveInUse, "hndLoc", "BusSide"
                ]
        def insert_items(into_where: int, new_keys: list):
            for idx, key in enumerate(new_keys):
                keys.insert(into_where + idx, key)
        #
        if values_from_excel[7] == 1:
            list_for_differences = [CG_dMinTF, CG_dPickupF, CG_dTimeAddF, CG_dTimeMultF, CG_sTypeFast]
            insert_items(8, list_for_differences)
        else:
            list_for_differences = [CG_dMinTS, CG_dPickupS, CP_dTimeAddS, CG_dTimeMultS, CG_sTypeSlow]
            insert_items(8, list_for_differences)

        if values_from_excel[12] == 1:
            list_for_differences = [CG_dMinTF, CG_dPickupF, CG_dTimeAddF, CG_dTimeMultF, CG_sTypeFast]
            insert_items(13, list_for_differences)
        else:
            list_for_differences = [CG_dMinTS, CG_dPickupS, CG_dTimeAddS, CG_dTimeMultS, CG_sTypeSlow]
            insert_items(13, list_for_differences)

        #makes an attempt at adding in the material
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        self.use_dict(my_dict, TC_RLYOCP, False)



    def addOCPRelay(self, values_from_excel) -> None:
        """
        adds a overcurrent relay to a line
        """

        keys = [OP_sID, OP_dCT, OP_nByCTConnect, OP_dMinTripTime, OP_dResetTime, 
                OP_sType, OP_sTapType, OP_dTap, OP_dTDial, OP_dTimeAdd, OP_nDirectional, OP_nSignalOnly, OP_nVoltControl,
                OP_vdDTPickup, OP_vdDTDelay, OP_dInst, OP_nDCOffset, OP_dTimeAdd2, OP_nIDirectional, OP_nPolar,
                "hndLoc", "BusSide"]
        
        list504 = list(map(float, values_from_excel[13].split(",")))
        list506 = list(map(float, values_from_excel[14].split(",")))
        values_from_excel[13] = list504
        values_from_excel[14] = list506

        #makes an attempt at adding in the material
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        self.use_dict(my_dict, TC_RLYOCP, False)


    def addOCGRelay(self, values_from_excel) -> None:
        """
        adds a overcurrent relay to a line
        """
        keys = [OG_sID,OG_dCT, OG_vdDTPickup, OG_vdDTDelay, OG_dInst, OG_nCTLocation, "hndLoc", "BusSide"]
        self.check_length(values_from_excel, keys)
        my_dict = dict(zip(keys, values_from_excel))
        self.use_dict(my_dict, TC_RLYOCG, False)
    

    def addDSGRelay(self, ID: str, type: str, CTR : int, VTR : int, DSParams : str, hndLn: int) -> None:

        keys = [DG_sID, DG_sDSType, DG_dCT, DG_sParam]
        # Add a DS relay to the line on bus1 side
        sID = c_char_p(OlxAPI.encode3(ID))# pyright: ignore[reportArgumentType]
        self.params[0] = cast(pointer(sID),c_void_p)
        self.tokens[0] = DG_sID
        DSType = c_char_p(OlxAPI.encode3(type))# pyright: ignore[reportArgumentType]
        self.params[1] = cast(pointer(DSType),c_void_p)
        self.tokens[1] = DG_sDSType
        self.params[2] = cast(pointer(c_double(CTR)),c_void_p)
        self.tokens[2] = DG_dCT
        self.params[3] = cast(pointer(c_double(VTR)),c_void_p)
        self.tokens[3] = DG_dVT
        self.params[4] = cast(pointer(c_char_p(OlxAPI.encode3(DSParams))),c_void_p)# pyright: ignore[reportArgumentType]
        self.tokens[4] = DG_sParam
        self.tokens[5] = 0  # Must terminate the param list with zero
        hndDS = OlxAPI.AddDevice(c_int(TC_RLYDSG),c_int(hndLn),0,self.tokens,self.params)
        if hndDS == 0:
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        logging.info( OlxAPI.PrintObj1LPF(hndDS) + " had been added successfully")


    def addDSPRelay(self, ID: str, type: str, CTR : int, VTR : int, DSParams : str, hndLn: int) -> None:
        # Add a DS relay to the line on bus1 side
        sID = c_char_p(OlxAPI.encode3(ID))# pyright: ignore[reportArgumentType]
        self.params[0] = cast(pointer(sID),c_void_p)
        self.tokens[0] = DP_sID
        DSType = c_char_p(OlxAPI.encode3(type))# pyright: ignore[reportArgumentType]
        self.params[1] = cast(pointer(DSType),c_void_p)
        self.tokens[1] = DP_sDSType
        self.params[2] = cast(pointer(c_double(CTR)),c_void_p)
        self.tokens[2] = DP_dCT
        self.params[3] = cast(pointer(c_double(VTR)),c_void_p)
        self.tokens[3] = DP_dVT
        self.params[4] = cast(pointer(c_char_p(OlxAPI.encode3(DSParams))),c_void_p)# pyright: ignore[reportArgumentType]
        self.tokens[4] = DP_sParam
        self.tokens[5] = 0  # Must terminate the param list with zero
        hndDS = OlxAPI.AddDevice(c_int(TC_RLYDSG),c_int(hndLn),0,self.tokens,self.params)
        if hndDS == 0:
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        logging.info( OlxAPI.PrintObj1LPF(hndDS) + " had been added successfully")


    def saveModel(self, olrFilePathNew: str) -> None:
        OlxAPI.Run1LPFCommand('<SNAPSPC OPTIONS="7"/>')
        if OLXAPI_FAILURE == OlxAPI.SaveDataFile(olrFilePathNew):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print( olrFilePathNew + " had been saved successfully")
