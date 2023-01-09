#!/usr/bin/env python3

# Importations

import os
import pickle
import pandas as pd

import logging
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    datefmt='%H:%M:%S',
    handlers=[
        # logging.FileHandler("debug.log"), 
        logging.StreamHandler()
    ]
)


import warnings
warnings.filterwarnings("ignore")

from EnergyAlternativesPlaning.f_consumptionModels import *
from Data_processing_functions import *
from Electric_System_model import *

from pyomo.opt import SolverFactory



def Simulation_multinode(xls_file, serialize=False, resultfilename="Result"):
    """Load the input data from the xlsx file then build the pyomo model and call the solver

    Args:
        xls_file : xslx input file, read with pandas
        serialize (bool, optional): Serialize the output of the model. Defaults to False.
        resultfilename (str, optional): Defaults to "Result".

    Returns:
        Output of the solver
    """
    
    start_time = datetime.now()
    year=2018

    ###############
    # Data import #
    ###############

    #Import generation and storage technologies data
    logging.info("Import generation and storage technologies data")
    TechParameters = pd.read_excel(xls_file,"TECHNO_AREAS")
    TechParameters.dropna(inplace=True)
    TechParameters.set_index(["AREAS", "TECHNOLOGIES"], inplace=True)
    StorageParameters = pd.read_excel(xls_file,"STOCK_TECHNO_AREAS").set_index(["AREAS", "STOCK_TECHNO"])

    #Import consumption data
    logging.info("Import consumption data")
    areaConsumption = pd.read_excel(xls_file,"areaConsumption")
    areaConsumption.dropna(inplace=True)
    areaConsumption["Date"] = pd.to_datetime(areaConsumption["Date"])
    areaConsumption.set_index(["AREAS", "Date"], inplace=True)


    #Import availibility factor data
    logging.info("Import availibility factor data")
    availabilityFactor = pd.read_excel(xls_file,"availability_factor")
    availabilityFactor.loc[availabilityFactor.availabilityFactor > 1, "availabilityFactor"] = 1
    availabilityFactor.dropna(inplace=True)
    availabilityFactor["Date"]=pd.to_datetime(availabilityFactor["Date"])
    availabilityFactor.set_index(["AREAS", "Date", "TECHNOLOGIES"],inplace=True)

    #Import interconnections data
    logging.info("Import interconnections data")
    ExchangeParameters = pd.read_excel(xls_file,"interconnexions").set_index(["AREAS", "AREAS.1"])

    ###################
    # Data adjustment #
    ###################
    logging.info("Data adjustment")
    # Temperature sensitivity inclusion
    areaConsumption = Thermosensibility(areaConsumption, xls_file)
    # CHP inclusion for France
    areaConsumption = CHP_processing(areaConsumption, xls_file)
    #Flexibility data inclusion
    ConsoParameters_,labour_ratio,to_flex_consumption=Flexibility_data_processing(areaConsumption,year,xls_file)

    ##############################
    # Model creation and solving #
    ##############################
    
    end_time = datetime.now()
    logging.info('Model creation at {}'.format(end_time - start_time))
    
    model = GetElectricitySystemModel(
        Parameters={
            "areaConsumption": areaConsumption,
            "availabilityFactor": availabilityFactor,
            "TechParameters": TechParameters,
            "StorageParameters": StorageParameters,
            "ExchangeParameters": ExchangeParameters,
            "to_flex_consumption": to_flex_consumption,
            "ConsoParameters_": ConsoParameters_,
            "labour_ratio": labour_ratio
        })
    
    solver = 'mosek'  # 'mosek'  ## no need for solverpath with mosek.
    tee_value = True
    solver_native_list = ["mosek", "glpk"]

    end_time = datetime.now()
    logging.info('Start solving at {}'.format(end_time - start_time))

    if solver in solver_native_list:
        opt = SolverFactory(solver)
        # opt.options['threads'] = 8
    else:
        opt = SolverFactory(solver,executable=solver_path,tee=tee_value)

    results=opt.solve(model)
    end_time = datetime.now()
    logging.info('Solved at {}'.format(end_time - start_time))
    
    ##############################
    # Data extraction and saving #
    ##############################
    Variables = getVariables_panda_indexed(model)

    if serialize:
        with open(resultfilename+'.pickle', 'wb') as f:
            logging.info(f"Saving serialized solver outputs in {resultfilename}.pickle")
            pickle.dump(Variables, f, protocol=pickle.HIGHEST_PROTOCOL)

    end_time = datetime.now()
    logging.info('Total duration: {}'.format(end_time - start_time))
    return Variables


if __name__ == "__main__":
    
    xls_file_path = "Multinode_input.xlsx"
    xls_file = pd.ExcelFile(xls_file_path)
    
    Simulation_multinode(xls_file, serialize=True)
    
    


