#!/usr/bin/env python3

# Importations

import sys
import os
import pickle
import pandas as pd
import logging

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
    
    dt0 = datetime(1,1,1)
    ref_time = datetime.now()
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
    
    logging.info('Input data loading time : ' + (dt0+(datetime.now()-ref_time)).strftime('%Mm:%Ss'))
    ref_time = datetime.now()
    
    logging.info('Start creating model')

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

    logging.info('Model creation time : ' + (dt0+(datetime.now()-ref_time)).strftime('%Mm:%Ss'))
    ref_time = datetime.now()

    logging.info('Start solving')

    if solver in solver_native_list:
        opt = SolverFactory(solver)
        # opt.options['threads'] = 8
    else:
        opt = SolverFactory(solver,executable=solver_path,tee=tee_value)

    results=opt.solve(model)
    
    logging.info('Solving time : ' + (dt0+(datetime.now()-ref_time)).strftime('%Mm:%Ss'))

    ##############################
    # Data extraction and saving #
    ##############################
    Variables = getVariables_panda_indexed(model)

    if serialize:
        with open(resultfilename+'.pickle', 'wb') as f:
            logging.info(f"Saving serialized solver outputs in {resultfilename}.pickle")
            pickle.dump(Variables, f, protocol=pickle.HIGHEST_PROTOCOL)

    logging.info('Solving time : ' + (dt0+(datetime.now()-start_time)).strftime('%Mm:%Ss'))
    
    return Variables


if __name__ == "__main__":
    
    # Reading input data file
    try:
        xls_file_path = sys.argv[1]
    except IndexError:
        print(f"ERROR : missing input file\nUsage : {sys.argv[0]} <input_data.xlsx>")
        exit()
        
    try:
        xls_file = pd.ExcelFile(xls_file_path)
    except ValueError as error:
        print("ERROR : issue with input data file format, original message:\n")
        print(error)
        
    # get the working directory
    working_dir = os.path.dirname(xls_file_path)
    if working_dir == "" : working_dir = os.getcwd()
    os.chdir(working_dir)

    # configure logging
    logging.basicConfig(
        level=logging.INFO,
        format="[%(asctime)s] [%(levelname)s] %(message)s",
        datefmt='%H:%M:%S',
        handlers=[
            logging.FileHandler("simulation.log"), 
            logging.StreamHandler()
        ]
    )

    # running the simulation
    Simulation_multinode(xls_file, serialize=True)
    
    


