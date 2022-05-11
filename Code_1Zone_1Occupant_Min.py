#%% importing requirements
#%%% Reading libraries
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import random
import re
import gurobipy as gp
from gurobipy import GRB
import random
import pwlf
import warnings
import seaborn as sns
warnings.filterwarnings("ignore")

#%%% Importing electricity price
Electricity_Price = pd.read_excel("DDRC implementation.xlsx",
                                  sheet_name = "Electricity Price")
Electricity_Price = Electricity_Price["Elecricity Price"]/1000
Electricity_Price = Electricity_Price[0:96]
#%%% Importing thermal comfort profiles
Profiles_Dataset = pd.read_csv("G:\Shared drives\_Research Repository - Mostafa Meimand\Working Papers\+HVAC controller integrated with Personal Thermal Comfort and Real Time Price\\ComfortProfiling\min_profiles.csv")

#%%
n_segments = 4 # best number based on different tests
mu = 0.5
#%%
def approx_PWLF(Subj1,Subj2): # approaximating objective function and constraints
    Multi = Profiles_Dataset["Probability" + str(Subj1)] + Profiles_Dataset["Probability" + str(Subj2)]
    Multi = Multi - Multi.min()
    Multi = Multi / Multi.max()
    
    my_pwlf = pwlf.PiecewiseLinFit(Profiles_Dataset["Temperature"], Multi)
    breaks = my_pwlf.fit(n_segments).round(2)
    y_breaks = my_pwlf.predict(breaks).round(2)
    
    maxConstraint = Profiles_Dataset["Temperature"][Multi.round(1) == mu].max()
    minConstraint = Profiles_Dataset["Temperature"][Multi.round(1) == mu].min()

    return breaks, y_breaks, maxConstraint, minConstraint

#%%
def ReadExcel(name):
    Output = pd.read_csv("C:/Users/tianzhi - adm/Desktop/ReactiveController_Commercial_9.0 -TwoOccupants_min/" + name + ".csv")
    
    # Adding dates to the dataset
    delimiters = " ", ":", "/"
    regexPattern = '|'.join(map(re.escape, delimiters))
    Output["Month"] = None
    Output["Day"] = None
    Output["Hour"] = None
    Output["Minutes"] = None
    
    for i in range(Output.shape[0]):
      Output["Month"][i] = int(re.split(regexPattern,Output["Date/Time"][i])[1])
      Output["Day"][i] = int(re.split(regexPattern,Output["Date/Time"][i])[2])
      Output["Hour"][i] = int(re.split(regexPattern,Output["Date/Time"][i])[4])
      Output["Minutes"][i] = int(re.split(regexPattern,Output["Date/Time"][i])[5])
    
    Output = Output[Output["Day"] == 1]
    Output["PSZ-AC:1:Air System Total Cooling Energy [J](TimeStep)"] *= 2.77778e-7
    Output["PSZ-AC:2:Air System Total Cooling Energy [J](TimeStep)"] *= 2.77778e-7
    Output["PSZ-AC:3:Air System Total Cooling Energy [J](TimeStep)"] *= 2.77778e-7
    Output["PSZ-AC:4:Air System Total Cooling Energy [J](TimeStep)"] *= 2.77778e-7
    Output["PSZ-AC:5:Air System Total Cooling Energy [J](TimeStep) "] *= 2.77778e-7

    Output["time"] = Output.index
    return Output

def textGenerator(Zone):
    k = 0
    String = "EnergyManagementSystem:Program," + "\n"
    String += "MyComputedCoolingSetpointProg_" + str(Zone + 1) + "," + "\n"
    String += "IF (Hour == " + str(X["Hours"][k]) + ") && (Minute  <=  " + str(X["Minutes"][k]) + "),  Set myCLGSETP_SCH_Override_" + str(Zone + 1) + " = " + str(X["Zone " + str(Zone)][k]) + "," 
    
    for i in range(1,96):
        String += "ELSEIF (Hour == " + str(X["Hours"][k]) + ") && (Minute  <=  " + str(X["Minutes"][k]) + "),  Set myCLGSETP_SCH_Override_" + str(Zone + 1) + " = " + str(X["Zone " + str(Zone)][k]) + "," 
        k += 1
    String += "ENDIF;"
    return String

def CoSimulation():
    text = open("OfficeSmall_main.txt").read()
    NextFile = open("OfficeSmall_1.IDF","wt")
    NextFile.write(text[:300896] + '\n' + '\n' + textGenerator(0) + '\n' + '\n' + textGenerator(1) + '\n'
                   + '\n' + textGenerator(2) + '\n' + '\n' + textGenerator(3) + '\n' + '\n' + textGenerator(4) + '\n'
                   + '\n' +  text[300899:])
    NextFile.close()
    os.system("energyplus -w USA_TX_Austin-Mueller.Muni.AP.722540_TMY3.epw -r OfficeSmall_1.idf")

def X_Preparation():
    X = ReadExcel("OfficeSmall_main")
    X["Index"] = X.index
    X["Minutes"] = X["Index"] % 4 * 15 + 15
    temp = []
    for i in range(0,24):
        temp.append([i] * 4)
    X["Hours"] = np.reshape(temp, (1,96))[0]
    for zone in range(5):
        X["Previous Temperature_" + str(zone)] = None
        X["Previous Temperature_" + str(zone)][0] = 23
        X["Zone " + str(zone)] = 29.44
        X["Zone " + str(zone)][32:36] = 20
    CoSimulation()
    X["Previous Temperature_0"] = ReadExcel("eplusout")["CORE_ZN:Zone Mean Air Temperature [C](TimeStep)"]
    X["Previous Temperature_1"] = ReadExcel("eplusout")["PERIMETER_ZN_1:Zone Mean Air Temperature [C](TimeStep)"]
    X["Previous Temperature_2"] = ReadExcel("eplusout")["PERIMETER_ZN_2:Zone Mean Air Temperature [C](TimeStep)"]
    X["Previous Temperature_3"] = ReadExcel("eplusout")["PERIMETER_ZN_3:Zone Mean Air Temperature [C](TimeStep)"]
    X["Previous Temperature_4"] = ReadExcel("eplusout")["PERIMETER_ZN_4:Zone Mean Air Temperature [C](TimeStep)"]
    X["Energy_0"] = ReadExcel("eplusout")['PSZ-AC:1:Air System Total Cooling Energy [J](TimeStep)']
    X["Energy_1"] = ReadExcel("eplusout")['PSZ-AC:2:Air System Total Cooling Energy [J](TimeStep)']
    X["Energy_2"] = ReadExcel("eplusout")['PSZ-AC:3:Air System Total Cooling Energy [J](TimeStep)']
    X["Energy_3"] = ReadExcel("eplusout")['PSZ-AC:4:Air System Total Cooling Energy [J](TimeStep)']
    X["Energy_4"] = ReadExcel("eplusout")['PSZ-AC:5:Air System Total Cooling Energy [J](TimeStep) ']

    return X
#%%
X = ReadExcel("OfficeSmall_main")
X["Index"] = X.index
X["Minutes"] = X["Index"] % 4 * 15 + 15
temp = []
for i in range(0,24):
    temp.append([i] * 4)
X["Hours"] = np.reshape(temp, (1,96))[0]
for zone in range(5):
    X["Previous Temperature_" + str(zone)] = None
    X["Previous Temperature_" + str(zone)][0] = 23
    X["Zone " + str(zone)] = 29.44
    X["Zone " + str(zone)][32:36] = 20

#%%
Subjects = pd.read_csv("2_min_subjects.csv")
for counter in range(120):
    X = X_Preparation()

    #Controling for a day
    Agents = Subjects.loc[counter][1:]
    
    for timestep in range(35,72):
        eta = 1
        k_p = 0.1
        ##########################################################################        
        #### Problem formulation for zone 0
        model = gp.Model("optim")
        z = model.addVar(name="z") # value of the objective function
        x = model.addVar(name="x") # next temperature
        delta_u = model.addVar(name="delta_u",lb=-20, ub=+20)
        # model.addConstr(X["Zone 0"][timestep-1] + delta_u >= 16)
        # setpoint of the building
        # Adding constraints
        model.addConstr(x == -0.0010 * X["Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)"][timestep] 
                        - 0.0035 * X["Previous Temperature_0"][timestep-1] + delta_u * 0.4254 + 0.1020 + 
                        X["Previous Temperature_0"][timestep-1])
    
        my_pwlf = approx_PWLF(Agents[0],Agents[1])
        model.addConstr(x <= my_pwlf[2])
        model.addConstr(x >= my_pwlf[3])
        
        # Auxilary varialbes for the second term
        x1 = model.addVar(name="x1") 
        x2 = model.addVar(name="x2")
        x3 = model.addVar(name="x3")
        x4 = model.addVar(name="x4")
        x5 = model.addVar(name="x5")
        model.addConstr(x == my_pwlf[0][0] * x1 + my_pwlf[0][1] * x2 + my_pwlf[0][2] * x3 + my_pwlf[0][3] * x4 + my_pwlf[0][4] * x5)
        model.addConstr(x1 + x2 + x3 + x4 + x5 == 1)
        model.addSOS(GRB.SOS_TYPE2, [x1, x2 , x3, x4, x5])
        
        # defining opjective function       
        model.addConstr(z == (0.0435 * X["Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)"][timestep]
                              - 0.1011 * x 
                              - 0.3524 * (X["Zone 0"][timestep-1] + delta_u) + 10.5841) * Electricity_Price[timestep]
                              - eta * (my_pwlf[1][0] * x1 + my_pwlf[1][1] * x2 + my_pwlf[1][2] * x3 + 
                                       my_pwlf[1][3] * x4 + my_pwlf[1][4] * x5)
                              + 10000)
        
        model.setObjective(z, GRB.MINIMIZE)
        model.optimize()
        
        X["Zone 0"][timestep + 1] = k_p * round(model.getVars()[2].x,2) + X["Zone 0"][timestep]
        ##########################################################################  
        #### Problem formulation for zone 1
        model = gp.Model("optim")
        z = model.addVar(name="z") # value of the objective function
        x = model.addVar(name="x") # next temperature
        delta_u = model.addVar(name="delta_u",lb=-20, ub=+20)
        # model.addConstr(X["Zone 1"][timestep-1] + delta_u >= 16)
        # setpoint of the building
        # Adding constraints
        model.addConstr(x == -0.00185 * X["Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)"][timestep] + 
                        0.0076 * X["Previous Temperature_1"][timestep-1] + delta_u * 0.4430 - 0.1351 + 
                        X["Previous Temperature_1"][timestep-1])
    
        my_pwlf = approx_PWLF(Agents[2],Agents[3])
        model.addConstr(x <= my_pwlf[2])
        model.addConstr(x >= my_pwlf[3])
        
        # Auxilary varialbes for the second term
        x1 = model.addVar(name="x1") 
        x2 = model.addVar(name="x2")
        x3 = model.addVar(name="x3")
        x4 = model.addVar(name="x4")
        x5 = model.addVar(name="x5")
        model.addConstr(x == my_pwlf[0][0] * x1 + my_pwlf[0][1] * x2 + my_pwlf[0][2] * x3 + my_pwlf[0][3] * x4 + my_pwlf[0][4] * x5)
        model.addConstr(x1 + x2 + x3 + x4 + x5 == 1)
        model.addSOS(GRB.SOS_TYPE2, [x1, x2 , x3, x4, x5])
        
        # defining opjective function       
        model.addConstr(z == (0.0664 * X["Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)"][timestep]
                              + 0.0288 * x 
                              - 0.3228 * (X["Zone 1"][timestep-1] + delta_u) + 5.9760) * Electricity_Price[timestep]
                              - eta * (my_pwlf[1][0] * x1 + my_pwlf[1][1] * x2 + my_pwlf[1][2] * x3 + 
                                       my_pwlf[1][3] * x4 + my_pwlf[1][4] * x5)
                              + 10000)
        
        model.setObjective(z, GRB.MINIMIZE)
        model.optimize()
        
        X["Zone 1"][timestep + 1] = k_p * round(model.getVars()[2].x,2) + X["Zone 1"][timestep]
    
        ##########################################################################  
        #### Problem formulation for zone 2
        model = gp.Model("optim")
        z = model.addVar(name="z") # value of the objective function
        x = model.addVar(name="x") # next temperature
        delta_u = model.addVar(name="delta_u",lb=-20, ub=+20)
        # model.addConstr(X["Zone 2"][timestep-1] + delta_u >= 16)
        # setpoint of the building
        # Adding constraints
        model.addConstr(x == 0.0065 * X["Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)"][timestep]
                        - 0.1338 * X["Previous Temperature_2"][timestep-1] + delta_u * 0.3635 + 2.8921 +  
                        X["Previous Temperature_2"][timestep-1])
        
        my_pwlf = approx_PWLF(Agents[4],Agents[5])
        model.addConstr(x <= my_pwlf[2])
        model.addConstr(x >= my_pwlf[3])
        
        # Auxilary varialbes for the second term
        x1 = model.addVar(name="x1") 
        x2 = model.addVar(name="x2")
        x3 = model.addVar(name="x3")
        x4 = model.addVar(name="x4")
        x5 = model.addVar(name="x5")
        model.addConstr(x == my_pwlf[0][0] * x1 + my_pwlf[0][1] * x2 + my_pwlf[0][2] * x3 + my_pwlf[0][3] * x4 + my_pwlf[0][4] * x5)
        model.addConstr(x1 + x2 + x3 + x4 + x5 == 1)
        model.addSOS(GRB.SOS_TYPE2, [x1, x2 , x3, x4, x5])
        
        # defining opjective function       
        model.addConstr(z == (0.0325 * X["Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)"][timestep]
                              - 0.0839 * x
                              - 0.1158 * (X["Zone 2"][timestep-1] + delta_u) + 4.1980) * Electricity_Price[timestep]
                              - eta * (my_pwlf[1][0] * x1 + my_pwlf[1][1] * x2 + my_pwlf[1][2] * x3 + 
                                       my_pwlf[1][3] * x4 + my_pwlf[1][4] * x5)
                              + 10000)
        
        model.setObjective(z, GRB.MINIMIZE)
        model.optimize()
        
        X["Zone 2"][timestep + 1] = k_p * round(model.getVars()[2].x,2) + X["Zone 2"][timestep]
    
        ##########################################################################  
        #### Problem formulation for zone 3
        model = gp.Model("optim")
        z = model.addVar(name="z") # value of the objective function
        x = model.addVar(name="x") # next temperature
        delta_u = model.addVar(name="delta_u",lb=-20, ub=+20)
        # model.addConstr(X["Zone 3"][timestep-1] + delta_u >= 16)
        # setpoint of the building
        # Adding constraints
        model.addConstr(x == 0.0038 * X["Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)"][timestep]
                         + 0.0168 * X["Previous Temperature_3"][timestep-1] + delta_u * 0.4559 - 0.4987 + 
                        X["Previous Temperature_3"][timestep-1])
        my_pwlf = approx_PWLF(Agents[6],Agents[7])
        model.addConstr(x <= my_pwlf[2])
        model.addConstr(x >= my_pwlf[3])
        
        # Auxilary varialbes for the second term
        x1 = model.addVar(name="x1")
        x2 = model.addVar(name="x2")
        x3 = model.addVar(name="x3")
        x4 = model.addVar(name="x4")
        x5 = model.addVar(name="x5")
        model.addConstr(x == my_pwlf[0][0] * x1 + my_pwlf[0][1] * x2 + my_pwlf[0][2] * x3 + my_pwlf[0][3] * x4 + my_pwlf[0][4] * x5)
        model.addConstr(x1 + x2 + x3 + x4 + x5 == 1)
        model.addSOS(GRB.SOS_TYPE2, [x1, x2 , x3, x4, x5])
        
        # defining opjective function       
        model.addConstr(z == (0.0452 * X["Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)"][timestep]
                              - 0.1492 * x
                              - 0.2293 * (X["Zone 3"][timestep-1] + delta_u) + 8.5402) * Electricity_Price[timestep]
                              - eta * (my_pwlf[1][0] * x1 + my_pwlf[1][1] * x2 + my_pwlf[1][2] * x3 + 
                                       my_pwlf[1][3] * x4 + my_pwlf[1][4] * x5)
                              + 10000)
        model.setObjective(z, GRB.MINIMIZE)
        model.optimize()
    
        X["Zone 3"][timestep + 1] = k_p * round(model.getVars()[2].x,2) + X["Zone 3"][timestep]
    
        ##########################################################################  
        #### Problem formulation for zone 4
        model = gp.Model("optim")
        z = model.addVar(name="z") # value of the objective function
        x = model.addVar(name="x") # next temperature
        delta_u = model.addVar(name="delta_u",lb=-20, ub=+20)
        # model.addConstr(X["Zone 4"][timestep-1] + delta_u >= 16)
        # setpoint of the building
        # Adding constraints
        model.addConstr(x == 0.0051 * X["Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)"][timestep] 
                        - 0.0299 * X["Previous Temperature_4"][timestep-1] + delta_u * 0.4656 + 0.5700 + 
                        X["Previous Temperature_4"][timestep-1])
    
        my_pwlf = approx_PWLF(Agents[8],Agents[9])
        model.addConstr(x <= my_pwlf[2])
        model.addConstr(x >= my_pwlf[3])
        
        # Auxilary varialbes for the second term
        x1 = model.addVar(name="x1") 
        x2 = model.addVar(name="x2")
        x3 = model.addVar(name="x3")
        x4 = model.addVar(name="x4")
        x5 = model.addVar(name="x5")
        model.addConstr(x == my_pwlf[0][0] * x1 + my_pwlf[0][1] * x2 + my_pwlf[0][2] * x3 + my_pwlf[0][3] * x4 + my_pwlf[0][4] * x5)
        model.addConstr(x1 + x2 + x3 + x4 + x5 == 1)
        model.addSOS(GRB.SOS_TYPE2, [x1, x2 , x3, x4, x5])
        
        # defining opjective function       
        model.addConstr(z == (0.0263 * X["Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)"][timestep]
                              - 0.0408 * x 
                              - 0.1692 * (X["Zone 4"][timestep-1] + delta_u) + 4.8648) * Electricity_Price[timestep]
                              - eta * (my_pwlf[1][0] * x1 + my_pwlf[1][1] * x2 + my_pwlf[1][2] * x3 + 
                                       my_pwlf[1][3] * x4 + my_pwlf[1][4] * x5)
                              + 10000)
        
        model.setObjective(z, GRB.MINIMIZE)
        model.optimize()
        
        X["Zone 4"][timestep + 1] = k_p * round(model.getVars()[2].x,2) + X["Zone 4"][timestep]
    
        
        ###################################
        CoSimulation()
        X["Previous Temperature_0"] = ReadExcel("eplusout")["CORE_ZN:Zone Mean Air Temperature [C](TimeStep)"]
        X["Previous Temperature_1"] = ReadExcel("eplusout")["PERIMETER_ZN_1:Zone Mean Air Temperature [C](TimeStep)"]
        X["Previous Temperature_2"] = ReadExcel("eplusout")["PERIMETER_ZN_2:Zone Mean Air Temperature [C](TimeStep)"]
        X["Previous Temperature_3"] = ReadExcel("eplusout")["PERIMETER_ZN_3:Zone Mean Air Temperature [C](TimeStep)"]
        X["Previous Temperature_4"] = ReadExcel("eplusout")["PERIMETER_ZN_4:Zone Mean Air Temperature [C](TimeStep)"]
        X["Energy_0"] = ReadExcel("eplusout")['PSZ-AC:1:Air System Total Cooling Energy [J](TimeStep)']
        X["Energy_1"] = ReadExcel("eplusout")['PSZ-AC:2:Air System Total Cooling Energy [J](TimeStep)']
        X["Energy_2"] = ReadExcel("eplusout")['PSZ-AC:3:Air System Total Cooling Energy [J](TimeStep)']
        X["Energy_3"] = ReadExcel("eplusout")['PSZ-AC:4:Air System Total Cooling Energy [J](TimeStep)']
        X["Energy_4"] = ReadExcel("eplusout")['PSZ-AC:5:Air System Total Cooling Energy [J](TimeStep) ']
    
    s = "min_results_"
    for i in range(10):
        s += str(Agents[i]) + "_"
    s += str(k_p) + ".csv"

    X.to_csv(s)
#%%
pd.DataFrame(Subjects).to_csv("Two_Min.csv")