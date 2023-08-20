#Turbinedisk_ThermoMechanical
1. Entire thermomechanical simulation converted into a Python compiler
2. This code will import the .stp file and define the design variables on it
3. In the current project, I have solved four different load cases as **Mechanical, Thermal_displacement, Thermal load, and Thermal load and fixed constrain load cases**.
4. 100 designs using LHS generated for 8 designs and 8 material & boundary variables
5. Using **runscript** option in the ABAQUS software, we can run the code. It will load the design variable and material properties from the Excel sheet and simulate all 100 simulations.
6. Responses of simulation such as **Von Mises stress, displacement, hoop stress, thermal flux, and Longitudinal stress** are saved in the new Excel sheet.

