# Seminar Supply Chain Management and Optimisation
Programming Code and Data Files used for the Seminar in Supply Chain Management and Optimisation 2024 by Femke van den Berg, Sammie Zwaan, Selene Bruni and Mumbi Gachara.

This project is about forecasting the 2028 demand of white goods in Europe for Haier, and modelling the supply chain network to transport the white goods from Haier's factory in China all the way to Haier's customers all around Europe. The Cplex Solver is used for the MILP problems. The following MILP models are used:

Model 1: Unimodal LDCs only 

Model 2: Unimodal EDCs only

Model 3: Unimodal Hybrid

Model 4: Multimodal minCost Problem

Model 5: Multimodal minEmissions Problem

Model 6: Multimodal minCost+minEmissionsCost Problem

The files in this repository include:

1.) 'trucksOnly' -> Java code for Modela 1-3

2.) 'minCostProblem' -> Java code for Model 4

3.) 'minEmissions' -> Java code for Model 5

4.) 'minCostsEmissions' -> Java code for Model 6

5.) 'forPareto' -> Java code for Pareto Analysis

6.) 'Data4' -> Input data for Model 4

7.) 'Data4Em' -> Input data for Model 5

8.) 'Data4Both_new' -> Input data for Model 6

9.) 'DataTruckOnly' -> Input data for Models 1-3
