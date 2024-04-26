import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
//import java.util.Arrays;
//import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ilog.concert.IloException;
import ilog.concert.IloNumExpr;
import ilog.concert.IloNumVar;
import ilog.cplex.IloCplex;
/**
 * @author makena mumbi
 *
 */
public class minCostsEmissions {

	/**
	 * @param args
	 */
	public static void main(String[] args) throws IOException {
		//put the correct data
		FileInputStream file = new FileInputStream(new File("C:\\Users\\maken\\OneDrive - Erasmus University Rotterdam\\Documents\\UPT Master\\Seminar Supply Chain\\Data4Both_new.xlsx"));
		Workbook workbook = new XSSFWorkbook(file);
		// get the multiple sheets
		Sheet sheet1 = workbook.getSheet("PortEDCBoth");
		Sheet sheet2 = workbook.getSheet("PortDCBoth");
		Sheet sheet3 = workbook.getSheet("EDCDemandBoth");
		Sheet sheet4 = workbook.getSheet("DCArbitraryBoth");
		Sheet sheet5 = workbook.getSheet("Demand");
		Sheet sheet6 = workbook.getSheet("variableEDC");
		Sheet sheet7 = workbook.getSheet("EDCcap");
		Sheet sheet8 = workbook.getSheet("WarehouseCostsLDCperContainer");
		Sheet sheet9 = workbook.getSheet("fixedEDC"); 
		Sheet sheet10 = workbook.getSheet("WarehouseEmissionsCost"); //emissions*99.8eur/CO2t
		
		
		//Ports = 12
				//EDCs = 11
				//DCs/Arbitrary = 29
				
				int numPorts = sheet1.getLastRowNum()+1;
				int numEDCs = sheet1.getRow(0).getLastCellNum();
				int numLDCs = sheet2.getRow(0).getLastCellNum();
				

				//inland transport China + shipping&Terminal handling costs + land Transportation costs + Inventory costs
				//predetermined mixed mode of transport
				//both EDCs and LDCs
				double[][] costPortEDC = new double[numPorts][numEDCs]; 
				double[][] costPortLDC = new double[numPorts][numLDCs]; //
				
				//fixed cost of opening an (E/L)DC (warehouse costs)
			//	double fixedDC = 621.0;
				double[][] variableEDC = new double[numEDCs][numLDCs];
				double [] warehouseLDC = new double [numLDCs];
				double [] edcCap = new double[numEDCs];
				double [] fixedEDC = new double[numEDCs];
				
				
				//costs from EDC-> Demand Region 
				//predetermined mode of transport
				double[][] costEDCDemandRegion = new double [numEDCs][numLDCs];  //
				//cost from LDC to Arbitrary city
				double[] costLDCArbitrary = new double [numLDCs]; //
				
				//Demanded number of containers
				int [] demand = new int [numLDCs];  //
				
				//warehouse emissions costs
				double [] warehouseEmissionsCosts = new double[numLDCs];
				
				//update the sheets
				for (int j = 0; j < numLDCs; j++) {
						demand[j] = (int)sheet5.getRow(0).getCell(j).getNumericCellValue();
						costLDCArbitrary[j] = sheet4.getRow(0).getCell(j).getNumericCellValue();
						warehouseLDC[j] = sheet8.getRow(0).getCell(j).getNumericCellValue();
						warehouseEmissionsCosts[j] = sheet10.getRow(0).getCell(j).getNumericCellValue();
						
					for (int i = 0; i < numPorts; i++) {
						costPortLDC[i][j] = sheet2.getRow(i).getCell(j).getNumericCellValue();
					}
					for(int i=0; i<numEDCs; i++) {
						costEDCDemandRegion[i][j] = sheet3.getRow(i).getCell(j).getNumericCellValue();
					}
				}
				
				for(int i=0; i<numPorts; i++) {
					for(int j=0; j<numEDCs; j++) {
						costPortEDC[i][j] = sheet1.getRow(i).getCell(j).getNumericCellValue();
					}
				}
				for(int i=0; i<numEDCs; i++) {
					fixedEDC[i] = sheet9.getRow(0).getCell(i).getNumericCellValue();
					edcCap[i] = sheet7.getRow(0).getCell(i).getNumericCellValue();
					for(int j=0; j<numLDCs; j++) {
						variableEDC[i][j] = sheet6.getRow(i).getCell(j).getNumericCellValue();
					}
				}
				workbook.close();

				try {
				//	for(int i=0; i<capacity.length; i++) {
						solveModel(costPortEDC, costPortLDC, fixedEDC, variableEDC, warehouseLDC, costEDCDemandRegion, costLDCArbitrary, demand, edcCap, warehouseEmissionsCosts);
				//	}
					

				} catch (IloException e) {
					System.out.println("A Cplex exception occured:" + e.getMessage());
					e.printStackTrace();
				}
			}
			
			public static void solveModel(double[][] costPortEDC, double[][] costPortLDC, double[] fixedEDC, double[][] variableEDC, double[] warehouseLDC, double[][] costEDCDemandRegion, double[] costLDCArbitrary,
					int[] demand, double[] edcCap, double[] warehouseEmissionsCosts) throws IloException {

				IloCplex cplex = new IloCplex();
				
				int numPorts = costPortEDC.length;
				int numEDCs = costPortEDC[0].length;
				int numLDCs = costPortLDC[0].length;
				int numDemandRegions = costEDCDemandRegion[0].length;
				
				int totDemand = 0;
				for(int k=0; k<numLDCs; k++) {
					totDemand += demand[k];
				}
				
				//Variable: if port i is in use
				//IloNumVar[] y1 = new IloNumVar[numPorts];
				//If EDC  is in use
				IloNumVar[] y1 = new IloNumVar[numEDCs]; 
				IloNumVar[] y2 = new IloNumVar[numLDCs];
				//if port op
				IloNumVar[] y3 = new IloNumVar[numPorts];
				
				
				
				//if EDC/LDC serves demand region
				IloNumVar[][] x1 = new IloNumVar[numEDCs][numDemandRegions];
				IloNumVar[][] x2 = new IloNumVar[numLDCs][numDemandRegions];
				//if port serves E/LDC
				IloNumVar[][] x3 = new IloNumVar[numPorts][numEDCs];
				IloNumVar[][] x4 = new IloNumVar[numPorts][numLDCs];
				
				
				//number containers Port to EDCs
				IloNumVar[][] z1 = new IloNumVar[numPorts][numEDCs];
				//number containers Port to LDCs
				IloNumVar[][] z2 = new IloNumVar[numPorts][numLDCs];
				//number of containers EDC to demand region
				IloNumVar[][] z3 = new IloNumVar[numEDCs][numDemandRegions];
				//number of containers LDC to arbitrary city
				IloNumVar[][] z4 = new IloNumVar[numLDCs][numDemandRegions];
				
				for(int i=0; i<numPorts; i++) {
					y3[i] = cplex.boolVar("y3(" + i + ")");
					for(int j=0; j<numEDCs; j++) {
						z1[i][j] = cplex.intVar(0, totDemand);
						x3[i][j] = cplex.boolVar();
					}
					for(int j=0; j<numLDCs; j++) {
						z2[i][j] = cplex.intVar(0, totDemand);
						x4[i][j] = cplex.boolVar();
						
					}
				}
				for(int j=0; j<numDemandRegions; j++) {
					for(int i=0; i<numEDCs; i++) {
						z3[i][j] = cplex.intVar(0, totDemand);
						x1[i][j] = cplex.boolVar();
					}
					for(int i=0; i<numLDCs; i++) {
						z4[i][j] = cplex.intVar(0, totDemand);
						x2[i][j] = cplex.boolVar();
					}
				}
				for(int i=0; i< numEDCs; i++) {
					y1[i] = cplex.boolVar("y1"
							+ "(" + i + ")");
				}
				for(int i=0; i< numLDCs; i++) {
					y2[i] = cplex.boolVar("y2"
							+ "(" + i + ")");
				}
				
				
				//Objective Function
				IloNumExpr objExpr = cplex.constant(0);
				// variable cost Port -> EDC/LDC
				for(int i=0; i<numPorts; i++) {
					for(int j=0; j<numEDCs; j++) {
						objExpr = cplex.sum(objExpr, cplex.prod(costPortEDC[i][j], z1[i][j]));
					}
					for(int j=0; j<numLDCs; j++) {
						objExpr = cplex.sum(objExpr, cplex.prod(costPortLDC[i][j], z2[i][j]));
					}	
				}
				
				//fixed cost opening an EDC
				for(int i=0; i<numEDCs; i++) {
						objExpr = cplex.sum(objExpr, cplex.prod(fixedEDC[i], y1[i]));
					
				}
				//warehouse cost an LDC
				for(int i=0; i<numLDCs; i++) {
					objExpr = cplex.sum(objExpr, cplex.prod(warehouseLDC[i], y2[i]));
				}
				//warehouse Emissions cost for LDC
				for(int i=0; i<numLDCs; i++) {
					for(int j=0; j<numDemandRegions; j++) {
						objExpr = cplex.sum(objExpr, cplex.prod(warehouseEmissionsCosts[j], x2[i][j]));
					}
				}
				
				//variable warehouse cost EDC + warehouse emissions Cost
				for(int i=0; i<numEDCs; i++) {
					for(int j=0; j<numDemandRegions; j++) {
						objExpr = cplex.sum(objExpr, cplex.prod(variableEDC[i][j], x1[i][j]), cplex.prod(warehouseEmissionsCosts[j], x1[i][j]));
					}
				}
				
				//variable cost EDC to Demand region
				for(int i=0; i<numEDCs; i++) {
					for(int j=0; j<numDemandRegions; j++) {
						objExpr = cplex.sum(objExpr, cplex.prod(costEDCDemandRegion[i][j], z3[i][j]));
					}
				}
				//variable cost LDC to arbitrary city
				for(int i=0; i<numLDCs; i++) {
					for(int j=0; j<numDemandRegions; j++) {
						objExpr = cplex.sum(objExpr, cplex.prod(costLDCArbitrary[i], z4[i][j]));
					}
				}
				
				
				cplex.addMinimize(objExpr);
				
				//constraint 1: Demand must be satisfied
				for(int i=0; i<numDemandRegions; i++) {
					IloNumExpr demandExpr = cplex.constant(0);
					for(int j=0; j<numEDCs; j++) {
						demandExpr = cplex.sum(demandExpr, z3[j][i]);
					}
					for(int j=0; j<numLDCs; j++) {
						demandExpr = cplex.sum(demandExpr, z4[j][i]);
					}
					cplex.addEq(demandExpr, demand[i]);
				}
				
				//constraint 2: Capacity of an EDC must be satisfied
				for(int j=0; j<numEDCs; j++) {
					IloNumExpr capExpr = cplex.constant(0);
					for(int i=0; i<numPorts; i++) {
						capExpr = cplex.sum(capExpr, z1[i][j]);
					}
					cplex.addLe(capExpr, edcCap[j]); 
				}
				for(int j=0; j<numLDCs; j++) {
					IloNumExpr capExpr2 = cplex.constant(0);
					for(int i=0; i<numPorts; i++) {
						capExpr2 = cplex.sum(capExpr2, z2[i][j]);
					}
					cplex.addLe(capExpr2, demand[j]); 
				}
				
				//constraint 3: containers through EDC only if its open
				for(int j=0; j<numEDCs; j++) {
					for(int i=0; i<numPorts; i++) {
						cplex.addLe(z1[i][j], cplex.prod(totDemand, y1[j]));
					}
				}
				for(int j=0; j<numEDCs; j++) {
					for(int k=0; k<numDemandRegions; k++) {
						cplex.addLe(z3[j][k], cplex.prod(totDemand, y1[j]));
					}
				}  
				
				//constraint 3: containers through LDC only if its open
				for(int j=0; j<numLDCs; j++) {
					for(int i=0; i<numPorts; i++) {
						cplex.addLe(z2[i][j], cplex.prod(totDemand, y2[j]));
					}
				}
				for(int j=0; j<numLDCs; j++) {
					for(int k=0; k<numDemandRegions; k++) {
						cplex.addLe(z4[j][k], cplex.prod(totDemand, y2[j]));
					}
				}  
				
				//constraint 4: all demand must leave the port
				IloNumExpr allDemand = cplex.constant(0);
				for(int i=0; i<numPorts; i++) {
					for(int j=0; j<numEDCs; j++) {
						allDemand = cplex.sum(allDemand, z1[i][j]);
					}
					for(int j=0; j<numLDCs; j++) {
						allDemand = cplex.sum(allDemand, z2[i][j]);
					}
				}
				
				cplex.addEq(allDemand, totDemand);
				
				//constraint 5: everything that comes into the (E)DC must come out
				for(int k=0; k<numEDCs; k++) {
					IloNumExpr inflow = cplex.constant(0);
					IloNumExpr outflow = cplex.constant(0);
					for(int i=0; i<numPorts; i++) {
						inflow = cplex.sum(inflow, z1[i][k]);
					}
					for(int j=0; j<numDemandRegions; j++) {
						outflow = cplex.sum(outflow, z3[k][j]);
					}
					cplex.addEq(inflow, outflow);
				}
				for(int k=0; k<numLDCs; k++) {
					IloNumExpr inflow = cplex.constant(0);
					IloNumExpr outflow = cplex.constant(0);
					for(int i=0; i<numPorts; i++) {
						inflow = cplex.sum(inflow, z2[i][k]);
					}
					for(int j=0; j<numDemandRegions; j++) {
						outflow = cplex.sum(outflow, z4[k][j]);
					}
					cplex.addEq(inflow, outflow);
				}
				
				//constraint 6: can only be served by either an EDC or LDC, not both
				for(int k=0; k<numEDCs; k++) {
					for(int i=0; i<numDemandRegions; i++) {
						IloNumExpr expr = cplex.prod(totDemand, x1[k][i]);
						cplex.addLe(z3[k][i], expr);
					}
				}
				for(int k=0; k<numLDCs; k++) {
					for(int i=0; i<numDemandRegions; i++) {
						cplex.addLe(z4[k][i], cplex.prod(totDemand, x2[k][i]));
					}
				}
				
				for(int j=0; j<numDemandRegions; j++) {
					IloNumExpr first = cplex.constant(0);
					IloNumExpr second = cplex.constant(0);
					for(int k=0; k<numEDCs; k++) {
						first = cplex.sum(first, x1[k][j]);
					}
					for(int k=0; k<numLDCs; k++) {
						second = cplex.sum(second, x2[k][j]);
					}
					cplex.addEq(cplex.sum(first, second), 1);
				}
					
				//an LDC can only serve 1 demand region
				for(int i=0; i<numLDCs; i++) {
					IloNumExpr ldc = cplex.constant(0);
					for(int j=0; j<numDemandRegions; j++) {
						ldc = cplex.sum(ldc, x2[i][j]);
					}
					cplex.addLe(ldc, 1);
				}
				
				//an LDC can only serve its local demand
				for(int i=0; i<numLDCs; i++) {
					for(int j=0; j<numDemandRegions; j++) {
						if(i != j) {
							cplex.addEq(x2[i][j], 0);
						}
					}	
				}
				//(E)DC can only have 1 port
				for(int j=0; j<numEDCs; j++) {
					IloNumExpr edc = cplex.constant(0);
					for(int k=0; k<numPorts; k++) {
						edc = cplex.sum(edc, x3[k][j]);
					}
					cplex.addLe(edc, 1);
				}
				for(int j=0; j<numLDCs; j++) {
					IloNumExpr ldc = cplex.constant(0);
					for(int k=0; k<numPorts; k++) {
						ldc = cplex.sum(ldc, x4[k][j]);
					}
					cplex.addLe(ldc, 1);
				}
				//containers only move through port-dc route if its being served
				for(int i=0; i<numPorts; i++) {
					for(int j=0; j<numEDCs; j++) {
						cplex.addLe(z1[i][j], cplex.prod(totDemand, x3[i][j]));
					}
				}
				for(int i=0; i<numPorts; i++) {
					for(int j=0; j<numLDCs; j++) {
						cplex.addLe(z2[i][j], cplex.prod(totDemand, x4[i][j]));
					}
				}
				//
				cplex.addEq(x4[0][11], 0);
				
				
				
				cplex.setOut(null);
				cplex.solve();

				// Query the solution
				// Print objective value and relevant variable values
				if (cplex.getStatus() == IloCplex.Status.Optimal) {
					
					System.out.println("Found optimal solution!");
					System.out.println("Objective = " + cplex.getObjValue());
					System.out.println("Solution: ");
					for(int i=0; i<numEDCs; i++) {
						if(cplex.getValue(y1[i]) >= 0.5) {
							System.out.println("EDC: " +i + " is used");
						}
					}
					for(int i=0; i<numLDCs; i++) {
						if(cplex.getValue(y2[i]) >= 0.5) {
							System.out.println("LDC: " +i + " is used");
						}
					}
					
					for(int i=0; i<numPorts; i++) {
						System.out.print("Port " + i + " serves EDCs: { ");
						for(int j=0; j<numEDCs; j++) {
							if (cplex.getValue(z1[i][j]) != 0) {
								System.out.print(j + " " /*+" with demand " + cplex.getValue(z1[i][j])+", "*/ );
							}
						}
						System.out.println("}");
					}
					
					for(int i=0; i<numPorts; i++) {
						System.out.print("Port " + i + " serves DCs: { ");
						for(int j=0; j<numLDCs; j++) {
							if (cplex.getValue(z2[i][j]) != 0) {
								System.out.print(j+" ");
							}
						}
						System.out.println("}");
					}
					
					for(int i=0; i<numEDCs; i++) {
						System.out.print("EDC " + i + " serves Demand Regions: { ");
						for(int j=0; j<numDemandRegions; j++) {
							if(cplex.getValue(z3[i][j]) > 1) {
								System.out.print(j + " " /*+" with demand " + cplex.getValue(z3[i][j])  + " demand required: " + demand[j]*/);
							}
						}
						System.out.println("}");
					}
					
					for(int i=0; i<numLDCs; i++) {
						for(int j=0; j<numDemandRegions; j++) {
							if(cplex.getValue(z4[i][j]) >= 0.5) {
								System.out.println("LDC " +i+" serves demand region "+j);
							}
						}
					}
				} 
				else {
					System.out.println("No optimal solution found");
				}
				// Close the model
				cplex.close();
				
			}

}
