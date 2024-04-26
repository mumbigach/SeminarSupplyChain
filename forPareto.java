import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

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
public class forPareto {

	public static void main(String[] args) throws IOException {
		//put the correct data
		FileInputStream file = new FileInputStream(new File("C:\\Users\\maken\\OneDrive - Erasmus University Rotterdam\\Documents\\UPT Master\\Seminar Supply Chain\\Data4.xlsx"));
		//FileInputStream file2 = new FileInputStream(new File("C:\\Users\\maken\\OneDrive - Erasmus University Rotterdam\\Documents\\UPT Master\\Seminar Supply Chain\\Data4Em.xlsx"));
		Workbook workbook = new XSSFWorkbook(file);
		// get the multiple sheets
		Sheet sheet1 = workbook.getSheet("PortEDCcosts");
		Sheet sheet2 = workbook.getSheet("PortDCosts");
		Sheet sheet3 = workbook.getSheet("EDCDemandCosts");
		Sheet sheet4 = workbook.getSheet("DCArbitraryCosts");
		Sheet sheet5 = workbook.getSheet("Demand");
		Sheet sheet6 = workbook.getSheet("variableEDC"); 
		Sheet sheet7 = workbook.getSheet("EDCcap");
		Sheet sheet8 = workbook.getSheet("WarehouseCostsLDCperContainer");
		Sheet sheet9 = workbook.getSheet("fixedEDC"); 
		
		Sheet sheet10 = workbook.getSheet("PortEDCEm");
		Sheet sheet11 = workbook.getSheet("PortDCEm");
		Sheet sheet12 = workbook.getSheet("EDCDemandEm");
		Sheet sheet13 = workbook.getSheet("DCArbitraryEm");
		Sheet sheet14 = workbook.getSheet("WarehouseEmissions");
		
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
		double[][] costPortLDC = new double[numPorts][numLDCs]; 
		double[][] emPortEDC = new double[numPorts][numEDCs]; 
		double[][] emPortLDC = new double[numPorts][numLDCs];
		
		//fixed cost of opening an (E/L)DC (warehouse costs)
		double[][] variableEDC = new double[numEDCs][numLDCs];
		double [] warehouseLDC = new double [numLDCs];
		double [] edcCap = new double[numEDCs];
		double [] fixedEDC = new double[numEDCs];
		
		
		//costs from EDC-> Demand Region 
		//predetermined mode of transport
		double[][] costEDCDemandRegion = new double [numEDCs][numLDCs];  
		//cost from LDC to Arbitrary city
		double[] costLDCArbitrary = new double [numLDCs]; 
		
		double[][] emEDCDemandRegion = new double [numEDCs][numLDCs];
		double[] emLDCArbitrary = new double [numLDCs]; 
		
		//Demanded number of containers
		int [] demand = new int [numLDCs];  
		
		//warehouse emissions
		double[] warehouseEmissions = new double[numLDCs];
		
		//update the sheets
		for (int j = 0; j < numLDCs; j++) {
				demand[j] = (int)sheet5.getRow(0).getCell(j).getNumericCellValue();
				costLDCArbitrary[j] = sheet4.getRow(0).getCell(j).getNumericCellValue();
				warehouseLDC[j] = sheet8.getRow(0).getCell(j).getNumericCellValue();
				emLDCArbitrary[j] = sheet13.getRow(0).getCell(j).getNumericCellValue();
				warehouseEmissions[j]  = sheet14.getRow(0).getCell(j).getNumericCellValue();
				
			for (int i = 0; i < numPorts; i++) {
				costPortLDC[i][j] = sheet2.getRow(i).getCell(j).getNumericCellValue();
				emPortLDC[i][j] = sheet11.getRow(i).getCell(j).getNumericCellValue();
			}
			for(int i=0; i<numEDCs; i++) {
				costEDCDemandRegion[i][j] = sheet3.getRow(i).getCell(j).getNumericCellValue();
				emEDCDemandRegion[i][j] = sheet12.getRow(i).getCell(j).getNumericCellValue();
			}
		}
		
		for(int i=0; i<numPorts; i++) {
			for(int j=0; j<numEDCs; j++) {
				costPortEDC[i][j] = sheet1.getRow(i).getCell(j).getNumericCellValue();
				emPortEDC[i][j] = sheet10.getRow(i).getCell(j).getNumericCellValue();
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
		
		double[] prices = new double[50];
		prices[0] = 20;
		for(int i=1; i<50; i++) {
			prices[i] = prices[i-1]+20;
			
		}
		for(int i=0; i<50; i++) {
			prices[i] = prices[i]/1000;
		}
		

		try {
			for(int i=0; i<prices.length; i++) {
				solveModel(costPortEDC, emPortEDC, costPortLDC, emPortLDC, fixedEDC, variableEDC, warehouseLDC, costEDCDemandRegion, emEDCDemandRegion, costLDCArbitrary, emLDCArbitrary, demand, edcCap, prices[i], warehouseEmissions);
			}
			

		} catch (IloException e) {
			System.out.println("A Cplex exception occured:" + e.getMessage());
			e.printStackTrace();
		}
	}
	
	public static void solveModel(double[][] costPortEDC, double[][] emPortEDC, double[][] costPortLDC, double[][] emPortLDC, double[] fixedEDC, double[][] variableEDC, double[] warehouseLDC, double[][] costEDCDemandRegion, double[][] emEDCDemandRegion, double[] costLDCArbitrary,double[] emLDCArbitrary,
			int[] demand, double[] edcCap, double price, double[] warehouseEmissions) throws IloException {

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
		IloNumVar[] y1 = new IloNumVar[numEDCs]; //check above if need to separate this because of constraint flow through EDC can only happen if its in use
		IloNumVar[] y2 = new IloNumVar[numLDCs];
		
		//if EDC/LDC serves demand region
		IloNumVar[][] x1 = new IloNumVar[numEDCs][numDemandRegions];
		IloNumVar[][] x2 = new IloNumVar[numLDCs][numDemandRegions];
		
		//number containers Port to EDCs
		IloNumVar[][] z1 = new IloNumVar[numPorts][numEDCs];
		//number containers Port to LDCs
		IloNumVar[][] z2 = new IloNumVar[numPorts][numLDCs];
		//number of containers EDC to demand region
		IloNumVar[][] z3 = new IloNumVar[numEDCs][numDemandRegions];
		//number of containers LDC to arbitrary city
		IloNumVar[][] z4 = new IloNumVar[numLDCs][numDemandRegions];
		
		for(int i=0; i<numPorts; i++) {
			//y1[i] = cplex.boolVar("y1(" + i + ")");
			for(int j=0; j<numEDCs; j++) {
				z1[i][j] = cplex.intVar(0, totDemand);
			}
			for(int j=0; j<numLDCs; j++) {
				z2[i][j] = cplex.intVar(0, totDemand);
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
		IloNumExpr objExpr2 = cplex.constant(0);
		// variable cost Port -> EDC/LDC
		for(int i=0; i<numPorts; i++) {
			for(int j=0; j<numEDCs; j++) {
				objExpr = cplex.sum(objExpr, cplex.prod(costPortEDC[i][j], z1[i][j]));
				objExpr2 = cplex.sum(objExpr2, cplex.prod(emPortEDC[i][j], z1[i][j]));
			}
			for(int j=0; j<numLDCs; j++) {
				objExpr = cplex.sum(objExpr, cplex.prod(costPortLDC[i][j], z2[i][j]));
				objExpr2 = cplex.sum(objExpr2, cplex.prod(emPortLDC[i][j], z2[i][j]));
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
		
		//warehouse Emissions for LDC
		for(int i=0; i<numLDCs; i++) {
			for(int j=0; j<numDemandRegions; j++) {
				objExpr2 = cplex.sum(objExpr2, cplex.prod(warehouseEmissions[j], x2[i][j]));
			}
		}
		//variable warehouse cost EDC + warehouse emissions
		for(int i=0; i<numEDCs; i++) {
			for(int j=0; j<numDemandRegions; j++) {
				objExpr = cplex.sum(objExpr, cplex.prod(variableEDC[i][j], x1[i][j]));
				objExpr2 = cplex.sum(objExpr2, cplex.prod(warehouseEmissions[j], x1[i][j]));
			}
		}
		
		//variable cost EDC to Demand region
		for(int i=0; i<numEDCs; i++) {
			for(int j=0; j<numDemandRegions; j++) {
				objExpr = cplex.sum(objExpr, cplex.prod(costEDCDemandRegion[i][j], z3[i][j]));
				objExpr2 = cplex.sum(objExpr2, cplex.prod(emEDCDemandRegion[i][j], z3[i][j]));
			}
		}
		//variable cost LDC to arbitrary city
		for(int i=0; i<numLDCs; i++) {
			for(int j=0; j<numDemandRegions; j++) {
				objExpr = cplex.sum(objExpr, cplex.prod(costLDCArbitrary[i], z4[i][j]));
				objExpr2 = cplex.sum(objExpr2, cplex.prod(emLDCArbitrary[i], z4[i][j]));
			}
		}
		
		cplex.addMinimize(cplex.sum(objExpr, cplex.prod(price, objExpr2)));
		
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
		
		
		
		cplex.setOut(null);
		cplex.solve();

		// Query the solution
		// Print objective value and relevant variable values
		if (cplex.getStatus() == IloCplex.Status.Optimal) {
			
		//	System.out.println("Price: "+price*1000);
		//	System.out.println("Total Objective = " + cplex.getObjValue());
			
			double totTranCosts = 0.0;
			double totEmCosts = 0.0;
			// variable cost Port -> EDC/LDC
			for(int i=0; i<numPorts; i++) {
				for(int j=0; j<numEDCs; j++) {
					if(cplex.getValue(z1[i][j]) > 0) {
						totTranCosts += costPortEDC[i][j]*cplex.getValue(z1[i][j]);
						totEmCosts += emPortEDC[i][j]*cplex.getValue(z1[i][j]);
					}
				}
				for(int j=0; j<numLDCs; j++) {
					if(cplex.getValue(z2[i][j]) > 0) {
						totTranCosts += costPortLDC[i][j]*cplex.getValue(z2[i][j]);
						totEmCosts += emPortLDC[i][j]*cplex.getValue(z2[i][j]);
					}
					
				}	
			}
			
			//variable cost EDC to Demand region
			for(int i=0; i<numEDCs; i++) {
				for(int j=0; j<numDemandRegions; j++) {
					if(cplex.getValue(z3[i][j]) > 0) {
						totTranCosts += costEDCDemandRegion[i][j]*cplex.getValue(z3[i][j]);
						totEmCosts += emEDCDemandRegion[i][j]*cplex.getValue(z3[i][j]);
					}	
				}
			}
			//variable cost LDC to arbitrary city
			for(int i=0; i<numLDCs; i++) {
				for(int j=0; j<numDemandRegions; j++) {
					if(cplex.getValue(z4[i][j]) > 0) {
						totTranCosts += costLDCArbitrary[i]*cplex.getValue(z4[i][j]);
						totEmCosts += emLDCArbitrary[i]*cplex.getValue(z4[i][j]);
					}
				}
			}
			
			//warehouse emissions
			for(int j=0; j<numDemandRegions; j++) {
				for(int i=0; i<numEDCs; i++) {
					if(cplex.getValue(x1[i][j]) > 0.5) {
						totEmCosts += warehouseEmissions[j];
					}
				}
				for(int i=0; i<numLDCs; i++) {
					if(cplex.getValue(x2[i][j]) > 0.5) {
						totEmCosts += warehouseEmissions[j];
					}
				}
			}
			
			System.out.println(/*"Total Transportation costs: " +*/ totTranCosts /*+"euros"*/);
	//		System.out.println(/*"Total Emissions are: " +*/ (totEmCosts/price)/1000 /*+"tCO2"*/);
	//		System.out.println();
			
		/*	System.out.println("Solution: ");
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
						System.out.print(j + " "  );
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
						System.out.print(j + " " );
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
			} */
		} 
		else {
			System.out.println("No optimal solution found");
		}
		// Close the model
		cplex.close();
		
	}

}
