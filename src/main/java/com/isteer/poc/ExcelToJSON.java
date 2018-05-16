package com.isteer.poc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import org.json.JSONObject;


public class ExcelToJSON
{
	private static final String CONFIG = "CONFIG_";
	private static final String CONF = "config";
	private static final String TEST_CASES = "testCases";
	private static final String SCENARIOS = "scenarios";
	private static final String EXCECUTIONS = "executions";
	private static final String TESTCASE_DATA = "testcaseData";
	private static final String EXCECUTIONS_SCENARIO_MAP = "executionsscenariomap";
	private static final String SCENARIOS_TESTCASE_MAP = "scenariosTestcasemap";
	private static final String FILE_PATH = "D://POC.json";
	private static final String CONFIG_SHEET = "config";
	private static final String TESTCASE1_SHEET = "testcases_1";
	private static final String SCENARIOS_SHEET = "scenarios";
	private static final String TESTCASE2_SHEET = "testcases_2";
	private static final String SUMMARY_SHEET = "summary";
	private static final String TESTCASEDATA_SHEET = "testcaseData";
	
	
	static Map<String, JSONObject> CONFIG_MAP = new HashMap<>();
	static Map<String, JSONObject> TESTCASE_MAP = new  HashMap<>();
	static Map<String, JSONObject> SCENARIO_MAP = new  HashMap<>();
	static Map<String, JSONObject> SUMMARY_MAP = new  HashMap<>();
	static Map<String, JSONObject> TESTCASE_DATA_MAP = new  HashMap<>();
	static Map<String, JSONArray> EXCECUTION_SCENARIO_MAP = new  HashMap<>();
	static Map<String, JSONArray> SCENARIO_TESTCASE_MAP = new  HashMap<>();

	
	public static final Logger LOGGER = Logger.getLogger(ExcelToJSON.class.getName());
	
	public static void main( String[] args ){
		try{ 
			File excelFile = new File("TMTEST.xls");
			FileInputStream excelInputStream = new FileInputStream(excelFile);
			Workbook excelWorkbook = null;

			excelWorkbook = new HSSFWorkbook(excelInputStream); // XSSFWorkbook(excelInputStream);
			Sheet config = excelWorkbook.getSheet(CONFIG_SHEET);
			Sheet sheetTestcase1 = excelWorkbook.getSheet(TESTCASE1_SHEET);
			Sheet scenarios = excelWorkbook.getSheet(SCENARIOS_SHEET);
			Sheet sheetTestCases2 = excelWorkbook.getSheet(TESTCASE2_SHEET);
			Sheet summary = excelWorkbook.getSheet(SUMMARY_SHEET);
			Sheet testCaseData = excelWorkbook.getSheet(TESTCASEDATA_SHEET);


			int length = 0;

			JSONObject jsonObjects = new JSONObject();

			jsonObjects.put(CONF, configMethod(length, config));
			jsonObjects.put(TEST_CASES, testCase1Method(length, sheetTestcase1));
			jsonObjects.put(SCENARIOS, moduleScenario(length, scenarios));
			jsonObjects.put(EXCECUTIONS, moduleSummary(length, summary));
			jsonObjects.put(TESTCASE_DATA, moduleTestCaseData(length, testCaseData, excelWorkbook));
			jsonObjects.put(EXCECUTIONS_SCENARIO_MAP, moduleExcecutionScenarios(length, scenarios));
			jsonObjects.put(SCENARIOS_TESTCASE_MAP, moduleScenariosTestcase(length, scenarios));
			LOGGER.info(jsonObjects.toString());

			File file = new File(FILE_PATH);
			FileWriter fileWriter = new FileWriter(file);
			fileWriter.write(jsonObjects.toString());
			LOGGER.info("JSON file created");
			fileWriter.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * 
	 * @param cellLength
	 * @param sheetTestcase1
	 * @return TESTCASE_MAP
	 */
	private static Map<String, JSONObject> testCase1Method(int cellLength, Sheet sheetTestcase1){
		//parsing test case 1 excel to json
		for(int i = 0; i <= sheetTestcase1.getLastRowNum();i++){
			ArrayList<String> testCase1Values = new  ArrayList<String>();
			Row row = sheetTestcase1.getRow(i);
			if(i == 0){
				cellLength = row.getLastCellNum();
			}	

			if(i>0){
				for(int j = 0; j< cellLength; j++ ){
					if (row.getCell(j) != null) {
						if(row.getCell(j).getCellType() == Cell.CELL_TYPE_BOOLEAN){
							testCase1Values.add(String.valueOf(row.getCell(j).getBooleanCellValue()));
						}else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
							testCase1Values.add(String.valueOf(row.getCell(j).getNumericCellValue()));
						}
						else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_STRING){
							testCase1Values.add(String.valueOf(row.getCell(j).getStringCellValue()));
						}
						
					}
					else{
						testCase1Values.add("");
					}

				}

				JSONObject objects = new JSONObject();
				Row testCase1row = sheetTestcase1.getRow(0); 
				objects.put(getRowValue(testCase1row, 2), testCase1Values.get(2));
				objects.put(getRowValue(testCase1row, 3), testCase1Values.get(3));
				objects.put(getRowValue(testCase1row, 4), testCase1Values.get(4));
				objects.put(getRowValue(testCase1row, 5), testCase1Values.get(5));
				objects.put(getRowValue(testCase1row, 6), testCase1Values.get(6));
				objects.put(getRowValue(testCase1row, 7), testCase1Values.get(7));
				objects.put(getRowValue(testCase1row, 8), testCase1Values.get(8));
				objects.put(getRowValue(testCase1row, 9), testCase1Values.get(9));

				JSONObject testCase1 = new JSONObject();
				testCase1.put("objects", objects);
				testCase1.put(testCase1row.getCell(0).getStringCellValue(),testCase1Values.get(0));
				testCase1.put(testCase1row.getCell(1).getStringCellValue(), testCase1Values.get(1));

				TESTCASE_MAP.put("testcases_1_"+testCase1Values.get(0), testCase1);

			}

		}
		return TESTCASE_MAP;
	}

	/**
	 * 
	 * @param cellLength
	 * @param config
	 * @return
	 */
	public static Map<String, JSONObject> configMethod(int cellLength, Sheet config){
		
		for(int i = 0; i <= config.getLastRowNum();i++){
			ArrayList<String> cellValues = new  ArrayList<String>();
			Row row = config.getRow(i);
			if(i == 0){
				cellLength = row.getLastCellNum();
			}

			if(i>0){
				for(int j = 0; j< cellLength; j++ ){
					if (row.getCell(j) != null) {
						if(row.getCell(j).getCellType() == Cell.CELL_TYPE_BOOLEAN){
							cellValues.add(String.valueOf(row.getCell(j).getBooleanCellValue()));
						}else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
							cellValues.add(String.valueOf(row.getCell(j).getNumericCellValue()));
						}
						else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_STRING){
							cellValues.add(String.valueOf(row.getCell(j).getStringCellValue()));
						}
					}
					else{
						cellValues.add("");
					}

				}

				JSONObject services=new JSONObject();

				//writing inner objects
				JSONObject devices = new JSONObject();
				Row configRows = config.getRow(0);
				devices.put(getRowValue(configRows, 0),CONFIG+cellValues.get(0));
				devices.put(getRowValue(configRows, 8),cellValues.get(8));
				devices.put(getRowValue(configRows, 9),cellValues.get(9));
				devices.put(getRowValue(configRows, 10),cellValues.get(10));
				devices.put(getRowValue(configRows, 3),cellValues.get(3));
				devices.put(getRowValue(configRows, 4),cellValues.get(4));
				devices.put(getRowValue(configRows, 5),cellValues.get(5));

				JSONObject configObj = new JSONObject();
				configObj.put(getRowValue(configRows, 1), cellValues.get(1));
				configObj.put(getRowValue(configRows, 0), CONFIG+cellValues.get(0));
				configObj.put("services", services);
				configObj.put(configRows.getCell(7).getStringCellValue(), cellValues.get(7));
				configObj.put("devices",devices);

				CONFIG_MAP.put(CONFIG + i, configObj);

			}
		}
		LOGGER.info("Config : "+CONFIG_MAP.toString());
		return CONFIG_MAP;
	}

	/**
	 * 
	 * @param cellLength
	 * @param scenarios
	 * @return
	 */
	public static Map<String, JSONObject> moduleScenario(int cellLength, Sheet scenarios){
	
		// Parsing from scenario excel to json
		for(int i = 0; i <= scenarios.getLastRowNum();i++){
			ArrayList<String> scenariosList = new  ArrayList<String>();
			Row row = scenarios.getRow(i);
			if(i == 0){
				cellLength = row.getLastCellNum();
			}

			if(i>0){
				for(int j = 0; j< cellLength; j++ ){
					if (row.getCell(j) != null) {
						if(row.getCell(j).getCellType() == Cell.CELL_TYPE_BOOLEAN){
							scenariosList.add(String.valueOf(row.getCell(j).getBooleanCellValue()));
						}else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
							scenariosList.add(String.valueOf(row.getCell(j).getNumericCellValue()));
						}
						else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_STRING){
							scenariosList.add(String.valueOf(row.getCell(j).getStringCellValue()));
						}
					}
					else{
						scenariosList.add("");
					}

				}

				JSONObject scenariosColumns = new JSONObject();
				Row scenarioRow = scenarios.getRow(0);
				scenariosColumns.put(getRowValue(scenarioRow, 0), scenariosList.get(0));
				scenariosColumns.put(getRowValue(scenarioRow, 1), scenariosList.get(1));
				scenariosColumns.put(getRowValue(scenarioRow, 2), scenariosList.get(2));
				scenariosColumns.put(getRowValue(scenarioRow, 3), scenariosList.get(3));

				SCENARIO_MAP.put(scenariosList.get(3), scenariosColumns);
				
			}

		}
		LOGGER.info("Scenarios :"+SCENARIO_MAP.toString());
		return SCENARIO_MAP;
	}

	/**
	 * 
	 * @param len
	 * @param summary
	 * @return
	 */
	private static Map<String, JSONObject> moduleSummary(int len,Sheet summary){
		//parsing summary table to json //executions
		for(int i = 0; i <= summary.getLastRowNum();i++){
			ArrayList<String> summaryValues = new  ArrayList<String>();
			Row row = summary.getRow(i);
			if(i == 0){
				len = row.getLastCellNum();
			}

			if(i>0){
				for(int j = 0; j< len; j++ ){
					if (row.getCell(j) != null) {
						if(row.getCell(j).getCellType() == Cell.CELL_TYPE_BOOLEAN){
							summaryValues.add(String.valueOf(row.getCell(j).getBooleanCellValue()));
						}else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
							summaryValues.add(String.valueOf(row.getCell(j).getNumericCellValue()));
						}
						else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_STRING){
							summaryValues.add(String.valueOf(row.getCell(j).getStringCellValue()));
						}
					}
					else{
						summaryValues.add("");
					}

				}

				JSONObject suiteValue = new JSONObject();
				Row summaryRows = summary.getRow(0);
				suiteValue.put(getRowValue(summaryRows, 0), summaryValues.get(0));
				suiteValue.put(getRowValue(summaryRows, 1), summaryValues.get(1));
				suiteValue.put(getRowValue(summaryRows, 2), summaryValues.get(2));
				suiteValue.put(getRowValue(summaryRows, 3), summaryValues.get(3));
				suiteValue.put(getRowValue(summaryRows, 4), summaryValues.get(4));
				suiteValue.put(getRowValue(summaryRows, 5), summaryValues.get(5));

				SUMMARY_MAP.put("suite_value", suiteValue);

			}

		}
		LOGGER.info("Summary : "+SUMMARY_MAP.toString());
		return SUMMARY_MAP;
	}

	/**
	 * 
	 * @param len
	 * @param testCaseData
	 * @param excelWorkbook
	 * @return
	 */
	private static Map<String, JSONObject> moduleTestCaseData(int len, Sheet testCaseData, Workbook excelWorkbook){
		//parsing testcaseData table to json
		for(int i = 0; i <= testCaseData.getLastRowNum();i++){
			ArrayList<String> testCaseValues = new  ArrayList<String>();
			Row row = testCaseData.getRow(i);
			if(i == 0){
				len = row.getLastCellNum();
			}

			if(i>0){
				for(int j = 0; j< len; j++ ){
					if (row.getCell(j) != null) {
						if(row.getCell(j).getNumericCellValue() == Cell.CELL_TYPE_BOOLEAN){
							testCaseValues.add(String.valueOf(row.getCell(j).getBooleanCellValue()));
						}else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
							testCaseValues.add(String.valueOf(row.getCell(j).getNumericCellValue()));
						}
						else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_STRING){
							testCaseValues.add(String.valueOf(row.getCell(j).getStringCellValue()));
						}
					}
					else{
						testCaseValues.add("");
					}

				}

				JSONObject suiteValue = new JSONObject();
				Row testCaseRows = testCaseData.getRow(0);
				suiteValue.put(getRowValue(testCaseRows, 0), testCaseValues.get(0)+"."+testCaseValues.get(1));
				suiteValue.put(getRowValue(testCaseRows, 1), (testCaseValues.get(0)+"."+testCaseValues.get(1)).toUpperCase());
				suiteValue.put(getRowValue(testCaseRows, 2), testCaseValues.get(2));
				suiteValue.put(getRowValue(testCaseRows, 4), excelWorkbook.getSheetName(4));
				suiteValue.put("index", testCaseData.getPhysicalNumberOfRows());

				TESTCASE_DATA_MAP.put("keyValue", suiteValue);
			}

		}
		LOGGER.info("TESTCASE_DATA_MAP : "+TESTCASE_DATA_MAP.toString());
		return TESTCASE_DATA_MAP;
	}

	/**
	 * 
	 * @param len
	 * @param scenarios
	 * @return
	 */
	private static Map<String, JSONArray> moduleExcecutionScenarios(int len, Sheet scenarios){
		//parsing excecutionScenarioMap table to json
		for(int i = 0; i <= scenarios.getLastRowNum();i++){
			ArrayList<String> scenarioValues = new  ArrayList<String>();
			Row row = scenarios.getRow(i);
			if(i == 0){
				len = row.getLastCellNum();
			}

			if(i>0){
				for(int j = 0; j< len; j++ ){
					if (row.getCell(j) != null) {
						if(row.getCell(j).getCellType() == Cell.CELL_TYPE_BOOLEAN){
							scenarioValues.add(String.valueOf(row.getCell(j).getBooleanCellValue()));
						}else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
							scenarioValues.add(String.valueOf(row.getCell(j).getNumericCellValue()));
						}
						else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_STRING){
							scenarioValues.add(String.valueOf(row.getCell(j).getStringCellValue()));
						}
					}
					else{
						scenarioValues.add("");
					}

				}

				JSONObject scenariosColumns = new JSONObject();
				Row scenarioRow = scenarios.getRow(0);
				scenariosColumns.put(getRowValue(scenarioRow, 0), scenarioValues.get(0));
				scenariosColumns.put(getRowValue(scenarioRow, 1), scenarioValues.get(1));
				scenariosColumns.put(getRowValue(scenarioRow, 2), scenarioValues.get(2));
				scenariosColumns.put(getRowValue(scenarioRow, 3), scenarioValues.get(3));

				JSONArray scenariosValue = new JSONArray();
				scenariosValue.put(scenarioValues.get(1));

				JSONObject sce = new JSONObject();
				sce.put(scenarioValues.get(3), scenariosValue);

				EXCECUTION_SCENARIO_MAP.put(scenarioValues.get(3), scenariosValue);
			}
		}
		LOGGER.info("EXCECUTION_SCENARIO_MAP : "+ EXCECUTION_SCENARIO_MAP.toString());
		return EXCECUTION_SCENARIO_MAP;
	}

	/**
	 * 
	 * @param len
	 * @param scenarios
	 * @return
	 */
	private static Map<String, JSONArray> moduleScenariosTestcase(int len, Sheet scenarios){
		//parsing scenariosTestcasemap table to json
		for(int i = 0; i <= scenarios.getLastRowNum();i++){
			ArrayList<String> scenarioValues = new  ArrayList<String>();
			Row row = scenarios.getRow(i);
			if(i == 0){
				len = row.getLastCellNum();
			}

			if(i>0){
				for(int j = 0; j< len; j++ ){
					if (row.getCell(j) != null) {
						if(row.getCell(j).getCellType() == Cell.CELL_TYPE_BOOLEAN){
							scenarioValues.add(String.valueOf(row.getCell(j).getBooleanCellValue()));
						}else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
							scenarioValues.add(String.valueOf(row.getCell(j).getNumericCellValue()));
						}
						else if(row.getCell(j).getCellType() == Cell.CELL_TYPE_STRING){
							scenarioValues.add(String.valueOf(row.getCell(j).getStringCellValue()));
						}
					}
					else{
						scenarioValues.add("");
					}

				}

				JSONObject scenariosColumns = new JSONObject();
				Row scenarioRow = scenarios.getRow(0);
				scenariosColumns.put(getRowValue(scenarioRow, 0), scenarioValues.get(0));
				scenariosColumns.put(getRowValue(scenarioRow, 1), scenarioValues.get(1));
				scenariosColumns.put(getRowValue(scenarioRow, 2), scenarioValues.get(2));
				scenariosColumns.put(getRowValue(scenarioRow, 3), scenarioValues.get(3));

				JSONArray testcases = new JSONArray();
				testcases.put(scenarioValues.get(3));


				SCENARIO_TESTCASE_MAP.put("SCENARIOS_"+scenarioValues.get(1), testcases);

			}

		}
		LOGGER.info("SCENARIO_TESTCASE_MAP : "+SCENARIO_TESTCASE_MAP.toString());
		return SCENARIO_TESTCASE_MAP;
	}

	/**
	 * 
	 * @param row
	 * @param index
	 * @return
	 */
	private static String getRowValue(Row row, int index) {
		return  row.getCell(index).getStringCellValue();
	}




}
