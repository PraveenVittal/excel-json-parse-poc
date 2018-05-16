package com.isteer.poc;


import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.jackson.JsonParseException;
import org.codehaus.jackson.map.JsonMappingException;
import org.codehaus.jackson.map.ObjectMapper;

import com.google.gson.Gson;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;

public class ReadJsonToExcel {

	private static final String FILE_WRITE_PATH = "JsonToExcel.xlsx";
	private static final String FILE_READ_PATH = "D:\\POC.json";
	public static final Logger LOGGER = Logger.getLogger(ReadJsonToExcel.class.getName());

	public static void main(String[] args) throws Exception {

		//Create blank workbook
		Workbook workbook = new XSSFWorkbook(); 

		//Create a blank sheet
		Sheet spreadsheet = workbook.createSheet("summary");
		Sheet scenarios = workbook.createSheet("scenarios");

		Gson gson = new Gson();

		JsonElement json = gson.fromJson(new FileReader(FILE_READ_PATH), JsonElement.class);

		summary(spreadsheet, json);
		scenario(scenarios, json);

		//Write the workbook in file system
		FileOutputStream out = new FileOutputStream(new File(FILE_WRITE_PATH));
		workbook.write(out);
		out.close();
		LOGGER.info("File written successfully");
	}

	/**
	 * 
	 * @param spreadsheet
	 * @param json
	 * @throws IOException
	 * @throws JsonParseException
	 * @throws JsonMappingException
	 */
	@SuppressWarnings("unchecked")

	private static void summary(Sheet spreadsheet, JsonElement json)
			throws IOException, JsonParseException, JsonMappingException {

		JsonObject jsonObject = json.getAsJsonObject(); 
		JsonElement suite = jsonObject.get("executions");
		JsonObject suiteElement = suite.getAsJsonObject();

		Map<String,String> suiteMap = new HashMap<String, String>();

		ObjectMapper objectMapper = new ObjectMapper();
		suiteMap = objectMapper.readValue(suiteElement.get("suite_value").toString(), HashMap.class);

		LOGGER.info(suiteMap.toString());

		int cellid = 0;
		Row row0 = spreadsheet.createRow(0);
		Row row1 = spreadsheet.createRow(1);

		for(Entry<String, String> maps : suiteMap.entrySet()) {
			LOGGER.info(maps.getKey() +"  "+ maps.getValue());
			int cellId = cellid++;

			Cell cell = row0.createCell(cellId);
			cell.setCellValue(maps.getKey());

			Cell cell1 = row1.createCell(cellId);
			cell1.setCellValue(maps.getValue());
		}
	}
	/**
	 * 
	 * @param scenarios
	 * @param json
	 * @throws IOException
	 * @throws JsonParseException
	 * @throws JsonMappingException
	 */
	@SuppressWarnings("unchecked")
	private static void scenario(Sheet scenarios, JsonElement json)
			throws IOException, JsonParseException, JsonMappingException {

		JsonObject jsonObject = json.getAsJsonObject();
		JsonElement suiteScenarios = jsonObject.get("scenarios");
		JsonObject suiteScenarios1 = suiteScenarios.getAsJsonObject();


		Map<String,String> tc3Map = new HashMap<String, String>();
		Map<String,String> suiteValMap = new HashMap<String, String>();

		ObjectMapper objectMapper1 = new ObjectMapper();
		ObjectMapper objectMapper2 = new ObjectMapper();
		tc3Map = objectMapper1.readValue(suiteScenarios1.get("TC3").toString(), HashMap.class);
		suiteValMap = objectMapper2.readValue(suiteScenarios1.get("TC1,TC2").toString(), HashMap.class);


		int cellid1 = 0;
		int cellid2 = 0;

		Row headerRow = scenarios.createRow(0);
		Row rowS1 = scenarios.createRow(1);
		Row rowS2 = scenarios.createRow(2);

		for(Entry<String, String> maps : tc3Map.entrySet()) {
			int cellId = cellid1++;

			Cell cell = headerRow.createCell(cellId);
			cell.setCellValue(maps.getKey());

			Cell cell1 = rowS1.createCell(cellId);
			cell1.setCellValue(maps.getValue());
		}

		for(Entry<String, String> maps : suiteValMap.entrySet()) {
			int cellId = cellid2++;
			Cell cell2 = rowS2.createCell(cellId);
			cell2.setCellValue(maps.getValue());
		}

	}
}