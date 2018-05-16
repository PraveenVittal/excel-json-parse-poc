package com.isteer.poc;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;

import javax.swing.LookAndFeel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONObject;
import org.junit.Assert;
import org.junit.Test;


public class ExcelToJsonTestCase
{
	static final Logger LOGGER = Logger.getLogger(ExcelToJsonTestCase.class.getName());
	
	ExcelToJSON exceltoJson;

	@Test
	public void configMethod() throws Exception{
		File excelFile = new File("TMTEST.xls");
		FileInputStream excelInputStream = new FileInputStream(excelFile);
		Workbook excelWorkbook = null;

		excelWorkbook = new HSSFWorkbook(excelInputStream); // XSSFWorkbook(excelInputStream);
		Sheet config = excelWorkbook.getSheet("config");
		Map<String, JSONObject> configMap = new HashMap<String, JSONObject>();
		configMap = exceltoJson.configMethod(0, config);
		LOGGER.info(configMap.toString());
		Assert.assertNotNull(configMap);
	}
	@Test
	public void moduleScenario() throws Exception{
		File excelFile = new File("TMTEST.xls");
		FileInputStream excelInputStream = new FileInputStream(excelFile);
		Workbook excelWorkbook = null;

		excelWorkbook = new HSSFWorkbook(excelInputStream); // XSSFWorkbook(excelInputStream);
		Sheet scenario = excelWorkbook.getSheet("scenarios");
		Map<String, JSONObject> configMap = new HashMap<String, JSONObject>();
		configMap = exceltoJson.moduleScenario(0, scenario);
		LOGGER.info(configMap.toString());
		Assert.assertNotNull(configMap);
	}
	
}
