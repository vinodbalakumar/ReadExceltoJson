package com.kiran.ExceltoJson;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

/**
 * 
 * @author Vinod.Panchaparvala
 * 
 *   This class is used to read excel data and convert excel data into  json.
 *        
 *
 *
 */

@RestController
@RequestMapping(value = "/")
public class ExcelToJsonController {
	/**
	 * This method is used to get Excel data.
	 * 
	 * @return ran method.
	 * @throws JSONException
	 * @throws IOException
	 */
	@RequestMapping(value = "convert", method = RequestMethod.GET)
	public String Excel_() throws JSONException, IOException {

		return ran();
	}
	
	/**
	 * This method is used to get Excel data and convert it into Json format.
	 * @return sucess or not.
	 * @throws IOException
	 * @throws JSONException
	 */
	public static String ran() throws IOException, JSONException {
		String excelFilePath = "D://ANMOL-2.0_API(1).xlsx";
		FileInputStream inp = new FileInputStream(excelFilePath);
		Workbook workbook = new XSSFWorkbook(inp);
		System.out.println(workbook);
		Sheet firstSheet = workbook.getSheetAt(0);

		// Start constructing JSON.
		JSONObject json = new JSONObject();

		// Iterate through the rows.
		JSONArray rows = new JSONArray();
		for (Iterator<Row> iterator = firstSheet.rowIterator(); iterator.hasNext();) {
			Row row = iterator.next();
			JSONObject jRow = new JSONObject();

			// Iterate through the cells.
			JSONArray cells = new JSONArray();
			for (Iterator<Cell> cellIterator = row.cellIterator(); cellIterator.hasNext();) {
				Cell cell = cellIterator.next();
				cells.put(cell.getStringCellValue());

			}
			jRow.put("cell", cells);
			rows.put(jRow);

		}

		// Create the JSON.
		json.put("rows", rows);
		String Response = ConvertIntoFile(json);

		// Get the JSON text.
		System.out.println(Response);
		return "sucess\n" /*+ json*/;

	}
/**
 * This method is used get Json data from above method and keep it in one text file
 * @param json
 * @return success or not.
 * @throws IOException
 */
	private static String ConvertIntoFile(JSONObject json) throws IOException {

		try (FileWriter file = new FileWriter("C:\\Users\\vinod.panchaparvala\\Documents\\file1.txt")) {
			file.write(json.toString());
			System.out.println("Successfully Copied JSON Object to File...");
			System.out.println("\nJSON Object: " + json);

			return "Sucess";
		}

	}
}
