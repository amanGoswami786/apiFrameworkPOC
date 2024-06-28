package jsonProject;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

public class ExcelToJsonDynamic {
	private static ObjectMapper objectMapper = new ObjectMapper();
	//static boolean isRootSheet = true;
	static String prevSheet = null;
	//static String rootSheetName = null;

	public static void main(String[] args) throws IOException {
		String templatePath = "C:\\Users\\2219691\\OneDrive - Cognizant\\Documents\\JsonTemplate.json";
		String excelPath = "C:\\Users\\2219691\\OneDrive - Cognizant\\Documents\\Book0.xlsx";
		FileInputStream excelFile = new FileInputStream(new File(excelPath));
		Workbook workbook = new XSSFWorkbook(excelFile);
		// Read the JSON template
		JsonNode jsonTemplate = objectMapper.readTree(new File(templatePath));
		// Convert the JSON template and Excel data to JSON
		JsonNode resultJson = processJsonNode(jsonTemplate, workbook, null, null);
		// Print the output JSON
		System.out.println(objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(resultJson));
	}

	private static JsonNode processJsonNode(JsonNode node, Workbook workbook, String currentSheetName,
			Map.Entry<Integer, String> currentTcid) throws IOException {

		// In template if its a node object
		if (node.isObject()) {
			ObjectNode objectNode = (ObjectNode) node;
			ObjectNode resultNode = objectMapper.createObjectNode();
			Iterator<Map.Entry<String, JsonNode>> fields = objectNode.fields();
			// Iterate the template
			while (fields.hasNext()) {
				Map.Entry<String, JsonNode> field = fields.next();
				String fieldName = field.getKey();
				JsonNode fieldValue = field.getValue();
				if (fieldValue.isTextual() && fieldValue.asText().startsWith("${")) {
					// Replace placeholder with actual value from Excel
					String placeholder = fieldValue.asText().substring(2, fieldValue.asText().length() - 1);
					prevSheet = currentSheetName;
					/*
					 * if(currentSheetName!= currentSheetName) { prevSheet = currentSheetName; }else
					 * { prevSheet = null; }
					 */
					String sheetName = currentSheetName != null ? currentSheetName : getSheetName(node);
					if (sheetName != null) {
						resultNode.put(fieldName, getCellValue(sheetName, currentTcid, placeholder, workbook));
					}
				} else {
					resultNode.set(fieldName, processJsonNode(fieldValue, workbook, fieldName, currentTcid));
				}
			}
			return resultNode;
		} else if (node.isArray()) {
			ArrayNode arrayNode = (ArrayNode) node;
			/*
			 * JsonNode json = original.findValue(currentSheetName); String name =
			 * json.traverse().getParsingContext().getParent().getCurrentName();
			 */
			ArrayNode resultArray = objectMapper.createArrayNode();
			if (currentSheetName != null) {
				Sheet sheet = prevSheet == null ? workbook.getSheet(currentSheetName) : workbook.getSheet(prevSheet);
				/*
				 * if(prevSheet == null) { rootSheetName = currentSheetName; }
				 */
				if (sheet != null) {
					/*
					 * if (rootSheet == null) { rootSheet = currentSheetName; }
					 */
					// if (!rootSheet.equalsIgnoreCase(currentSheetName)) {
					LinkedHashMap<Integer, String> Tcids = getUniqueTcids(sheet);
					for (Map.Entry<Integer, String> entry : Tcids.entrySet()) {
						if (prevSheet != null) {
							sheet = workbook.getSheet(currentSheetName);
							LinkedHashMap<Integer, String> childTcids = getUniqueTcids(sheet);
							// for(Map.Entry<Integer, String> childEntry : childTcids.entrySet()) {
							Map<Integer, String> filteredChildId = childTcids.entrySet().stream()
									.filter(map -> map.getValue().equals(entry.getValue()))
									.collect(Collectors.toMap(map -> map.getKey(), map -> map.getValue()));

							for (Map.Entry<Integer, String> childEntry : filteredChildId.entrySet()) {
								resultArray
										.add(processJsonNode(arrayNode.get(0), workbook, currentSheetName, childEntry));
							}
							// return resultArray;
						} else {
							resultArray.add(processJsonNode(arrayNode.get(0), workbook, currentSheetName, entry));
						}
						// return resultArray;
						/*
						 * if (currentSheetName == prevSheet) { currentSheetName = rootSheetName;
						 * prevSheet = null; }
						 */
					}
					/*
					 * } else { resultArray.add(processJsonNode(arrayNode.get(0), workbook,
					 * currentSheetName, null)); }
					 */

					/*
					 * else { isRootSheet = false; resultArray.add(processJsonNode(arrayNode.get(0),
					 * workbook, currentSheetName, null)); }
					 */

				}
			}
			return resultArray;
		}
		return node;

	}

	private static String getCellValue(String sheetName, Map.Entry<Integer, String> tcid, String columnName,
			Workbook workbook) {
		Sheet sheet = workbook.getSheet(sheetName);
		if (sheet != null) {
			for (Row row : sheet) {
				Cell cell = row.getCell(0);
				if (cell != null && row.getRowNum() != 0 && (tcid == null
						|| (tcid.getKey() == row.getRowNum()) && tcid.getValue().equals(cell.toString()))) {
					for (Cell c : row) {
						if (columnName.equals(sheet.getRow(0).getCell(c.getColumnIndex()).toString())) {
							return c.toString();
						}
					}
				}
			}
		}
		return "";
	}

	/*
	 * private static String getCellValue(String sheetName, Map.Entry<Integer,
	 * String> tcid, String columnName, Workbook workbook) { Sheet sheet =
	 * workbook.getSheet(sheetName); if (sheet != null) { for (Row row : sheet) {
	 * Cell cell = row.getCell(0); if (cell != null && row.getRowNum() != 0 && (tcid
	 * == null || (tcid.getKey() == row.getRowNum()) &&
	 * tcid.getValue().equals(cell.toString()))) { for (Cell c : row) { if
	 * (columnName.equals(sheet.getRow(0).getCell(c.getColumnIndex()).toString())) {
	 * return c.toString(); } } } } } return ""; }
	 */

	private static LinkedHashMap<Integer, String> getUniqueTcids(Sheet sheet) {
		LinkedHashMap<Integer, String> Tcids = new LinkedHashMap<Integer, String>();
		boolean isFirstRow = true;
		for (Row row : sheet) {
			if (!isFirstRow) {
				Cell cell = row.getCell(0);
				if (cell != null && !cell.toString().equals("")) {
					Tcids.put(row.getRowNum(), cell.toString());
				}
			}
			isFirstRow = false;
		}
		return Tcids;
	}

	private static String getSheetName(JsonNode node) {
		Iterator<String> fieldNames = node.fieldNames();
		if (fieldNames.hasNext()) {
			return fieldNames.next();
		}
		return null;
	}
}