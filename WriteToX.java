import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class WriteToX {

	private static int columnNum = 3;

	public static void main(String[] args) throws IOException, InvalidFormatException {

		// Create a Workbook
		Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file
		Sheet bookSheet = workbook.createSheet("My Sheet");
			
		Row bookRow = bookSheet.createRow(0); // Create a Row
		Cell bookCell = bookRow.createCell(2);
		bookCell.setCellValue("Test Log");

		for (int j = 1; j<=columnNum; j++) {
			
			bookRow = bookSheet.createRow(j);
			bookCell = bookRow.createCell(2);
			
			//bookCell.setCellValue("Pass");
			
			bookCell.setCellValue(giveString());
			
		}
		
		for (int i = 0; i < columnNum; i++) {
			bookSheet.autoSizeColumn(i);
		}

		FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx"); // Write the output to a file
		workbook.write(fileOut);
		fileOut.close();

		workbook.close(); // Closing the workbook

	}
	
	public static String giveString() {
		
		String aString = "can be anything";
		
		return aString;
		
		
	}
}

/*
 * Write to an existing File
 * 
 * private static void modifyExistingWorkbook() throws InvalidFormatException,
 * IOException { // Obtain a workbook from the excel file Workbook workbook =
 * WorkbookFactory.create(new File("existing-spreadsheet.xlsx"));
 * 
 * // Get Sheet at index 0 Sheet sheet = workbook.getSheetAt(0);
 * 
 * // Get Row at index 1 Row row = sheet.getRow(1);
 * 
 * // Get the Cell at index 2 from the above row Cell cell = row.getCell(2);
 * 
 * // Create the cell if it doesn't exist if (cell == null) cell =
 * row.createCell(2);
 * 
 * // Update the cell's value cell.setCellType(CellType.STRING);
 * cell.setCellValue("Updated Value");
 * 
 * // Write the output to the file FileOutputStream fileOut = new
 * FileOutputStream("existing-spreadsheet.xlsx"); workbook.write(fileOut);
 * fileOut.close();
 * 
 * // Closing the workbook workbook.close(); }
 * 
 */
