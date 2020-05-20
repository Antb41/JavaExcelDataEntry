package JavaExcelDataEntry.main;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.*;
import org.apache.commons.collections4.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class Main {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		
		//Location where Excel files are stored for this project, can be any destination
		String excellFolder = "C://Users//Anthony//Desktop//";
		
		//Create an output file for data to be written to
		String filePath = excellFolder + "out_test.xlsx";

		try (FileOutputStream outputFile = new FileOutputStream(filePath)){
			
	        //Create new blank workbook
	        XSSFWorkbook workbook = new XSSFWorkbook();
	        
	        //Create new blank sheet in workbook 
	        Sheet sheet = workbook.createSheet("First sheet");
	        
	        //Create a new row within sheet, first row index will be 0
	        Row row = sheet.createRow(0);
	        
	        //Create cells within the rows on the new sheet
	        Cell cell = row.createCell(0);
	        cell.setCellValue("Column A");
	        
	        cell = row.createCell(1);
	        cell.setCellValue("Column B");
	        
	        cell = row.createCell(2);
	        cell.setCellValue("Column C");
	        
	        cell = row.createCell(3);
	        cell.setCellValue("Column D");
	        
	        cell = row.createCell(4);
	        cell.setCellValue("Column E");
	        
			//Save the workbook to the file system
			workbook.write(outputFile);
			workbook.close();
			System.out.println("Saved Excell file to: " + filePath);
			
		}
		
		catch(IOException ex) {
			System.out.println("The file could not be written: " + ex.getMessage());
		}
		
	}

}
