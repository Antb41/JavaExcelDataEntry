package JavaExcelDataEntry.main;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.*;
import org.apache.commons.collections4.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Scanner;

import javax.swing.text.html.HTMLDocument.Iterator;

public class Main {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		
		//Location where Excel files are stored for this project, can be any destination
		String excellFolder = "C://Users//Anthony//Desktop//";
		
		//Create an output file for data to be written to
		String filePath = excellFolder + "default.xlsx";
		
		//Check if workbook needs to be updated or created
		Scanner update = new Scanner(System.in);
		String updateCreateAnswer = "";
		System.out.println("Would you like to create a new workbook or update the existing workbook"
				+ " at " + filePath + "\n" + "1. Update" + "\n" + "2. Create");
		System.out.print("Please enter (1) or (2) and hit ENTER: ");
		updateCreateAnswer = update.nextLine();
		System.out.println();
		
        //Create new blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();				
        
        //Create new blank sheet in workbook 
        Sheet sheet = workbook.createSheet("temp_sheet");
		
		if(updateCreateAnswer.equals("2")){
			
			//create
			Scanner newFolder = new Scanner(System.in);
			String newFolderPath = "";
			System.out.print("Please enter the folder locaton where you would like you excell doc to be created: ");
			newFolderPath = newFolder.nextLine();
			System.out.println();
			filePath = excellFolder;
			Scanner newDoc = new Scanner(System.in);
			String newDocName = "";
			System.out.print("Please enter the name of your excell doc (Do not include .xlsx): ");
			newDocName = newDoc.nextLine();
			newDocName = newDocName + ".xlsx";
			System.out.println();
			
			//Concatenate file path
			filePath = excellFolder + newDocName;
			
			//give temp sheet a customizable name
	        String sheetName = "";
	        Scanner sheetNameInput = new Scanner(System.in);
	        System.out.print("Please enter the name of a new sheet for your workbook: ");
	        sheetName = sheetNameInput.nextLine();
	        System.out.println();
	        workbook.setSheetName(workbook.getSheetIndex(sheet), sheetName);
	        		
			try (FileOutputStream outputFile = new FileOutputStream(filePath)){
		        
				//input for allowing the user to stop creating headers
				String answer = "";
				Scanner answerStop = new Scanner(System.in);
				
				//input for allowing user to customize header names
				String headerNames = "";
				Scanner createHeaderNames = new Scanner(System.in);
				int z = 0;
				int headerRows = 0;
				int dataRows = 0;
				int numberOfSheets = 1;
				//Create new sheets if needed 
				int w = 0;
				for(w = 0; w <= numberOfSheets; w++) {
			        //Create a new row within sheet, first row index will be 0
			        Row header = sheet.createRow(0);
			        for(z = 0; z <= headerRows; z++) {
			        	
				        //Create header cells within the rows on the new sheet
				        Cell cell_header = header.createCell(z);
				        System.out.print("Please enter the name of a new header: ");
				        headerNames = createHeaderNames.nextLine();
				        System.out.println();
				        cell_header.setCellValue(headerNames);
				        System.out.print("Would you like to create another header? (y/n): ");
				        answer = answerStop.nextLine();
				        System.out.println();
				        headerRows++;
				        
				        //stop creation of new headers (makes it so z wont be less than headerRows)
				        if(answer.equals("n")) {
				        	dataRows = headerRows;
				        	headerRows = 0;
				        }
			        }
			        
			        //Create new rows and cells and enter data into cells
				    String dataValues = "";
			        Scanner input = new Scanner(System.in);
			        int numberOfDataRows = 0;
			        int k = 0;
			        int x = 0;
			        for(k = 0; k <= numberOfDataRows; k++) {
				        Row rows = sheet.createRow(k + 1);
				        for(x = 0; x <= (dataRows - 1); x++) {
					        Cell cells = rows.createCell(x);
					        System.out.print("Please enter a/an " 
					        + header.getCell(x).getStringCellValue() + ": ");
					        dataValues = input.nextLine();
					        System.out.println();
					        cells.setCellValue(dataValues);
				        }
			        }
			        
			        String sheetCreationAnswer = "";
			        Scanner sheetCreationAnswerInput = new Scanner(System.in);
					System.out.print("Would you like to create another sheet? (y/n): ");
					sheetCreationAnswer = sheetCreationAnswerInput.nextLine();
					System.out.println();
					if(sheetCreationAnswer.equals("y")) {
						String newSheetName = "";
						Scanner newSheetNameInput = new Scanner(System.in);
				        System.out.print("Please enter the name of a new sheet for your workbook: ");
				        newSheetName = newSheetNameInput.nextLine();
						System.out.println();
						sheet = workbook.createSheet(newSheetName);
						numberOfSheets++;
					}
				}
		        
				//Save the workbook to the file system
				workbook.write(outputFile);
				workbook.close();
				System.out.println("Saved Excell file to: " + filePath);
			}
			
			catch(IOException ex) {
				System.out.println("The file could not be written: " + ex.getMessage());
			}
		
		}
		
		if(updateCreateAnswer.equals("1")){
			//update			

			Scanner newFolder = new Scanner(System.in);
			String newFolderPath = "";
			System.out.print("Please enter the folder locaton of the excell file you would like to update: ");
			newFolderPath = newFolder.nextLine();
			System.out.println();
			filePath = excellFolder;
			Scanner newDoc = new Scanner(System.in);
			String newDocName = "";
			System.out.print("Please enter the name of the excell doc you would like to update (Do not include .xlsx): ");
			newDocName = newDoc.nextLine();
			newDocName = newDocName + ".xlsx";
			System.out.println();
			
			filePath = excellFolder + newDocName;
			
			//Get xlxs file that has already been created in specified file path
			File readFile = new File(filePath);
			FileInputStream inputStream = new FileInputStream(readFile);
			XSSFWorkbook readWorkbook = new XSSFWorkbook(inputStream); 
				
			try (FileOutputStream outputFile = new FileOutputStream(new File(filePath))){
				
				//Get specific sheet in workbook
				int numberOfSheets = readWorkbook.getNumberOfSheets();
				int sheetSelection = 0;
				Scanner sheetSelectionInput = new Scanner(System.in);
				int i = 0;
				Sheet readSheet = null;
				for(i = 0; i <= (numberOfSheets - 1); i++) {
					System.out.println((i + 1) + ". " + readWorkbook.getSheetAt(i).getSheetName());
					if(i == (numberOfSheets - 1)) {
						System.out.print("Please select with sheet you would like to update: ");
						sheetSelection = sheetSelectionInput.nextInt();
						System.out.println();
						readSheet = readWorkbook.getSheetAt((sheetSelection - 1));
					}
				}
				
				//Update cell within selected Excell doc
				int numberOfRows = readSheet.getRow(0).getPhysicalNumberOfCells();
				Scanner cellSelectionInput = new Scanner(System.in);
				int cellSelection = 0;
				int j = 0;
				Cell cellUpdate;
				for(j = 0; j <= (numberOfRows - 1); j++) {
					System.out.println((j + 1) + ". " + readSheet.getRow(0).getCell(j) + " (Current: "
							+ readSheet.getRow(1).getCell(j) + ")");
					if(j == (numberOfRows - 1)) {
						System.out.print("Please select which column you would like to update: ");
						cellSelection = cellSelectionInput.nextInt();
						System.out.println();
						cellUpdate = readSheet.getRow(1).getCell((cellSelection - 1));
						String updateCell = "";
				        System.out.print("Please Enter A New " + readSheet.getRow(0).getCell((cellSelection - 1)) 
				        		+ ": ");
						Scanner updateCellInput = new Scanner(System.in);
				        updateCell = updateCellInput.nextLine();
						cellUpdate.setCellValue(updateCell);
					}
				}
				
				inputStream.close();	
				
				//Save the workbook to the file system
				readWorkbook.write(outputFile);
				readWorkbook.close();
				System.out.println("Saved Excell file to: " + filePath);
			}
			
			catch(IOException ex) {
				System.out.println("The file could not be written: " + ex.getMessage());
			}
		}
		
	}

}
