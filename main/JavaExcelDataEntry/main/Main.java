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
        Sheet sheet = workbook.createSheet("First sheet");
		
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
			
			filePath = excellFolder + newDocName;
	        
			
				
			try (FileOutputStream outputFile = new FileOutputStream(filePath)){
		        
		        //Create a new row within sheet, first row index will be 0
		        Row header = sheet.createRow(0);
		        
		        //Create header cells within the rows on the new sheet
		        Cell cell_header = header.createCell(0);
		        cell_header.setCellValue("ID");
		        
		        cell_header = header.createCell(1);
		        cell_header.setCellValue("FIRSTNAME");
		        
		        cell_header = header.createCell(2);
		        cell_header.setCellValue("LASTNAME");
		        
		        cell_header = header.createCell(3);
		        cell_header.setCellValue("PHONENUMBER");
		        
		        //Create data row and cells for header cells, allow for user input using scanner
		        Scanner input = new Scanner(System.in);
		        
		        Row row_1 = sheet.createRow(1);  
		        
		        Cell cell_1 = row_1.createCell(0);
		        String userId = "";
		        System.out.print("Please Enter Your ID number: ");
		        userId = input.nextLine();
		        System.out.println();
		        cell_1.setCellValue(userId);
		        
		        cell_1 = row_1.createCell(1);
		        String userFirstname = "";
		        System.out.print("Please Enter Your Firstname: ");
		        userFirstname = input.nextLine();
		        System.out.println();
		        cell_1.setCellValue(userFirstname);
		        
		        cell_1 = row_1.createCell(2);
		        String userLastname = "";
		        System.out.print("Please Enter Your Lastname: ");
		        userLastname = input.nextLine();
		        System.out.println();
		        cell_1.setCellValue(userLastname);
		        
		        cell_1 = row_1.createCell(3);
		        String userPhonenumber = "";
		        System.out.print("Please Enter Your Phonenumber: ");
		        userPhonenumber = input.nextLine();
		        System.out.println();
		        cell_1.setCellValue(userPhonenumber);
		        
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
				Sheet readSheet = readWorkbook.getSheetAt(0);
				
				//Update cell within selected Excell doc
				Scanner cellSelection = new Scanner(System.in);
				String selection = "";
				System.out.println("\n" + "1. ID (Current: " + readSheet.getRow(1).getCell(0) + ")" +
				"\n" + "2. Firstname (Current: " + readSheet.getRow(1).getCell(1) + ")" +
						"\n" + "3. Lastname (Current: " + readSheet.getRow(1).getCell(2) + ")" +
						"\n" + "4. Phonenumber (Current: " + readSheet.getRow(1).getCell(3) + ")" );
				System.out.print("Please select with cell you would like to update: ");
				selection = cellSelection.nextLine();
				System.out.println();
				Cell cellUpdate;
				
				if(selection.equals("1")) {
					cellUpdate = readSheet.getRow(1).getCell(0);
					Scanner updateCell = new Scanner(System.in);
			        String userId = "";
			        System.out.print("Please Enter A New ID number: ");
			        userId = updateCell.nextLine();
					cellUpdate.setCellValue(userId);
					inputStream.close();
			        
					//Save the workbook to the file system
					readWorkbook.write(outputFile);
					readWorkbook.close();
					System.out.println("Saved Excell file to: " + filePath);
				}else if(selection.equals("2")) {
					cellUpdate = readSheet.getRow(1).getCell(1);
					Scanner updateCell = new Scanner(System.in);
			        String userFirstname = "";
			        System.out.print("Please Enter A New Firstname: ");
			        userFirstname = updateCell.nextLine();
					cellUpdate.setCellValue(userFirstname);
					inputStream.close();
			        
					//Save the workbook to the file system
					readWorkbook.write(outputFile);
					readWorkbook.close();
					System.out.println("Saved Excell file to: " + filePath);
				}else if(selection.equals("3")) {
					cellUpdate = readSheet.getRow(1).getCell(2);
					Scanner updateCell = new Scanner(System.in);
			        String userLastname = "";
			        System.out.print("Please Enter A New Firstname: ");
			        userLastname = updateCell.nextLine();
					cellUpdate.setCellValue(userLastname);
					inputStream.close();
			        
					//Save the workbook to the file system
					readWorkbook.write(outputFile);
					readWorkbook.close();
					System.out.println("Saved Excell file to: " + filePath);
				}else if(selection.equals("4")) {
					cellUpdate = readSheet.getRow(1).getCell(2);
					Scanner updateCell = new Scanner(System.in);
			        String userPhonenumber = "";
			        System.out.print("Please Enter A New Lastname: ");
			        userPhonenumber = updateCell.nextLine();
					cellUpdate.setCellValue(userPhonenumber);
					inputStream.close();
			        
					//Save the workbook to the file system
					readWorkbook.write(outputFile);
					readWorkbook.close();
					System.out.println("Saved Excell file to: " + filePath);
				}
			}
			
			catch(IOException ex) {
				System.out.println("The file could not be written: " + ex.getMessage());
			}
		}
		
	}

}
