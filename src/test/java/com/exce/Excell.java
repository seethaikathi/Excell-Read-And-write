package com.exce;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excell {
	public static void main(String[] args) throws Exception {
		File f = new File("C:\\Users\\ADMIN\\Desktop\\Book.xlsx");
		
	    FileOutputStream file = new FileOutputStream(f);
	    
	    Workbook w = new XSSFWorkbook();
	    Sheet sheet = w.createSheet("New sheet");
	    Row row = sheet.createRow(0);
	    Cell cell = row.createCell(0);
	    Cell cell2 = row.createCell(1);
	    Cell cell5 = row.createCell(2);
	    cell.setCellValue("Seethai");
	    cell2.setCellValue("Kathi");
	    cell5.setCellValue("TCS");
	    Row row2 = sheet.createRow(1);
	    Cell cell3 = row2.createCell(0);
	    Cell cell6 = row2.createCell(2);
	    cell3.setCellValue("Pudukkottai");
	    cell6.setCellValue("Chennai");
	    Row row3 = sheet.createRow(2);
	    Cell cell4 = row3.createCell(0);
	    cell4.setCellValue("Career Guidance");
	    
	    w.write(file);
		
			}
			
		
		
		
	}	
