package org.framworkaug;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Dtata_FramWork {
	
	public static void main(String[] args) throws IOException
	{
		File file = new File("C:\\Users\\kumar\\OneDrive\\Desktop\\Suriya-jave everyday documents\\sample.xlsx");
		
		Workbook book = new XSSFWorkbook();
		
		Sheet createSheet = book.createSheet("logindata");
		
		System.out.println("test");
		
		Row createRow = createSheet.createRow(2);
		
		Cell createCell = createRow.createCell(0);
		
		createCell.setCellValue("Suriya@123");
		
		FileOutputStream out = new FileOutputStream(file);
		
		book.write(out);

		
	}
	
}
