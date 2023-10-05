package org.framworkaug;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class Dtata_FramWork2 {

	public static void main(String[] args) throws IOException
	{
		File file = new File("C:\\Users\\kumar\\OneDrive\\Desktop\\Suriya-jave everyday documents\\Book1.xlsx");

		FileInputStream input = new FileInputStream(file);

		Workbook book = new XSSFWorkbook(input);

		Sheet sheet = book.getSheet("login");
		
//		Row row = sheet.getRow(2);
//		
//		Cell cell = row.getCell(0);
//		
//		String stringCellValue = cell.getStringCellValue();
//		System.out.print(stringCellValue+"    \t");
//		
		
		

		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {

			Row row = sheet.getRow(i);

			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {

				Cell cell = row.getCell(j);

				CellType cellType = cell.getCellType();

				switch (cellType) {
				case STRING:
					String stringCellValue = cell.getStringCellValue();

					if(stringCellValue.equals("suriyaravi@gmail.com"))         //suriyaravirani@gamil.com
					{
						cell.setCellValue("suriyarani@gmail.com");
						FileOutputStream out = new FileOutputStream(file);
						book.write(out);

					}else if(stringCellValue.equals("suriyaravirani@gamil.com")) {
						cell.setCellValue("suriyarani@gmail.com");
						FileOutputStream out1 = new FileOutputStream(file);
						book.write(out1);

					}
					System.out.print(stringCellValue+"    \t");
					break;
				case NUMERIC:
					if(DateUtil.isCellDateFormatted(cell)) {
						Date dateCellValue = cell.getDateCellValue();
						SimpleDateFormat simple = new SimpleDateFormat("dd/MMMM/yyyy");
						String format = simple.format(dateCellValue);
						System.out.print(format);
					} else {
						double numericCellValue = cell.getNumericCellValue();

						long l = (long)numericCellValue;
						BigDecimal valueOf = BigDecimal.valueOf(l);
						String string = valueOf.toString();
						System.out.print(string+"   \t");

					}

					break;

				default:
					System.out.println("None of the above");
					break;
				}
			}
			System.out.println();
		}



	}
}
