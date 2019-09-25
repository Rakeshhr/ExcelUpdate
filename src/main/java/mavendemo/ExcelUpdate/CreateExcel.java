package mavendemo.ExcelUpdate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {
	
	public void updateexcel(String num,String res){
		try { 
			FileInputStream file = new FileInputStream(new File("F:/JavaBooks.xlsx")); 

			// Create Workbook instance holding reference to .xlsx file 
			XSSFWorkbook workbook = new XSSFWorkbook(file); 

			// Get first/desired sheet from the workbook 
			XSSFSheet sheet = workbook.getSheetAt(0); 

			// Iterate through each rows one by one 
			Iterator<Row> rowIterator = sheet.iterator(); 
			while (rowIterator.hasNext()) { 
				Row row = rowIterator.next(); 
				// For each row, iterate through all the columns 
				Iterator<Cell> cellIterator = row.cellIterator(); 

				while (cellIterator.hasNext()) { 
					Cell cell = cellIterator.next(); 
					// Check the cell type and format accordingly 
					if((cell.getColumnIndex()==0))//for example of c
					{

						if(cell.toString().equals(num))
						{
							int rowtoupdate = row.getRowNum();
							int coltoupdate = cell.getColumnIndex()+2;						
							Cell cell2Update = sheet.getRow(rowtoupdate).getCell(coltoupdate);
							cell2Update.setCellValue(res);
							FileOutputStream outputStream = new FileOutputStream("F:/JavaBooks.xlsx"); 
								workbook.write(outputStream);
						
						}
					}

				} 

			} 
			file.close(); 
		} 
		catch (Exception e) { 
			e.printStackTrace(); 
		}

	}

}
