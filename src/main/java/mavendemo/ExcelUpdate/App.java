package mavendemo.ExcelUpdate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Hello world!
 *
 */
public class App 
{
	public static void main(String[] args)  {

		CreateExcel excel = new CreateExcel();
		excel.updateexcel("3","fail");
	}
}
