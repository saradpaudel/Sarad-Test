package demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriteDemo {
	
	public static void main(String[] args) throws Exception  {


			// Specify path of file
			File sourcefile = new File("C:\\Users\\Sarad\\OneDrive\\Documents\\testdata.xlsx");

			FileInputStream fis = new FileInputStream(sourcefile);

			// load workbook
			XSSFWorkbook wb = new XSSFWorkbook(fis);

			// Load Sheet - we are going to load sheet1
			XSSFSheet sh1 = wb.getSheet("Sheet1");

			sh1.getRow(4).getCell(1).setCellValue("howru");

			FileOutputStream fout = new FileOutputStream(sourcefile);

			wb.write(fout);
			

			
			//sh1.getRow(4).getCell(2).setCellValue("Hw r U?");

			//FileOutputStream fout1 = new FileOutputStream(sourcefile);

			//wb.write(fout1);
			
			}
			}
			