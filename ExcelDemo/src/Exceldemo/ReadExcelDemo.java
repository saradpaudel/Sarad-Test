package Exceldemo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDemo {
	

	public static void main(String[] args) throws Exception {

	//Specify path of file
	File sourcefile = new File ("C:\\Users\\Sarad\\OneDrive\\Documents\\testdata.xlsx");

	//load file
	FileInputStream fis = new FileInputStream(sourcefile);

	//load workbook
	XSSFWorkbook wb = new XSSFWorkbook(fis);

	//Load Sheet - we are going to load sheet1
	XSSFSheet sh1 = wb.getSheet("Sheet1");
	//XSSFSheet sh1 = wb.getSheetAt(0);
	
	String value1= sh1.getRow(0).getCell(2).getStringCellValue();

	System.out.println("Username = "+value1);

	Integer value2= (int)sh1.getRow(1).getCell(2).getNumericCellValue();

	System.out.println("Password = "+value2);
	
	String value3= sh1.getRow(2).getCell(2).getStringCellValue();

	System.out.println("Phone = "+value3);
	
	
	}

	}


