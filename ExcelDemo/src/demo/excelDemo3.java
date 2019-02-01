package demo;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelDemo3 {
	
	public static void main(String[] args) throws Exception {

		XSSFWorkbook wb = new XSSFWorkbook();

		XSSFSheet sh = wb.createSheet("firstSheet");

		XSSFRow row = sh.createRow(0);

		row.createCell(1).setCellValue("1234");


		File destination = new File("C:\\Users\\Sarad\\OneDrive\\Documents\\Book1.xlsx");

		FileOutputStream fout = new FileOutputStream(destination);

		wb.write(fout);

		}

		}


