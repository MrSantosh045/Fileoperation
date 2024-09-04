package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading_Excel
{

	public static void main(String[] args) throws IOException
	{

		// Specify the location of Excel File
		File src= new File("C:\\Users\\Mrsan\\My_Test\\TestSheet.xlsx");


		// Load File

		FileInputStream fis =new FileInputStream(src);


		// Load Workbook

		XSSFWorkbook wb =new XSSFWorkbook(fis);


		// Load WorkSheet
		XSSFSheet sh =wb.getSheet("User_data");

		// Print the name of loaded sheet System.out.println(sh.getSheetName());


		// Print the name of loaded sheet System.out.println(sh.getSheetName());

		// Print UserName from Excel Sheet
		System.out.println(sh.getRow(0).getCell(0).getStringCellValue());
		// Print p2 from Excel Sheet
		System.out.println(sh.getRow(2).getCell(1).getStringCellValue());

	}

}


