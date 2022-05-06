package org.UsingApachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.ListIterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class WriteExcel {
	public static void main(String[] args) throws IOException {

		
		
		
		File files=new File("C:\\Users\\ADMIN\\Desktop\\user Data.xlsx");
		
		FileInputStream fs=new FileInputStream(files);
		
		Workbook workbook=new XSSFWorkbook(fs);
		
		Sheet sheet = workbook.getSheet("Sheet1");
		
		/*for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j <row.getPhysicalNumberOfCells();  j++) {
				Cell cell = row.getCell(j);
				System.out.println(cell.toString());
			}
		}	*/	
		
		List<String> namelist=new ArrayList<String>();
		List<String> surnamelist=new ArrayList<String>();
		List<String> dateofbirth=new ArrayList<String>();
		List<String> age=new ArrayList<String>();
		List<String> address=new ArrayList<String>();
		
		Iterator<Row> rowiterator = sheet.iterator();

		while(rowiterator.hasNext())
		{
			Row title = rowiterator.next();
			Iterator<Cell> celliterator = title.iterator();
			int i=2;	
			while(celliterator.hasNext())
			{
			if(i==2)
			{
				namelist.add(celliterator.next().toString());
			}else if(i==3) {
				surnamelist.add(celliterator.next().toString());
			}else if(i==4) {
				dateofbirth.add(celliterator.next().toString());
			}else if(i==5) {
				age.add(celliterator.next().toString());
			}else if(i==6) {
				address.add(celliterator.next().toString());
				break;
			}
			i++;
		}
		}
			for (int i = 0; i < namelist.size(); i++) {
				System.out.println(namelist.get(i));
				System.out.println(surnamelist.get(i));
				System.out.println(dateofbirth.get(i));
				System.out.println(age.get(i));
				System.out.println(address.get(i));

			}
		
	
	
	}}	


