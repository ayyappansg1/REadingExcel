package org.UsingApachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
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
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class SabaTask {
	public static void main(String[] args) throws IOException {

		System.setProperty("webdriver.chrome.driver", "E:\\java files\\ReadingExcel\\Driver\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.get("http://demo.automationtesting.in/Register.html");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));

		List<String> skillsList=new ArrayList<String>();

		List<WebElement> findElements = driver.findElements(By.xpath("//select[@id='Skills']//option"));

		for (WebElement webElement : findElements) {
			skillsList.add(webElement.getText());
		}
		System.out.println(skillsList);

		
		Workbook workbook=new XSSFWorkbook();
		Sheet createSheet = workbook.createSheet("SabaTask");
		
	

		for (int i = 0; i < skillsList.size(); i++) {
			FileOutputStream filewrite=new FileOutputStream("E:\\SabasTask.xlsx");
			Row createRow = createSheet.createRow(i);
			Cell createCell = createRow.createCell(0);
			createCell.setCellValue(skillsList.get(i));
			workbook.write(filewrite);
			
		}


		driver.quit();


	}}	


