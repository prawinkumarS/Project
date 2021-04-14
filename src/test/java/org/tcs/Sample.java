package org.tcs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Sample {

	public static void main(String[] args) throws IOException {
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\91790\\eclipse-workspace\\maven\\Driver\\chromedriver.exe");

		WebDriver driver = new ChromeDriver();

		driver.get("http://demo.automationtesting.in/Register.html");

		driver.manage().window().maximize();

		WebElement year = driver.findElement(By.id("countries"));
		Select s=new Select(year);
		List<WebElement> op = s.getOptions();
		File f=new File("C:\\Users\\91790\\eclipse-workspace\\maven\\Excel Sheet\\dd.xlsx");
		Workbook w=new XSSFWorkbook();
		Sheet s1 = w.createSheet("DropDown");
		Row r = s1.createRow(op.size());
		Cell c = r.createCell(0);
		for (int i = 0; i < op.size(); i++) {
			c.setCellValue(op.get(i).getText());
			System.out.println(op.get(i).getText());
		}
		w.write(new FileOutputStream(f));
	}
}
