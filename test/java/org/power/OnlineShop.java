package org.power;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class OnlineShop {
	public static void main(String[] args) throws Throwable, IOException {
		System.setProperty("webdriver.chrome.driver", "E:\\Eclipseworkspace\\OnlineShope\\Driver\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.candy.com/");

		WebElement msecourse = driver.findElement(By.xpath("/html/body/div[1]/header/div/ul/li[3]/a/span[2]"));
		Actions acc = new Actions(driver);
		acc.moveToElement(msecourse).perform();

		WebElement Loginbtn = driver
				.findElement(By.xpath("/html/body/div[1]/header/div/ul/li[3]/div/div/div[1]/p[1]/a"));
		Loginbtn.click();

		WebElement FirstName = driver
				.findElement(By.xpath("/html/body/div[1]/div[5]/div/div/form/div[1]/div[2]/ul/li[1]/div/input"));
		FirstName.sendKeys(getData(1, 1));

		WebElement Password = driver
				.findElement(By.xpath("/html/body/div[1]/div[5]/div/div/form/div[1]/div[2]/ul/li[2]/div/input"));
		Password.sendKeys(getData(1, 2));

		driver.findElement(By.xpath("/html/body/div[1]/div[5]/div/div/form/div[1]/div[3]/div[2]/button")).click();

		// driver.navigate().to("https://www.candy.com/");
		WebElement searchBox1 = driver.findElement(By.xpath("//input[@id='search']"));

		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].value='Chocolates';", searchBox1);
		WebElement searchBtn = driver.findElement(By.xpath("//*[@id=\"search_mini_form\"]/div/button"));
		js.executeScript("arguments[0].click()", searchBtn);
		WebElement Ghirardelli = driver
				.findElement(By.xpath("//*[@id=\"searchspring-search_results\"]/div[17]/div/div[2]/div[1]/h3/a"));
		js.executeScript("arguments[0].scrollIntoView(true)", Ghirardelli);
		js.executeScript("arguments[0].click()", Ghirardelli);
		WebElement Addkart1 = driver
				.findElement(By.xpath("//*[@id=\"product_addtocart_form\"]/div[2]/div[2]/div[4]/button"));
		Addkart1.sendKeys(Keys.ENTER);
		driver.findElement(By.xpath("/html/body/div[1]/div[5]/div/div/div[2]/div[2]/form/div[2]/button")).click();

		WebElement fstnme = driver.findElement(By.xpath("//*[@id=\"shipping:firstname\"]"));
		fstnme.clear();
		fstnme.sendKeys(getData(1, 3));
		WebElement lstnme = driver.findElement(By.xpath("//*[@id=\"shipping:lastname\"]"));
		lstnme.clear();
		lstnme.sendKeys(getData(1, 4));
		WebElement phneno = driver.findElement(By.xpath("//*[@id=\"shipping:street1\"]"));
		phneno.clear();
		phneno.sendKeys(getData(1, 5));
		WebElement addrs = driver.findElement(By.xpath("//*[@id=\"shipping:street1\"]"));
		addrs.clear();
		addrs.sendKeys(getData(1, 6));
		WebElement zipcod = driver.findElement(By.xpath("//*[@id=\"shipping:postcode\"]"));
		zipcod.clear();
		zipcod.sendKeys(getData(1, 7));
		WebElement stedropdwn = driver.findElement(By.xpath("//*[@id=\"shipping:region_id\"]"));
		Select s = new Select(stedropdwn);
		s.selectByValue("2");
		WebElement city = driver.findElement(By.xpath("//*[@id=\"shipping:city\"]"));
		city.clear();
		city.sendKeys(getData(1, 8));
		WebElement contry = driver.findElement(By.xpath("//*[@id=\"shipping:country_id\"]"));
		Select s1 = new Select(contry);
		s1.selectByVisibleText("United States");
		Thread.sleep(2000);
		driver.findElement(By.xpath("//*[@id=\"shipping-buttons-container\"]/button")).click();

		Thread.sleep(2000);
		driver.findElement(By.xpath("/html/body/div[1]/div/div/div[1]/div/ol/li[1]/div[3]/form/div/button")).click();

		TakesScreenshot ts = (TakesScreenshot) driver;
		File f = ts.getScreenshotAs(OutputType.FILE);
		File g = new File("E:\\Eclipseworkspace\\OnlineShop\\screenshot.jpeg");
		FileUtils.copyFile(f, g);

		driver.findElement(By.xpath("//*[@id=\"shipping-address-additions\"]/div/div[2]/button")).click();
		driver.findElement(By.xpath("//*[@id=\"delivery-option-standard\"]/div[1]/div[1]/div[1]/span[1]/label"))
				.click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("//*[@id=\"shipping-method-buttons-container\"]/button")).click();
	}

	private static String getData(int rowNo, int cellno) throws Throwable {
		String v = null;
		File f = new File("E:\\Eclipseworkspace\\OnlineShop\\File\\MSK.xlsx");
		FileInputStream st = new FileInputStream(f);

		Workbook w = new XSSFWorkbook(st);
		Sheet s = w.getSheet("Sheet1");
		Row r = s.getRow(rowNo);
		Cell c = r.getCell(cellno);
		v = c.getStringCellValue();
		return v;
	}
}
