package com.exelatech.linkchecker;

import java.net.HttpURLConnection;
import java.net.URL;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;


public class BaseTestClass {

	public static WebDriver driver;
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"\\drivers\\chromedriver.exe");
		
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		
		String sheetname = System.getProperty("user.dir")+"\\drivers\\QueryDataSheet.xlsx";
		
		driver.navigate().to("https://www.google.co.in/"); //open the google home page
		
		Thread.sleep(5000);
		int rowscount = ExcelReader.setExcelFile(sheetname);  //get the total number of keywords
		
		int QueryColNum=ExcelReader.getColumnNumber(0,"Query");
		int statusCode;
		
		
		String keyword = (String) ExcelReader.getCellData(1, QueryColNum);
		//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input
		driver.findElement(By.xpath("//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input")).sendKeys(keyword);
		driver.findElement(By.xpath("//div[@class='FPdoLc VlcLAe']/center/input[@value='Google Search']")).click();
		
		//execute the first for loop for the number of keywords
		for(int x=2; x<=rowscount; x++)
		{
			driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
			
			List<WebElement> itr = driver.findElements(By.xpath(".//*[@id='rso']/div/div/div/div/div/div[1]/a"));
			
			//for loop for iterating through the pages of google
			for(int y=0; y<4; y++)
			{
				//loop for iterating through the number of links present on a single google page
				for(int z=0; z<itr.size(); z++)
				{
					//itr.get(z).click();
					String href;
					href = itr.get(z).getAttribute("href");
					HttpURLConnection huc;
					huc = (HttpURLConnection)(new URL(href).openConnection());
					Thread.sleep(100);
					huc.setRequestMethod("HEAD");
					Thread.sleep(100);
					huc.connect();
					Thread.sleep(100);
					statusCode = huc.getResponseCode();
					Thread.sleep(100);
					if(statusCode >= 400) {
					    System.out.println(href+" is a broken link");
					} else {
					    System.out.println(href+" is a valid link");
					}
					
				}
			}
			keyword = (String) ExcelReader.getCellData(x, QueryColNum);
			//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input
			driver.findElement(By.xpath("//*[@id='tsf']/div[2]/div[1]/div[2]/div/div[2]/input")).clear();
			driver.findElement(By.xpath("//*[@id='tsf']/div[2]/div[1]/div[2]/div/div[2]/input")).sendKeys(keyword);
			driver.findElement(By.xpath("//*[@id='tsf']/div[2]/div[1]/div[2]/button")).click();
			System.out.println("========================================================================");
			System.out.println("New Keyword");
			System.out.println("========================================================================");
		}
	}

}
