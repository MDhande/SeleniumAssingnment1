package Cases;

import org.openqa.selenium.By;
import org.openqa.selenium.By.ByClassName;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelData 
{
	WebDriver driver = null;

	@BeforeTest
	public void setupTest()
	{
		String path =System.getProperty("user.dir");
		System.setProperty("webdriver.chrome.driver",path+ "\\Drivers\\Chromedriver\\chromedriver.exe");
		driver= new ChromeDriver();  
	}

	@Test(dataProvider = "test1data")
	public void test1(String username, String password) throws InterruptedException
	{
		System.out.println(username+ "  | " +password);

		driver.get("https://j2store.org/login.html");
		driver.findElement(By.id("username")).sendKeys(username);
		driver.findElement(By.name("password")).sendKeys(password);
		driver.findElement(By.xpath("//*[@id=\"t3-content\"]/div[2]/div[1]/form/fieldset/div[4]/div/button")).sendKeys(Keys.ENTER);


		Thread.sleep(3000);
	}

	@AfterTest
	public void downuptest()
	{

		System.out.println("User Login Successful");

	}


	@DataProvider(name="test1data")
	public static Object[][] getData()
	{
		String excelpath="D:\\JavaProjects\\SeleniumAssignment\\Excel\\Data.xlsx";
		Object data[][] = test(excelpath, "Sheet1");
		return data;

	}


	public static Object[][] test(String excelpath, String sheetname)	
	{

		ExcelUtils excel= new ExcelUtils( excelpath, sheetname);

		int rowcount=excel.getRow();
		int Colcount=excel.getCol();

		Object data[][] =new Object[rowcount-1][Colcount];

		for(int i=1;i<rowcount;i++)
		{
			for(int j=0;j<Colcount;j++)
			{
				String cellData=excel.getcellDataString(i, j);
				data[i-1][j] = cellData;


			}

		}
		return data;



	}
}

