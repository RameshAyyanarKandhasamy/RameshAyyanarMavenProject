package driver;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class FrameworkWithExcelSheet {
	
	DesiredCapabilities capabilities;
	WebDriver driver;
	WebElement element;	
	boolean Status;
	int i=1;
	File file;
	FileInputStream inputStream;
	XSSFWorkbook InputWorkbook;
	FileOutputStream out;
	XSSFSheet InputSheet;
	Row rows;
	DateFormat dateFormat;
	Date date;
	
	
@BeforeTest
	
	public void BeforeTestmethod(){
//		capabilities = DesiredCapabilities.internetExplorer();
//		capabilities.setCapability("ignoreZoomSetting", true);
		dateFormat = new SimpleDateFormat("yyyy_MM_dd HH_mm_ss");
		date = new Date();
//		System.out.println(dateFormat.format(date)); //2016/11/16 12:08:43
	System.out.println("Before Test");
	}
	
@BeforeMethod
	
	public void Beforemethod(){
//	System.setProperty("webdriver.chrome.driver","D:\\workspace\\Com.Maven.New.Selenium\\driver\\chromedriver.exe");
//	driver = new ChromeDriver();
//	driver.manage().window().maximize();
	System.out.println("Before method");
	}

@Test(dataProvider="datapro")
public void Testmethod1(String Username,String pwd) throws InterruptedException{
//	driver.get("https://www.facebook.com/");
//	element=driver.findElement(By.name("email"));
//	element.sendKeys(Username);
//	element=driver.findElement(By.name("pass"));
//	element.sendKeys(pwd);
//	driver.findElement(By.name("login")).click();
//	Thread.sleep(5000);
	System.out.println(Username);
	System.out.println(pwd);
	rows.createCell(3).setCellValue("Pass");
}

@AfterMethod

public void AfterMethod() throws IOException{
//	driver.quit();
}


@AfterTest

public void AfterTest() throws IOException{
	out=new FileOutputStream("D:\\workspace\\Com.Maven.New.Selenium\\output\\Ouputsheet_"+dateFormat.format(date)+".xlsx");
    InputWorkbook.write(out);
    out.close(); 
}

@DataProvider(name="datapro")
public Object[][] getData() throws IOException
{
	
	file=new File("D:\\workspace\\Com.Maven.New.Selenium\\input\\inputsheet.xlsx");
    //Create an object of FileInputStream class to read excel file
    inputStream = new FileInputStream(file);
   //Create an object for workbook
    InputWorkbook = new XSSFWorkbook(inputStream);
   //Read sheet inside the workbook by its name
    InputSheet = InputWorkbook.getSheetAt(0);
    //Find number of rows in excel file
    int rowCount = InputSheet.getLastRowNum();
    
  //Rows - Number of times your test has to be repeated.
  //Columns - Number of parameters in test data.
    Object[][] data=new Object[rowCount][2];
    for(int row=1;row<=rowCount;row++)
    {
        rows=InputSheet.getRow(row);
       	// 1st row
    	data[row-1][0]=rows.getCell(0).getStringCellValue();
    	data[row-1][1]=rows.getCell(1).getStringCellValue();

    }	
return data;
}
}
