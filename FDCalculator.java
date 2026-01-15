package Excelworks;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class FDCalculator {

	public static void main(String[] args) throws IOException, InterruptedException {
		// TODO Auto-generated method stub
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		
		driver.get("https://www.calculatorsoup.com/calculators/financial/simple-interest-calculator.php");
		driver.manage().window().maximize();
		
		String filepath = System.getProperty("user.dir")+"\\data\\caladata.xlsx";
		int rows = Excellutils.getRowCount(filepath, "Sheet1");
		for(int i =1; i<=rows;i++)
		{
			//read data from excel
			String pric = Excellutils.getCellData(filepath,"Sheet1",i,0);
			String rateofinterest = Excellutils.getCellData(filepath,"Sheet1",i,1);
			String per1 = Excellutils.getCellData(filepath,"Sheet1",i,2);
			String per2 = Excellutils.getCellData(filepath,"Sheet1",i,3);
			String fre = Excellutils.getCellData(filepath,"Sheet1",i,4);
			String exp_mvalue = Excellutils.getCellData(filepath,"Sheet1",i,5);
			//pass above data into applicatopn
			driver.findElement(By.xpath("//input[@name='P']")).sendKeys(pric);
			driver.findElement(By.xpath("//input[@name='R']")).sendKeys(rateofinterest);
			driver.findElement(By.xpath("//input[@name='t']")).sendKeys(per1);
			
			Select perdrop=new Select(driver.findElement(By.xpath("//Select[@name='time_t']")));
			perdrop.selectByVisibleText(per2);
			
			driver.findElement(By.xpath("//input[@type='submit']")).click();
			//capture the result
			String act_mvalue=driver.findElement(By.xpath("//div[@id='answer']/strong[1]")).getText();
			if(Double.parseDouble(exp_mvalue)==Double.parseDouble(act_mvalue))
			{
				System.out.println("Test passed");
				Excellutils.SetCellData(filepath,"Sheet1",i, 7,"Passed");
				Excellutils.fillGreenColor(filepath,"Sheet1",i, 7);
			}
			else
			{
				System.out.println("Test failed");
				Excellutils.SetCellData(filepath,"Sheet1",i, 7,"failed");
				Excellutils.fillRedColor(filepath,"Sheet1",i, 7);
			}
			Thread.sleep(3000);
			driver.findElement(By.xpath("//input[@type='reset']")).click();
			//write data into excel
		}
		driver.quit();

	}

}
