package com.functionlibrary;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.excelplugin.ReadExcelData;

public class FunctionLibrary {

	public WebDriver driver;
	
	
	
	
	public void Login()
	{
		
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		
		driver=new ChromeDriver();
		ReadExcelData data=new ReadExcelData();
		
		String url=data.getData("Login", "Url", 1);
		driver.get(url);
		
		
		
		driver.manage().window().maximize();
	}
	
	
}
