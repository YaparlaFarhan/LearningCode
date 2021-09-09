package com.atmecs.test;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.Test;

public class AtmecsHomePageTest {

	@Test
	public static void TestAtmecsHome() {
		System.setProperty("webdriver.chrome.driver", "D:\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(3000, TimeUnit.SECONDS);

		String url = "https://www.atmecs.com/";
		String title = "Home | Atmecs, Inc. | Digital Solutions & Product Engineering Services";
		driver.get(url);
		String act_url = driver.getCurrentUrl();
		Assert.assertEquals(url, act_url);

		String act_title = driver.getTitle();
		Assert.assertEquals(title, act_title);
		driver.close();

	}

}
