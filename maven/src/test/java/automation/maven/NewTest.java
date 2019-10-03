package automation.maven;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class NewTest {
	WebDriver driver;
  @Test
  public void f() {
	  System.out.println("hi");
	  System.setProperty("webdriver.chrome.driver","E:\\seleniumsoftware\\chromedriver.exe");
	  driver=new ChromeDriver();
	  driver.get("http://www.google.com");
  }
  
}
