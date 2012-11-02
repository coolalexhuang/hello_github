package selenium;

import java.util.concurrent.TimeUnit;
import org.junit.*;
import junit.framework.TestCase;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;


@SuppressWarnings("unused")
public class Test1 extends TestCase{
	private WebDriver driver;
	private String baseUrl;
	private StringBuffer verificationErrors = new StringBuffer();
	@Before
	public void setUp() throws Exception {
		//For IE,Firefox
		//driver = new InternetExplorerDriver(); //For IE
		driver = new FirefoxDriver(); //For fire fox
		baseUrl = "http://www.moodys.com";
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		 
		//For Chrome
		//Optional, if not specified, WebDriver will search your path for chromedriver.
		/*System.setProperty("webdriver.chrome.driver", "/Selenium/chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		baseUrl = "http://www.moodys.com";
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		driver.get(baseUrl + "/");
		assertTrue(isElementPresent(By.id("RegisterText")));
		driver.findElement(By.id("LoginText")).click();
		driver.findElement(By.id("MdcUserName")).clear();
		driver.findElement(By.id("MdcUserName")).sendKeys("huangjia");
		driver.findElement(By.id("MdcPassword")).clear();
		driver.findElement(By.id("MdcPassword")).sendKeys("alex1234");
		driver.findElement(By.cssSelector("#LoginImageButton > span")).click();
		assertTrue(isElementPresent(By.id("LogOutText")));
		driver.findElement(By.id("kw")).click();
		driver.findElement(By.id("kw")).clear();
		driver.findElement(By.id("kw")).sendKeys("nike");
		driver.get(baseUrl + "/credit-ratings/NIKE-Inc-credit-rating-40400");	*/	 
	}

	@Test
	public void testJava() throws Exception {
		driver.get(baseUrl + "/");
		assertTrue(isElementPresent(By.id("RegisterText")));
		driver.findElement(By.id("LoginText")).click();
		driver.findElement(By.id("MdcUserName")).clear();
		driver.findElement(By.id("MdcUserName")).sendKeys("huangjia");
		driver.findElement(By.id("MdcPassword")).clear();
		driver.findElement(By.id("MdcPassword")).sendKeys("alex1234");
		driver.findElement(By.cssSelector("#LoginImageButton > span")).click();
		assertTrue(isElementPresent(By.id("LogOutText")));
	}

	@After
	public void tearDown() throws Exception {
		driver.quit();
		String verificationErrorString = verificationErrors.toString();
		if (!"".equals(verificationErrorString)) {
			fail(verificationErrorString);
		}
	}

	private boolean isElementPresent(By by) {
		try {
			driver.findElement(by);
			return true;
		} catch (NoSuchElementException e) {
			return false;
		}
	}
}
