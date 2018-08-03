
import java.lang.reflect.Method;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class WordDocWithScreenshotTest {
	// Declaring variables
	private WebDriver driver;
	private String baseUrl;
	

	@BeforeMethod
	public void setUp() throws Exception {
		// Selenium version 3 beta releases require system property set up
		System.setProperty("webdriver.gecko.driver",
				"C:\\Users\\Soumyansh\\eclipse-workspace\\ScreeenshotsInWord\\Executables\\geckodriver.exe");
		// Create a new instance for the class FirefoxDriver
		// that implements WebDriver interface
		driver = new FirefoxDriver();
		// Implicit wait for 5 seconds
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		// Assign the URL to be invoked to a String variable
		baseUrl = "https://accounts.google.com/SignUp";
	}

	@Test
	public void testPageTitle(Method m) throws Exception {
		String testCaseName=m.getName();
		// Open baseUrl in Firefox browser window
		driver.get(baseUrl);
		SaveScreenshot.capture(m.getName() + "0", driver);
		// Locate First Name text box by id and
		// assign it to a variable of type WebElement
		WebElement firstName = driver.findElement(By.id("firstName"));
		// Clear the default placeholder or any value present
		firstName.clear();
		// Enter/type the value to the text box
		firstName.sendKeys("fname01");
		// Take a screenshot
		SaveScreenshot.capture(m.getName() + "1", driver);
		// Locate last name text box by name
		WebElement lastName = driver.findElement(By.name("lastName"));
		// Clear and enter a value
		lastName.clear();
		lastName.sendKeys("lname01");
		// Take a screenshot
		SaveScreenshot.capture(m.getName() + "2", driver);
		// Create a word document and include all screenshots
		SaveDocument.createDoc("TestCase_"+m.getName(), new String[] {m.getName() + "0", m.getName() + "1", m.getName() + "2" });
	}

	@AfterMethod
	public void tearDown() throws Exception {
		// Close the Firefox browser
		driver.close();
	}
}
