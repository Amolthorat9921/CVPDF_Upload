package UAT.AOS;


import java.io.IOException;
import java.time.Duration;
import java.util.Properties;


import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.testng.asserts.SoftAssert;

import Pagefactory.utility;

public class LoginPage extends utility 
{
	
	Properties prop =new Properties();
	
	public WebDriver wd;

	@FindBy(id = "userName")
	WebElement username;

	@FindBy(id = "loginPassword")
	WebElement password;

	@FindBy(xpath = "//select[@id='rooms']")
	WebElement ddproject;

	@FindBy(xpath = "//button[@id='submitid']")
	WebElement btnlogin;

	@FindBy(xpath = "//div[@id='header-main']")
	WebElement header;

	@FindBy(xpath = "//span[@id='loginError']")
	WebElement lbllogin;

	@FindBy(xpath = "//button[@id='cvModelLoginValidationOk']")
	WebElement sesionmgr;

	@FindBy(xpath = "//div[@class='nameValidation']")
	WebElement lblusername;

	@FindBy(xpath = "//div[@class='passwordValidation']")
	WebElement lblpwd;

	@FindBy(xpath = "//span[@id='cvModelLoginValidationMessage']")
	WebElement validationalert;

	public LoginPage(WebDriver wd) 
	{
		this.wd = wd;
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		PageFactory.initElements(wd, this);
	}

	public void login() throws IOException //String User, String Pass
	
	
	{
		SoftAssert softAssert = new SoftAssert();
		username.sendKeys(ConfigReader.getUsername());
		password.sendKeys(ConfigReader.getPassword());

		utility.Dropdownbyindex(ddproject, 1);

		btnlogin.click();

		if (utility.isDisaplyedW(sesionmgr, wd, 5)) 
		{
			sesionmgr.click();
			System.out.println("Another Session Ended");
		}

		if (utility.isVisible(header, wd, 10)) 
		{
			System.out.println("Login sucess");
			softAssert.assertAll("Login With Valid data Pass");

		} else 
		{
			String Msg = lbllogin.getText();
			System.out.println("Login failed with reason:- " + Msg);
		}

	}

}
