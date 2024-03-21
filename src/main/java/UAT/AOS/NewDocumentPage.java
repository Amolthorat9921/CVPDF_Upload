package UAT.AOS;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import Pagefactory.utility;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NewDocumentPage extends utility 
{
	public WebDriver wd;

	public NewDocumentPage(WebDriver wd) 
	{
		this.wd = wd;
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		PageFactory.initElements(wd, this);
	}

	@FindBy(xpath = "//div[@id='addPagesDropDown']")
	WebElement SubMenuNwdocument;

	@FindBy(xpath = "//input[@id='viewDocumentAddPages']")
	WebElement UploadNewFile;

	@FindBy(xpath = "//div[@id='saveAddedPages']")
	WebElement btnSaveDocument;

	@FindBy(xpath = "//span[@id='messageContent42']")
	WebElement lblSaveDocument;

	@FindBy(xpath = "//button[@id='messageButtonOK42']")
	WebElement btnokSaveDocument;

	@FindBy(xpath = "//div[@id='cvDocumentClose']")
	WebElement btnCloseDocument;

	@FindBy(xpath = "(//div[@id='progressMsg'])[2]")
	WebElement DocumentUploadProgressBar;

	@FindBy(xpath = "(//*[contains(text(), 'ASK')]/ancestor::li/ins[@class='jstree-icon'])[1]")
	WebElement IconPlusCabinet;

	@FindBy(xpath = "(//*[contains(text(), 'ASK')]/ancestor::li/ins[@class='jstree-icon'])[2]")
	WebElement IconPlusDrawer;

	@FindBy(linkText = "ASKUpload")
	WebElement FolderNameWord;

	@FindBy(linkText = "ASKUpload")//SalesReportPDF
	WebElement FolderNamePDF;

	@FindBy(xpath = "//input[@id='tableFilter']")
	WebElement searchbarPdfDocument;

	@FindBy(xpath = "//td[@class=' customDocName']")
	WebElement clickDocument;

	@FindBy(xpath = "//select[@id='docTypeList']")
	WebElement ddDocumenttype;

	@FindBy(xpath = "//button[@id='createDocumentSubmit']")
	WebElement btnCreateDocument;

	@FindBy(xpath = "//input[@id='retainBtn']")
	WebElement chkRetain;

	@FindBy(xpath = "//input[@id='indices_5']")
	WebElement txtReportName;

	@FindBy(xpath = "(//div[@class='spinner-border spinner-border-lg'])[2]")
	WebElement DocumentUploadloader;

	@FindBy(xpath = "//button[@id='modelNewDocument']")
	WebElement btnNewDocument;

	@FindBy(xpath = "//span[@id='createDocumentMessage']")
	WebElement FileUploadlblMessage;

	@FindBy(xpath = "//a[@id='createDocument']")
	WebElement SubMenuCreateNewDocument;

	@FindBy(xpath = "//input[@id='indices_46']")
	WebElement txtCompanyName;

	@FindBy(xpath = "//div[@class='e-toast-message']")
	WebElement errorwindow;
	
	@FindBy(xpath="//button[@class='e-small e-lib e-btn e-control e-primary e-toast-danger']")
	WebElement btnCloseError; 
			
	
	
	@SuppressWarnings("resource")


	public void pdf() throws IOException 
	{
		FileInputStream fis = new FileInputStream(System.getProperty("user.dir")+ "\\data\\Pan_Card.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);

		XSSFSheet sheet = wb.getSheet("Sheet2");
		int rowCount = sheet.getPhysicalNumberOfRows() - 1;

		Actions a = new Actions(wd);
		utility.isVisible(IconPlusCabinet, wd, 5);
		IconPlusCabinet.click();
		utility.isVisible(IconPlusDrawer, wd, 5);
		IconPlusDrawer.click();
		utility.isVisible(FolderNamePDF, wd, 5);
		FolderNamePDF.click();
		SubMenuCreateNewDocument.click();
		utility.Dropdownbytxt(ddDocumenttype, "Sales Report");
		
		if (chkRetain.isSelected()) {

		} else 
		{
			chkRetain.click();
		}

		System.out.println("No of Record Found Into Excel :- " + rowCount);

		for (int i = 1; i <= rowCount; i++) 
		{

			try {
				Thread.sleep(2000);
				utility.isDisaplyedW(SubMenuNwdocument, wd, 10);
				a.moveToElement(SubMenuNwdocument).perform();
				System.out.println(sheet.getRow(i).getCell(0).getStringCellValue());
				UploadNewFile.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
				
				Thread.sleep(2000);

				if (utility.isAlertPresent(wd) == true) 
				{
					wd.switchTo().alert().accept();
					Thread.sleep(4000);
				} 
				else 
				{
					
				}
				
				utility.isVisible(txtCompanyName, wd, 15);
				txtCompanyName.clear();
				txtCompanyName.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());

				btnCreateDocument.click();
				Thread.sleep(3000);
				
				utility.isInvisible(DocumentUploadloader, wd, 10);
				String FileUploadStatusMsg = FileUploadlblMessage.getText();

				XSSFRow row = sheet.getRow(i);
				XSSFCell cell = row.createCell(2);
				FileOutputStream fos = new FileOutputStream(System.getProperty("user.dir")+ "\\data\\Pan_Card.xlsx");

				if (FileUploadStatusMsg.contains("Document created successfully")) 
				{
					cell.setCellValue("File Upload");
					wb.write(fos);

				} else 
				{
					cell.setCellValue("Fail");
					wb.write(fos);
				}

				btnNewDocument.click();
				FileUploadStatusMsg = "";

				} catch (Exception e) 
				{
					if (utility.isDisaplyedW(errorwindow, wd, 1)) 
				{
					System.out.println(sheet.getRow(i).getCell(1).getStringCellValue()+":- This File is zero KB");
					btnCloseError.click();
				}
				continue;

			}

		}	System.out.println("Its Done");

	}
	
}
