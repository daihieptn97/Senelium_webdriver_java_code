
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class CommentTest {

	private WebDriver driver;
	private int timeDelay = 2; // seconds
	public static final String TEST_FILE_PATH = "./TestCase.xlsx";
	public static final String exePath = "./chromedriver.exe";
	public int numberTest = 6;
	public int numberTestSucces = 0;
	public int numberTestFail = 0;

	@BeforeClass
	public void beforeClass() {
		// Khởi tạo trình duyệt Firefox
//		System.setProperty("webdriver.gecko.driver","./geckodriver.exe");
//		driver = new FirefoxDriver();
		
//		System.setProperty("webdriver.edge.driver","./MicrosoftWebDriver.exe");
//		 driver =  new EdgeDriver();
//		
//		System.setProperty("webdriver.ie.driver", "./IEDriverServer.exe");
//		 driver = new InternetExplorerDriver();
		
		System.setProperty("webdriver.chrome.driver", exePath);
		driver = new ChromeDriver();

		driver.get("http://localhost/news/public/user/news/8");
		// Mở rộng cửa sổ trình duyệt lớn nhất
		driver.manage().window().maximize();
		// Wait 10s cho page được load thành công
		driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
		// In ra thông báo theo mong muốn
		System.out.println("Open Success!");
	}

	@Test
	public void handllerOpenWeb() throws InterruptedException, IOException {

		File excelFile = new File(TEST_FILE_PATH);
		FileInputStream fis = new FileInputStream(excelFile);

		// we create an XSSF Workbook object for our XLSX Excel File
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		// lay sheet dau tien
		XSSFSheet sheet = workbook.getSheetAt(0);

		// lay cot input 6
		for (int i = 2; i < 6; i++) {
			XSSFCell cell = sheet.getRow(i).getCell(3);
			String name, content, contentInit, messges;
			contentInit = cell.getStringCellValue();
			String[] parts = contentInit.split("\n"); // tach chuoi lay content
			name = parts[1].substring(parts[0].indexOf(":") + 2);
			content = parts[0].substring(parts[0].indexOf(":") + 2);
			messges = sheet.getRow(i).getCell(4).getStringCellValue();
			System.out.println("TestCase " + (i - 1) + " : \n" + contentInit + "\nEo : " + messges);

			handllerTestCaseComment(content, name, messges);

			System.out.println();
		}

		workbook.close();
		fis.close();

	}

	@AfterClass
	public void afterClass() {
		// Đóng trình duyệt
		try {
			System.out.println(" -------- Test result  -------- ");
			System.out.println("Test Succes  : " + numberTestSucces);
			System.out.println("Test fail : " + numberTestFail);
			System.out.println(" -------- -----------  -------- ");
			TimeUnit.SECONDS.sleep(5);
			driver.quit();
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void handllerTestCaseComment(String content, String name, String messges) throws InterruptedException {
		if (content.trim().length() > 0) {
			driver.findElement(By.cssSelector("textarea.form-control.resize-textarea.input-check")).sendKeys(content);
			TimeUnit.SECONDS.sleep(timeDelay);
		}

		driver.findElement(By.cssSelector(".btn.btn-outline-primary.float-right")).click();
		TimeUnit.SECONDS.sleep(timeDelay);

		if (name.trim().length() > 0) {
			driver.findElement(By.cssSelector("input.form-control.input-group-sm.input-check")).sendKeys(name);
			TimeUnit.SECONDS.sleep(timeDelay);
		}

		driver.findElement(By.cssSelector("button.btn.btn-primary.w-100")).click();

		TimeUnit.SECONDS.sleep(timeDelay);
		String messgesTestResult = null;
		try {
			messgesTestResult = driver.findElement(By.cssSelector(".toast")).getText();
			messgesTestResult = messgesTestResult.substring(messgesTestResult.indexOf("\n") + 1);
		} catch (Exception e) {
			messgesTestResult = driver.findElement(By.cssSelector(".sweet-alert.showSweetAlert.visible h2")).getText();
			Actions actionObject = new Actions(driver);
			actionObject.keyDown(Keys.CONTROL).sendKeys(Keys.F5);
		}

		if (messgesTestResult.equals(messges)) { // so sanhh messsger ban dau va cai duoc tra ve
			System.out.println("Passed");
			numberTestSucces += 1;
		} else {
			numberTestFail += 1;
			System.out.println("Fail");
		}

		TimeUnit.SECONDS.sleep(2);
	}

}
