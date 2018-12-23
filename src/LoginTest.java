import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.TestNG;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class LoginTest {

	private WebDriver driver;
	private int timeDelay = 1; // seconds
	public static final String TEST_FILE_PATH = "./TestCase.xlsx";
	public static final String exePath = "./chromedriver.exe";
	public int numberTest = 6;
	public int numberTestSucces = 0;
	public int numberTestFail = 0;

	@BeforeClass
	public void beforeClass() {
		// Khởi tạo trình duyệt Firefox
//		System.setProperty("webdriver.gecko.driver","./geckodriver.exe");
//		 driver = new FirefoxDriver();
		
//		System.setProperty("webdriver.edge.driver","./MicrosoftWebDriver.exe");
//		 driver =  new EdgeDriver();
		
		System.setProperty("webdriver.ie.driver", "./IEDriverServer.exe");
		driver = new InternetExplorerDriver();
		
//		System.setProperty("webdriver.chrome.driver", exePath);
//		driver = new ChromeDriver();

		driver.get("http://localhost/news/public/admin/dashboard");
		// Mở rộng cửa sổ trình duyệt lớn nhất
		driver.manage().window().maximize();
		// Wait 10s cho page được load thành công
		driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
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
		XSSFSheet sheet = workbook.getSheetAt(1);

		// lay cot input 6
		for (int i = 1; i < 2; i++) {
			XSSFCell cell = sheet.getRow(i).getCell(2);
			String username, passwrod, contentInit, messges;
			contentInit = cell.getStringCellValue();
			String[] parts = contentInit.split("\n"); // tach chuoi lay content
			username = parts[0].substring(parts[0].indexOf(":") + 1);
			passwrod = parts[1].substring(parts[0].indexOf(":") + 1);
			messges = sheet.getRow(i).getCell(3).getStringCellValue();

			System.out.println("TestCase " + i + " : \n" + contentInit + "\nEo : " + messges);

			handllerTestCaseComment(username, passwrod, messges, i);

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
			System.out.println("Test result  : ");
			System.out.println("Test Succes  : " + numberTestSucces);
			System.out.println("Test fail : " + numberTestFail);
			System.out.println(" -------- -----------  -------- ");
			TimeUnit.SECONDS.sleep(5);
//			driver.quit();
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void handllerTestCaseComment(String username, String passwrod, String messges, int i)
			throws InterruptedException {
	//	System.out.println(driver.findElement(By.cssSelector("input[name=username]")));
		if (username.trim().length() > 0) {
			driver.findElement(By.cssSelector("input[name=username]")).sendKeys(username);
			TimeUnit.SECONDS.sleep(timeDelay);
		}

		if (passwrod.trim().length() > 0) {
			driver.findElement(By.cssSelector("input[name=password]")).sendKeys(passwrod);
			TimeUnit.SECONDS.sleep(timeDelay);
		}

		driver.findElement(By.cssSelector("button#btn-login")).click();
		
		// i = 1 la dong dau tien co pass dung va mat khau dung
		if (i == 1) {
			driver.findElement(By.cssSelector("i.mdi.mdi-power")).click(); // dang xuat ra de con test tiep
			numberTestSucces += 1;
			System.out.println("Passed");
		} else {
			// neu dang nhap khong thanh cong thi bat dau kiem tra cac messges co giong
			// trong kich ban
			TimeUnit.SECONDS.sleep(timeDelay);
			String messgesTestResult = null;
			try {
				messgesTestResult = driver.findElement(By.cssSelector(".sweet-alert.showSweetAlert.visible h2"))
						.getText();
				driver.get(driver.getCurrentUrl());

			} catch (Exception e) {
				messgesTestResult = driver.findElement(By.cssSelector(".alert.alert-danger")).getText();
				driver.get(driver.getCurrentUrl());
			}

			if (messgesTestResult.equals(messges)) { // so sanhh messsger ban dau va cai duoc tra ve
				System.out.println("Passed");
				numberTestSucces += 1;

			} else {
				numberTestFail += 1;
				System.out.println("Fail");
			}
		}

		TimeUnit.SECONDS.sleep(timeDelay);
	}

}
