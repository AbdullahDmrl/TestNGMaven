package Utils;

import UITest_YourStore._02_POM;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeDriverService;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.io.FileHandler;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.*;
import org.testng.annotations.*;
import java.io.File;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

public class SingleDriver {
    public static WebDriver driver;
    public static WebDriverWait wait;

    @BeforeClass
    public void beforeClass()
    {
        Logger logger= Logger.getLogger("");
        logger.setLevel(Level.SEVERE);
        System.setProperty(ChromeDriverService.CHROME_DRIVER_SILENT_OUTPUT_PROPERTY, "true");
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");
        driver = new ChromeDriver(options);
        driver.manage().window().maximize(); // Ekranı max yapıyor.
        Duration dr=Duration.ofSeconds(30);
        driver.manage().timeouts().pageLoadTimeout(dr);
        driver.manage().timeouts().implicitlyWait(dr);
        wait=new WebDriverWait(driver, Duration.ofSeconds(30));
        driver.get("http://opencart.abstracta.us/index.php?route=account/login");
        LoginTest();
    }

    @DataProvider(name = "getDataFromExcel")
    public Object[][] getDataExcell() {
        List<List<String>> excellData= ExcelUtility.getlistData("src/test/ApachePOI/resource/searchTable.xlsx","Sheet1",1);

        Object[][] dataExcell=new Object[excellData.size()][excellData.get(0).size()];
        for (int i = 0; i < excellData.size(); i++) {
            for (int j = 0; j < excellData.get(i).size(); j++) {

                dataExcell[i][j]=excellData.get(i).get(j);
            }
        }
        return dataExcell;
    }


    @AfterMethod
    public void printError(ITestResult result) {
        LocalDateTime date=LocalDateTime.now();
        DateTimeFormatter formatter=DateTimeFormatter.ofPattern("dd.MM.yyyy");

        if (result.getStatus() == ITestResult.FAILURE) {
            String clsName = result.getTestClass().getName();
            clsName = clsName.substring(clsName.lastIndexOf(".") + 1);
            System.out.println("Test " + clsName + "." +
                    result.getName() + " FAILED");

            TakesScreenshot ts= (TakesScreenshot) driver;
            File screenShots=ts.getScreenshotAs(OutputType.FILE);
            try {
                FileHandler.copy(screenShots, new File("src/testResults/ScreenShots"+date.format(formatter)+".png"));

            } catch (IOException e) {
                e.printStackTrace();
            }
           }
       ExcelUtility.writeExcel("src/testResults/results.xlsx",result,"chrome",date.format(formatter));

    }

    @AfterClass
    public void afterClass() {

        driver.quit();
    }
    void LoginTest()
    {

        _02_POM elements=new _02_POM();
        elements.sendKeysFunction(elements.inputEmail,"jiyowir527@richdn.com" );
       elements.sendKeysFunction(elements.password,"12345");
       elements.clickFunction(elements.loginBtn);
    }


}
