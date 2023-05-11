package service;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import utility.ConfigUtility;

import java.io.File;
import java.io.FileOutputStream;
import java.time.Duration;

@Component
public class Template {

    @Autowired
    ConfigUtility configUtility;
    @Autowired
    Template1 template1;
    @Autowired
    Template2 template2;
    @Autowired
    Template3 template3;
    @Autowired
    Template4 template4;
    @Autowired
    Template5 template5;
    @Autowired
    Template6 template6;
    @Autowired
    Template7 template7;
    @Autowired
    Template8 template8;
    @Autowired
    Template9 template9;

    public static XSSFWorkbook workbook;

    public void createTemplate() {
        try {
            WebDriver driver = Login.driver;
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofMinutes(2));
            workbook = new XSSFWorkbook();

            //Click AIS section - TDS/TCS information
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("AIS_SECTION"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("AIS_SECTION"))).click();


            Thread.sleep(1000);


            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"dropdownMenuButton\"]")));
            driver.findElement(By.xpath("//*[@id=\"dropdownMenuButton\"]")).click();

            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"content-page\"]/div/app-ais-form26/div/div/div[1]/div[2]/app-dropdown-btn/div/div/button[2]")));
            driver.findElement(By.xpath("//*[@id=\"content-page\"]/div/app-ais-form26/div/div/div[1]/div[2]/app-dropdown-btn/div/div/button[2]")).click();

            Thread.sleep(2000);


            File file = new File(Login.file.getPath() + File.separator + "Template");
            file.mkdir();


            FileOutputStream fileOut = new FileOutputStream(file.getPath() + File.separator + "Template.xlsx");
            //template1.createSheet();
            //--template2.createSheet();
            //--template3.createSheet();
            //--template4.createSheet();
            //--template5.createSheet();
            //--template6.createSheet();
            //--template7.createSheet();
            template8.createSheet();
            //template9.createSheet();
            workbook.write(fileOut);
            workbook.close();
            fileOut.close();

            System.out.println("Templates created successfully");

        } catch (Exception exception) {
            System.out.println("@createTemplate Exception = " + exception);
            exception.printStackTrace();
        }
    }

}
