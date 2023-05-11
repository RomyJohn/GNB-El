package service;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import utility.ConfigUtility;

import java.io.File;
import java.io.FileInputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

@Component
public class Login {

    @Autowired
    ConfigUtility configUtility;
    @Autowired
    Scrapping scrapping;
    @Autowired
    Template template;
    Logger logger = LoggerFactory.getLogger(Login.class);
    public static WebDriver driver;
    private FileInputStream fileIn;
    private XSSFWorkbook workbook;
    public static File file = null;

    public void getLoginCredentials() {
        try {
            WebDriverManager.chromedriver().setup();
            driver = new ChromeDriver();
            driver.get(configUtility.getProperty("WEBSITE_ADDRESS"));
            driver.manage().window().maximize();

            fileIn = new FileInputStream(configUtility.getProperty("LOGIN_CREDENTIALS_FILE"));
            workbook = new XSSFWorkbook(fileIn);
            XSSFSheet sheet = workbook.getSheetAt(0);

            int rowCount = 1;
            int totalRowCount = sheet.getLastRowNum();

            logger.info("checking");

            while (rowCount <= totalRowCount) {
                Row row = sheet.getRow(rowCount);
                Cell username = row.getCell(1);
                Cell password = row.getCell(2);
                startExecution(username.toString(), password.toString());
                rowCount++;
            }
            System.out.println("Reports Generated Of All Users");
        } catch (Exception exception) {
            System.out.println("@getLoginCredentials Exception = " + exception);

        } finally {
            try {
                fileIn.close();
                workbook.close();
                driver.close();
                driver.quit();
            } catch (Exception exception) {
                System.out.println(exception.getMessage());
            }
        }
    }

    public void startExecution(String username, String password) {
        try {
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofMinutes(2));

            //Click login button
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("LOGIN_BUTTON"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("LOGIN_BUTTON"))).click();

            //Enter username
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(configUtility.getProperty("USERNAME_INPUT"))));
            driver.findElement(By.id(configUtility.getProperty("USERNAME_INPUT"))).sendKeys(username);

            //Click continue button
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("CONTINUE_BUTTON_USERNAME"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("CONTINUE_BUTTON_USERNAME"))).click();

            //Click password checkbox
            wait.until(ExpectedConditions.elementToBeClickable(By.id(configUtility.getProperty("PASSWORD_CHECKBOX"))));
            driver.findElement(By.id(configUtility.getProperty("PASSWORD_CHECKBOX"))).click();

            //Enter password
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(configUtility.getProperty("PASSWORD_INPUT"))));
            driver.findElement(By.id(configUtility.getProperty("PASSWORD_INPUT"))).sendKeys(password);

            //Click continue button
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("CONTINUE_BUTTON_PASSWORD"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("CONTINUE_BUTTON_PASSWORD"))).click();

            try {
                //Login alert box, click login here
                Thread.sleep(2000);
                boolean loginHere = driver.findElement(By.cssSelector(configUtility.getProperty("LOGIN_ALERTBOX_BUTTON"))).isDisplayed();
                if (loginHere)
                    driver.findElement(By.cssSelector(configUtility.getProperty("LOGIN_ALERTBOX_BUTTON"))).click();
            } catch (Exception exception) {
                System.out.println("Info : Login alert box is not present.");
            }

            Thread.sleep(2000);
            ((JavascriptExecutor) driver).executeScript("window.scrollBy(0,-400);");

            //Hover services in nav main menu
            Actions action = new Actions(driver);
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("SERVICES_MAIN_MENU"))));
            WebElement navMainMenu = driver.findElement(By.cssSelector(configUtility.getProperty("SERVICES_MAIN_MENU")));
            action.moveToElement(navMainMenu).build().perform();

            //Click AIS in nav services sub menu
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(configUtility.getProperty("AIS_SUB_MENU"))));
            WebElement navSubMenu = driver.findElement(By.xpath(configUtility.getProperty("AIS_SUB_MENU")));
            action.moveToElement(navSubMenu).click().perform();

            //AIS alert box, click proceed
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("AIS_ALERTBOX_BUTTON"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("AIS_ALERTBOX_BUTTON"))).click();

            //To switch window to AIS tab
            Thread.sleep(2000);
            List<String> windowId = new ArrayList<>(driver.getWindowHandles());
            driver.switchTo().window(windowId.get(1));

            //Click AIS tab
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("AIS_BUTTON"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("AIS_BUTTON"))).click();

            file = new File("D:/" + username);
            file.mkdir();

            scrapping.startScrapping();
            //driver.navigate().back();
            //template.createTemplate();

            //To click AIS logout dropdown button
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("AIS_LOGOUT_DROPDOWN_BUTTON"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("AIS_LOGOUT_DROPDOWN_BUTTON"))).click();

            //To click AIS logout button
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("AIS_LOGOUT_BUTTON"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("AIS_LOGOUT_BUTTON"))).click();

            //To click AIS logout confirmation button
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath(configUtility.getProperty("AIS_LOGOUT_CONFIRMATION_BUTTON"))));
            driver.findElement(By.xpath(configUtility.getProperty("AIS_LOGOUT_CONFIRMATION_BUTTON"))).click();

            //To click AIS successfully logged out info button
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("AIS_LOGGED_OUT_INFORMATION_BUTTON"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("AIS_LOGGED_OUT_INFORMATION_BUTTON"))).click();

            //To switch window to main dashboard tab
            Thread.sleep(1000);
            windowId = new ArrayList<>(driver.getWindowHandles());
            driver.switchTo().window(windowId.get(0));

            //To click dashboard logout dropdown button
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("DASHBOARD_LOGOUT_DROPDOWN_BUTTON"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("DASHBOARD_LOGOUT_DROPDOWN_BUTTON"))).click();

            //To click dashboard logout button
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("DASHBOARD_LOGOUT_BUTTON"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("DASHBOARD_LOGOUT_BUTTON"))).click();

        } catch (Exception exception) {
            System.out.println("@startExecution Exception = " + exception);
        }
    }

}
