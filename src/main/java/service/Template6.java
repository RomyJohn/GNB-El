package service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import utility.ConfigUtility;

import java.time.Duration;
import java.util.List;

@Component
public class Template6 {

    @Autowired
    ConfigUtility configUtility;
    private WebDriver driver;
    private WebDriverWait wait;
    private XSSFSheet sheet;

    public void createSheet() {
        try {
            this.driver = Login.driver;
            wait = new WebDriverWait(driver, Duration.ofMinutes(2));
            XSSFWorkbook workbook = Template.workbook;
            sheet = workbook.createSheet("Taxes Paid");

            //Click Payment Of Taxes
            wait.until(ExpectedConditions.elementToBeClickable(By.id(configUtility.getProperty("PAYMENT_TAXES"))));
            driver.findElement(By.id(configUtility.getProperty("PAYMENT_TAXES"))).click();
            iteratePartBMenu();

        } catch (Exception exception) {
            System.out.println("@createSheet Exception = " + exception);
        }
    }

    public void iteratePartBMenu() {
        try {
            int rowCount = 0;

            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_HEADER"))));
            WebElement partBMainTableHeadRows = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_HEADER")));
            List<WebElement> partBMainTableHeadColumns = partBMainTableHeadRows.findElements(By.tagName("th"));
            List<WebElement> partBMainTableBodyRows = null;
            try {
                Thread.sleep(2000);
                driver.findElement(By.cssSelector(configUtility.getProperty("TAXES_PAID_NO_TAX_AVAILABLE_PARAGRAPH"))).isDisplayed();
            } catch (Exception exception) {
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_ROW"))));
                partBMainTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_ROW")));
            }
            if (partBMainTableBodyRows != null) {
                setExcelData(partBMainTableHeadColumns, rowCount++);
                for (int i = 0; i < partBMainTableBodyRows.size(); i++) {
                    List<WebElement> partBMainTableBodyColumns = partBMainTableBodyRows.get(i).findElements(By.tagName("td"));
                    setExcelData(partBMainTableBodyColumns, rowCount++);
                }
            }
        } catch (Exception exception) {
            System.out.println("@iteratePartBMenu Exception = " + exception);
        }
    }

    public void setExcelData(List<WebElement> tableColumns, int rowCount) {
        int cellCount = 0;
        Row row = sheet.createRow(rowCount);
        Cell cell = null;
        for (int st = 0; st < tableColumns.size(); st++) {
            sheet.autoSizeColumn(cellCount);
            cell = row.createCell(cellCount);
            cell.setCellValue(tableColumns.get(st).getText());
            cellCount++;
        }
    }

}
