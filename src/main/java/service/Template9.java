package service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
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
import java.util.Arrays;
import java.util.List;

@Component
public class Template9 {

    @Autowired
    ConfigUtility configUtility;
    private WebDriver driver;
    private WebDriverWait wait = null;
    private XSSFWorkbook workbook = null;
    private XSSFSheet sheet;
    private int dividend = 0;
    private int refund = 0;
    private int sb = 0;
    private int deposits = 0;

    public void createSheet() {
        try {
            this.driver = Login.driver;
            wait = new WebDriverWait(driver, Duration.ofMinutes(2));
            workbook = Template.workbook;
            sheet = workbook.createSheet("Income from Other Sources");

            //Click SFT information
            wait.until(ExpectedConditions.elementToBeClickable(By.id(configUtility.getProperty("SFT"))));
            driver.findElement(By.id(configUtility.getProperty("SFT"))).click();
            //iteratePartBMenu_SFT();

            //Click Other Information
            wait.until(ExpectedConditions.elementToBeClickable(By.id(configUtility.getProperty("OTHER_INFORMATION"))));
            driver.findElement(By.id(configUtility.getProperty("OTHER_INFORMATION"))).click();
            //iteratePartBMenu_OtherInfo();

            setExcelData(2, 2, Arrays.asList("Gross income chargeable to tax at normal applicable rates"));
            setExcelData(3, 0, Arrays.asList("a", "Dividends, Gross", "", "1a"), 3, 3, 1, 2);
            setExcelData(4, 0, Arrays.asList("ai", "Dividend income [other than (ii)]", "", "1ai", String.format("%,d", dividend)), 4, 4, 1, 2);
            setExcelData(5, 0, Arrays.asList("aii", "Dividend income u/s 2(22)e", "", "1aii"), 5, 5, 1, 2);
            setExcelData(6, 0, Arrays.asList("b", "Interest, Gross", "", "1b"), 6, 6, 1, 2);
            setExcelData(7, 1, Arrays.asList("i", "From Savings Bank", "1bi", String.format("%,d", sb)));
            setExcelData(8, 1, Arrays.asList("ii", "From Deposit (Bank/ Post Office/ Co-operative Society)", "1bii", String.format("%,d", deposits)));
            setExcelData(9, 1, Arrays.asList("iii", "From Interest on Income Tax refund", "1biii", String.format("%,d", refund)));
            setExcelData(10, 1, Arrays.asList("iv", "In the nature of  Pass through income\\loss", "1biv"));
            setExcelData(11, 1, Arrays.asList("v", "others", "1bv"));

        } catch (Exception exception) {
            System.out.println("@createSheet Exception = " + exception);
            exception.printStackTrace();
        }
    }

    public void iteratePartBMenu_SFT() {
        try {
            int rowCount = 15;
            int cellCount = 0;

            //To get partB main table thead columns
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_HEAD"))));
            List<WebElement> partBMainTableHeadColumns = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_HEAD")));
            //To get partB main table tbody rows
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY"))));
            List<WebElement> partBMainTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY")));

            int partBMainTableBodySize = 0;

            WebElement partBMainTablePaginationButton = null;
            try {
                //To get partB main table pagination button
                partBMainTablePaginationButton = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_PAGINATION_BUTTON")));
            } catch (Exception exception) {
            }
            if (partBMainTablePaginationButton != null && partBMainTablePaginationButton.isEnabled() && partBMainTablePaginationButton.isDisplayed()) {
                //To get partB main table total row count text
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_ROW_COUNT_TEXT"))));
                WebElement partBMainTableRowCount = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_ROW_COUNT_TEXT")));
                String[] totalRowCount = partBMainTableRowCount.getText().split(" ");
                partBMainTableBodySize = Integer.parseInt(totalRowCount[4]) * 2;
            } else {
                partBMainTableBodySize = partBMainTableBodyRows.size();
            }

            for (int i = 0; i < partBMainTableBodySize; i++) {
                if (i % 20 == 19) {
                    if (partBMainTablePaginationButton.isEnabled() && partBMainTablePaginationButton.isDisplayed()) {
                        wait.until(ExpectedConditions.elementToBeClickable(partBMainTablePaginationButton));
                        partBMainTablePaginationButton.click();
                        //To get partB main table tbody rows
                        wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY"))));
                        partBMainTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY")));
                    }
                }
                int remainder = i % 20;
                if (remainder % 2 == 0) {
                    wait.until(ExpectedConditions.elementToBeClickable(partBMainTableBodyRows.get(i % 20)));
                    List<WebElement> partBMainTableBodyColumns = partBMainTableBodyRows.get(i % 20).findElements(By.tagName("td"));
                    if (partBMainTableBodyColumns.get(2).getText().equals("SFT-015") || partBMainTableBodyColumns.get(2).getText().contains("SFT-016")) {
                        partBMainTableBodyRows.get(i % 20).click();
                        //To get partB sub table thead row
                        wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_HEAD_2"))));
                        WebElement partBSubTableHeadRow = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_HEAD_2")));
                        //To get partB sub table thead columns
                        List<WebElement> partBSubTableHeadColumns = partBSubTableHeadRow.findElements(By.tagName("th"));
                        setScrappedData(partBMainTableHeadColumns, partBSubTableHeadColumns, rowCount, cellCount);
                        cellCount = 0;
                        rowCount++;
                        while (true) {
                            //To get partB sub table tbody rows
                            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_BODY"))));
                            List<WebElement> partBSubTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_BODY")));
                            //To get partB sub table pagination button
                            WebElement partBSubTablePaginationButton = null;
                            try {
                                partBSubTablePaginationButton = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_PAGINATION_BUTTON")));
                            } catch (Exception exception) {
                            }
                            if (partBSubTablePaginationButton != null && partBSubTablePaginationButton.isDisplayed() && partBSubTablePaginationButton.isEnabled()) {
                                for (WebElement partBSubTableBodyRowsItem : partBSubTableBodyRows) {
                                    //To get partB sub table tbody columns
                                    List<WebElement> partBSubTableBodyColumns = partBSubTableBodyRowsItem.findElements(By.tagName("td"));
                                    if (partBSubTableBodyColumns.get(partBSubTableBodyColumns.size() - 2).getText().equals("Active")) {
                                        setScrappedData(partBMainTableBodyColumns, partBSubTableBodyColumns, rowCount, cellCount);
                                        cellCount = 0;
                                        rowCount++;
                                    }
                                }
                                wait.until(ExpectedConditions.elementToBeClickable(partBSubTablePaginationButton));
                                partBSubTablePaginationButton.click();
                            } else {
                                for (WebElement partBSubTableBodyRowsItem : partBSubTableBodyRows) {
                                    //To get partB sub table tbody columns
                                    List<WebElement> partBSubTableBodyColumns = partBSubTableBodyRowsItem.findElements(By.tagName("td"));
                                    if (partBSubTableBodyColumns.get(partBSubTableBodyColumns.size() - 2).getText().equals("Active")) {
                                        setScrappedData(partBMainTableBodyColumns, partBSubTableBodyColumns, rowCount, cellCount);
                                        cellCount = 0;
                                        rowCount++;
                                    }
                                }
                                break;
                            }
                        }
                        sheet.createRow(rowCount++);
                    }
                }
            }
        } catch (Exception exception) {
            System.out.println("@iteratePartBMenu Exception = " + exception);
            exception.printStackTrace();
        }
    }

    public void iteratePartBMenu_OtherInfo() {
        try {
            List<WebElement> partBMainTableBodyRows = null;
            try {
                Thread.sleep(2000);
                driver.findElement(By.cssSelector(configUtility.getProperty("TAXES_PAID_NO_TAX_AVAILABLE_PARAGRAPH"))).isDisplayed();
            } catch (Exception exception) {
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_ROW"))));
                partBMainTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_ROW")));
            }
            if (partBMainTableBodyRows != null) {
                for (int i = 0; i < partBMainTableBodyRows.size(); i++) {
                    if (i % 2 == 0) {
                        List<WebElement> partBMainTableBodyColumns = partBMainTableBodyRows.get(i).findElements(By.tagName("td"));
                        if (partBMainTableBodyColumns.get(1).getText().equals("")) {

                        }
                    }
                }
            }
        } catch (Exception exception) {
            System.out.println("@iteratePartBMenu Exception = " + exception);
            exception.printStackTrace();
        }
    }

    public void setScrappedData(List<WebElement> mainTableColumns, List<WebElement> subTableColumns, int rowCount, int cellCount) {
        Row row = sheet.createRow(rowCount);
        Cell cell = null;
        for (int mt = 0; mt < mainTableColumns.size() - 3; mt++) {
            sheet.autoSizeColumn(mt);
            cell = row.createCell(mt);
            cell.setCellValue(mainTableColumns.get(mt).getText());
            cellCount++;
        }
        sheet.autoSizeColumn(cellCount);
        cell = row.createCell(cellCount);
        cell.setCellValue(mainTableColumns.get(mainTableColumns.size() - 1).getText());
        cellCount++;
        for (int st = 0; st < subTableColumns.size() - 2; st++) {
            if (st == 0) {
                System.out.println("Data::" + mainTableColumns.get(2).getText());
                if (mainTableColumns.get(2).getText().equals("SFT-015"))
                    dividend += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
                if (mainTableColumns.get(2).getText().equals("SFT-016(SB)"))
                    sb += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
                if (mainTableColumns.get(2).getText().equals("SFT-016(FD)") || mainTableColumns.get(2).getText().equals("SFT-016(RD)") || mainTableColumns.get(2).getText().equals("SFT-016(TD)"))
                    deposits += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
                if (mainTableColumns.get(2).getText().equals("INT-REF-001"))
                    refund += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
                sheet.autoSizeColumn(0);
                cell = row.createCell(0);
            } else {
                sheet.autoSizeColumn(cellCount - 1);
                cell = row.createCell(cellCount - 1);
            }
            cell.setCellValue(subTableColumns.get(st).getText());
            cellCount++;
        }
    }

    public void setExcelData(int rowNumber, int column, List list, int... merge) {
        if (merge.length != 0)
            sheet.addMergedRegion(new CellRangeAddress(merge[0], merge[1], merge[2], merge[3]));
        Row row = sheet.createRow(rowNumber);
        Cell cell = null;
        for (int i = 0; i < list.size(); i++) {
            sheet.autoSizeColumn(column + i);
            cell = row.createCell(column + i);
            cell.setCellValue(list.get(i).toString());
            if (rowNumber == 2) {
                CellStyle cellStyle = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyle.setFont(font);
                cell.setCellStyle(cellStyle);
            }
        }
    }

}
