package service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import utility.ConfigUtility;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.Arrays;
import java.util.List;

@Component
public class Template8 {

    @Autowired
    ConfigUtility configUtility;
    private WebDriver driver;
    private WebDriverWait wait;

    private XSSFSheet sheet;

    private boolean isDataPresent = false;

    int tdscAmount = 0;
    int tdsjAmount = 0;
    int tdsjaAmount = 0;
    int tdsjbAmount = 0;

    public void createSheet(List<WebElement> subTableColumns) {
        try {
            XSSFWorkbook workbook = Template.workbook;
            sheet = workbook.createSheet("PGBP");

            FileInputStream fileIn = new FileInputStream("D:/ABAPP0998H/Scrapped Data/Salary/CustomersDetail3.csv");
            Workbook workbook1 = new XSSFWorkbook(fileIn);
            XSSFSheet sheet_sft = (XSSFSheet) workbook1.getSheetAt(0);

            for (int i = 0; i <= sheet_sft.getLastRowNum(); i++) {
                Row row = sheet_sft.getRow(i);
                if (row.getFirstCellNum() != -1) {

                    if (row.getCell(2).toString().equals("TDS-194C"))
                        tdscAmount += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
                    if (row.getCell(2).equals("TDS-194J"))
                        tdsjAmount += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
                    if (row.getCell(2).toString().equals("TDS-194JA"))
                        tdsjaAmount += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
                    if (row.getCell(2).equals("TDS-194JB"))
                        tdsjbAmount += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));

                }
            }


            int totalAMount = tdsjAmount + tdsjaAmount + tdsjbAmount;

            drawBox(3, 22, 0, 5);

            sheet.addMergedRegion(new CellRangeAddress(3, 3, 1, 3));
            setExcelData(3, 0, Arrays.asList(" i ", " For assessee carrying on Business"), 3, 9, 0, 0);
            sheet.addMergedRegion(new CellRangeAddress(4, 4, 2, 3));
            setExcelData(4, 0, Arrays.asList("", " a", " Gross turnover/Gross receipts (a1+a2)"), 4, 6, 1, 1);
            setExcelData(5, 0, Arrays.asList("", "", " 1", " Through a/c payee cheque or a/c payee bank draft or bank electronic clearing system received or other prescribed electronic modes before specified date", String.format("%,d", tdscAmount), tdscAmount > 20000000 ? "Gross Receipt is more than 2 crore" : ""));
            setExcelData(6, 0, Arrays.asList("", "", " 2", " Any other mode", tdscAmount > 0 ? "0" : String.format("%,d", tdscAmount * 8 / 100)));
            setExcelData(7, 0, Arrays.asList("", " b", " Gross profit", "", tdscAmount > 0 ? String.format("%,d", tdscAmount) : String.format("%,d", tdscAmount + (tdscAmount * 8 / 100))), 7, 7, 2, 3);
            setExcelData(8, 0, Arrays.asList("", " c", " Expenses"), 8, 8, 2, 3);
            setExcelData(9, 0, Arrays.asList("", " d", " Net profit"), 9, 9, 2, 3);
            sheet.addMergedRegion(new CellRangeAddress(10, 10, 1, 3));
            setExcelData(10, 0, Arrays.asList(" ii ", " For assessee carrying on Profession"), 10, 16, 0, 0);
            sheet.addMergedRegion(new CellRangeAddress(11, 11, 2, 3));
            setExcelData(11, 0, Arrays.asList("", " a", "  Gross receipts (a1 + a2)"), 11, 13, 1, 1);
            setExcelData(12, 0, Arrays.asList("", "", " 1", " Through a/c payee cheque or a/c payee bank draft or bank electronic clearing system received or other prescribed electronic modes before specified date", String.format("%,d", totalAMount), totalAMount > 5000000 ? "Gross Receipt is more than 50 lakhs" : ""));
            setExcelData(13, 0, Arrays.asList("", "", " 2", " Any other mode", totalAMount > 0 ? "0" : String.format("%,d", totalAMount * 8 / 100)));
            setExcelData(14, 0, Arrays.asList("", " b", " Gross profit", "", totalAMount > 0 ? String.format("%,d", totalAMount) : String.format("%,d", totalAMount + (totalAMount * 8 / 100))), 14, 14, 2, 3);
            setExcelData(15, 0, Arrays.asList("", " c", " Expenses"), 15, 15, 2, 3);
            setExcelData(16, 0, Arrays.asList("", " d", " Net profit"), 16, 16, 2, 3);
            setExcelData(17, 0, Arrays.asList(" iii ", " Total Profit (64(i)d+ 64(ii)d)"), 17, 17, 1, 3);
            setExcelData(18, 0, Arrays.asList(" iv ", " Turnover From Speculative Activity"), 18, 18, 1, 3);
            setExcelData(19, 0, Arrays.asList(" v ", " Gross Profit"), 19, 19, 1, 3);
            setExcelData(20, 0, Arrays.asList(" vi ", " Expenditure, if any"), 20, 20, 1, 3);
            setExcelData(21, 0, Arrays.asList(" vii ", " Net Income From Speculative Activity (65ii-65iii)"), 21, 21, 1, 3);
        } catch (
                Exception exception) {
            System.out.println("@createSheet Exception = " + exception);
            exception.printStackTrace();
        }

    }

    public void iteratePartBMenu() {
        try {
            int rowCount = 25;
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
                    if (partBMainTableBodyColumns.get(2).getText().equals("TDS-194J") || partBMainTableBodyColumns.get(2).getText().equals("TDS-194JB") || partBMainTableBodyColumns.get(2).getText().equals("TDS-194C")) {
                        isDataPresent = true;
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
                if (mainTableColumns.get(2).getText().equals("TDS-194J"))
                    tdsjAmount += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
                if (mainTableColumns.get(2).getText().equals("TDS-194JB"))
                    tdsjbAmount += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
                if (mainTableColumns.get(2).getText().equals("TDS-194C"))
                    tdscAmount += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
                if (mainTableColumns.get(2).getText().equals("TDS-194JA"))
                    tdsjaAmount += Integer.parseInt(subTableColumns.get(subTableColumns.size() - 3).getText().replace(",", ""));
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
            if (!isDataPresent)
                sheet.autoSizeColumn(column + i);
            cell = row.createCell(column + i);
            cell.setCellValue(list.get(i).toString());
        }
    }

    public void drawBox(int row1, int row2, int column1, int column2) {
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, column1, row1, column2, row2);
        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
        XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
        shape.setShapeType(ShapeTypes.RECT);
        shape.setLineWidth(1);
        shape.setLineStyleColor(0, 0, 0);
    }

}
