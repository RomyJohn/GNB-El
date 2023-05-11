package service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
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

import java.io.File;
import java.io.FileOutputStream;
import java.time.Duration;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

@Component
public class Scrapping {

    @Autowired
    ConfigUtility configUtility;
    private WebDriver driver;
    private WebDriverWait wait;
    private XSSFWorkbook workbook_tds, workbook_sft, workbook_other;
    public static XSSFSheet sheet_tds, sheet_sft, sheet_other;
    private int tds, sft, other, sessionCount, rowCount, fileCount;
    private ScheduledExecutorService executor;

    public void startScrapping() {
        try {
            driver = Login.driver;
            wait = new WebDriverWait(driver, Duration.ofMinutes(2));

            workbook_tds = new XSSFWorkbook();
            workbook_sft = new XSSFWorkbook();
            workbook_other = new XSSFWorkbook();

            sheet_tds = workbook_tds.createSheet("Data");
            sheet_sft = workbook_sft.createSheet("Data");
            sheet_other = workbook_other.createSheet("Data");

            tds = 0;
            sft = 0;
            other = 0;
            sessionCount = 0;

            //Click TIS section
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("TIS_SECTION"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("TIS_SECTION"))).click();

            Thread.sleep(1000);
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"dropdownMenuButton\"]")));
            driver.findElement(By.xpath("//*[@id=\"dropdownMenuButton\"]")).click();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"content-page\"]/div/app-ais-summary/div/div/div[1]/div[2]/app-dropdown-btn/div/div/button[2]")));
            driver.findElement(By.xpath("//*[@id=\"content-page\"]/div/app-ais-summary/div/div/div[1]/div[2]/app-dropdown-btn/div/div/button[2]")).click();
            Thread.sleep(2000);

            executor = Executors.newScheduledThreadPool(1);
            Runnable task = () -> {
                try {
                    wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#mat-dialog-" + sessionCount + configUtility.getProperty("CONTINUE_SESSION_BUTTON"))));
                    driver.findElement(By.cssSelector("#mat-dialog-" + sessionCount + configUtility.getProperty("CONTINUE_SESSION_BUTTON"))).click();
                    sessionCount++;
                    Thread.sleep(1000);
                    System.out.println("Session Continued");
                } catch (Exception exception) {
                    System.out.println("Info : Continue session alert box is not present.");
                }
            };
            executor.scheduleWithFixedDelay(task, 18, 18, TimeUnit.MINUTES);

            iterateAccordion();

        } catch (Exception exception) {
            System.out.println("@startScrapping Exception = " + exception);
        }
    }

    public void iterateAccordion() {
        try {
            int accordionCurrentSize = 0;
            int accordionSubMenuCurrentSize = 0;
            int accordionSize = 1;
            int accordionSubMenuSize = 1;
            int accordionSubMenuPaginationSize = 1;
            int accordionSubMenuPaginationCurrentSize = 0;

            File file = new File(Login.file.getPath() + File.separator + "Scrapped Data");
            file.mkdir();
            File file1 = null;

            while (true) {
                //To get accordion list
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(configUtility.getProperty("ACCORDION"))));
                List<WebElement> accordion = driver.findElements(By.id(configUtility.getProperty("ACCORDION")));
                if (!accordion.isEmpty() && accordion.size() != 0) {
                    if ((accordionSubMenuCurrentSize < accordionSubMenuSize) && (accordionSubMenuCurrentSize != 0) && (accordionSubMenuSize > 1)) {
                        wait.until(ExpectedConditions.elementToBeClickable(accordion.get(accordionCurrentSize - 1)));
                        accordion.get(accordionCurrentSize - 1).click();
                        wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("ACCORDION_SUBMENU_1") + accordionCurrentSize + configUtility.getProperty("ACCORDION_SUBMENU_2"))));
                        List<WebElement> accordionSubMenu = driver.findElements(By.cssSelector(configUtility.getProperty("ACCORDION_SUBMENU_1") + accordionCurrentSize + configUtility.getProperty("ACCORDION_SUBMENU_2")));
                        if (accordionSubMenuSize > 10 && accordionSubMenuCurrentSize > 9) {
                            //To get accordion submenu pagination button
                            WebElement accordionSubMenuPaginationButton = driver.findElement(By.cssSelector("div[id='collapse" + (accordionCurrentSize - 1) + configUtility.getProperty("ACCORDION_SUBMENU_PAGINATION_BUTTON")));
                            if (accordionSubMenuPaginationButton.isDisplayed() && accordionSubMenuPaginationButton.isEnabled()) {
                                accordionSubMenuPaginationSize = (int) Math.floor(accordionSubMenuSize / 10);
                                if (accordionSubMenuPaginationCurrentSize < accordionSubMenuPaginationSize) {
                                    for (int i = 0; i <= accordionSubMenuPaginationCurrentSize; i++) {
                                        wait.until(ExpectedConditions.elementToBeClickable(accordionSubMenuPaginationButton));
                                        accordionSubMenuPaginationButton.click();
                                    }
                                    //To get accordion submenu list
                                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("ACCORDION_SUBMENU_1") + accordionCurrentSize + configUtility.getProperty("ACCORDION_SUBMENU_2"))));
                                    accordionSubMenu = driver.findElements(By.cssSelector(configUtility.getProperty("ACCORDION_SUBMENU_1") + accordionCurrentSize + configUtility.getProperty("ACCORDION_SUBMENU_2")));
                                }
                                if (accordionSubMenuCurrentSize % 10 == 9) {
                                    accordionSubMenuPaginationCurrentSize++;
                                }
                            }
                        }
                        int accordionSubMenuIndex = accordionSubMenuCurrentSize % 10;
                        checkMenuType(accordionSubMenu, accordionSubMenuIndex, file1);
                        accordionSubMenuCurrentSize++;
                    } else {
                        accordionSubMenuCurrentSize = 0;
                        accordionSubMenuPaginationCurrentSize = 0;
                        if (accordionCurrentSize < accordionSize) {
                            String accordionTitle = accordion.get(accordionCurrentSize).getText().replaceAll("\\d", "").replaceAll(",", "").replaceAll("/", " - ").trim();
                            fileCount = 1;
                            file1 = new File(file.getPath() + File.separator + accordionTitle);
                            file1.mkdir();
                            accordionSize = accordion.size();
                            wait.until(ExpectedConditions.elementToBeClickable(accordion.get(accordionCurrentSize)));
                            accordion.get(accordionCurrentSize).click();
                            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("ACCORDION_SUBMENU_1") + (accordionCurrentSize + 1) + configUtility.getProperty("ACCORDION_SUBMENU_2"))));
                            List<WebElement> accordionSubMenu = driver.findElements(By.cssSelector(configUtility.getProperty("ACCORDION_SUBMENU_1") + (accordionCurrentSize + 1) + configUtility.getProperty("ACCORDION_SUBMENU_2")));
                            WebElement accordionSubMenuPaginationButton = null;
                            try {
                                //To get accordion submenu pagination button
                                accordionSubMenuPaginationButton = driver.findElement(By.cssSelector("div[id='collapse" + accordionCurrentSize + configUtility.getProperty("ACCORDION_SUBMENU_PAGINATION_BUTTON")));
                            } catch (Exception exception) {
                                System.out.println("Info : Pagination button is not present.");
                            }
                            if (accordionSubMenuPaginationButton != null && accordionSubMenuPaginationButton.isDisplayed() && accordionSubMenuPaginationButton.isEnabled()) {
                                //To get accordion submenu total row count text
                                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("ACCORDION_SUBMENU_1") + (accordionCurrentSize + 1) + configUtility.getProperty("ACCORDION_SUBMENU_ROW_COUNT_TEXT"))));
                                WebElement accordionSubMenuRowCount = driver.findElement(By.cssSelector(configUtility.getProperty("ACCORDION_SUBMENU_1") + (accordionCurrentSize + 1) + configUtility.getProperty("ACCORDION_SUBMENU_ROW_COUNT_TEXT")));
                                String[] totalRowCount = accordionSubMenuRowCount.getText().split(" ");
                                accordionSubMenuSize = Integer.parseInt(totalRowCount[4]);
                            } else {
                                accordionSubMenuSize = accordionSubMenu.size();
                            }
                            checkMenuType(accordionSubMenu, accordionSubMenuCurrentSize, file1);
                            accordionSubMenuCurrentSize++;
                            accordionCurrentSize++;
                        } else {
                            workbook_tds.close();
                            workbook_sft.close();
                            workbook_other.close();
                            if (!executor.isShutdown())
                                executor.shutdown();
                            System.out.println("Scrapping Completed");
                            break;
                        }
                    }
                }
            }
        } catch (Exception exception) {
            System.out.println("@iterateAccordion Exception = " + exception);
        }
    }

    public void checkMenuType(List<WebElement> accordionSubMenu, int index, File file) {
        try {
            List<WebElement> accordionSubMenuColumns = accordionSubMenu.get(index).findElements(By.tagName("td"));
            String tabMenuText = accordionSubMenuColumns.get(0).getText();

            String filename = file.getPath() + File.separator + "CustomersDetail" + (fileCount++) + ".csv";

            if (tabMenuText.equals("TDS/TCS") && tds == 0) {
                clickAccordionSubMenu(accordionSubMenu, index, filename, "tds", workbook_tds, true);
                tds++;
            } else if (tabMenuText.equals("TDS/TCS") && tds != 0) {
                clickAccordionSubMenu(accordionSubMenu, index, filename, "tds", workbook_tds, false);
            } else if (tabMenuText.equals("SFT") && sft == 0) {
                clickAccordionSubMenu(accordionSubMenu, index, filename, "sft", workbook_sft, true);
                sft++;
            } else if (tabMenuText.equals("SFT") && sft != 0) {
                clickAccordionSubMenu(accordionSubMenu, index, filename, "sft", workbook_sft, false);
            } else if (tabMenuText.equals("Other") && other == 0) {
                clickAccordionSubMenu(accordionSubMenu, index, filename, "other", workbook_other, true);
                other++;
            } else if (tabMenuText.equals("Other") && other != 0) {
                clickAccordionSubMenu(accordionSubMenu, index, filename, "other", workbook_other, false);
            }

        } catch (Exception exception) {
            System.out.println("@checkMenuType Exception = " + exception);
        }
    }

    public void clickAccordionSubMenu(List<WebElement> accordionSubMenu, int accordionSubMenuCurrentSize, String filename, String menuType, Workbook workbook, boolean iteration) {
        try {
            wait.until(ExpectedConditions.elementToBeClickable(accordionSubMenu.get(accordionSubMenuCurrentSize)));
            accordionSubMenu.get(accordionSubMenuCurrentSize).click();

            if (iteration)
                iteratePartBMenu(menuType);

            FileOutputStream fileOut = new FileOutputStream(filename);
            workbook.write(fileOut);
            fileOut.close();

            driver.navigate().back();
            //Click TIS section
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(configUtility.getProperty("TIS_SECTION"))));
            driver.findElement(By.cssSelector(configUtility.getProperty("TIS_SECTION"))).click();
        } catch (Exception exception) {
            System.out.println("@clickAccordionSubMenu Exception = " + exception);
        }
    }

    public void iteratePartBMenu(String menuType) {
        try {
            int partBMainTableBodySize = 0;
            rowCount = 0;
            Thread.sleep(1000);

            List<WebElement> partBMainTableHeadColumns = null;
            List<WebElement> partBMainTableBodyRows = null;
            if (menuType.equals("sft") || menuType.equals("tds")) {
                //To get partB main table thead columns
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_HEAD"))));
                partBMainTableHeadColumns = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_HEAD")));
                //To get partB main table tbody rows
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY"))));
                partBMainTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY")));
            } else {
                //To get partB main table thead columns
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_HEAD_OTHER"))));
                partBMainTableHeadColumns = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_HEAD_OTHER")));
                //To get partB main table tbody rows
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY_OTHER"))));
                partBMainTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY_OTHER")));
            }

            WebElement partBMainTablePaginationButton = null;
            try {
                if (menuType.equals("sft") || menuType.equals("tds")) {
                    //To get partB main table pagination button
                    partBMainTablePaginationButton = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_PAGINATION_BUTTON")));
                } else {
                    //To get partB main table pagination button
                    partBMainTablePaginationButton = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_PAGINATION_BUTTON_OTHER")));
                }
            } catch (Exception exception) {
                System.out.println("Info : Pagination button is not present.");
            }
            if (partBMainTablePaginationButton != null && partBMainTablePaginationButton.isEnabled() && partBMainTablePaginationButton.isDisplayed()) {
                WebElement partBMainTableRowCount = null;
                if (menuType.equals("sft") || menuType.equals("tds")) {
                    //To get partB main table total row count text
                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_ROW_COUNT_TEXT"))));
                    partBMainTableRowCount = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_ROW_COUNT_TEXT")));
                } else {
                    //To get partB main table total row count text
                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_ROW_COUNT_TEXT_OTHER"))));
                    partBMainTableRowCount = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_ROW_COUNT_TEXT_OTHER")));
                }
                String[] totalRowCount = partBMainTableRowCount.getText().split(" ");
                partBMainTableBodySize = Integer.parseInt(totalRowCount[4]) * 2;
            } else {
                partBMainTableBodySize = partBMainTableBodyRows.size();
            }

            if (!partBMainTableBodyRows.isEmpty() && partBMainTableBodyRows.size() != 0) {
                for (int i = 0; i < partBMainTableBodySize; i++) {
                    if (i % 20 == 19) {
                        if (partBMainTablePaginationButton.isEnabled() && partBMainTablePaginationButton.isDisplayed()) {
                            wait.until(ExpectedConditions.elementToBeClickable(partBMainTablePaginationButton));
                            partBMainTablePaginationButton.click();
                            if (menuType.equals("sft") || menuType.equals("tds")) {
                                //To get partB main table tbody rows
                                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY"))));
                                partBMainTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY")));
                            } else {
                                //To get partB main table tbody rows
                                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY_OTHER"))));
                                partBMainTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_MAIN_TABLE_BODY_OTHER")));
                            }
                        }
                    }
                    int remainder = i % 20;
                    if (remainder % 2 == 0) {
                        wait.until(ExpectedConditions.elementToBeClickable(partBMainTableBodyRows.get(i % 20)));
                        partBMainTableBodyRows.get(i % 20).click();
                        if (menuType.equals("tds"))
                            scrapData(partBMainTableBodyRows, partBMainTableHeadColumns, i, menuType, sheet_tds);
                        else if (menuType.equals("sft"))
                            scrapData(partBMainTableBodyRows, partBMainTableHeadColumns, i, menuType, sheet_sft);
                        else if (menuType.equals("other"))
                            scrapData(partBMainTableBodyRows, partBMainTableHeadColumns, i, menuType, sheet_other);
                    }
                }
            }
        } catch (Exception exception) {
            System.out.println("@iteratePartBMenu Exception = " + exception);
        }
    }

    public void scrapData(List<WebElement> partBMainTableBodyRows, List<WebElement> partBMainTableHeadColumns, int i, String menuType, XSSFSheet sheet) {
        try {
            //To get partB main table tbody columns
            List<WebElement> partBMainTableBodyColumns = partBMainTableBodyRows.get(i % 20).findElements(By.tagName("td"));
            WebElement partBSubTableHeadRow = null;
            if (menuType.equals("sft") || menuType.equals("tds")) {
                //To get partB sub table thead row
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_HEAD_2"))));
                partBSubTableHeadRow = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_HEAD_2")));
            } else {
                //To get partB sub table thead row
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_OTHER_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_HEAD_OTHER_2"))));
                partBSubTableHeadRow = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_OTHER_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_HEAD_OTHER_2")));
            }

            //To get partB sub table thead columns
            List<WebElement> partBSubTableHeadColumns = partBSubTableHeadRow.findElements(By.tagName("th"));
            setExcelData(partBMainTableHeadColumns, partBSubTableHeadColumns, rowCount, 0, sheet);
            rowCount++;

            while (true) {
                List<WebElement> partBSubTableBodyRows = null;
                if (menuType.equals("sft") || menuType.equals("tds")) {
                    //To get partB sub table tbody rows
                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_BODY"))));
                    partBSubTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_BODY")));
                } else {
                    //To get partB sub table tbody rows
                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_OTHER_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_BODY_OTHER"))));
                    partBSubTableBodyRows = driver.findElements(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_OTHER_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_BODY_OTHER")));
                }
                if (!partBSubTableBodyRows.isEmpty() && partBSubTableBodyRows.size() != 0) {
                    //To get partB sub table pagination button
                    WebElement partBSubTablePaginationButton = null;
                    try {
                        if (menuType.equals("sft") || menuType.equals("tds"))
                            partBSubTablePaginationButton = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_PAGINATION_BUTTON")));
                        else
                            partBSubTablePaginationButton = driver.findElement(By.cssSelector(configUtility.getProperty("PARTB_SUB_TABLE_HEAD_OTHER_1") + ((i % 20) + 2) + configUtility.getProperty("PARTB_SUB_TABLE_OTHER_PAGINATION_BUTTON")));
                    } catch (Exception exception) {
                        System.out.println("Info : Pagination button is not present.");
                    }
                    if (partBSubTablePaginationButton != null && partBSubTablePaginationButton.isDisplayed() && partBSubTablePaginationButton.isEnabled()) {
                        for (WebElement partBSubTableBodyRowsItem : partBSubTableBodyRows) {
                            //To get partB sub table tbody columns
                            List<WebElement> partBSubTableBodyColumns = partBSubTableBodyRowsItem.findElements(By.tagName("td"));
                            setExcelData(partBMainTableBodyColumns, partBSubTableBodyColumns, rowCount, 0, sheet);
                            rowCount++;
                        }
                        wait.until(ExpectedConditions.elementToBeClickable(partBSubTablePaginationButton));
                        partBSubTablePaginationButton.click();
                    } else {
                        for (WebElement partBSubTableBodyRowsItem : partBSubTableBodyRows) {
                            //To get partB sub table tbody columns
                            List<WebElement> partBSubTableBodyColumns = partBSubTableBodyRowsItem.findElements(By.tagName("td"));
                            setExcelData(partBMainTableBodyColumns, partBSubTableBodyColumns, rowCount, 0, sheet);
                            rowCount++;
                        }
                        break;
                    }
                }
            }
            sheet.createRow(rowCount++);
        } catch (Exception exception) {
            System.out.println("@scrapData Exception = " + exception);
        }
    }

    public void setExcelData(List<WebElement> mainTableColumns, List<WebElement> subTableColumns, int rowCount, int cellCount, XSSFSheet sheet) {
        try {
            Row row = sheet.createRow(rowCount);
            Cell cell = null;
            for (int mt = 0; mt < mainTableColumns.size(); mt++) {
                sheet.autoSizeColumn(mt);
                cell = row.createCell(mt);
                cell.setCellValue(mainTableColumns.get(mt).getText());
                cellCount++;
            }
            for (int st = 0; st < subTableColumns.size(); st++) {
                if (st == 0) {
                    sheet.autoSizeColumn(0);
                    cell = row.createCell(0);
                } else {
                    sheet.autoSizeColumn(cellCount - 1);
                    cell = row.createCell(cellCount - 1);
                }
                cell.setCellValue(subTableColumns.get(st).getText());
                cellCount++;
            }
        } catch (Exception exception) {
            System.out.println("@setExcelData Exception = " + exception);
        }
    }

}
