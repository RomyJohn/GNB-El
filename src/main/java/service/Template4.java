package service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.Arrays;
import java.util.List;

@Component
public class Template4 {

    private XSSFSheet sheet;

    public void createSheet() {
        try {
            XSSFWorkbook workbook = Template.workbook;
            sheet = workbook.createSheet("Schedule Salary");
            XSSFSheet sheet_sft = Scrapping.sheet_sft;

            int totalGrossSalary = 0;
            int rowCount = 1;
            int employerCount = 1;

            for (int i = 0; i <= sheet_sft.getLastRowNum(); i++) {
                Row row = sheet_sft.getRow(i);
                if (row.getFirstCellNum() != -1) {
                    if (row.getCell(1).toString().equals("Salary")) {
                        StringBuilder employerInfoText = new StringBuilder(row.getCell(4).toString());
                        employerInfoText.insert(employerInfoText.length() - 12, "%%");
                        String[] employerInfo = employerInfoText.toString().split("%%");
                        String employerName = employerInfo[0];
                        String employerTAN = employerInfo[1].replace("(", "").replace(")", "");

                        int section17_1 = Integer.parseInt(row.getCell(10).toString().replace(",", ""));
                        int section17_2 = Integer.parseInt(row.getCell(11).toString().replace(",", ""));
                        int section17_3 = Integer.parseInt(row.getCell(12).toString().replace(",", ""));
                        int grossSalary = section17_1 + section17_2 + section17_3;
                        totalGrossSalary += grossSalary;

                        rowCount++;
                        sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
                        setExcelData(rowCount++, 4, Arrays.asList("Employer " + employerCount));
                        rowCount++;
                        setExcelData(rowCount++, 5, Arrays.asList("Employer Name", employerName));
                        setExcelData(rowCount++, 5, Arrays.asList("TAN of the Employer", employerTAN));
                        rowCount++;
                        sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
                        setExcelData(rowCount++, 3, Arrays.asList(employerCount == 1 ? " 1" : "", "Gross Salary"));
                        setExcelData(rowCount++, 4, Arrays.asList("a", "Salary as per Section 17(1)", String.format("%,d", section17_1)));
                        setExcelData(rowCount++, 4, Arrays.asList("b", "Salary as per Section 17(2)", String.format("%,d", section17_2)));
                        setExcelData(rowCount++, 4, Arrays.asList("c", "Profit in leiu of Salary as Section 17(3)", String.format("%,d", section17_3)));
                        setExcelData(rowCount++, 4, Arrays.asList("d", "Income from retirement benefit account maintained in notified country", "-"));
                        setExcelData(rowCount++, 4, Arrays.asList("e", "Income from retirement benefit account maintained other than  notified country", "-"));
                        setExcelData(rowCount++, 5, Arrays.asList("Total Gross Salary", String.format("%,d", grossSalary)));

                        employerCount++;
                    }
                }
            }

            if (rowCount == 1) {
                rowCount++;
                sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
                setExcelData(rowCount++, 4, Arrays.asList("Employer " + 1));
                rowCount++;
                setExcelData(rowCount++, 5, Arrays.asList("Employer Name"));
                setExcelData(rowCount++, 5, Arrays.asList("TAN of the Employer"));
                rowCount++;
                sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
                setExcelData(rowCount++, 3, Arrays.asList(" 1", "Gross Salary"));
                setExcelData(rowCount++, 4, Arrays.asList("a", "Salary as per Section 17(1)", "-"));
                setExcelData(rowCount++, 4, Arrays.asList("b", "Salary as per Section 17(2)", "-"));
                setExcelData(rowCount++, 4, Arrays.asList("c", "Profit in leiu of Salary as Section 17(3)", "-"));
                setExcelData(rowCount++, 4, Arrays.asList("d", "Income from retirement benefit account maintained in notified country", "-"));
                setExcelData(rowCount++, 4, Arrays.asList("e", "Income from retirement benefit account maintained other than  notified country", "-"));
                setExcelData(rowCount++, 5, Arrays.asList("Total Gross Salary", "-"));
            }
            rowCount++;
            setExcelData(rowCount++, 3, Arrays.asList(" 2", "", "Total Gross Salary from all employers", totalGrossSalary == 0 ? "-" : String.format("%,d", totalGrossSalary)));
            rowCount++;
            setExcelData(rowCount++, 3, Arrays.asList(" 2a ", "", "Income claimed for relief from taxation u/s 89A", "-"));
            rowCount++;
            setExcelData(rowCount++, 3, Arrays.asList(" 3", "Less :", "Allowances to the extent u/s 10", "-"));
            rowCount++;
            setExcelData(rowCount++, 3, Arrays.asList(" 4", "", "Net Salary", "-"));
            rowCount++;
            setExcelData(rowCount++, 3, Arrays.asList(" 5", "", "Deductions u/s 16"));
            setExcelData(rowCount++, 4, Arrays.asList("a", "Standard Deduction", "-"));
            setExcelData(rowCount++, 4, Arrays.asList("b", "Entertainment Allowance", "-"));
            setExcelData(rowCount++, 4, Arrays.asList("c", "Professional Tax", "-"));
            rowCount++;
            setExcelData(rowCount++, 3, Arrays.asList(" 6", "", "Income chargeable under the Head \" Salaries \"", "-"));

        } catch (Exception exception) {
            System.out.println("@createSheet Exception = " + exception);
        }
    }

    public void setExcelData(int rowNumber, int column, List list) {
        Row row = sheet.createRow(rowNumber);
        Cell cell = null;
        for (int i = 0; i < list.size(); i++) {
            cell = row.createCell(column + i);
            cell.setCellValue(list.get(i).toString());
            sheet.autoSizeColumn(column + i);
        }
    }

}
