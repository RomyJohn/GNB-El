package service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.Arrays;
import java.util.List;

@Component
public class Template2 {

    private XSSFSheet sheet;

    public void createSheet() {
        try {
            XSSFWorkbook workbook = Template.workbook;
            sheet = workbook.createSheet("Chapter VI-A");

            sheet.addMergedRegion(new CellRangeAddress(1, 1, 2, 3));
            setExcelData(1, 1, Arrays.asList("Particulars", "Allowable"));
            setExcelData(3, 1, Arrays.asList("80C - LIC"));
            setExcelData(4, 1, Arrays.asList("80C - Provident Fund"));
            setExcelData(5, 1, Arrays.asList("80C - Housing Loan Repayment"));
            setExcelData(6, 1, Arrays.asList("80C - Tax Saver MF"));
            setExcelData(9, 1, Arrays.asList("1. Part B -Deduction in respect of certain payments"));
            setExcelData(10, 1, Arrays.asList("80 C - Life insurance premia, deferred annuity, contributions to provident fund, subscription to certain equity shares or debentures, etc.", "a"));
            setExcelData(11, 1, Arrays.asList("80 CCC - Payment in respect Pension Fund", "b"));
            setExcelData(12, 1, Arrays.asList("80 CCD(1) - Contribution to pension scheme of Central Government", "c"));
            setExcelData(13, 1, Arrays.asList("80 CCD(1B) - Contribution to pension scheme of Central Government", "d"));
            setExcelData(14, 1, Arrays.asList("80 CCD(2) - Contribution to pension scheme of Central Government by the Employer", "e"));
            setExcelData(15, 1, Arrays.asList("80 CCG - Investment made under an equity savings scheme", "f"));
            setExcelData(16, 1, Arrays.asList("80 CCF", "f"));
            setExcelData(17, 1, Arrays.asList("80D - Deduction in respect of Health Insurance premia", "f"));
            setExcelData(18, 1, Arrays.asList("(A) Health Insurance Premium"));
            setExcelData(19, 1, Arrays.asList("(B) Medical expenditure"));
            setExcelData(20, 1, Arrays.asList("(C) Preventive health check -up"));
            setExcelData(21, 1, Arrays.asList("80 DD - Maintenance including medical treatment of a dependant who is a person with disability", "g"));
            setExcelData(22, 1, Arrays.asList("80 DDB - Medical treatment of specified disease", "h"));
            setExcelData(23, 1, Arrays.asList("80 E - Interest on loan taken for higher education", "i"));
            setExcelData(24, 1, Arrays.asList("80E E - Interest on loan taken for residential house property", "j"));
            setExcelData(25, 1, Arrays.asList("80E EA - Deduction in respect of interest on loan taken for certain house property", "k"));
            setExcelData(26, 1, Arrays.asList("80E EB - Deduction in respect of purchase of electric vehicle", "l"));
            setExcelData(27, 1, Arrays.asList("80 G - Donations to certain funds, charitable institutions, etc (Please fill 80 G Schedule.This field is auto -populated from schedule.)", "m"));
            setExcelData(28, 1, Arrays.asList("80 GG - Rent paid", "n"));
            setExcelData(29, 1, Arrays.asList("80 GGA", "m"));
            setExcelData(30, 1, Arrays.asList("80 GGC - Donation to Political party", "o"));
            setExcelData(31, 1, Arrays.asList("Total Deduction under Part B(total of a to o)"));
            setExcelData(32, 1, Arrays.asList("3. Part CA and D Deduction in respect of certain incomes /other Deductions"));
            setExcelData(33, 1, Arrays.asList("80 TTA - Interest on saving bank Accounts incase of other than Resident senior citizens", "Y"));
            setExcelData(34, 1, Arrays.asList("80 TTB - Interest on deposits in case of Resident senior citizens", "z"));
            setExcelData(35, 1, Arrays.asList("80 U - In case of a person with disability.", "i"));
            setExcelData(36, 1, Arrays.asList("Total Deduction under Part CA and D(total of I, ii and iii)", "3"));
            setExcelData(37, 1, Arrays.asList("Total deductions under Chapter VI - A(1 + 2 + 3)", "4"));
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
