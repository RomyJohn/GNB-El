package service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import utility.ConfigUtility;

@Component
public class Template5 {

    @Autowired
    private ConfigUtility configUtility;
    private static XSSFWorkbook workbook = null;

    public void createSheet() {
        try {
            workbook = Template.workbook;
            XSSFSheet sheet1 = workbook.createSheet("Capital Gains - Long Term");
            XSSFSheet sheet2 = workbook.createSheet("Capital Gains - Short Term");
            XSSFSheet sheet_sft = Scrapping.sheet_sft;

            String[] header = configUtility.getProperty("CAPITAL_GAINS_HEADER_DATA").split(",");

            setExcelHeader(sheet1, header);
            setExcelHeader(sheet2, header);

            int totalRowCount = sheet_sft.getLastRowNum();
            int rowCount1 = 1;
            int rowCount2 = 1;

            for (int i = 0; i <= totalRowCount; i++) {
                Row row = sheet_sft.getRow(i);
                if (row.getFirstCellNum() != -1) {
                    if (row.getCell(2).toString().equals("SFT-012") ||
                            row.getCell(2).toString().equals("SFT-17-LES(M)") ||
                            row.getCell(2).toString().equals("SFT-17-EMF(M)") ||
                            row.getCell(2).toString().equals("SFT-18-EMF(M)") ||
                            row.getCell(2).toString().equals("SFT-17-OTU(M)") ||
                            row.getCell(2).toString().equals("SFT-18-OTU(M)")) {
                        if (row.getCell(13).toString().equals("Long term")) {
                            Row row1 = sheet1.createRow(rowCount1);
                            setExcelData(row, row1, sheet1);
                            rowCount1++;
                        } else if (row.getCell(13).toString().equals("Short term")) {
                            Row row2 = sheet2.createRow(rowCount2);
                            setExcelData(row, row2, sheet2);
                            rowCount2++;
                        }
                    }
                }
            }

        } catch (Exception exception) {
            System.out.println("@createSheet Exception = " + exception);
        }
    }

    public static void setExcelHeader(XSSFSheet sheet, String[] header) {
        Row row = sheet.createRow(0);
        Cell cell = null;
        for (int i = 0; i < header.length; i++) {
            sheet.autoSizeColumn(i);
            cell = row.createCell(i);
            cell.setCellValue(header[i]);
        }
    }

    public static void setExcelData(Row row, Row newRow, Sheet sheet) {
        int columnCount = row.getLastCellNum();
        for (int j = 0; j < row.getLastCellNum(); j++) {
            newRow.createCell(j).setCellValue(row.getCell(j).toString());
            sheet.autoSizeColumn(j);
        }

        int salesConsideration = Integer.parseInt(row.getCell(16).toString().replace(",", ""));
        int costOfAcquisition = Integer.parseInt(row.getCell(17).toString().replace(",", ""));
        newRow.createCell(columnCount++).setCellValue(String.format("%,d", salesConsideration - costOfAcquisition));
        sheet.autoSizeColumn(columnCount);

        if (row.getCell(12).toString().equals("Off market")) {
            newRow.createCell(columnCount++).setCellValue("Cost of acquisition is zero");
            setBackgroundColor(newRow);
            sheet.autoSizeColumn(columnCount);
        }
    }

    public static void setBackgroundColor(Row row) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        for (int i = 0; i < row.getLastCellNum(); i++)
            row.getCell(i).setCellStyle(cellStyle);
    }

}
