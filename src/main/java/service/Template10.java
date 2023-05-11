package service;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

@Component
public class Template10 {

    public void createSheet() {
        try {
            XSSFWorkbook workbook = Template.workbook;
            XSSFSheet sheet_tds = Scrapping.sheet_tds;
            XSSFSheet sheet = workbook.createSheet("TDSTCS");

            for (int i = 0; i <= sheet_tds.getLastRowNum(); i++) {
                Row row = sheet_tds.getRow(i);
                Row row1 = sheet.createRow(i);
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    row1.createCell(j).setCellValue(row.getCell(j).toString());
                    sheet.autoSizeColumn(j);
                }
            }

        } catch (Exception exception) {
            System.out.println("@createSheet10 Exception = " + exception);
            exception.printStackTrace();
        }
    }

}
