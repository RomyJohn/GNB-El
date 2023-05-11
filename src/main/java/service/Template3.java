package service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Component
public class Template3 {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<Integer> rentReceivedAmountList;

    public void createSheet() {
        try {
            workbook = Template.workbook;
            sheet = workbook.createSheet("Schedule House Property");
            XSSFSheet sheet_tds = Scrapping.sheet_tds;
            XSSFSheet sheet_other = Scrapping.sheet_other;

            rentReceivedAmountList = new ArrayList<>();

            int totalRowCount = sheet_tds.getLastRowNum();

            boolean rentReceivedPresent = extractData(totalRowCount, sheet_tds);

            if (!rentReceivedPresent) {
                totalRowCount = sheet_other.getLastRowNum();
                extractData(totalRowCount, sheet_other);
            }

            Row row = sheet.createRow(11);
            row.createCell(0).setCellValue("Rent Amount(For 12 Months)");
            int columnCount = 2;

            for (Integer item : rentReceivedAmountList) {
                Cell cell = row.createCell(columnCount);
                cell.setCellValue(String.format("%,d", item));
                columnCount++;
            }

            drawBox(4, 9, 0, 6);
            drawBox(10, 23, 0, 6);
            drawBox(4, 23, 0, 1);
            drawBox(4, 23, 2, 3);
            drawBox(4, 23, 4, 5);
            drawBox(4, 23, 5, 6);

            sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 5));
            setExcelData(3, 0, Arrays.asList("House Property Income Statement"));
            setExcelData(4, 0, Arrays.asList("Type of Property", "House Property-1", "House Property-2", "House Property-3", "House Property-4", "House Property-5"));
            setExcelData(5, 0, Arrays.asList("Let Out Property"));
            setExcelData(6, 0, Arrays.asList("Self Occupied Property"));
            setExcelData(7, 0, Arrays.asList("Deemed Let Out Property"));
            setExcelData(9, 0, Arrays.asList("Particulars", "Amount(Rs.)", "Amount(Rs.)", "Amount(Rs.)", "Amount(Rs.)", "Amount(Rs.)"));
            setExcelData(10, 0, Arrays.asList("Income from House Property"));
            setExcelData(12, 0, Arrays.asList("Less: Property Tax Paid"));
            setExcelData(13, 0, Arrays.asList("Annual Value of Property"));
            setExcelData(14, 0, Arrays.asList("Your percentage of share in the property", "100%", "100%", "100%", "100%", "100%"));
            setExcelData(15, 0, Arrays.asList("Annual value of the property owned (own percentage share * Annual Value)"));
            setExcelData(17, 0, Arrays.asList("Deductions u/s 24"));
            setExcelData(18, 0, Arrays.asList("30% of Net Annual Value"));
            setExcelData(19, 0, Arrays.asList("Interest on Borrowed Capital (Cannot Exceed Rs.2,00,000/- in case of Self Occupied Property)"));
            setExcelData(21, 0, Arrays.asList("Income from House Property"));

        } catch (Exception exception) {
            System.out.println("@createSheet Exception = " + exception);
        }
    }

    public boolean extractData(int totalRowCount, XSSFSheet sheet) {
        boolean rentReceivedPresent = false;
        int rentReceived = 0;
        for (int i = 0; i <= totalRowCount; i++) {
            Row row = sheet.getRow(i);
            Row row1 = sheet.getRow(i - 1);
            if (row.getFirstCellNum() != -1) {
                if (row.getCell(1).toString().equals("Rent received") && row1.getCell(0).toString().equals("Sr.No")) {
                    rentReceived = Integer.parseInt(row.getCell(7).toString().replace(",", ""));
                    rentReceivedAmountList.add(rentReceived);
                    rentReceivedPresent = true;
                }
            }
        }
        return rentReceivedPresent;
    }

    public void drawBox(int row1, int row2, int column1, int column2) {
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, column1, row1, column2, row2);
        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
        XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
        shape.setShapeType(ShapeTypes.RECT);
        shape.setLineWidth(1.5);
        shape.setLineStyleColor(0, 0, 0);
    }

    public void setExcelData(int rowNumber, int column, List list) {
        Row row = sheet.createRow(rowNumber);
        Cell cell = null;
        for (int i = 0; i < list.size(); i++) {
            cell = row.createCell(column + i);
            cell.setCellValue(list.get(i).toString());
            sheet.autoSizeColumn(column + i);
            if (rowNumber == 3 || rowNumber == 4 || rowNumber == 9) {
                CellStyle cellStyle = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                if (rowNumber == 3 || rowNumber == 9)
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyle.setFont(font);
                cell.setCellStyle(cellStyle);
            }
        }
    }

}
