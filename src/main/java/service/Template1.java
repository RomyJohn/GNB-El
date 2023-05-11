package service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.Arrays;
import java.util.List;

@Component
public class Template1 {

    private XSSFSheet sheet;

    public void createSheet() {
        try {
            XSSFWorkbook workbook = Template.workbook;
            sheet = workbook.createSheet("Computation 2122");

            drawBox(7, 134, 4, 9);
            drawBox(135, 168, 4, 9);

            sheet.addMergedRegion(new CellRangeAddress(7, 7, 4, 8));
            setExcelData(8, 4, Arrays.asList("Computation of Income for FY 2021-22"), 4, 8);
            setExcelData(9, 4, Arrays.asList("Particulars", "", "", "Amount", "Amount"), 4, 6);
            setExcelData(11, 4, Arrays.asList("1", "Income from Salary"), 5, 6);
            setExcelData(13, 5, Arrays.asList("Gross Salary"), 5, 6);
            setExcelData(14, 5, Arrays.asList("Less : Allowances under Section 10"), 5, 6);
            setExcelData(15, 5, Arrays.asList("Less : Standard Deduction of Rs 50000"), 5, 6);
            setExcelData(16, 5, Arrays.asList("Less : Professional Tax"), 5, 6);
            setExcelData(18, 4, Arrays.asList("2", "Income from House Property"), 5, 6);
            setExcelData(20, 4, Arrays.asList("3", "Income from Business or Profession"), 5, 6);
            setExcelData(22, 5, Arrays.asList("Presumptive Profession Income under 44ADA"), 5, 6);
            setExcelData(23, 5, Arrays.asList("Gross Receipts"), 5, 6);
            setExcelData(24, 5, Arrays.asList("a", "Gross Receipts Received"));
            setExcelData(26, 5, Arrays.asList("Presumptive income under section 44ADA"), 5, 6);
            setExcelData(27, 5, Arrays.asList("a", "50% or the amount claimed to have been earned"));
            setExcelData(29, 5, Arrays.asList("Business Income"), 5, 6);
            setExcelData(30, 6, Arrays.asList("Futures"));
            setExcelData(31, 6, Arrays.asList("Options"));
            setExcelData(33, 5, Arrays.asList("Speculative Income/Activity"), 5, 6);
            setExcelData(35, 5, Arrays.asList("Turnover from speculative activity"), 5, 6);
            setExcelData(36, 5, Arrays.asList("Gross Profit"), 5, 6);
            setExcelData(37, 5, Arrays.asList("Expenditure, if any"), 5, 6);
            setExcelData(38, 5, Arrays.asList("Net Income From Speculative Activity"), 5, 6);
            setExcelData(40, 4, Arrays.asList("4", "Income from Capital Gains"), 5, 6);
            setExcelData(41, 5, Arrays.asList("Long Term Capital Gain - Residential Property"), 5, 6);
            setExcelData(42, 5, Arrays.asList("Long Term Capital Gain on Sale of Equity Shares"), 5, 6);
            setExcelData(43, 5, Arrays.asList("Long Term Capital Gain on Sale of Equity Oriented Mutual Fund"), 5, 6);
            setExcelData(44, 5, Arrays.asList("Long Term Capital Gain on Sale of Other Units / Debt Funds"), 5, 6);
            setExcelData(45, 5, Arrays.asList("Long Term Capital Gain on Sale of Unlisted Shares"), 5, 6);
            setExcelData(47, 5, Arrays.asList("Short Term Capital Gain -Residential Property"), 5, 6);
            setExcelData(48, 5, Arrays.asList("Short Term Capital Gain on Sale of Equity Shares"), 5, 6);
            setExcelData(49, 5, Arrays.asList("Short Term Capital Gain on Sale of Equity Oriented Mutual Fund"), 5, 6);
            setExcelData(50, 5, Arrays.asList("Short Term Capital Gain on Sale of Other Units / Debt Funds"), 5, 6);
            setExcelData(51, 5, Arrays.asList("Short Term Capital Gain on Sale of Unlisted Shares"), 5, 6);
            setExcelData(54, 4, Arrays.asList("5", "Income from Other Sources"), 5, 6);
            setExcelData(55, 5, Arrays.asList("Dividend Income"), 5, 6);
            setExcelData(56, 5, Arrays.asList("Interest from Term Deposit"), 5, 6);
            setExcelData(57, 5, Arrays.asList("Interest from Recurring Deposit"), 5, 6);
            setExcelData(58, 5, Arrays.asList("Interest from Saving Bank"), 5, 6);
            setExcelData(59, 5, Arrays.asList("Interest on Income Tax Refund"), 5, 6);
            setExcelData(60, 5, Arrays.asList("Any Other Income - Pls Specify"), 5, 6);
            setExcelData(61, 5, Arrays.asList("Any Other Income - Pls Specify"), 5, 6);
            setExcelData(62, 5, Arrays.asList("Any Other Income - Pls Specify"), 5, 6);
            setExcelData(64, 4, Arrays.asList("6", "Total Income -Head Wise"), 5, 6);
            setExcelData(66, 4, Arrays.asList("7", "Losses of current year set off against 6", "", "", "10"), 5, 6);
            setExcelData(67, 5, Arrays.asList("Housing Loan Interest - Restricted Rs 2 Lacs", "", "1"), 5, 6);
            setExcelData(68, 5, Arrays.asList("Business Loss", "", "2"), 5, 6);
            setExcelData(69, 5, Arrays.asList("Short Term Capital Loss", "", "3"), 5, 6);
            setExcelData(70, 5, Arrays.asList("Long Term Capital Loss", "", "4"), 5, 6);
            setExcelData(72, 4, Arrays.asList("8", "Balance after set off current year losses(6 - 7)", "", "", "(10)"), 5, 6);
            setExcelData(74, 4, Arrays.asList("9", "Brought forward losses set off against 8", "", "", "4"), 5, 6);
            setExcelData(75, 5, Arrays.asList("Housing Loan Interest - Restricted Rs 2 Lacs", "", "1"), 5, 6);
            setExcelData(76, 5, Arrays.asList("Business Loss", "", "1"), 5, 6);
            setExcelData(77, 5, Arrays.asList("Short Term Capital Loss", "", "1"), 5, 6);
            setExcelData(78, 5, Arrays.asList("Long Term Capital Loss", "", "1"), 5, 6);
            setExcelData(80, 4, Arrays.asList("10", "Gross Total Income(8 - 9)", "", "", "(14)"), 5, 6);
            setExcelData(82, 4, Arrays.asList("11", "Income chargeable to tax at special rates"), 5, 6);
            setExcelData(83, 5, Arrays.asList("Section 111 - Tax on accumulated balance of recognised PF @ 1 %"), 5, 6);
            setExcelData(84, 5, Arrays.asList("Section 111 A - STCG on shares where STT is paid taxable @ 15 %"), 5, 6);
            setExcelData(85, 5, Arrays.asList("Section 115 AD(1) (b) (ii) - STCG on shares where STT is paid taxable @ 15 %"), 5, 6);
            setExcelData(86, 5, Arrays.asList("Section 112 - LTCG with Indexation @ 20 %"), 5, 6);
            setExcelData(87, 5, Arrays.asList("Section 112 Proviso - LTCG on listed securities / without indexation @ 10 % (IFSC)"), 5, 6);
            setExcelData(88, 5, Arrays.asList("Section 112 (1) (c) (iii) - LTCG on unlisted shares incase of non -residents @ 10 %"), 5, 6);
            setExcelData(89, 5, Arrays.asList("Section 112 A - LTCG on sale of shares where STT is paid @ 10 %"), 5, 6);
            setExcelData(92, 4, Arrays.asList("12", "Deductions under Chapter VI A", "", "", "50,000"), 5, 6);
            setExcelData(94, 4, Arrays.asList("13", "Deductions under 10 AA"), 5, 6);
            setExcelData(96, 4, Arrays.asList("14", "Total Income", "", "", "(50, 014)"), 5, 6);
            setExcelData(98, 4, Arrays.asList("15", "Income which is included in Total Income chargeable at special rates"), 5, 6);
            setExcelData(100, 4, Arrays.asList("16", "Net Agricultural Income,if any"), 5, 6);
            setExcelData(102, 4, Arrays.asList("17", "Aggregrate Income (14 - 15 + 16)", "", "", "(50, 014)"), 5, 6);
            setExcelData(104, 4, Arrays.asList("18", "Losses of current year to be carried forward"), 5, 6);
            setExcelData(105, 4, Arrays.asList("19", "Deemed income under section 115 JC", "", "", "(50, 010)"), 5, 6);
            setExcelData(107, 5, Arrays.asList("Tax Payable on Aggregrate Income"), 5, 6);
            setExcelData(108, 5, Arrays.asList("a", "Tax at normal rates", "", "0"));
            setExcelData(109, 5, Arrays.asList("b1", "Tax on accumulated balance of recognised PF @ 1 %"));
            setExcelData(110, 5, Arrays.asList("b2", "Tax on STCG on shares where STT is paid taxable @ 15 %"));
            setExcelData(111, 5, Arrays.asList("b3", "Tax on STCG on shares where STT is paid taxable @ 15 % -Non Resident"));
            setExcelData(112, 5, Arrays.asList("b4", "Tax on LTCG with Indexation @ 20 %"));
            setExcelData(113, 5, Arrays.asList("b5", "Tax on LTCG on listed securities/without indexation @ 10 % (IFSC)"));
            setExcelData(114, 5, Arrays.asList("b6", "LTCG on unlisted shares incase of non - residents @ 10 %"));
            setExcelData(115, 5, Arrays.asList("b7", "LTCG on sale of shares where STT is paid @ 10 %"));
            setExcelData(116, 5, Arrays.asList("b8", "Tax on agricultural income"));
            setExcelData(118, 5, Arrays.asList("c", "Rebate on agricultural income,if applicable"));
            setExcelData(120, 5, Arrays.asList("d", "Tax Payable on Total Income (a + b1tob8 –c)"));
            setExcelData(121, 5, Arrays.asList("e", "Rebate under section 87 A"));
            setExcelData(122, 5, Arrays.asList("f", "Tax Payable after rebate(d - e)"));
            setExcelData(124, 5, Arrays.asList("g 1", "Surcharge @applicable rates"));
            setExcelData(125, 5, Arrays.asList("g 2", "Surcharge @ 15 %"));
            setExcelData(127, 5, Arrays.asList("h", "Health and Education Cess"));
            setExcelData(128, 5, Arrays.asList("i", "Gross Tax Liability (f + g1 + g2 + h)"));
            setExcelData(130, 5, Arrays.asList("Credit under Section 115JD"), 5, 6);
            setExcelData(131, 5, Arrays.asList("Tax Relief under Section 89/90/90A/91"), 5, 6);
            setExcelData(132, 5, Arrays.asList("Net Tax Liability"), 5, 6);
            setExcelData(136, 5, Arrays.asList("Interest and Fee Payable"), 5, 6);
            setExcelData(137, 6, Arrays.asList("Interest for default in furnishing the return (234A)"));
            setExcelData(138, 6, Arrays.asList("Interest for default in payment of advance tax (234B)"));
            setExcelData(139, 6, Arrays.asList("Interest for deferment of advance tax (234C)"));
            setExcelData(140, 6, Arrays.asList("Interest under 234F"));
            setExcelData(142, 5, Arrays.asList("Less :"), 5, 6);
            setExcelData(143, 5, Arrays.asList("Advance Tax - 100", "", ""), 5, 6);
            setExcelData(144, 5, Arrays.asList("Self Assessment Tax - 300", "", ""), 5, 6);
            setExcelData(145, 5, Arrays.asList("Regular Assessment Tax - 400", "", ""), 5, 6);
            setExcelData(146, 5, Arrays.asList("Tax Collected @ Source", "", ""), 5, 6);
            setExcelData(147, 5, Arrays.asList("Tax Deducted @ Source", "", ""), 5, 6);
            setExcelData(149, 5, Arrays.asList("Amount Payable/Refund", "", ""), 5, 6);
            setExcelData(152, 6, Arrays.asList("Chapter VI-A Deductions"));
            setExcelData(154, 6, Arrays.asList("80C - Provident Fund"));
            setExcelData(155, 6, Arrays.asList("80C - Life Insurance Premiums, If any"));
            setExcelData(156, 6, Arrays.asList("80C - Term Insurance"));
            setExcelData(157, 6, Arrays.asList("80C - Pension Scheme"));
            setExcelData(158, 6, Arrays.asList("Total ( Restricted to Rs 1.5Lacs)"));
            setExcelData(160, 6, Arrays.asList("80D - Health Insurance Premium"));
            setExcelData(162, 6, Arrays.asList("80TTB - Saving Bank & FD Interest", "", "50,000"));
            setExcelData(164, 6, Arrays.asList("80G - Donations"));
            setExcelData(166, 6, Arrays.asList("Total Allowable Deduction", "", "50,000"));
        } catch (Exception exception) {
            System.out.println("@createSheet Exception = " + exception);
        }
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

    public void setExcelData(int rowNumber, int column, List list, int... columns) {
        if (columns.length != 0)
            sheet.addMergedRegion(new CellRangeAddress(rowNumber, rowNumber, columns[0], columns[1]));
        Row row = sheet.createRow(rowNumber);
        Cell cell = null;
        for (int i = 0; i < list.size(); i++) {
            cell = row.createCell(column + i);
            cell.setCellValue(list.get(i).toString());
            sheet.autoSizeColumn(column + i);
        }
    }

}
