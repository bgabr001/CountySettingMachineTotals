//
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Arrays;

public class CountySettingTotals {

    public static void main(String[] args) throws Exception {

        String inputFolder = "2026PrimaryCounties";
        String outputFile = "2026PrimaryCountyMachineTotals.xlsx";

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("County Totals");

        CellStyle headerStyle = createHeaderStyle(outputWorkbook);
        CellStyle normalStyle = createNormalStyle(outputWorkbook);
        CellStyle alternateStyle = createAlternateStyle(outputWorkbook);

        Row header = outputSheet.createRow(0);

        String[] headers = {
                "County",
                "Scan",
                "ADA",
                "Print",
                "Paper Rolls",
                "Red Seals",
                "Blue/Yellow/Black",
                "Notes"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        File folder = new File(inputFolder);

        File[] files = folder.listFiles((dir, name) -> {
            String lowerName = name.toLowerCase();

            return lowerName.endsWith(".xlsx")
                    && !lowerName.startsWith("~$")
                    && !lowerName.contains("county_machine_totals");
        });

        if (files == null || files.length == 0) {
            System.out.println("No Excel files found in counties folder.");
            outputWorkbook.close();
            return;
        }

        Arrays.sort(files, (a, b) ->
                a.getName().compareToIgnoreCase(b.getName()));

        int outputRowNum = 1;

        for (File file : files) {

            MachineTotals totals = readCountyFile(file);

            int totalMachines = totals.scans + totals.ada + totals.prints;

            Row row = outputSheet.createRow(outputRowNum);

            CellStyle rowStyle = (outputRowNum % 2 == 0)
                    ? alternateStyle
                    : normalStyle;

            createStyledCell(row, 0, getCountyName(file), rowStyle);
            createStyledCell(row, 1, totals.scans, rowStyle);
            createStyledCell(row, 2, totals.ada, rowStyle);
            createStyledCell(row, 3, totals.prints, rowStyle);
            createStyledCell(row, 4, totalMachines, rowStyle);
            createStyledCell(row, 5, totalMachines, rowStyle);
            createStyledCell(row, 6, totals.scans, rowStyle);
            createStyledCell(row, 7, "", rowStyle);

            outputRowNum++;

            System.out.println("Processed: " + file.getName());
        }

        outputSheet.createFreezePane(0, 1);

        outputSheet.setAutoFilter(new CellRangeAddress(
                0,
                outputRowNum - 1,
                0,
                7
        ));

        for (int i = 0; i < 8; i++) {
            outputSheet.autoSizeColumn(i);
        }

        // Add extra spacing to columns
        for (int i = 0; i < 8; i++) {

            int currentWidth = outputSheet.getColumnWidth(i);

            outputSheet.setColumnWidth(i, currentWidth + 1500);
        }

// Wider Notes column
        outputSheet.setColumnWidth(7, 9000);

        outputSheet.setColumnWidth(7, 6000);

        FileOutputStream fos = new FileOutputStream(outputFile);
        outputWorkbook.write(fos);
        fos.close();
        outputWorkbook.close();

        System.out.println();
        System.out.println("DONE!");
        System.out.println("Created spreadsheet: " + outputFile);
    }

    private static MachineTotals readCountyFile(File file) throws Exception {

        MachineTotals totals = new MachineTotals();

        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);

        for (Sheet sheet : workbook) {
            for (Row row : sheet) {

                String rowText = getWholeRowText(row).toLowerCase();

                if (rowText.contains("machines used for election")) {

                    totals.scans = getNumericValue(row.getCell(2));
                    totals.ada = getNumericValue(row.getCell(3));
                    totals.prints = getNumericValue(row.getCell(4));

                    workbook.close();
                    fis.close();

                    return totals;
                }
            }
        }

        System.out.println("WARNING: Could not find totals row in " + file.getName());

        workbook.close();
        fis.close();

        return totals;
    }

    private static void createStyledCell(Row row, int column, String value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static void createStyledCell(Row row, int column, int value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {

        CellStyle style = workbook.createCellStyle();

        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());

        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.TEAL.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    private static CellStyle createNormalStyle(Workbook workbook) {

        CellStyle style = workbook.createCellStyle();

        // Pastel blue color
        style.setFillForegroundColor(
                IndexedColors.PALE_BLUE.getIndex());

        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        return style;
    }

    private static CellStyle createAlternateStyle(Workbook workbook) {

        CellStyle style = workbook.createCellStyle();

        style.setFillForegroundColor(
                IndexedColors.WHITE.getIndex());

        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        style.setAlignment(HorizontalAlignment.CENTER);

        return style;
    }
    private static String getWholeRowText(Row row) {

        StringBuilder text = new StringBuilder();

        for (Cell cell : row) {
            text.append(getCellText(cell)).append(" ");
        }

        return text.toString();
    }

    private static int getNumericValue(Cell cell) {

        if (cell == null) {
            return 0;
        }

        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return (int) cell.getNumericCellValue();
            }

            if (cell.getCellType() == CellType.FORMULA) {
                if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                    return (int) cell.getNumericCellValue();
                }

                if (cell.getCachedFormulaResultType() == CellType.STRING) {
                    String text = cell.getStringCellValue().trim();
                    text = text.replace(",", "");

                    try {
                        return Integer.parseInt(text);
                    } catch (Exception e) {
                        return 0;
                    }
                }
            }

            if (cell.getCellType() == CellType.STRING) {
                String text = cell.getStringCellValue().trim();
                text = text.replace(",", "");

                try {
                    return Integer.parseInt(text);
                } catch (Exception e) {
                    return 0;
                }
            }

        } catch (Exception e) {
            return 0;
        }

        return 0;
    }

    private static String getCellText(Cell cell) {

        if (cell == null) {
            return "";
        }

        try {
            if (cell.getCellType() == CellType.STRING) {
                return cell.getStringCellValue();
            }

            if (cell.getCellType() == CellType.NUMERIC) {
                return String.valueOf((int) cell.getNumericCellValue());
            }

            if (cell.getCellType() == CellType.FORMULA) {
                if (cell.getCachedFormulaResultType() == CellType.STRING) {
                    return cell.getStringCellValue();
                }

                if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            }

        } catch (Exception e) {
            return "";
        }

        return "";
    }

    private static String getCountyName(File file) {

        return file.getName()
                .replace(".xlsx", "")
                .replace(".xls", "")
                .replace(" Election Plan", "");
    }

    static class MachineTotals {
        int scans;
        int ada;
        int prints;
    }
}