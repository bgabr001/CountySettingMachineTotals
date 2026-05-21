import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class CountySettingTotals {

    public static void main(String[] args) throws Exception {

        String inputFolder = "counties";
        String outputFile = "County_Machine_Totals.xlsx";

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("County Totals");

        Row header = outputSheet.createRow(0);
        header.createCell(0).setCellValue("County");
        header.createCell(1).setCellValue("Scans");
        header.createCell(2).setCellValue("ADA");
        header.createCell(3).setCellValue("Prints");
        header.createCell(4).setCellValue("Total Machines");

        File folder = new File(inputFolder);

        File[] files = folder.listFiles((dir, name) -> {
            String lowerName = name.toLowerCase();
            return lowerName.endsWith(".xlsx")
                    && !lowerName.startsWith("~$")
                    && !lowerName.contains("county_machine_totals");
        });

        java.util.Arrays.sort(files,
                (a, b) -> a.getName().compareToIgnoreCase(b.getName()));

        if (files == null || files.length == 0) {
            System.out.println("No Excel files found in counties folder.");
            return;
        }

        int outputRowNum = 1;

        for (File file : files) {

            MachineTotals totals = readCountyFile(file);

            Row row = outputSheet.createRow(outputRowNum++);
            row.createCell(0).setCellValue(getCountyName(file));
            row.createCell(1).setCellValue(totals.scans);
            row.createCell(2).setCellValue(totals.ada);
            row.createCell(3).setCellValue(totals.prints);
            row.createCell(4).setCellValue(totals.total);

            System.out.println("Processed: " + file.getName()
                    + " | Scans: " + totals.scans
                    + " | ADA: " + totals.ada
                    + " | Prints: " + totals.prints
                    + " | Total: " + totals.total);
        }

        for (int i = 0; i < 5; i++) {
            outputSheet.autoSizeColumn(i);
        }

        FileOutputStream fos = new FileOutputStream(outputFile);
        outputWorkbook.write(fos);
        fos.close();
        outputWorkbook.close();

        System.out.println("DONE! Created: " + outputFile);
    }

    private static MachineTotals readCountyFile(File file) throws Exception {

        MachineTotals totals = new MachineTotals();

        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);

        for (Sheet sheet : workbook) {
            for (Row row : sheet) {

                String rowText = getWholeRowText(row).toLowerCase();

                if (rowText.contains("machines used for election")) {

                    totals.scans = getNumericValue(row.getCell(2));   // C
                    totals.ada = getNumericValue(row.getCell(3));     // D
                    totals.prints = getNumericValue(row.getCell(4));  // E
                    totals.total = getNumericValue(row.getCell(5));   // F

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
        int total;
    }
}