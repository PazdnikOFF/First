package first.writer;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {
    private static XSSFWorkbook workbook = new XSSFWorkbook();
    private static int rowcount=1;
    private static Sheet sheet;
    public static ExcelWriter writer;
    public static ExcelWriter writerprice;

    public static void create(String filename) throws IOException {
        sheet = createSheet(workbook);
        createHeader(workbook);
    }

    public static void close(String filename) throws IOException {
        try (var outputStream = new FileOutputStream(filename)) {
          workbook.write(outputStream);
        }
        workbook.close();
    }

    public static boolean isNumeric(String str) {
        return str.matches("-?\\d+(\\.\\d+)?");  //match a number with optional '-' and decimal.
    }

    private static Sheet createSheet(XSSFWorkbook workbook) {
        sheet = workbook.createSheet("СпецИнтер");
        sheet.setColumnWidth(0, 8000);
        sheet.setColumnWidth(1, 4000);
        sheet.setColumnWidth(2, 17000);
        sheet.setColumnWidth(3, 4000);
        sheet.setColumnWidth(4, 2000);
        sheet.setColumnWidth(5, 4000);
        sheet.setColumnWidth(6, 4000);
        sheet.setColumnWidth(7, 4000);
        sheet.setColumnWidth(8, 3000);
        return sheet;
    }

    private static void createHeader(XSSFWorkbook workbook) {
        var header = sheet.createRow(0);

        var headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        var font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        headerStyle.setFont(font);

        var headerCell = header.createCell(0);
        headerCell.setCellValue("Производитель");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Артикул детали");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(2);
        headerCell.setCellValue("Наименование детали");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(3);
        headerCell.setCellValue("Цена");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(4);
        headerCell.setCellValue("Наличие");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(5);
        headerCell.setCellValue("Артикул поставщика");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(6);
        headerCell.setCellValue("Код товара");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(7);
        headerCell.setCellValue("Поставщик");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(8);
        headerCell.setCellValue("Город");
        headerCell.setCellStyle(headerStyle);
    }

    public static void createCells(String manufacture, String articul, String name, String price, String avaliable, String supplierArticul, String code, String supplier, String city){
            var style = workbook.createCellStyle();
            style.setWrapText(true);
            var createHelper = workbook.getCreationHelper();
            sheet.setDefaultColumnStyle(1, style);
    if (isNumeric(price)) {
            var row = sheet.createRow(rowcount);

            var cell = row.createCell(0);
            cell.setCellValue(manufacture);

            cell = row.createCell(1);
            cell.setCellValue(articul);

            cell = row.createCell(2);
            cell.setCellValue(name);

            row.createCell(3).setCellValue((double) (isNumeric(price) ? Double.valueOf(price) : 0));

            row.createCell(4).setCellValue((double) (isNumeric(avaliable) ? Double.valueOf(avaliable) : 1));

            cell = row.createCell(5);
            cell.setCellValue(supplierArticul);

            cell = row.createCell(6);
            cell.setCellValue(code);

            cell = row.createCell(7);
            cell.setCellValue(supplier);

            cell = row.createCell(8);
            cell.setCellValue(city);
            rowcount++;
    }
        }
    }

