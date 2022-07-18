package first.writer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;

public class ExcelReader {
    public class xlsInfo {

        private String word = "";
        private int startRow;
        private int manufactureRow;
        private int articulRow;
        private int nameRow;
        private int priceRow;
        private int avaliableRow;
        private int supplierArticulRow;
        private int codeRow;
        private String supplierRow;
        private String cityRow;

        public xlsInfo(String iword, int istartRow, int imanufactureRow, int iarticulRow, int inameRow, int ipriceRow, int iavaliableRow, int isupplierArticulRow, int icodeRow, String isupplierRow, String icityRow) {
            this.word = iword;
            this.startRow = istartRow;
            this.manufactureRow = imanufactureRow;
            this.articulRow = iarticulRow;
            this.nameRow = inameRow;
            this.priceRow = ipriceRow;
            this.avaliableRow = iavaliableRow;
            this.supplierArticulRow = isupplierArticulRow;
            this.codeRow = icodeRow;
            this.supplierRow = isupplierRow;
            this.cityRow = icityRow;
        }

        public String getWord() {
            return word;
        }
        public int getStartRow() {
            return startRow;
        }
        public int getManufactureRow() {
            return manufactureRow;
        }
        public int getArticulRow() {
            return articulRow;
        }
        public int getNameRow() {
            return nameRow;
        }
        public int getPriceRow() {return priceRow;}
        public int getAvaliableRow() {
            return avaliableRow;
        }
        public int getSupplierArticulRow() {
            return supplierArticulRow;
        }
        public int getCodeRow() {
            return codeRow;
        }
        public String getSupplierRow() {
            return supplierRow;
        }
        public String getCityRow() {
            return cityRow;
        }
    }

    private xlsInfo xlsfile;

    public void read(String filename) throws IOException {
        System.out.println(filename);
        if (filename.indexOf("STAL") > 0) {
            xlsfile = new xlsInfo("STAL", 8, -1, 0, 5, 13, -1, -1, -1, "STAL","ЕКБ");
        }
        if (filename.indexOf("ЕКБ") > 0) {
            xlsfile = new xlsInfo("ЕКБ", 13, -1, 3, 1, 4, 2, -1, -1, "ООО Фастгеар","ЕКБ");
        }
        if (filename.indexOf("ШАКМАН") > 0) {
            xlsfile = new xlsInfo("ШАКМАН", 1, -1, 1, 2, 4, 3, 0, -1, "ШАКМАН","ЕКБ");
        }
        if (filename.indexOf("Урал промо") > 0) {
            xlsfile = new xlsInfo("Урал промо", 7, 0, 2, 1, 4, 5, 6, -1, "Урал промо","ЕКБ");
        }
        Workbook workbook = loadWorkbook(filename);
        if (xlsfile != null) {
            var sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                processSheet(sheet);
            }
        }
        xlsfile = new xlsInfo("", 0, 0, 0, 0, 0, 0, 0, 0, "","");
    }

    private Workbook loadWorkbook(String filename) throws IOException {
        var extension = filename.substring(filename.lastIndexOf(".") + 1).toLowerCase();
        var file = new FileInputStream(new File(filename));
        switch (extension) {
            case "xls":
                return new HSSFWorkbook(file);
            case "xlsx":
                return new XSSFWorkbook(file);
            default:
                throw new RuntimeException("Unknown Excel file extension: " + extension);
        }
    }

    private void processSheet(Sheet sheet) {
        var iterator = sheet.rowIterator();
        if (xlsfile.startRow != -1) {
            for (var rowIndex = 0; iterator.hasNext(); rowIndex++) {
                var row = iterator.next();
                if (row != null && rowIndex >= xlsfile.startRow) {
                    processRow(rowIndex, row);
                }
            }
        }
    }

    private void processRow(int rowIndex, Row row) {
        try
        {
            if (processCell(row.getCell(xlsfile.getPriceRow()))!="")
        first.writer.conn.WriteDB(
                xlsfile.getManufactureRow() == -1 ? "" : processCell(row.getCell(xlsfile.getManufactureRow()))
                , xlsfile.getArticulRow() == -1 ? "" : processCell(row.getCell(xlsfile.getArticulRow()))
                , xlsfile.getNameRow() == -1 ? "" : processCell(row.getCell(xlsfile.getNameRow())).replace("'","`")
                , xlsfile.getPriceRow() == -1 ? "0" : processCell(row.getCell(xlsfile.getPriceRow()))
                , xlsfile.getAvaliableRow() == -1 ? "0" : processCell(row.getCell(xlsfile.getAvaliableRow()))
                , xlsfile.getSupplierArticulRow() == -1 ? "" : processCell(row.getCell(xlsfile.getSupplierArticulRow()))
                , xlsfile.getCodeRow() == -1 ? "" : processCell(row.getCell(xlsfile.getCodeRow()))
                , xlsfile.getSupplierRow()
                , xlsfile.getCityRow());
        }
        catch(SQLException e)
        {
            System.err.println(e.getMessage());
        }
    }

    private String processCell(Cell cell) {
        String ss = "";
        if (cell != null)
            switch (cell.getCellType()){
                case STRING:
                    ss = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        ss = " ";
                    } else {
                        ss = NumberToTextConverter.toText(cell.getNumericCellValue());
                    }
                    break;
                case ERROR:
                    ss = " ";
                    break;
                default:
                    ss = " ";}
        return ss.trim();
    }


 }
