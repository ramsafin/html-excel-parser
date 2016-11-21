package ru.kpfu.itis.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.kpfu.itis.html.HTMLTableService;
import ru.kpfu.itis.table.ExcelTable;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.function.Consumer;

import static org.apache.poi.ss.usermodel.CellType.*;


/**
 * Excel - Table converter
 */
public final class ExcelTableService {


    public void writeTwoTables(CellData[][] tableLeft, ExcelTable tableRight) {

        System.out.println(Arrays.deepToString(tableLeft));

        System.out.println("\n\n\n\n\n");

        System.out.println(tableRight);
    }

    //reads first table
    private CellData[][] readTable1(String path) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook(path);

        XSSFSheet sheet = workbook.getSheetAt(0);

        if (sheet == null) throw new IllegalArgumentException("There is no sheets in the document");

        final int rows = sheet.getPhysicalNumberOfRows(); //NOTICE: Might be exceptions

        CellData[][] cellData = new CellData[rows][3];

        sheet.forEach(new Consumer<Row>() {
            private int rowIdx = 0;

            @Override
            public void accept(Row row) {
                for (int k = 0; k < 3; k++) {
                    Cell cell = row.getCell(k);
                    cellData[rowIdx][k] = readCellValue1(cell);
                }
                ++rowIdx;
            }
        });
        close(workbook);

        return cellData;
    }


    //reads second table
    public ExcelTable readTable2(String path) throws IOException, InvalidFormatException {

        XSSFWorkbook workbook = readWorkbook(path);

        XSSFSheet sheet = workbook.getSheetAt(0); //get first sheet

        if (sheet == null) throw new IllegalArgumentException("There is no sheets in the document");

        final int rows = sheet.getLastRowNum() + 1;
        final int columns = sheet.getRow(0).getLastCellNum();

        ExcelTable table = new ExcelTable(rows, columns - 3);

        sheet.forEach(row -> {
            List<String> values = new ArrayList<>(columns - 3); //cell values of row
            for (int k = 3, colIdx = k - 3; k < columns; k++, colIdx++) {
                Cell cell = row.getCell(k);
                if (k == 3 && readCellValue2(cell).valueString().equals("")) {
                    return;
                }
                CellData value = readCellValue2(cell);
                values.add(value.valueString());
            }
            table.addRow(values.get(0), values.toArray(new String[0])); //in each row the key is data-id (first cell)
        });
        close(workbook);

        return table;
    }


    //for table's first part
    private CellData readCellValue1(Cell cell) {
        if (cell == null) return new CellData("", BLANK);
        switch (cell.getCellTypeEnum()) {
            case BLANK:
                return new CellData("", BLANK);
            case STRING:
                return new CellData(cell.getStringCellValue(), STRING);
            case NUMERIC:
                if (isInteger(cell.getNumericCellValue())) {
                    return new CellData((int) cell.getNumericCellValue(), NUMERIC);
                }
                return new CellData(cell.getNumericCellValue(), NUMERIC);
            case ERROR:
                return new CellData(cell.getErrorCellValue(), ERROR);
            case BOOLEAN:
                return new CellData(cell.getBooleanCellValue(), BOOLEAN);
            case FORMULA:
                return new CellData(cell.getCellFormula(), FORMULA);
            default:
                return new CellData("", STRING);
        }
    }


    //for table's second part
    private CellData readCellValue2(Cell cell) {
        if (cell == null) return new CellData("", BLANK);
        switch (cell.getCellTypeEnum()) {
            case BLANK:
                return new CellData("", STRING);
            case STRING:
                return new CellData(cell.getStringCellValue(), STRING);
            case NUMERIC:
                if (isInteger(cell.getNumericCellValue())) {
                    return new CellData((int) cell.getNumericCellValue(), NUMERIC);
                }
                return new CellData(cell.getNumericCellValue(), NUMERIC);
            case ERROR:
                return new CellData(Byte.toString(cell.getErrorCellValue()), STRING);
            case BOOLEAN:
                return new CellData(Boolean.toString(cell.getBooleanCellValue()), STRING);
            case FORMULA:
                return new CellData(cell.getCellFormula(), FORMULA);
            default:
                return new CellData("", STRING);
        }
    }


    private boolean isInteger(Double d) {
        return d.toString().matches("^[0-9]+(.0)?$");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        HTMLTableService htmlTableService = new HTMLTableService();
        ExcelTableService excelTableService = new ExcelTableService();

        ExcelTable htmlTable = htmlTableService.createTable("/Users/Ramil/Desktop/site.html"); //create table from html

//        excelTableService.writeTable(htmlTable.sort(0), "/Users/Ramil/Desktop/doc.xlsx");

//        CellData[][] table1 = excelTableService.readTable1("/Users/Ramil/Desktop/doc.xlsx"); //read 1
        ExcelTable table2 = excelTableService.readTable2("/Users/Ramil/Desktop/doc.xlsx"); // and 2 tables from file

        //merge html and table2
        table2.merge(htmlTable, 3); // 3 columns
//
        System.out.println(table2);

//        ExcelTable table2 = wrapper.getTable();
//        excelTable1.merge(excelTable, 3);
//        excelTableConverter.writeTwoTable(excelTable1.sort(sortColumn), wrapper.getCellData(), excelFile.getPath());
    }


    private XSSFWorkbook readWorkbook(String path) throws IOException, InvalidFormatException {
        try (BufferedInputStream in = new BufferedInputStream(new FileInputStream(path))) {
            return new XSSFWorkbook(in); //we should close it when stop working with it, see docs
        }
    }


    private void setUpColumnWidth(Sheet sheet) {
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            sheet.autoSizeColumn(i);
        }
    }


    private void close(Workbook workbook) {
        try {
            workbook.close();
        } catch (IOException e) {
            System.err.println(e.getMessage());
        } finally {
            try {
                workbook.close();
            } catch (IOException e) { /* nothing to do  */}
        }
    }


    public static class CellData {

        private final Object data; //data of cell
        private final CellType cellType;

        public CellData(Object data, CellType cellType) {
            this.data = data;
            this.cellType = cellType;
        }

        public Object getData() {
            return data;
        }

        public CellType getCellType() {
            return cellType;
        }

        public String valueString() {
            return data.toString();
        }

        @Override
        public String toString() {
            return String.format("{ %s, %s }\n", data.toString(), cellType);
        }
    }


    //writes to new file
    public void writeTable(ExcelTable excelTable, String path) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook(); //create new workbook

        XSSFSheet sheet = workbook.createSheet(); //create new sheet with index 0

        int currentRow = 0;

        for (String keyRow : excelTable.rowKeys()) {
            int columnCount = 3;
            CellType type; //cell type
            Row row = sheet.createRow(currentRow++); //increment rowCount

            for (String keyColumn : excelTable.columnKeys()) {
                type = (columnCount >= 3 && isNumericString(excelTable.getValue(keyRow, keyColumn))) ? CellType.NUMERIC : CellType.STRING;
                Cell cell = row.createCell(columnCount++, type); //increment columnCount
                String value = excelTable.getValue(keyRow, keyColumn);

                if (type == CellType.NUMERIC) {
                    cell.setCellValue(getNumericString(value));
                } else {
                    cell.setCellValue((value == null) ? "" : value); //to be sure
                }
            }
        }

        setUpColumnWidth(sheet);

        try (BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream(path, false))) {
            workbook.write(out); //write to file all workbook
        } finally {
            close(workbook);
        }
    }


    private int getNumericString(String value) {
        return value.isEmpty() ? 0 : Integer.parseInt(value);
    }

    private boolean isNumericString(String value) {
        try {
            if (value.isEmpty()) return true;
            Integer.parseInt(value);
            return true;
        } catch (RuntimeException e) {
            return false;
        }
    }


}
