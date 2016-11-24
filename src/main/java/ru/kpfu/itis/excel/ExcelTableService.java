package ru.kpfu.itis.excel;

import com.google.common.math.DoubleMath;
import com.google.common.primitives.Doubles;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;
import ru.kpfu.itis.table.ExcelTable;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;

import static org.apache.poi.ss.usermodel.CellType.*;


/**
 * Excel - Table converter
 */
public final class ExcelTableService {

    private static final String BLANK_VALUE = "";

    public ExcelTableService() {
        ZipSecureFile.setMinInflateRatio(0.0001);
    }

    public void writeTwoTables(CellData[][] tableLeft, ExcelTable tableRight, String path) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet();

        int rightRowsMax = tableRight.rowCount();

        tableRight.rowKeys().forEach(new Consumer<String>() {
            private int rowIdx = 0;

            @Override
            public void accept(String rowKey) {

                Row row = sheet.createRow(rowIdx);

                for (int i = 0; i < 3; i++) {
                    Cell cell = row.createCell(i);
                    setupCell(cell, workbook);
                    writeCellValue1(cell, tableLeft[rowIdx][i]); //write rowIdx row, 3 columns (0 ... 2)
                }

                for (int i = 0; i < tableRight.columnCount(); i++) {
                    CellData cellData = getCellData(tableRight.getValue(rowKey, tableRight.getColumnKey(i)), i, 3);
                    Cell cell = row.createCell(i + 3, cellData.getCellType());
                    setupCell(cell, workbook);
                    if (cellData.getCellType() == NUMERIC) {
                        if (cellData.isInteger()) {
                            cell.setCellValue(cellData.getIntData());
                        } else {
                            cell.setCellValue(cellData.getDoubleData());
                        }
                    } else {
                        cell.setCellValue(cellData.getStringValue());
                    }
                }

                rowIdx++;

                if (rowIdx == rightRowsMax && rowIdx < tableLeft.length) {
                    for (int i = rowIdx; i < tableLeft.length; i++) {
                        row = sheet.createRow(i);
                        for (int j = 0; j < 3; j++) {
                            Cell cell = row.createCell(j);
                            setupCell(cell, workbook);
                            writeCellValue1(cell, tableLeft[i][j]);
                        }
                    }
                }

            }
        });

        setUpColumnWidth(sheet);

        try (BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream(path, false))) {
            workbook.setWorkbookType(XSSFWorkbookType.XLSX);
            workbook.write(out); //write to file all workbook
        } finally {
            close(workbook);
        }
    }


    private void writeCellValue1(Cell cell, CellData cellData) {
        cell.setCellType(cellData.getCellType()); // set cell type
        switch (cellData.getCellType()) {
            case STRING:
                cell.setCellValue(cellData.getStringValue());
                break;
            case NUMERIC:
                cell.setCellValue(cellData.getDoubleData());
                break;
            case FORMULA:
                cell.setCellFormula(cellData.getStringValue());
                break;
            case BLANK:
                cell.setCellValue(BLANK_VALUE);
                break;
            case BOOLEAN:
                cell.setCellValue(cellData.getBooleanValue());
                break;
            case ERROR:
                cell.setCellErrorValue(cellData.getErrorValue());
                break;
            default:
                throw new IllegalArgumentException("There is no such type of cell");
        }
    }



    //reads first table
    public CellData[][] readTable1(String path) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(path);
        XSSFSheet sheet = workbook.getSheetAt(0);
        if (sheet == null) throw new IllegalArgumentException("There is no sheets in the document");
        final int rows = sheet.getLastRowNum() + 1; //forgot +1 !!!
        CellData[][] cellData = new CellData[rows][3];
        for (int i = 0; i < rows; i++) {

            Row row = sheet.getRow(i);

            for (int k = 0; k < 3; k++) {
                Cell cell = row.getCell(k);
                cellData[i][k] = readCellValue(cell);
            }
        }
        close(workbook);
        return cellData;
    }


    //reads second table
    public ExcelTable readTable2(String path) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(path);
        XSSFSheet sheet = workbook.getSheetAt(0); //get first sheet
        if (sheet == null) throw new IllegalArgumentException("There is no sheets in the document");
        final int rows = sheet.getLastRowNum() + 1;
        final int columns = sheet.getRow(0).getLastCellNum(); //read all columns, not only 4
        ExcelTable table = new ExcelTable(rows, columns - 3);
        sheet.forEach(row -> {
            List<String> values = new ArrayList<>(columns - 3); //cell values of row
            for (int k = 3, colIdx = k - 3; k < columns; k++, colIdx++) {
                Cell cell = row.getCell(k);
                CellData value = readCellValue(cell);
                if (k == 3 && value.getCellType() == BLANK) return;
                if (value.isInteger() && k >= 6) {
                    values.add(Integer.toString(value.getIntData()));
                } else if (value.isInteger() && k == 5) {
                    values.add(Integer.toString(value.getIntData()));
                } else {
                    values.add(value.getStringValue());
                }
            }
            table.addRow(values.get(0), values.toArray(new String[0])); //in each row the key is data-id (first cell)
        });
        close(workbook);
        return table;
    }


    private CellData readCellValue(Cell cell) {
        if (cell == null) return new CellData(BLANK_VALUE, BLANK);
        switch (cell.getCellTypeEnum()) {
            case STRING:
                return new CellData(cell.getStringCellValue(), STRING);
            case NUMERIC:
                if (isInteger(cell.getNumericCellValue())) {
                    return new CellData(cell.getNumericCellValue(), NUMERIC, true);
                }
                return new CellData(cell.getNumericCellValue(), NUMERIC);
            case BLANK:
                return new CellData(BLANK_VALUE, BLANK);
            case BOOLEAN:
                return new CellData(cell.getBooleanCellValue(), BOOLEAN);
            case FORMULA:
                return new CellData(cell.getCellFormula(), FORMULA);
            case ERROR:
                return new CellData(cell.getErrorCellValue(), ERROR);
            default:
                throw new IllegalArgumentException("There is no such type of cell");
        }
    }


    private boolean isInteger(Double d) {
        return DoubleMath.isMathematicalInteger(d);
    }

    private void setUpColumnWidth(Sheet sheet) {
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            sheet.autoSizeColumn(i);
        }
        sheet.setHorizontallyCenter(true);
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
        private boolean isInteger;

        public CellData(Object data, CellType cellType) {
            this(data, cellType, false);
        }

        public CellData(Object data, CellType cellType, boolean isInteger) {
            this.data = data;
            this.cellType = cellType;
            this.isInteger = isInteger;
        }

        public Object getData() {
            return data;
        }

        public int getIntData() {
            return ((Double) data).intValue();
        }

        public double getDoubleData() {
            return (Double) data;
        }

        public String getStringValue() {
            return Objects.toString(data, BLANK_VALUE);
        }

        public Boolean getBooleanValue() {
            return (Boolean) data;
        }

        public Byte getErrorValue() {
            return (Byte) data;
        }

        public boolean isInteger() {
            return isInteger;
        }


        public CellType getCellType() {
            return cellType;
        }

        @Override
        public String toString() {
            return String.format("{ %s, %s, %b }", data.toString(), cellType, isInteger);
        }
    }


    //writes 3 + 4 columns only, creates new file
    public void writeTable(ExcelTable excelTable, String path) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(); //create new workbook
        XSSFSheet sheet = workbook.createSheet(); //create new sheet with index 0

        int currentRow = 0;

        for (String keyRow : excelTable.rowKeys()) {
            int columnCount = 3;
            Row row = sheet.createRow(currentRow++); //increment rowCount

            for (String keyColumn : excelTable.columnKeys()) {
                CellData cellData = getCellData(excelTable.getValue(keyRow, keyColumn), columnCount, 7);
                Cell cell = row.createCell(columnCount++, cellData.getCellType()); //increment columnCount
                if (cellData.getCellType() == NUMERIC) {
                    if (cellData.isInteger()) {
                        cell.setCellValue(cellData.getIntData());
                    } else {
                        cell.setCellValue(cellData.getDoubleData());
                    }
                } else {
                    cell.setCellValue(cellData.getStringValue());
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

    private void setupCell(Cell cell, Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(cellStyle);
    }

    private CellData getCellData(String value, int columnCount, int fromColumnInt) {
        if (columnCount >= fromColumnInt) {
            Double d = Doubles.tryParse(value);
            if (d != null && isInteger(d)) return new CellData(d, NUMERIC, true);
            if (d != null) return new CellData(d, NUMERIC);
        }
        return new CellData(value, STRING);
    }

}
