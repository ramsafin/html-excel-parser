package ru.kpfu.itis.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.kpfu.itis.table.ExcelTable;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel - Table converter
 */
public final class ExcelTableService {


    /**
     * Write table to the new .xlsx file
     *
     * @param excelTable - table to be written
     * @param path       - path to file
     * @throws IOException - if occurred an exception while we were writing to the file
     */
    public void writeTable(ExcelTable excelTable, String path) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook(); //create new workbook

        XSSFSheet sheet = workbook.createSheet(); //create new sheet with index 0

        CellStyle cellStyle = createCellStyle(sheet);
        int currentRow = 0;

        for (String keyRow : excelTable.rowKeys()) {
            int columnCount = 0;
            Row row = sheet.createRow(currentRow++); //increment rowCount

            for (String keyColumn : excelTable.columnKeys()) {

                Cell cell = row.createCell(columnCount++, CellType.STRING); //increment columnCount
                cell.setCellStyle(cellStyle); //style of cell
                String value = excelTable.getValue(keyRow, keyColumn);
                cell.setCellValue((value == null) ? "" : value); //to be sure
            }
        }

        setUpColumnWidth(sheet);

        try (BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream(path, false))) {
            workbook.write(out); //write to file all workbook
        } finally {
            close(workbook);
        }
    }


    /**
     * Reads table instance from an existing file
     *
     * @param path - path to an existing file
     * @return - created table
     */
    public ExcelTable readTable(String path) throws IOException {

        XSSFWorkbook workbook = readWorkbook(path);

        XSSFSheet sheet = workbook.getSheetAt(0); //get first sheet

        if (sheet == null) throw new IllegalArgumentException("There is no sheets in the document");

        final int columns = getColumnRealCount(sheet.getRow(0), false);
        final int rows = sheet.getLastRowNum() + 1;

        ExcelTable table = new ExcelTable(rows, columns);

        sheet.forEach(cells -> {
            if (cells.getFirstCellNum() == -1) return; //doesn't contain any cells

            List<String> values = new ArrayList<>(columns); //cell values of row

            for (int k = 0; k < columns; k++) {
                Cell cell = cells.getCell(k);
                values.add(readCellValue(cell));
            }

            table.addRow(values.get(0), values.toArray(new String[0])); //in each row the key is data-id (first cell)
        });
        close(workbook); //don't forget to close workbook

        return table;
    }


    /**
     * Returns real count of columns in the row
     *
     * @param row - excell file row
     * @return integer, real count of columns
     */
    private int getColumnRealCount(Row row, boolean hasEmpty) {

        if (row == null) return 0; //if there is no row

        int i; //to return

        for (i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell == null) return i;
            if (!hasEmpty && cell.getStringCellValue().isEmpty()) return i;
        }

        return row.getLastCellNum();
    }


    /**
     * Reads cell's value depends on data type in it
     *
     * @param cell - document cell
     * @return String representation of value in cell
     */
    @SuppressWarnings("deprecated")
    private String readCellValue(Cell cell) {

        if (cell == null) return "";

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_BLANK:
                return "";
            case Cell.CELL_TYPE_BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case Cell.CELL_TYPE_NUMERIC:
                return Integer.toString((int) cell.getNumericCellValue());
            default:
                return "undefined";
        }
    }


    /**
     * Reads workbook
     *
     * @param path - path to excel-file
     * @return xlsx workbook representation
     * @throws IOException - file read expetion
     */
    private XSSFWorkbook readWorkbook(String path) throws IOException {

        try (BufferedInputStream in = new BufferedInputStream(new FileInputStream(path))) {
            return new XSSFWorkbook(in);
        }
    }


    private void setUpColumnWidth(Sheet sheet) {
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            sheet.autoSizeColumn(i);
        }

    }

    private static final short DEFAULT_FONT_SIZE = 9;
    private static final short DEFAULT_INDENT = 1;


    private static CellStyle createCellStyle(Sheet sheet) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();

        //font
        Font font = sheet.getWorkbook().createFont();
        font.setFontHeightInPoints(DEFAULT_FONT_SIZE);
        cellStyle.setFont(font);

        //alignment
        cellStyle.setAlignment(HorizontalAlignment.LEFT);

        //indention
        cellStyle.setIndention(DEFAULT_INDENT);
        return cellStyle;
    }


    /**
     * Closes workbook
     */
    private void close(Workbook workbook) {
        if (workbook != null) {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                try {
                    workbook.close();
                } catch (IOException e) { /* nothing to do, can't close this shit */ }
            }
        }
    }

}
