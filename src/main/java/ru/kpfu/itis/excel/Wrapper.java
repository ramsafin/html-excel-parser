package ru.kpfu.itis.excel;

import ru.kpfu.itis.table.ExcelTable;

public class Wrapper {

    private ExcelTable table;

    private ExcelTableService.CellData[][] cellTable;

    public Wrapper(ExcelTable table, ExcelTableService.CellData[][] cellTable) {
        this.table = table;
        this.cellTable = cellTable;
    }

    public ExcelTable getTable() {
        return table;
    }

    public ExcelTableService.CellData[][] getCellTable() {
        return cellTable;
    }
}
