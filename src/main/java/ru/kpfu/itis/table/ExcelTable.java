package ru.kpfu.itis.table;

import com.google.common.collect.HashBasedTable;
import com.google.common.collect.Lists;
import com.google.common.collect.Ordering;
import com.google.common.collect.Table;

import java.util.*;
import java.util.function.Consumer;

/**
 * This class represents an excel table
 */
public final class ExcelTable {


    /**
     * Constants
     * Default values of table size
     **/
    public static final String HEADERS_KEY = "data-id";

    private static final int EXPECTED_ROWS = 600;
    private static final int EXPECTED_COLUMNS = 4;


    /**
     * Table structure, Map<R, Map<C, V>>, where R - row, C - column, V - value
     **/
    private Table<String, String, String> table;


    /**
     * Table's generated column keys
     * By default (0, 1, 2 ...)
     **/
    private List<String> generatedColumnKeys;


    /**
     * Default constructor
     * Creates table with default size
     */
    public ExcelTable() {
        this(EXPECTED_ROWS, EXPECTED_COLUMNS);
    }


    /**
     * Constructor
     * Creates HashBasedTable
     * @param rows - rows count
     * @param columns - columns
     */
    public ExcelTable(int rows, int columns) {
        this.table = HashBasedTable.create(rows, columns);
        init(columns);
    }


    public Table<String, String, String> getTable() {
        return this.table;
    }


    public List<String> getGeneratedColumnKeys() {
        return this.generatedColumnKeys;
    }


    /**
     * Returns cell's value
     * @param row - row counted from 0
     * @param column - column from 0
     * @return value of the cell
     */
    public String getValue(String row, String column) {
        return this.table.get(row, column);
    }


    /**
     * Returns column count int the table
     *
     * @return integer column count
     */
    public int columnCount() {
        return this.table.columnKeySet().size();
    }


    /**
     * Returns row count in the table
     *
     * @return integer row count
     */
    public int rowCount() {
        return this.table.rowKeySet().size();
    }


    /**
     * Returns columns key's / id's
     *
     * @return set of keys
     */
    public Set<String> columnKeys() {
        return this.table.columnKeySet();
    }


    /**
     * Returns rows key's / id's
     *
     * @return set of keys
     */
    public Set<String> rowKeys() {
        return this.table.rowKeySet();
    }


    /**
     * Adds row to an existing table with determined columnKeys
     *
     * @param rowKey - row key /id
     * @param values - String[] values
     */
    public void addRow(String rowKey, String[] values) {
        addRowWithCheck(rowKey, values);
    }


    /**
     * Adds column to an existing table
     * Column name will be generated automatically
     *
     * @param values - values in the column
     */
    public void addColumn(String[] values) {

        addColumnWithCheck(getNextColumnKey(), values);
    }


    /**
     * Merges with mergeTable
     *
     * @param mergeTable - table to be merged with
     * @param columns    - columns count to be merged (0 - none, 1 - first column,
     *                   2 - first two columns, ...)
     */
    public void merge(ExcelTable mergeTable, int columns) {

        String[] newLastColumn = mergeLastColumns(mergeTable); //create new max count column

        mergeTable.rowKeys().forEach(rowKey -> { //for each row key
            if (isRowExist(rowKey)) {

                //if row exists in the old table, merge with it
                mergeRow(mergeTable.getTable().row(rowKey), rowKey, columns);

            } else {

                //else add first column cells of row
                addRowWithoutCheck(rowKey, Arrays.copyOfRange(mergeTable.getTable()
                        .row(rowKey).values().toArray(new String[0]), 0, columns));
            }
        });

        addColumn(newLastColumn); //add last column with max count
    }


    /**
     * Merges last columns of the tables
     * @param table - table to be merged with
     * @return merged last column
     */
    private String[] mergeLastColumns(ExcelTable table) {
        List<String> newColumnValues = new ArrayList<>(rowCount());

        Map<String, String> thisCol = getLastColumn(); // last column of current table
        Map<String, String> tableCol = table.getLastColumn(); // last column of the table

        //firstly go through current table's column row keys
        for (String rowKey : thisCol.keySet()) {
            if (table.isRowExist(rowKey)) {
                newColumnValues.add(table.getValue(rowKey, table.getLastColumnKey())); //update it
            } else {
                newColumnValues.add(""); //clear, make it empty
            }
        }

        //go through table
        for (String rowKey : tableCol.keySet()) {
            if (!isRowExist(rowKey)) {
                newColumnValues.add(tableCol.get(rowKey)); //add not existing values
            }
        }

        return newColumnValues.toArray(new String[0]);
    }


    /**
     * Returns last column of the table
     * @return row key, value map of last column
     */
    private Map<String, String> getLastColumn() {
        return table.column(generatedColumnKeys.get(generatedColumnKeys.size() - 1));
    }


    /**
     * Returns last column key (0, 1, ...)
     * @return last column key
     */
    private String getLastColumnKey() {
        return generatedColumnKeys.get(generatedColumnKeys.size() - 1);
    }


    /**
     * Check if row with some key exists int this table
     * @param rowKey - row's key to be checked
     * @return true or false
     */
    private boolean isRowExist(String rowKey) {
        return table.containsRow(rowKey);
    }


    /**
     * Merges first 'columns' columns of row
     *
     * @param row     - row to be merged with
     * @param columns - columns count to be merged
     */
    private void mergeRow(Map<String, String> row, String rowKey, int columns) {
        if (row.size() <= columns) {
            throw new IllegalArgumentException("Column count is greater than row size");
        }

        Map<String, String> tableRow = this.table.row(rowKey); //get row of table by key

        for (String columnKey : row.keySet()) {
            if (columns-- == 0) return;
            tableRow.put(columnKey, row.get(columnKey));
        }
    }


    private void addRowWithCheck(String rowKey, String[] values) {

        if (this.generatedColumnKeys.size() > values.length) {
            throw new IllegalArgumentException(String.format("Expected : %d, but it is %d value(s)",
                    this.generatedColumnKeys.size(), values.length));
        }

        for (int i = 0; i < this.generatedColumnKeys.size(); i++) {
            table.put(rowKey, this.generatedColumnKeys.get(i), values[i]);
        }

    }


    private void addRowWithoutCheck(String rowKey, String[] values) {

        if (values.length < columnCount()) {
            int old = values.length;
            values = Arrays.copyOfRange(values, 0, columnCount());
            Arrays.fill(values, old, columnCount(), "");
        }

        for (int i = 0; i < values.length; i++) {
            try {

                table.put(rowKey, this.generatedColumnKeys.get(i), values[i]);
            } catch (NullPointerException e) {
                System.out.println(String.format("Index %d", i));
            }
        }
    }


    private void addColumnWithCheck(String columnKey, String[] values) {
        if (this.generatedColumnKeys.contains(columnKey)) {
            throw new IllegalArgumentException(String.format("Column key '%s' already exists", columnKey));
        }
        if (rowCount() > values.length) {
            throw new IllegalArgumentException(String.format("Expected : %d, but it is %d value(s)",
                    this.generatedColumnKeys.size(), values.length));
        }
        //for each row key add 'columnKey' column's values
        rowKeys().forEach(new Consumer<String>() {
            int i = 0;
            @Override
            public void accept(String rowKey) {
                table.put(rowKey, columnKey, values[i++]);
            }
        });
        this.generatedColumnKeys.add(columnKey); //update columns list
    }


    private void init(int columns) {
        this.generatedColumnKeys = new ArrayList<>(columns);
        for (Integer i = 0; i < columns; i++) {
            this.generatedColumnKeys.add(i, i.toString());
        }
    }


    private String getNextColumnKey() {
        return Integer.toString(generatedColumnKeys.size());
    }


    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();

        //all meta information
        sb.append("====== meta info =======\n");
        sb.append(String.format("Row count - %d\n", rowCount())); //row count
        sb.append(String.format("Column count - %d\n", columnCount())); //column count

        //all column keys and row keys
        sb.append("====== row and column keys =======\n");
        sb.append(table.rowKeySet().toString()).append("\n"); //row keys
        sb.append(table.columnKeySet().toString()).append("\n"); //column keys

        //all data row by row
        sb.append("====== rows data =======\n");
        table.rowKeySet().forEach(rowKey -> sb.append(table.row(rowKey)).append("\n"));

        return sb.toString();
    }


    public ExcelTable sort(int column) {

        List<Table.Cell<String, String, String>> filteredList = Lists.newArrayList();

        Set<Table.Cell<String, String, String>> cells = table.cellSet();

        String columnKey = generatedColumnKeys.get(column); // get columnKey
        String columnHeader = table.column(columnKey).get(HEADERS_KEY);

        for (Table.Cell<String, String, String> cell : cells) {
            if (columnKey.equals(cell.getColumnKey())) {
                filteredList.add(cell);
            }
        }
        Collections.sort(filteredList, compareByColumn(columnHeader));
        ExcelTable excelTable = new ExcelTable(rowCount(), columnCount()); //new table, sorted

        filteredList.forEach(cell -> excelTable.addRow(cell.getRowKey(), rowValues(cell.getRowKey())));

        return excelTable;
    }


    /**
     * Comparator that compare cells
     * @param column - column name
     * @return Comparator
     */
    public static Ordering<Table.Cell<String, String, String>> compareByColumn(final String column) {

        return new Ordering<Table.Cell<String, String, String>>() {

            @Override
            public int compare(
                    Table.Cell<String, String, String> cell1,
                    Table.Cell<String, String, String> cell2
            ) {
                if (column.equals(cell1.getValue())) return 1;
                if (column.equals(cell2.getValue())) return 1;

                String cell1Val =  cell1.getValue();
                String cell2Val = cell2.getValue();

                if (cell1Val == null && cell2Val == null) return 0;
                if (cell1Val == null) return -1;
                if (cell2Val == null) return 1;

                return cell1Val.compareTo(cell2Val);
            }
        };
    }


    /**
     * Gets all row values
     * @param rowKey - row's key
     * @return row values array
     */
    private String[] rowValues(String rowKey) {

        List<String> values = new ArrayList<>(columnCount());

        table.row(rowKey).forEach((columnKey, value) -> values.add(value));

        return values.toArray(new String[0]);
    }

}
