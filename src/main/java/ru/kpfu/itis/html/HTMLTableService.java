package ru.kpfu.itis.html;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import ru.kpfu.itis.table.ExcelTable;

import java.io.File;
import java.io.IOException;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.function.Consumer;

/**
 * Class that converts html page to ExcelTable
 */
public final class HTMLTableService {

    /**
     * Data table rows selector class (XPath)
     **/
    private static final String DATA_TABLE_CLASS = ".b2b-basket-prods-list table tr";


    /**
     * Creates ExcelTable from and existing HTML file
     *
     * @param path - path to HTML
     * @return ExcelTable instance converted form HTML
     * @throws IOException (SelectorParseException) - exception while loading DOM or if selector is invalid
     */
    public ExcelTable createTable(String path) throws IOException {

        Document document = loadDocument(path);

        Elements tableRows = document.select(DATA_TABLE_CLASS);

        return buildTable(tableRows);
    }


    /**
     * Builds table form selected rows
     *
     * @param tableRows - DOM rows
     * @return ExcelTable instance
     */
    private ExcelTable buildTable(Elements tableRows) {

        String[] headers = createHeaders(); //create headers

        final ExcelTable table = new ExcelTable(tableRows.size(), headers.length); //create table

        table.addRow(ExcelTable.HEADERS_KEY, headers); //add headers to table

        //iterating through DOM rows
        tableRows.forEach(new Consumer<Element>() {
            private byte rowspan = 0; //HTML rowspan attribute
            private String name; //product name

            @Override
            public void accept(Element row) {
                if (rowspan == 0) {
                    rowspan = Byte.parseByte(row.select(".c1").attr("rowspan"));
                    name = cleanText(row.select(".c1"));
                }
                String dataId = row.attr("data-id");
                --rowspan; //decrement rowspan

                //add row to table
                table.addRow(dataId, new String[]{
                        dataId, //data-id
                        name, //name
                        row.select(".c2").text(), //tone number
                        row.select(".c9 input").attr("max") //max count
                });
            }
        });

        return table;
    }


    private String cleanText(Elements text) {
        return text.text().replace("PAESE", "").replace("  ", "").replace("Paese", "").replace("paese", "");
    }


    /**
     * Loads document instance form HTML file
     *
     * @param path - path to an existing file
     * @return DOM (Document Object Model) representing HTML
     * @throws IOException (InvalidPathException) - if the file could not be found, or read,
     *                     or if the charsetName is invalid.
     */
    private Document loadDocument(String path) throws IOException {

        File inputHtml = Paths.get(path).toFile();
        return Jsoup.parse(inputHtml, null); //try to parse with http-equiv, otherwise UTF-8
    }


    /**
     * !!! NOTE: headers count is 4 (data-id, name, tone number, max count (current date) )
     **/
    private static String[] createHeaders() {
        String[] values = new String[ColumnHeaders.values().length];
        int i = 0;

        for (ColumnHeaders headers : ColumnHeaders.values()) {
            values[i++] = headers.getColumnName();
        }
        values[i - 1] = String.format("%s", LocalDateTime.now()
                .format(DateTimeFormatter.ofPattern("dd.MM.YYYY HH:mm:ss")));

        return values;
    }


    /**
     * Headers enum
     */
    public enum ColumnHeaders {

        DATA_ID("data-id"), NAME("название"), TONE_NUMBER("номер тона"), MAX_COUNT();

        private String name;

        ColumnHeaders() { }

        ColumnHeaders(String header) {
            this.name = header;
        }

        public String getColumnName() {
            return name;
        }
    }


}
