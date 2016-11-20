package ru.kpfu.itis;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import ru.kpfu.itis.excel.ExcelTableService;
import ru.kpfu.itis.html.HTMLTableService;
import ru.kpfu.itis.table.ExcelTable;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.io.IOException;

public class Application extends JFrame {

    /**
     * HTML and Excel Services
     **/
    private ExcelTableService excelTableConverter;
    private HTMLTableService htmlToExcelTableConverter;

    /**
     * Files
     */
    private File htmlFile; //html to be converted
    private File excelFile; //excel to be read
    private File newExcelFile; //excel to be created


    public Application() {
        super("HTML -> XLSX"); //window titile
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setResizable(false);
        this.setBounds(500, 300, 400, 300); //size of window
        createGUI();
        this.setVisible(true); //make it visivle

        //init services
        excelTableConverter = new ExcelTableService();
        htmlToExcelTableConverter = new HTMLTableService();
    }


    public void createGUI() {
        JButton updateExcelBtn = new JButton("Обновить таблицу");

        JFileChooser htmlFileChooser = new JFileChooser();
        htmlFileChooser.setAcceptAllFileFilterUsed(false);
        htmlFileChooser.addChoosableFileFilter(new FileNameExtensionFilter("HTML файл", "html"));

        JPanel mainPanel = new JPanel(null);

        JLabel chooseHtmlLabel = new JLabel("Файл не выбран");
        chooseHtmlLabel.setBounds(20, 20, 210, 30);
        mainPanel.add(chooseHtmlLabel);

        JButton chooseHtmlBtn = new JButton("Выбрать html");
        chooseHtmlBtn.setBounds(230, 20, 150, 30);
        mainPanel.add(chooseHtmlBtn);
        chooseHtmlBtn.addActionListener(e -> {
            int retVal = htmlFileChooser.showOpenDialog(mainPanel);
            if (retVal == JFileChooser.APPROVE_OPTION) {
                htmlFile = htmlFileChooser.getSelectedFile();
                chooseHtmlLabel.setText(htmlFile.getName());
            }
        });

        JFileChooser excelFileChooser = new JFileChooser();
        excelFileChooser.setAcceptAllFileFilterUsed(false);
        excelFileChooser.addChoosableFileFilter(new FileNameExtensionFilter("Документ Open XML Microsoft Excel", "xlsx"));

        JLabel chooseExcelLabel = new JLabel("Выбери *.xlsx файл");
        chooseExcelLabel.setBounds(20, 60, 210, 30);
        mainPanel.add(chooseExcelLabel);

        JButton chooseExcelBtn = new JButton("Выбрать excel");
        chooseExcelBtn.setBounds(230, 60, 150, 30);
        mainPanel.add(chooseExcelBtn);
        chooseExcelBtn.addActionListener(e -> {
            int retVal = excelFileChooser.showOpenDialog(mainPanel);
            if (retVal == JFileChooser.APPROVE_OPTION) {
                excelFile = excelFileChooser.getSelectedFile();
                chooseExcelLabel.setText(excelFile.getName());
            }
        });



        JButton createNewExcelBtn = new JButton("Создать новый excel из html");
        createNewExcelBtn.setBounds(90, 120, 210, 40);
        mainPanel.add(createNewExcelBtn);
        createNewExcelBtn.addActionListener(e -> {

            switchButtons(false, chooseHtmlBtn, createNewExcelBtn); //disable  buttons

            new Thread(() -> {
                if (htmlFile != null) {
                    JFileChooser save = new JFileChooser();
                    if (JFileChooser.APPROVE_OPTION == save.showSaveDialog(mainPanel)) {
                        newExcelFile = save.getSelectedFile();
                        try {
                            ExcelTable table = htmlToExcelTableConverter.createTable(htmlFile.getPath());
                            excelTableConverter.writeTable(table, newExcelFile.getPath());

                            htmlFile = null; newExcelFile = null;

                            chooseHtmlLabel.setText("Файл не выбран");
                            chooseExcelLabel.setText("Выбери *.xlsx файл");

                        } catch (IOException e1) {
                            JOptionPane.showMessageDialog(mainPanel, e1.getMessage(), "Ошибка", JOptionPane.ERROR_MESSAGE);
                            switchButtons(true, chooseHtmlBtn, createNewExcelBtn);
                            return;
                        }
                        JOptionPane.showMessageDialog(mainPanel, "Файл успешно сконвертирован!", "Успех", JOptionPane.INFORMATION_MESSAGE);
                    }

                } else {
                    JOptionPane.showMessageDialog(mainPanel, "Выбери html файл!", "Ошибка", JOptionPane.WARNING_MESSAGE);
                }
            }).start();
            switchButtons(true, chooseHtmlBtn, createNewExcelBtn);
        });


        updateExcelBtn.setBounds(90, 200, 210, 40);
        mainPanel.add(updateExcelBtn);

        updateExcelBtn.addActionListener(e -> {

            switchButtons(false, chooseHtmlBtn, chooseExcelBtn, updateExcelBtn, createNewExcelBtn);

            new Thread(() -> {
                if (excelFile != null && htmlFile != null) {
                    try {

                        ExcelTable excelTable = htmlToExcelTableConverter.createTable(htmlFile.getPath());
                        ExcelTable excelTable1 = excelTableConverter.readTable(excelFile.getPath());

                        excelTable1.merge(excelTable, 3);
                        excelTableConverter.writeTable(excelTable1.sort(1), excelFile.getPath());

                    } catch (IOException e1) {
                        JOptionPane.showMessageDialog(mainPanel, e1.getMessage(), "Ошибка", JOptionPane.ERROR_MESSAGE);
                        switchButtons(true, chooseHtmlBtn, chooseExcelBtn, updateExcelBtn, createNewExcelBtn);
                        return;
                    }

                    htmlFile = null; excelFile = null; newExcelFile = null;

                    chooseHtmlLabel.setText("Файл не выбран");
                    chooseExcelLabel.setText("Выбери *.xlsx файл");

                    JOptionPane.showMessageDialog(mainPanel, "Данные успешно обновлены!", "Успех", JOptionPane.INFORMATION_MESSAGE);

                } else {
                    JOptionPane.showMessageDialog(mainPanel, "Выбери html и excel файлы!", "Ошибка", JOptionPane.WARNING_MESSAGE);
                }

                switchButtons(true, chooseHtmlBtn, chooseExcelBtn, updateExcelBtn, createNewExcelBtn);

            }).start();

        });

        this.getContentPane().add(mainPanel);
    }

    private void switchButtons(boolean state, JButton ... buttons) {
        for (JButton b : buttons) {
            b.setEnabled(state);
        }
    }


    public static void main(String[] args) throws IOException, InvalidFormatException {
        SwingUtilities.invokeLater(Application::new);
    }

}
