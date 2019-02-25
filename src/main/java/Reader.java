
//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by Fernflower decompiler)
//

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.List;
import javax.swing.AbstractButton;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JEditorPane;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JRadioButton;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reader extends JPanel implements ActionListener {
    private FileChooser c;
    private JButton startButton;
    private JFrame container;
    private ButtonGroup group = new ButtonGroup();
    private JEditorPane label = new JEditorPane("text/html", "");
    private JProgressBar bar;
    private Thread t;
    private JTextField txt;
    private JTextArea console;

    Reader(JFrame container) {
        this.container = container;
        this.c = new FileChooser("Donner la liste de MELNumber", this);
        this.label.setText(" <b> Entrez le titre de l'excel à creer </b>");
        this.txt = new JTextField("new data sheet");
        this.startButton = new JButton("valider");
        this.startButton.addActionListener(this);
        this.console = new JTextArea("Console");
        this.setLayout(new BoxLayout(this, BoxLayout.Y_AXIS));
        this.add(this.c);
        this.add(this.label);
        this.add(this.txt);
        this.add(this.console);

        this.bar = new JProgressBar();
        this.bar.setMaximum(100);
        this.bar.setMinimum(0);
        this.bar.setStringPainted(true);
        this.add(this.bar, "Center");
    }

    public void actionPerformed(ActionEvent e) {
        try {
            if (e.getSource() == this.startButton) {

                this.addToConsole("start button pushed !");
                List<String> list = new ArrayList<String>();
                try {
                    FileInputStream flux = new FileInputStream(this.c.getFile());
                    InputStreamReader lecture = new InputStreamReader(flux);
                    BufferedReader buff = new BufferedReader(lecture);
                    String ligne;
                    while ((ligne = buff.readLine()) != null) {
                        // System.out.println(ligne);
                        list.add(ligne);
                    }
                    buff.close();
                } catch (Exception ex) {
                    System.out.println(e.toString());
                    this.addToConsole(ex.toString());
                }
                bar.setValue(20);

                // ici nous avons toutes les lignes comportant les mel number dans un array
                // il faut maintenant faire un appel a mongowritter pour creer le excel
                // en allant chercher les mel number correspondant dans la db, puis les
                // enregistrer dans le
                // excel selon les caractèristiques
                MongoWriter mw = new MongoWriter(this.txt.getText(), list, this);
                mw.generateMongo();

            }

        } catch (Exception ex2) {
            this.addToConsole(ex2.toString());
        }

        // MongoWriter mongoWriter = new MongoWriter(this.txt.getText(), wb, rowL,
        // this);
        // mongoWriter.generateMongo();

        bar.setValue(100);
    }

    private String getSelectedButtonLabel() {
        Enumeration<AbstractButton> buttons = this.group.getElements();

        while (buttons.hasMoreElements()) {
            AbstractButton button = buttons.nextElement();
            if (button.isSelected()) {
                return button.getText();
            }
        }

        return null;
    }

    public void getStarted() {
        System.out.println("event File triggered");
        this.addToConsole("event File triggered");
        Path currentRelativePath = Paths.get("");
        this.addToConsole("Current desktop path is: " + System.getProperty("user.home") + "/Desktop/");
        String excelFileName = System.getProperty("user.home") + "/Desktop/Generation_de_catalogue/" + txt;
        // name of excel file= "C:/"+txt+".xlsx";
        this.addToConsole("le fichier sera enregistré sous " + excelFileName);

        this.add(this.startButton);
        this.add(this.txt);
        this.container.pack();
        this.addToConsole("file choosen, start button availlable");
    }

    private void deleteRadioButtons() {
    }

    private void enableStart() {
        this.startButton.setEnabled(true);
    }

/*
    private HSSFSheet getSheetByName(HSSFWorkbook workbook, String sheetName) {
    
        for (Object aWorkbook : workbook) {
            HSSFSheet sheet = (HSSFSheet) aWorkbook;
            if (sheet.getSheetName().equals(sheetName)) {
                return sheet;
            }
        }
        
        return null;
    }
*/

    private XSSFSheet getSheetByName(XSSFWorkbook workbook, String sheetName) {

        for (Object aWorkbook : workbook) {
            XSSFSheet sheet = (XSSFSheet) aWorkbook;
            if (sheet.getSheetName().equals(sheetName)) {
                return sheet;
            }
        }

        return null;
    }

/*
    private String cellToString(HSSFCell cell, CellType type) {
        switch (type) {
        case STRING:
            return cell.getStringCellValue();
        case NUMERIC:
            return String.valueOf(cell.getNumericCellValue());
        case ERROR:
            return String.valueOf(cell.getErrorCellValue());
        case BLANK:
            return "[x]";
        case FORMULA:
            return this.cellToString(cell, cell.getCachedFormulaResultTypeEnum());
        default:
            return "----------------------------------------- " + type.toString();
        }
    }
*/

    void addToConsole(String txt) {
        this.console.append(" - " + txt + "\r\n");
        this.container.pack();
        this.repaint();
    }

    public FileChooser getC() {
        return c;
    }

    public void setC(FileChooser c) {
        this.c = c;
    }

    public JButton getStartButton() {
        return startButton;
    }

    public void setStartButton(JButton startButton) {
        this.startButton = startButton;
    }

    public JFrame getContainer() {
        return container;
    }

    public void setContainer(JFrame container) {
        this.container = container;
    }

    public ButtonGroup getGroup() {
        return group;
    }

    public void setGroup(ButtonGroup group) {
        this.group = group;
    }

    public JEditorPane getLabel() {
        return label;
    }

    public void setLabel(JEditorPane label) {
        this.label = label;
    }

    JProgressBar getBar() {
        return bar;
    }

    public void setBar(JProgressBar bar) {
        this.bar = bar;
    }

    public Thread getT() {
        return t;
    }

    public void setT(Thread t) {
        this.t = t;
    }

    public JTextField getTxt() {
        return txt;
    }

    public void setTxt(JTextField txt) {
        this.txt = txt;
    }

    public JTextArea getConsole() {
        return console;
    }

    public void setConsole(JTextArea console) {
        this.console = console;
    }
}