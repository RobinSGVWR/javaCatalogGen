import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import java.util.Enumeration;

import javax.swing.AbstractButton;
import javax.swing.BoxLayout;
import javax.swing.Box;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JEditorPane;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JTextArea;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reader extends JPanel implements ActionListener {
    private static final long serialVersionUID = 1L;
    
    private JFrame container;
    private ButtonGroup group;
    private Box buttonsBox;
    private JButton startButton;
    private JCheckBox imageCheckBox;
    private JEditorPane label;
    private JProgressBar bar;
    private JTextArea console;
    private Thread t;

    Reader(JFrame container) {
        this.container = container;
        this.label = new JEditorPane("text/html", "");
        this.label.setText("<b>Ajouter vos fichiers .TXT dans le dossier</b> puis lancez le programme");
        this.startButton = new JButton("Lancer le programme");
        this.startButton.addActionListener(this);
        this.imageCheckBox = new JCheckBox("Récupérer les images");
        this.buttonsBox = new Box(BoxLayout.X_AXIS);
        this.buttonsBox.add(this.startButton);
        this.buttonsBox.add(Box.createHorizontalStrut(50));
        this.buttonsBox.add(this.imageCheckBox);
        this.console = new JTextArea("Console");
        this.setLayout(new BoxLayout(this, BoxLayout.Y_AXIS));
        this.add(this.label);
        this.add(this.console);
        this.add(this.buttonsBox);
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

                MongoWriter mw = new MongoWriter(this);
                mw.generateMongo(this.imageCheckBox.isSelected());
            }
        } catch (Exception ex2) {
            this.addToConsole(ex2.toString());
        }

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

    private void enableStart() {
        this.startButton.setEnabled(true);
    }

    private XSSFSheet getSheetByName(XSSFWorkbook workbook, String sheetName) {
        for (Object aWorkbook : workbook) {
            XSSFSheet sheet = (XSSFSheet) aWorkbook;
            if (sheet.getSheetName().equals(sheetName)) {
                return sheet;
            }
        }

        return null;
    }

    void addToConsole(String txt) {
        this.console.append(" - " + txt + "\r\n");
        this.container.pack();
        this.repaint();
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

    public JTextArea getConsole() {
        return console;
    }

    public void setConsole(JTextArea console) {
        this.console = console;
    }
}