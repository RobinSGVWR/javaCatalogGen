import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.*;

public class FileChooser extends JPanel implements ActionListener {
    private static final long serialVersionUID = 1L;

    private JButton openButton;
    private JTextField path;
    private JFileChooser fc;
    private File file;
    private JLabel label;
    private Reader reader;

    FileChooser(String l, Reader reader) {

        super(new FlowLayout());
        this.reader = reader;

        path = new JTextField(15);
        fc = new JFileChooser();

        label = new JLabel(l);
        openButton = new JButton("Parcourir");
        openButton.addActionListener(this);

        this.add(label);
        this.add(path);
        this.add(openButton);
    }

    FileChooser(String l) {
        super(new FlowLayout());

        path = new JTextField(15);
        fc = new JFileChooser();

        label = new JLabel(l);
        openButton = new JButton("Parcourir");
        openButton.addActionListener(this);

        this.add(label);
        this.add(path);
        this.add(openButton);
    }

    public JFileChooser getJFileChooser() {
        return fc;
    }

    public void actionPerformed(ActionEvent e) {

        if (e.getSource() == openButton) {

            if (fc.showOpenDialog(FileChooser.this) == JFileChooser.APPROVE_OPTION) {
                file = fc.getSelectedFile();
                path.setText(file.getPath());
                reader.getStarted();
                reader.addToConsole("Into the file chooser");
            }
            path.setCaretPosition(path.getDocument().getLength());
        }
    }

    File getFile() {
        return file;
    }
}