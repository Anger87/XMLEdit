package Panel;

import ExelLogic.UpdateCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

public class Form extends JFrame {
    private JButton startButton;
    public JProgressBar progressBar1;
    private JFrame f;
    int i = 0, num = 0;
    String filePath = "";
    public static String fileName;

    public Form() {

        f = new JFrame();
        startButton = new JButton("Start");
        progressBar1 = new JProgressBar();
        f.setSize(400, 300);//400 width and 500 height
        f.setLayout(null);//using no layout managers
        f.setVisible(true);//making the frame visible
        startButton.setBounds(150, 150, 100, 40);
        f.add(startButton);
        f.add(progressBar1);

        final JLabel label = new JLabel("Selected file");
        label.setAlignmentX(CENTER_ALIGNMENT);
        label.setBounds(50, 50, 200, 40);
        f.add(label);

        JButton button = new JButton("Select the file");
        button.setAlignmentX(CENTER_ALIGNMENT);
        button.setBounds(50, 100, 200, 40);
        f.add(button);

        button.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileopen = new JFileChooser(System.getProperty("user.dir"));
                int ret = fileopen.showDialog(null, "Open the file");
                if (ret == JFileChooser.APPROVE_OPTION) {
                    File file = fileopen.getSelectedFile();
                    fileName = file.getName();
                    label.setText(fileName);
                    filePath = file.getAbsolutePath();
                }
            }
        });
/*
        startButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                try {
                    if (filePath.length() > 1) {
                        System.out.println("in startButton");
                        UpdateCell.ScanDoc(filePath);
                        JOptionPane.showMessageDialog(null, "File Scan Done");
                    } else {
                        JOptionPane.showMessageDialog(null, "Select the file");
                    }
                } catch (IOException e1) {
                    e1.printStackTrace();
                } catch (InvalidFormatException e1) {
                    e1.printStackTrace();
                }

            }
        });*/


    }
}
