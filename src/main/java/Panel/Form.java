package Panel;

import ExelLogic.UpdateCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

public class Form extends JFrame{
    private JButton startButton;
    public JProgressBar progressBar1;
    private JFrame f;
    int i=0,num=0;

    public Form() {

        f=new JFrame();
        startButton = new JButton("Start");
        progressBar1 = new JProgressBar();
        f.setSize(400,500);//400 width and 500 height
        f.setLayout(null);//using no layout managers
        f.setVisible(true);//making the frame visible
        startButton.setBounds(150,450,100, 40);
        f.add(startButton);
        f.add(progressBar1);

        startButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                System.out.println("Test");
                try {
                    UpdateCell.ScanDoc();
                } catch (IOException e1) {
                    e1.printStackTrace();
                } catch (InvalidFormatException e1) {
                    e1.printStackTrace();
                }
                JOptionPane.showMessageDialog(null,"File Scan Done");
            }
        });
    }

    public static void main(String[] args) {
        new Form();
    }
}
