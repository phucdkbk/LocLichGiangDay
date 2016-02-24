/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.main;

import com.phucdk.lichhoc.object.GeneralData;
import com.phucdk.lichhoc.util.ExcelExportUtil;
import com.phucdk.lichhoc.util.ExcelReadDataUtil;


import java.awt.BorderLayout;
import java.awt.Container;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextField;

public class ExtractInvidualSchedule extends JFrame {

    private final JTextField inputLabel = new JTextField();
    private final JTextField inputFile = new JTextField();
    private final JTextField inputBusyLabel = new JTextField();
    private final JTextField inputBusyFile = new JTextField();
    private final JTextField outputLabel = new JTextField();
    private final JTextField outputFolder = new JTextField();
    private final JTextField result = new JTextField();

    private final JButton open = new JButton("Open general file");
    private final JButton save = new JButton("Process filter");

    public ExtractInvidualSchedule() {
        JPanel p = new JPanel();
        save.addActionListener(new ProcessFilter());
        p.add(save);
        Container cp = getContentPane();
        cp.add(p, BorderLayout.SOUTH);

        p = new JPanel();
        p.setLayout(new GridLayout(9, 1));
        inputLabel.setEditable(false);
        inputLabel.setText("Input file:");
        p.add(inputLabel);
        inputFile.setEditable(false);
        p.add(inputFile);
        
        inputBusyLabel.setEditable(false);
        inputBusyLabel.setText("Input file:");
        p.add(inputBusyLabel);
        inputBusyFile.setEditable(false);
        p.add(inputBusyFile);
        
        open.addActionListener(new OpenL());
        p.add(open);
        outputLabel.setEditable(false);
        outputLabel.setText("Output folder:");
        p.add(outputLabel);
        outputFolder.setEditable(true);
        outputFolder.setText("D:\\Output");
        p.add(outputFolder);
        result.setEditable(false);
        p.add(result);

        cp.add(p, BorderLayout.NORTH);
    }

    class OpenL implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent e) {
            JFileChooser c = new JFileChooser();            
            int rVal = c.showOpenDialog(ExtractInvidualSchedule.this);
            if (rVal == JFileChooser.APPROVE_OPTION) {
                inputFile.setText(c.getCurrentDirectory().toString() + "\\" + c.getSelectedFile().getName());
            }
            if (rVal == JFileChooser.CANCEL_OPTION) {
                inputFile.setText("");
            }
        }
    }

    class ProcessFilter implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent e) {
            GeneralData generalData;
            try {
                generalData = ExcelReadDataUtil.readData(inputFile.getText(), inputBusyFile.getText());
                ExcelExportUtil.exportFile(generalData, outputFolder.getText());
                result.setText("Filter Done!");
            } catch (IOException ex) {
                Logger.getLogger(ExtractInvidualSchedule.class.getName()).log(Level.SEVERE, null, ex);
            } catch (Exception ex) {
                Logger.getLogger(ExtractInvidualSchedule.class.getName()).log(Level.SEVERE, null, ex);
            }

        }
    }

    public static void main(String[] args) {
        run(new ExtractInvidualSchedule(), 550, 350);
    }

    public static void run(JFrame frame, int width, int height) {
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(width, height);
        frame.setVisible(true);
    }
}
