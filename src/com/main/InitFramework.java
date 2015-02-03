/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.main;

import java.awt.BorderLayout;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 *
 * @author Mohit-Ankit
 */
public class InitFramework extends JFrame {

    static JFileChooser jExcelPathChooser;
    static JFileChooser jTestCasePathChooser;
    String gExcelPath = "";
    String gTestCaseFolder = "";

    public InitFramework(String title) {
        super(title);


        JPanel panel = new JPanel(new GridBagLayout());
        getContentPane().add(panel, BorderLayout.NORTH);
        GridBagConstraints gbConst = new GridBagConstraints();

        JLabel lbl_UsrName = new JLabel("Excel File Path: ");
        gbConst.gridy = 0;
        gbConst.gridx = 0;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(lbl_UsrName, gbConst);

        final JTextField txt_ExcelPath = new JTextField(20);
        gbConst.gridy = 0;
        gbConst.gridx = 1;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(txt_ExcelPath, gbConst);

        JButton btn_BrowseExcel = new JButton("Browse");
        gbConst.gridy = 0;
        gbConst.gridx = 2;
        //gbConst.gridwidth = 1;
        //gbConst.ipadx = 40;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(btn_BrowseExcel, gbConst);

        JLabel lbl_Psswd = new JLabel("Test Case Folder: ");
        gbConst.gridy = 1;
        gbConst.gridx = 0;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(lbl_Psswd, gbConst);

        final JTextField txt_TesctCaseFldr = new JTextField(20);
        gbConst.gridy = 1;
        gbConst.gridx = 1;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(txt_TesctCaseFldr, gbConst);

        JButton btn_BrowseTestCase = new JButton("Browse");
        gbConst.gridy = 1;
        gbConst.gridx = 2;
        //gbConst.gridwidth = 1;
        //gbConst.ipadx = 40;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(btn_BrowseTestCase, gbConst);

        JButton btn_Import = new JButton("Import");
        gbConst.gridy = 2;
        gbConst.gridx = 0;
        //gbConst.gridwidth = 1;
        gbConst.ipadx = 40;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(btn_Import, gbConst);

        btn_BrowseExcel.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent arg0) {
                jExcelPathChooser = new JFileChooser();

                jExcelPathChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
                jExcelPathChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                
                //Only allow .xls & .xlsx files
                jExcelPathChooser.addChoosableFileFilter(new FileNameExtensionFilter("Microsoft Excel Documents", "xls", "xlsx"));
                jExcelPathChooser.setAcceptAllFileFilterUsed(false);

                int result = jExcelPathChooser.showOpenDialog(null);

                if (result == JFileChooser.APPROVE_OPTION) {
                    File file = jExcelPathChooser.getSelectedFile();
                    if (!file.getName().endsWith("xls")) {
                        JOptionPane.showMessageDialog(null, "Please select only Excel file.", "Error", JOptionPane.ERROR_MESSAGE);
                    } else {
                        gExcelPath = file.getPath();
                        txt_ExcelPath.setText(gExcelPath);
                    }
                } else {
                    System.out.println("Cancel button clicked");
                }
            }
        });


        btn_BrowseTestCase.addActionListener(new ActionListener() {

            @Override
            public void actionPerformed(ActionEvent arg0) {
                jTestCasePathChooser = new JFileChooser();
                jTestCasePathChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                jTestCasePathChooser.setAcceptAllFileFilterUsed(false);
                int result = jTestCasePathChooser.showOpenDialog(null);

                if (result == JFileChooser.APPROVE_OPTION) {
                    if (jTestCasePathChooser.getCurrentDirectory() != null) {
                        gTestCaseFolder = jTestCasePathChooser.getCurrentDirectory().getPath();
                        txt_TesctCaseFldr.setText(gTestCaseFolder);
                    }
                } else {
                    System.out.println("Cancel button clicked");
                }

            }
        });

        //Import Button Listener
        btn_Import.addActionListener(new ActionListener() {

            @Override
            public void actionPerformed(ActionEvent arg0) {
                //Closing Current Screen
                InitFramework.this.dispose();

                //Redirecting to the home screen
                JFrame viewTestCasesFrame = new ViewTestCases("Test Cases", gExcelPath, gTestCaseFolder);
                
                viewTestCasesFrame.setLocationRelativeTo(null);
                viewTestCasesFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
                viewTestCasesFrame.setVisible(true);
                viewTestCasesFrame.setSize(700, 500);
            }
        });
    }

    public static void main(String[] args) {
        // TODO code application logic here
        JFrame mainFrame = new InitFramework("Init Settings");

        mainFrame.setLocationRelativeTo(null);
        mainFrame.setSize(500, 200);
        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mainFrame.setVisible(true);
    }
}
