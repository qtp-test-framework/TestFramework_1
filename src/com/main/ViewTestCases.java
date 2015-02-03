package com.main;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Vector;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ViewTestCases extends JFrame {

    static String gExcelPath;
    static String gTestCaseFolder;
    static Vector headers = new Vector();
    // Model is used to construct JTable 
    static DefaultTableModel model = null;
    // data is Vector contains Data from Excel File 
    static Vector data = new Vector();
    static JFileChooser jChooser;
    static int jx;
    static int trc;
    static JTable table;
    static JScrollPane scroll;

    public ViewTestCases(String vTitle, String vExcelPath, String vTestCaseFolder) {
        super(vTitle);
        try {
            gExcelPath = vExcelPath;
            gTestCaseFolder = vTestCaseFolder;

            fillData();
            table = new JTable();
            table.setAutoCreateRowSorter(true);
            model = new DefaultTableModel(data, headers);
            table.setModel(model);
            table.setBackground(Color.pink);
            table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
            table.setEnabled(false);
            table.setRowHeight(25);
            table.setRowMargin(4);
            int tableWidth = model.getColumnCount() * 150;
            int tableHeight = model.getRowCount() * 25;
            table.setPreferredSize(new Dimension(tableWidth, tableHeight));
            scroll = new JScrollPane(table);
            scroll.setBackground(Color.pink);
            scroll.setPreferredSize(new Dimension(tableWidth, 400));
            scroll.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
            scroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);


            JPanel panel = new JPanel(new GridBagLayout());
            GridBagConstraints gbConst = new GridBagConstraints();
            //panel.setSize(700, 500);            
            
            
            gbConst.gridy = 0;
            gbConst.gridx = 0;
            gbConst.gridheight = 3;
            //gbConst.insets = new Insets(10, 10, 10, 10);
            panel.add(scroll, gbConst);
            
            JLabel jLbl_ExeLogs = new JLabel("Execution Log");
            gbConst.gridy = 0;
            gbConst.gridx = 1;
            panel.add(jLbl_ExeLogs, gbConst);
            
            JTextArea jt_ShowLogs = new JTextArea(10,20);
            gbConst.gridy = 1;
            gbConst.gridx = 1;
            gbConst.gridwidth = 2;
            gbConst.gridheight = 2;
            panel.add(jt_ShowLogs, gbConst);

            JButton btn_Execute = new JButton("Execute");
            gbConst.gridy = 5;
            gbConst.gridx = 1;
            //gbConst.insets = new Insets(10, 10, 10, 10);
            panel.add(btn_Execute, gbConst);

            getContentPane().add(panel, BorderLayout.NORTH);
//        JFrame mainFrame = new InitFramework(vTitle);
//
//        mainFrame.setLocationRelativeTo(null);
//        mainFrame.setSize(700, 500);
//        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//        mainFrame.add(scroll, BorderLayout.CENTER);
//        JButton btn_Execute = new JButton("Execute");
//        mainFrame.add(btn_Execute, BorderLayout.SOUTH);
//
//        mainFrame.setSize(600, 600);
//        mainFrame.setResizable(true);
//        mainFrame.setVisible(true);



            btn_Execute.addActionListener(new ActionListener() {

                @Override
                public void actionPerformed(ActionEvent arg0) {
                    try {
                        Process p = Runtime.getRuntime().exec("wscript ../TestFramework_1/src/Resources/DriverScript.vbs " + gExcelPath);
                        p.waitFor();

                        System.out.println("Execution Complete and Exit Value = " + p.exitValue());
                    } catch (IOException ie) {
                        ie.printStackTrace();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public void fillData() {
        Workbook workbook = null;
        try {
            try {
                File file = new File(gExcelPath);
                workbook = Workbook.getWorkbook(file);
            } catch (IOException ex) {
                ex.printStackTrace();
                //Logger.getLogger(ExcelImport.class.getName()).log(Level.SEVERE, null, ex);
            }
            Sheet sheet = workbook.getSheet(0);
            headers.clear();
            // for (int i = 0; i < sheet.getColumns(); i++) {
            Cell cell1 = sheet.getCell(1, 0);
            // Cell cell1 = sheet.getColumn(1);
            headers.add(cell1.getContents());
            // }
            data.clear();
            for (int j = 1; j < sheet.getRows(); j++) {
                Vector d = new Vector();
                // for (int i = 0; i < sheet.getColumns(); i++) {
                Cell cell = sheet.getCell(1, j);
                d.add(cell.getContents());
                // }
                jx = j;
                d.add("\n");
                data.add(d);
            }
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }
}
