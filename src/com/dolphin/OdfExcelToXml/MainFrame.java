/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dolphin.OdfExcelToXml;

import java.awt.Color;
import java.awt.Point;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JMenuItem;
import javax.swing.JPopupMenu;
import javax.swing.JTextField;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;

/**
 * Main Frame
 *
 * @author hungd
 */
public class MainFrame extends javax.swing.JFrame {

    /**
     * Variables
     */
    private String defaultOdsPath;
    private String defaultXmlPath;

    private final MyJCombobox mCbSheets;
    private final int ALL_SHEETS_ASC = 0;
    private final int ALL_SHEETS_DESC = 1;
    private final int SELECTED_SHEETS = 2;

    /**
     * Creates new form MainFrame
     */
    public MainFrame() {
        initComponents();

        this.setLocationRelativeTo(null);
        pnCenter.setVisible(true);
        //this.setSize(this.getWidth(), this.getHeight() - pnCenter.getHeight());
        this.setIconImage(new ImageIcon(getClass().getClassLoader().getResource("resources/icon.png")).getImage());

        defaultOdsPath = "";
        defaultXmlPath = "";

        configOptions(tfRows);
        configOptions(tfSheets);
        configPopupMenus(tfSrcPath, pmOdsPath);
        configPopupMenus(tfXmlPath, pmXmlPath);
        configPopupMenus(tfRows, pmRows);
        configPopupMenus(tfSheets, pmSheets);
        configLogsMenu();

        mCbSheets = new MyJCombobox(cbSheets);
        pnSheets.setVisible(false);
        tfRows.setVisible(true);
        tfRows.setBackground(new Color(0xF5F5F5));
        tfSrcPath.requestFocus();
    }

    /**
     * Get system clipboard
     *
     * @return
     */
    private Clipboard getSystemClipboard() {
        Toolkit defaultToolkit = Toolkit.getDefaultToolkit();
        Clipboard systemClipboard = defaultToolkit.getSystemClipboard();
        return systemClipboard;
    }

    /**
     * Add listener to check validation of Options textfield
     */
    private void configOptions(JTextField textField) {

        textField.getDocument().addDocumentListener(new DocumentListener() {
            @Override
            public void insertUpdate(DocumentEvent e) {
                checkOptions();
            }

            @Override
            public void removeUpdate(DocumentEvent e) {
                checkOptions();
            }

            @Override
            public void changedUpdate(DocumentEvent e) {
                checkOptions();
            }

            // Check ranges of rows or sheet IDs
            public void checkOptions() {
                if (Functions.checkRanges(textField.getText())) {
                    setStatus("", false);
                } else {
                    setStatus("Invalid range: " + textField.getText(), true);
                }
            }
        });
    }

    /**
     * Configure popup menu for textField
     *
     * @param textField
     * @param popupMenu
     */
    private void configPopupMenus(JTextField textField, JPopupMenu popupMenu) {
        JMenuItem miCut = new JMenuItem(" Cut", new ImageIcon(getClass().getClassLoader().getResource("resources/cut.png")));
        miCut.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String fullText = textField.getText();
                String selectedText = textField.getSelectedText();
                if (selectedText != null && !selectedText.equalsIgnoreCase("")) {
                    getSystemClipboard().setContents(new StringSelection(selectedText), null);
                    int curPosition = fullText.indexOf(selectedText);
                    textField.setText(fullText.substring(0, fullText.indexOf(selectedText)) + ""
                            + fullText.substring(fullText.indexOf(selectedText) + selectedText.length()));
                    textField.setCaretPosition(curPosition);
                } else {
                    if (!fullText.equalsIgnoreCase("")) {
                        textField.setText("");
                        getSystemClipboard().setContents(new StringSelection(fullText), null);
                    }
                }
            }
        });
        popupMenu.add(miCut);

        JMenuItem miCopy = new JMenuItem(" Copy", new ImageIcon(getClass().getClassLoader().getResource("resources/copy.png")));
        miCopy.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String fullText = textField.getText();
                String selectedText = textField.getSelectedText();
                if (selectedText != null && !selectedText.equalsIgnoreCase("")) {
                    getSystemClipboard().setContents(new StringSelection(selectedText), null);
                } else {
                    if (!fullText.equalsIgnoreCase("")) {
                        getSystemClipboard().setContents(new StringSelection(fullText), null);
                    }
                }
            }
        });
        popupMenu.add(miCopy);

        JMenuItem miOdsPaste = new JMenuItem(" Paste", new ImageIcon(getClass().getClassLoader().getResource("resources/paste.png")));
        miOdsPaste.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    String fullText = textField.getText();
                    String selectedText = textField.getSelectedText();

                    Transferable trans = getSystemClipboard().getContents(null);
                    StringSelection nowData = new StringSelection("");
                    if (!trans.isDataFlavorSupported(DataFlavor.stringFlavor)) {
                        getSystemClipboard().setContents(nowData, nowData);
                    }
                    String pastedText = (String) getSystemClipboard().getContents(null).getTransferData(DataFlavor.stringFlavor);
                    int curPosition = textField.getCaretPosition();
                    if (pastedText != null) {
                        if (selectedText != null && !selectedText.equalsIgnoreCase("")) {
                            curPosition = curPosition - selectedText.length() + pastedText.length();
                            fullText = fullText.substring(0, fullText.indexOf(selectedText)) + pastedText
                                    + fullText.substring(fullText.indexOf(selectedText) + selectedText.length());
                        } else {
                            String lastText = "";
                            if (curPosition < fullText.length()) {
                                lastText = fullText.substring(curPosition + 1);
                            }
                            fullText = fullText.substring(0, curPosition) + pastedText + lastText;
                            curPosition = curPosition + pastedText.length();
                        }
                        textField.setText(fullText);
                        textField.setCaretPosition(curPosition);
                    }
                    textField.requestFocus();
                } catch (UnsupportedFlavorException ex) {
                    Logger.getLogger(MainFrame.class.getName()).log(Level.SEVERE, null, ex);
                    showLog(ex.toString());
                } catch (IOException ex) {
                    Logger.getLogger(MainFrame.class.getName()).log(Level.SEVERE, null, ex);
                    showLog(ex.toString());
                }
            }
        });
        popupMenu.add(miOdsPaste);

        textField.setComponentPopupMenu(popupMenu);
    }

    /**
     * Configure popup menu for Logs textfield
     */
    private void configLogsMenu() {
        JMenuItem miClear = new JMenuItem(" Clear", new ImageIcon(getClass().getClassLoader().getResource("resources/clear.png")));
        miClear.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                taLogs.setText("");
            }
        });
        pmLogs.add(miClear);
        pmLogs.addSeparator();

        JMenuItem miCopy = new JMenuItem(" Copy", new ImageIcon(getClass().getClassLoader().getResource("resources/copy.png")));
        miCopy.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String fullText = taLogs.getText();
                String selectedText = taLogs.getSelectedText();
                if (selectedText != null && !selectedText.equalsIgnoreCase("")) {
                    getSystemClipboard().setContents(new StringSelection(selectedText), null);
                } else {
                    if (!fullText.equalsIgnoreCase("")) {
                        getSystemClipboard().setContents(new StringSelection(fullText), null);
                    }
                }
            }
        });
        pmLogs.add(miCopy);

        JMenuItem miSelectAll = new JMenuItem(" Select All", new ImageIcon(getClass().getClassLoader().getResource("resources/select_all.png")));
        miSelectAll.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                taLogs.requestFocus();
                taLogs.selectAll();
                getSystemClipboard().setContents(new StringSelection(taLogs.getText()), null);
            }
        });
        pmLogs.add(miSelectAll);
        taLogs.setComponentPopupMenu(pmLogs);
    }

    private void showLog(String log) {
        System.out.println(log);
        taLogs.append("\n"+log);
    }

    private void setStatus(String status, boolean mark) {
        lbStatus.setText(status);
        if (mark) {
            lbStatus.setForeground(Color.red);
        } else {
            lbStatus.setForeground(Color.black);
        }
    }

    private void showOptions(boolean multiSheets, int id) {
        if (multiSheets) {
            rbFull.setText("Merged");
            rbFull.setToolTipText("Convert the selected sheets to a single XML file!");
            rbPart.setText("Separate");
            rbPart.setToolTipText("Convert every selected sheet to a separate XML file. XML Path must be a directory!");

            if (id == ALL_SHEETS_ASC || id == ALL_SHEETS_DESC) {
                pnRanges.setVisible(false);
            } else if (id == SELECTED_SHEETS) {
                pnRanges.setVisible(true);
                tfRows.setVisible(false);
                pnSheets.setVisible(true);
                tfSheets.requestFocus();
            }
            if (rbPart.isSelected()) {
                setStatus("NOTE: XML Path must be a directory path", false);
            } else {
                setStatus("NOTE: XML Path must be a file path", false);
            }
        } else {
            rbFull.setSelected(true);
            rbFull.setText("Full");
            rbFull.setToolTipText("Convert the whole selected sheet to the XML file!");
            rbPart.setText("Rows:");
            rbPart.setToolTipText("Convert a part of the selected sheet to the XML file!");
            pnRanges.setVisible(true);
            tfRows.setVisible(true);
            pnSheets.setVisible(false);
            setStatus("NOTE: XML Path must be a file path", false);
        }
    }

    private void loadSheets(File file) {
        boolean planArc = rbLinux.isSelected();
        showLog("LOADING ODF/Excel file: " + file.getPath());
        if (file.getPath().endsWith(".ods")) {
            OdfToXml odfToXml = new OdfToXml(planArc, file, null, lbStatus, taLogs);

            showLog(" > Sheet count: " + odfToXml.getSheetCount());
            mCbSheets.removeAllItems();
            if (odfToXml.getSheetCount() > 1) {
                mCbSheets.addItem(new SheetEntry("All sheets (Ascending)", ALL_SHEETS_ASC, false));
                mCbSheets.addItem(new SheetEntry("All sheets (Descending)", ALL_SHEETS_DESC, false));
                mCbSheets.addItem(new SheetEntry("Selected sheets", SELECTED_SHEETS, false));
            }
            for (int i = 0; i < odfToXml.getSheetCount(); i++) {
                mCbSheets.addItem(new SheetEntry("Sheet " + i + ": " + odfToXml.getSheet(i).getName(), i, true));
            }
            showLog(" > Loading completed!");
            setStatus("Loading completed!", false);
        } else if (file.getPath().endsWith(".xls") || file.getPath().endsWith(".xlsx")) {
            ExcelToXml excelToXml = new ExcelToXml(planArc, file, null, lbStatus, taLogs);

            showLog(" > Sheet count: " + excelToXml.getSheetCount());
            mCbSheets.removeAllItems();
            if (excelToXml.getSheetCount() > 1) {
                mCbSheets.addItem(new SheetEntry("All sheets (Ascending)", ALL_SHEETS_ASC, false));
                mCbSheets.addItem(new SheetEntry("All sheets (Descending)", ALL_SHEETS_DESC, false));
                mCbSheets.addItem(new SheetEntry("Selected sheets", SELECTED_SHEETS, false));
            }
            for (int i = 0; i < excelToXml.getSheetCount(); i++) {
                mCbSheets.addItem(new SheetEntry("Sheet " + i + ": " + excelToXml.getSheet(i).getSheetName(), i, true));
            }
            showLog(" > Loading completed!");
            setStatus("Loading completed!", false);
        } else {
            showLog(" > [ERROR] The specified input file is not ODF or Excel format!");
            setStatus("ERROR: " + "The specified input file is not ODF or Excel file!", true);
        }
        showOptions(mCbSheets.getItemCount() > 1, 0);
    }

    private ArrayList<Point> getRanges(String src) {
        ArrayList<Point> results = new ArrayList<>();

        char[] filter1 = {',', ';'};
        char[] filter2 = {' ', '-', ':'};
        ArrayList<String> ranges = Functions.parseString(src, filter1);
        for (int i = 0; i < ranges.size(); i++) {
            ArrayList<String> range = Functions.parseString(ranges.get(i), filter2);
            if (range.size() == 1) {
                try {
                    int fromRow = Integer.valueOf(range.get(0));
                    int toRow = Integer.valueOf(range.get(0));
                    results.add(new Point(fromRow, toRow));
                } catch (NumberFormatException ex) {
                    showLog(ex.toString());
                    setStatus(" > ERROR: " + ex.getMessage(), true);
                }
            }
            if (range.size() >= 2) {
                try {
                    int fromRow = Integer.valueOf(range.get(0));
                    int toRow = Integer.valueOf(range.get(1));
                    results.add(new Point(fromRow, toRow));
                } catch (NumberFormatException ex) {
                    showLog(ex.toString());
                    setStatus(" > ERROR: " + ex.getMessage(), true);
                }
            }
        }
        return results;
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        bgArchitect = new javax.swing.ButtonGroup();
        bgOptions = new javax.swing.ButtonGroup();
        pmOdsPath = new javax.swing.JPopupMenu();
        pmXmlPath = new javax.swing.JPopupMenu();
        pmRows = new javax.swing.JPopupMenu();
        pmSheets = new javax.swing.JPopupMenu();
        pmLogs = new javax.swing.JPopupMenu();
        mainPanel = new javax.swing.JPanel();
        pnNorth = new javax.swing.JPanel();
        pnTitle = new javax.swing.JPanel();
        jLabel5 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        pnInputs = new javax.swing.JPanel();
        rbLinux = new javax.swing.JRadioButton();
        rbWindows = new javax.swing.JRadioButton();
        tfSrcPath = new javax.swing.JTextField();
        tfXmlPath = new javax.swing.JTextField();
        cbSheets = new javax.swing.JComboBox<>();
        pnOptions = new javax.swing.JPanel();
        pnModes = new javax.swing.JPanel();
        rbFull = new javax.swing.JRadioButton();
        rbPart = new javax.swing.JRadioButton();
        pnRanges = new javax.swing.JPanel();
        tfRows = new javax.swing.JTextField();
        pnSheets = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        tfSheets = new javax.swing.JTextField();
        pnAction = new javax.swing.JPanel();
        jLabel7 = new javax.swing.JLabel();
        btOpenSrc = new javax.swing.JButton();
        btSaveXml = new javax.swing.JButton();
        btRefresh = new javax.swing.JButton();
        btConvert = new javax.swing.JButton();
        pnCenter = new javax.swing.JPanel();
        jScrollPane2 = new javax.swing.JScrollPane();
        taLogs = new javax.swing.JTextArea();
        pnSouth = new javax.swing.JPanel();
        jPanel2 = new javax.swing.JPanel();
        lbStatus = new javax.swing.JLabel();
        jPanel1 = new javax.swing.JPanel();
        btExit = new javax.swing.JButton();
        btShowLog = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("ODF/Excel to XML Converter");

        mainPanel.setLayout(new java.awt.BorderLayout());

        pnNorth.setLayout(new java.awt.BorderLayout());

        jLabel5.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        jLabel5.setText("ODF/Excel");

        jLabel3.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        jLabel3.setText("XML Path");
        jLabel3.setToolTipText("");

        jLabel6.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        jLabel6.setText("Sheet ID");

        jLabel4.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        jLabel4.setText("Options");

        jLabel2.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        jLabel2.setText("Architect");

        javax.swing.GroupLayout pnTitleLayout = new javax.swing.GroupLayout(pnTitle);
        pnTitle.setLayout(pnTitleLayout);
        pnTitleLayout.setHorizontalGroup(
            pnTitleLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(pnTitleLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(pnTitleLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(pnTitleLayout.createSequentialGroup()
                        .addGroup(pnTitleLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel3)
                            .addComponent(jLabel6)
                            .addComponent(jLabel4)
                            .addComponent(jLabel2))
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addComponent(jLabel5, javax.swing.GroupLayout.Alignment.TRAILING))
                .addGap(13, 13, 13))
        );

        pnTitleLayout.linkSize(javax.swing.SwingConstants.HORIZONTAL, new java.awt.Component[] {jLabel2, jLabel3, jLabel4, jLabel5, jLabel6});

        pnTitleLayout.setVerticalGroup(
            pnTitleLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(pnTitleLayout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel6, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pnTitleLayout.linkSize(javax.swing.SwingConstants.VERTICAL, new java.awt.Component[] {jLabel2, jLabel3, jLabel4, jLabel5, jLabel6});

        pnNorth.add(pnTitle, java.awt.BorderLayout.WEST);

        bgArchitect.add(rbLinux);
        rbLinux.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        rbLinux.setSelected(true);
        rbLinux.setText("Linux");
        rbLinux.setToolTipText("Convert to XML format on Linux");

        bgArchitect.add(rbWindows);
        rbWindows.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        rbWindows.setText("Windows");
        rbWindows.setToolTipText("Convert to XML format on Windows");

        tfSrcPath.setFont(new java.awt.Font("DejaVu Sans", 0, 14)); // NOI18N
        tfSrcPath.setToolTipText("Enter input ODF (.ods) or Excel (.xls, .xlsx) file");

        tfXmlPath.setFont(new java.awt.Font("DejaVu Sans", 0, 14)); // NOI18N
        tfXmlPath.setToolTipText("Enter output XML file or folder");

        cbSheets.setToolTipText("Select sheet(s) for conversion");
        cbSheets.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cbSheetsActionPerformed(evt);
            }
        });

        pnOptions.setLayout(new java.awt.BorderLayout());

        bgOptions.add(rbFull);
        rbFull.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        rbFull.setSelected(true);
        rbFull.setText("Full");

        bgOptions.add(rbPart);
        rbPart.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        rbPart.setText("Part");
        rbPart.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                rbPartItemStateChanged(evt);
            }
        });

        javax.swing.GroupLayout pnModesLayout = new javax.swing.GroupLayout(pnModes);
        pnModes.setLayout(pnModesLayout);
        pnModesLayout.setHorizontalGroup(
            pnModesLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(pnModesLayout.createSequentialGroup()
                .addComponent(rbFull, javax.swing.GroupLayout.PREFERRED_SIZE, 100, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(rbPart)
                .addGap(0, 0, Short.MAX_VALUE))
        );
        pnModesLayout.setVerticalGroup(
            pnModesLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(pnModesLayout.createSequentialGroup()
                .addGap(0, 0, Short.MAX_VALUE)
                .addGroup(pnModesLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.CENTER)
                    .addComponent(rbFull, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(rbPart, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)))
        );

        pnOptions.add(pnModes, java.awt.BorderLayout.WEST);

        pnRanges.setLayout(new java.awt.CardLayout());

        tfRows.setFont(new java.awt.Font("DejaVu Sans", 0, 14)); // NOI18N
        tfRows.setToolTipText("Enter row ranges of the selected sheet");
        tfRows.setEnabled(false);
        pnRanges.add(tfRows, "card2");

        jLabel1.setFont(new java.awt.Font("DejaVu Sans", 1, 13)); // NOI18N
        jLabel1.setText("Sheets:");

        tfSheets.setToolTipText("Enter ranges of ID of sheets");

        javax.swing.GroupLayout pnSheetsLayout = new javax.swing.GroupLayout(pnSheets);
        pnSheets.setLayout(pnSheetsLayout);
        pnSheetsLayout.setHorizontalGroup(
            pnSheetsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(pnSheetsLayout.createSequentialGroup()
                .addGap(25, 25, 25)
                .addComponent(jLabel1)
                .addGap(4, 4, 4)
                .addComponent(tfSheets, javax.swing.GroupLayout.DEFAULT_SIZE, 179, Short.MAX_VALUE))
        );
        pnSheetsLayout.setVerticalGroup(
            pnSheetsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(pnSheetsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.CENTER)
                .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(tfSheets, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE))
        );

        pnRanges.add(pnSheets, "card3");

        pnOptions.add(pnRanges, java.awt.BorderLayout.CENTER);

        javax.swing.GroupLayout pnInputsLayout = new javax.swing.GroupLayout(pnInputs);
        pnInputs.setLayout(pnInputsLayout);
        pnInputsLayout.setHorizontalGroup(
            pnInputsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(tfSrcPath)
            .addComponent(tfXmlPath)
            .addGroup(pnInputsLayout.createSequentialGroup()
                .addComponent(rbLinux, javax.swing.GroupLayout.PREFERRED_SIZE, 100, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(rbWindows)
                .addGap(0, 0, Short.MAX_VALUE))
            .addComponent(cbSheets, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addComponent(pnOptions, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        pnInputsLayout.setVerticalGroup(
            pnInputsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(pnInputsLayout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addGroup(pnInputsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(rbWindows, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(rbLinux, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(6, 6, 6)
                .addComponent(tfSrcPath, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(tfXmlPath, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(cbSheets, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(pnOptions, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pnNorth.add(pnInputs, java.awt.BorderLayout.CENTER);

        jLabel7.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);

        btOpenSrc.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        btOpenSrc.setText("Open...");
        btOpenSrc.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btOpenSrcActionPerformed(evt);
            }
        });

        btSaveXml.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        btSaveXml.setText("Save...");
        btSaveXml.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btSaveXmlActionPerformed(evt);
            }
        });

        btRefresh.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        btRefresh.setText("Refresh");
        btRefresh.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btRefreshActionPerformed(evt);
            }
        });

        btConvert.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        btConvert.setText("Convert");
        btConvert.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btConvertActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout pnActionLayout = new javax.swing.GroupLayout(pnAction);
        pnAction.setLayout(pnActionLayout);
        pnActionLayout.setHorizontalGroup(
            pnActionLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(pnActionLayout.createSequentialGroup()
                .addGap(4, 4, 4)
                .addGroup(pnActionLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(btConvert, javax.swing.GroupLayout.DEFAULT_SIZE, 90, Short.MAX_VALUE)
                    .addComponent(btOpenSrc, javax.swing.GroupLayout.DEFAULT_SIZE, 90, Short.MAX_VALUE)
                    .addComponent(btSaveXml, javax.swing.GroupLayout.DEFAULT_SIZE, 90, Short.MAX_VALUE)
                    .addComponent(btRefresh, javax.swing.GroupLayout.DEFAULT_SIZE, 90, Short.MAX_VALUE)
                    .addComponent(jLabel7, javax.swing.GroupLayout.DEFAULT_SIZE, 90, Short.MAX_VALUE))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pnActionLayout.linkSize(javax.swing.SwingConstants.HORIZONTAL, new java.awt.Component[] {btConvert, btOpenSrc, btRefresh, btSaveXml, jLabel7});

        pnActionLayout.setVerticalGroup(
            pnActionLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, pnActionLayout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(6, 6, 6)
                .addComponent(btOpenSrc, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btSaveXml, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btRefresh, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btConvert, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 11, Short.MAX_VALUE))
        );

        pnActionLayout.linkSize(javax.swing.SwingConstants.VERTICAL, new java.awt.Component[] {btConvert, btOpenSrc, btRefresh, btSaveXml, jLabel7});

        pnNorth.add(pnAction, java.awt.BorderLayout.EAST);

        mainPanel.add(pnNorth, java.awt.BorderLayout.NORTH);

        pnCenter.setToolTipText("");

        taLogs.setEditable(false);
        taLogs.setColumns(20);
        taLogs.setRows(5);
        jScrollPane2.setViewportView(taLogs);

        javax.swing.GroupLayout pnCenterLayout = new javax.swing.GroupLayout(pnCenter);
        pnCenter.setLayout(pnCenterLayout);
        pnCenterLayout.setHorizontalGroup(
            pnCenterLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(pnCenterLayout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 568, Short.MAX_VALUE)
                .addContainerGap())
        );
        pnCenterLayout.setVerticalGroup(
            pnCenterLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 117, Short.MAX_VALUE)
        );

        mainPanel.add(pnCenter, java.awt.BorderLayout.CENTER);

        pnSouth.setLayout(new java.awt.BorderLayout());

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(lbStatus, javax.swing.GroupLayout.DEFAULT_SIZE, 352, Short.MAX_VALUE))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(lbStatus, javax.swing.GroupLayout.PREFERRED_SIZE, 23, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(12, Short.MAX_VALUE))
        );

        pnSouth.add(jPanel2, java.awt.BorderLayout.CENTER);

        btExit.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        btExit.setText("Exit");
        btExit.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btExitActionPerformed(evt);
            }
        });

        btShowLog.setFont(new java.awt.Font("DejaVu Sans", 1, 12)); // NOI18N
        btShowLog.setText("Hide Logs");
        btShowLog.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btShowLogActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(btShowLog, javax.swing.GroupLayout.PREFERRED_SIZE, 112, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(4, 4, 4)
                .addComponent(btExit, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(5, 5, 5)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btShowLog, javax.swing.GroupLayout.DEFAULT_SIZE, 30, Short.MAX_VALUE)
                    .addComponent(btExit, javax.swing.GroupLayout.DEFAULT_SIZE, 30, Short.MAX_VALUE))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pnSouth.add(jPanel1, java.awt.BorderLayout.EAST);

        mainPanel.add(pnSouth, java.awt.BorderLayout.SOUTH);

        getContentPane().add(mainPanel, java.awt.BorderLayout.CENTER);

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void btExitActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btExitActionPerformed
        System.exit(0);
    }//GEN-LAST:event_btExitActionPerformed

    private void btConvertActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btConvertActionPerformed
        boolean planArc = rbLinux.isSelected();
        File odsFile = new File(tfSrcPath.getText());
        File xmlFile = new File(tfXmlPath.getText());
        int sheetID = mCbSheets.getSelectedItem().getId();
        boolean isSheet = mCbSheets.getSelectedItem().isSheet();
        String selectedName = mCbSheets.getSelectedItem().getName();

        showLog("\nCONVERTING ODF/Excel file: " + odsFile.getPath());
        showLog("- Options: " + (rbFull.isSelected() ? (isSheet ? "Single sheet | Full" : selectedName + " | Merged")
                : (isSheet ? "Single sheet | Ranges" : selectedName + " | Separate")) + "\n");

        boolean result = false;
        if (odsFile.getPath().endsWith(".ods")) {
            OdfToXml odfToXml = new OdfToXml(planArc, odsFile, xmlFile, lbStatus, taLogs);
            if (isSheet) {
                if (rbFull.isSelected()) {
                    result = odfToXml.writeXML(sheetID);
                } else {
                    ArrayList<Point> ranges = getRanges(tfRows.getText());
                    result = odfToXml.writeXML(sheetID, ranges);
                }
            } else {
                if (sheetID == ALL_SHEETS_ASC || sheetID == ALL_SHEETS_DESC) {
                    result = odfToXml.writeXML(rbFull.isSelected(), tfXmlPath.getText(), sheetID == ALL_SHEETS_ASC);
                } else if (sheetID == SELECTED_SHEETS) {
                    ArrayList<Point> ranges = getRanges(tfSheets.getText());
                    result = odfToXml.writeXML(rbFull.isSelected(), tfXmlPath.getText(), ranges);
                }
            }
            if (result) {
                showLog("CONVERSION COMPLETED!");
                setStatus("CONVERSION COMPLETED!", false);
            }
        } else if (odsFile.getPath().endsWith(".xls") || odsFile.getPath().endsWith(".xlsx")) {
            ExcelToXml excelToXml = new ExcelToXml(planArc, odsFile, xmlFile, lbStatus, taLogs);
            if (isSheet) {
                if (rbFull.isSelected()) {
                    result = excelToXml.writeXML(sheetID);
                } else {
                    ArrayList<Point> ranges = getRanges(tfRows.getText());
                    result = excelToXml.writeXML(sheetID, ranges);
                }
            } else {
                if (sheetID == ALL_SHEETS_ASC || sheetID == ALL_SHEETS_DESC) {
                    result = excelToXml.writeXML(rbFull.isSelected(), tfXmlPath.getText(), sheetID == ALL_SHEETS_ASC);
                } else if (sheetID == SELECTED_SHEETS) {
                    ArrayList<Point> ranges = getRanges(tfSheets.getText());
                    result = excelToXml.writeXML(rbFull.isSelected(), tfXmlPath.getText(), ranges);
                }
            }
            if (result) {
                showLog("CONVERSION COMPLETED!");
                setStatus("CONVERSION COMPLETED!", false);
            }
        } else {
            showLog("[ERROR] The specified input file is not ODF or Excel format!");
            setStatus("ERROR: " + "The specified input file is not ODF or Excel format!", true);
        }
    }//GEN-LAST:event_btConvertActionPerformed

    private void btRefreshActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btRefreshActionPerformed
        File odsFile = new File(tfSrcPath.getText());
        loadSheets(odsFile);
    }//GEN-LAST:event_btRefreshActionPerformed

    private void btSaveXmlActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btSaveXmlActionPerformed
        JFileChooser fileChooser = new JFileChooser(new File(defaultXmlPath));
        fileChooser.setDialogTitle("Select a xml file or container folder");
        fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

        int returnVal = fileChooser.showOpenDialog(MainFrame.this);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File xmlFile = fileChooser.getSelectedFile();
            defaultXmlPath = fileChooser.getSelectedFile().getParent();
            tfXmlPath.setText(xmlFile.getPath());
        }
    }//GEN-LAST:event_btSaveXmlActionPerformed

    private void btOpenSrcActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btOpenSrcActionPerformed
        JFileChooser fileChooser = new JFileChooser(new File(defaultOdsPath));
        fileChooser.setDialogTitle("Select a SpreadSheet file");
        int returnVal = fileChooser.showOpenDialog(MainFrame.this);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File odsFile = fileChooser.getSelectedFile();
            defaultOdsPath = fileChooser.getSelectedFile().getParent();
            tfSrcPath.setText(odsFile.getPath());
            if (tfXmlPath.getText().equalsIgnoreCase("")) {
                String xmlName = odsFile.getPath();
                xmlName = xmlName.substring(0, xmlName.lastIndexOf('.')) + ".xml";
                tfXmlPath.setText(xmlName);
            }
            loadSheets(odsFile);
        }
    }//GEN-LAST:event_btOpenSrcActionPerformed

    private void btShowLogActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btShowLogActionPerformed
        if (pnCenter.isVisible()) {
            pnCenter.setVisible(false);
            this.setSize(this.getWidth(), this.getHeight() - pnCenter.getHeight());
            btShowLog.setText("Show Logs");
        } else {
            pnCenter.setVisible(true);
            this.setSize(this.getWidth(), this.getHeight() + pnCenter.getHeight());
            btShowLog.setText("Hide Logs");
        }
    }//GEN-LAST:event_btShowLogActionPerformed

    private void rbPartItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_rbPartItemStateChanged
        SheetEntry sheet = mCbSheets.getSelectedItem();
        if (sheet.getId() >= 0 && !sheet.isSheet()) {
            int id = sheet.getId();
            if (id == ALL_SHEETS_ASC || id == ALL_SHEETS_DESC) {
                pnRanges.setVisible(false);
            } else if (id == SELECTED_SHEETS) {
                pnRanges.setVisible(true);
                tfRows.setVisible(false);
                pnSheets.setVisible(true);
                tfSheets.requestFocus();
            }
            if (rbPart.isSelected()) {
                setStatus("NOTE: XML Path must be a directory path", false);
            } else {
                setStatus("NOTE: XML Path must be a file path", false);
            }            
        } else {
            if (rbPart.isSelected()) {
                tfRows.setEnabled(true);
                tfRows.setBackground(Color.WHITE);
                tfRows.requestFocus();
                setStatus("Example: 5-10, 15:20; 25 30, 35, 50", false);
            } else {
                tfRows.setEnabled(false);
                tfRows.setBackground(new Color(0xF4F4F4));
                setStatus("NOTE: XML Path must be a file path", false);
            }
        }
    }//GEN-LAST:event_rbPartItemStateChanged

    private void cbSheetsActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cbSheetsActionPerformed
        SheetEntry sheet = mCbSheets.getSelectedItem();
        showLog(" > Selected sheet: " + sheet.getName());
        if (sheet.getId() >= 0 && !sheet.isSheet()) {
            int id = sheet.getId();
            showOptions(true, id);
        } else {
            showOptions(false, 0);
        }

        tfXmlPath.requestFocus();
        tfXmlPath.selectAll();
    }//GEN-LAST:event_cbSheetsActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;

                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);

        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);

        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);

        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new MainFrame().setVisible(true);
            }
        });

    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.ButtonGroup bgArchitect;
    private javax.swing.ButtonGroup bgOptions;
    private javax.swing.JButton btConvert;
    private javax.swing.JButton btExit;
    private javax.swing.JButton btOpenSrc;
    private javax.swing.JButton btRefresh;
    private javax.swing.JButton btSaveXml;
    private javax.swing.JButton btShowLog;
    private javax.swing.JComboBox<String> cbSheets;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JLabel lbStatus;
    private javax.swing.JPanel mainPanel;
    private javax.swing.JPopupMenu pmLogs;
    private javax.swing.JPopupMenu pmOdsPath;
    private javax.swing.JPopupMenu pmRows;
    private javax.swing.JPopupMenu pmSheets;
    private javax.swing.JPopupMenu pmXmlPath;
    private javax.swing.JPanel pnAction;
    private javax.swing.JPanel pnCenter;
    private javax.swing.JPanel pnInputs;
    private javax.swing.JPanel pnModes;
    private javax.swing.JPanel pnNorth;
    private javax.swing.JPanel pnOptions;
    private javax.swing.JPanel pnRanges;
    private javax.swing.JPanel pnSheets;
    private javax.swing.JPanel pnSouth;
    private javax.swing.JPanel pnTitle;
    private javax.swing.JRadioButton rbFull;
    private javax.swing.JRadioButton rbLinux;
    private javax.swing.JRadioButton rbPart;
    private javax.swing.JRadioButton rbWindows;
    private javax.swing.JTextArea taLogs;
    private javax.swing.JTextField tfRows;
    private javax.swing.JTextField tfSheets;
    private javax.swing.JTextField tfSrcPath;
    private javax.swing.JTextField tfXmlPath;
    // End of variables declaration//GEN-END:variables
}
