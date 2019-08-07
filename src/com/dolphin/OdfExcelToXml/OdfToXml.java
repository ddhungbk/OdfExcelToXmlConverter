/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dolphin.OdfExcelToXml;

import java.awt.Color;
import java.awt.Point;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JLabel;
import javax.swing.JTextArea;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

/**
 * ODF to XML Converter
 *
 * @author hungd
 */
public class OdfToXml {

    private final int MAX_COL = 100;
    private final int MAX_ROW = 100;

    private boolean planArc;
    private File odfFile;
    private File xmlFile;
    private SpreadSheet spreadSheet;

    JLabel lbStatus;
    JTextArea taLogs;

    public OdfToXml(boolean planArc, File odfFile, File xmlFile, JLabel lbStatus, JTextArea taLogs) {
        this.planArc = planArc;
        this.odfFile = odfFile;
        this.xmlFile = xmlFile;
        this.lbStatus = lbStatus;
        this.taLogs = taLogs;

        load();
    }

    public void load() {
        try {
            spreadSheet = SpreadSheet.createFromFile(odfFile);
        } catch (IOException e) {
            showLog(" > ERROR: Read spreadsheet error! (" + e.toString() + ")");
            setStatus("ERROR: " + "No such file or directory!", true);
            throw new RuntimeException(e);
        } catch (NullPointerException e) {
            showLog(" > ERROR: Read spreadsheet error! (" + e.toString() + ") " + odfFile.getPath());
            setStatus("ERROR: " + "Spreadsheet file format!", true);
            throw new RuntimeException(e);
        }
    }

    public int getSheetCount() {
        return spreadSheet.getSheetCount();
    }

    public Sheet getSheet(int id) {
        return spreadSheet.getSheet(id);
    }

    /**
     * Get content of a cell at position (col,row)
     *
     * @param sheet
     * @param col
     * @param row
     * @return
     */
    private String getCell(Sheet sheet, int col, int row) {
        MutableCell<SpreadSheet> cell = sheet.getCellAt(col, row);
        return cell.getValue().toString();
    }

    /**
     * Get content of a row at position (column,row)
     *
     * @param sheet
     * @param col
     * @param row
     * @return
     */
    private RowEntry getRowEntry(Sheet sheet, int col, int row, int columnCount) {
        ArrayList<String> contents = new ArrayList<>();
        for (int i = 0; i < columnCount; i++) {
            contents.add(getCell(sheet, col + i, row));
        }
        return new RowEntry(columnCount, contents);
    }

    /**
     * Detect the starting row and column index
     *
     * @param sheet
     * @param maxCol
     * @param maxRow
     * @return
     */
    private ArrayList<Integer> getBoundary(Sheet sheet, int maxCol, int maxRow) {
        ArrayList<Integer> results = new ArrayList<>();

        showLog("- Detect boundary positions:");
        boolean detectDone = false;
        for (int col = 0; col < maxCol; col++) {
            showLog("  > Detecting column " + col+"...");
            if (col < sheet.getColumnCount()) {
                for (int row = 0; row < maxRow; row++) {
                    if (row < sheet.getRowCount()) {
                        String cell = this.getCell(sheet, col, row);
                        if (!cell.trim().equalsIgnoreCase("")) {
                            results.add(col);
                            results.add(row);
                            detectDone = true;
                            break;
                        }
                    }
                }
            }
            if (detectDone) {
                break;
            }
        }

        if (detectDone) {
            detectDone = false;
            for (int col = results.get(0); col < sheet.getColumnCount(); col++) {
                String cell = this.getCell(sheet, col, results.get(1));
                if (cell.trim().equalsIgnoreCase("")) {
                    results.add(col - 1);
                    detectDone = true;
                    break;
                }
            }
            if (detectDone == false) {
                results.add(sheet.getColumnCount() - 1);
            }

            detectDone = false;
            for (int row = results.get(1); row < sheet.getRowCount(); row++) {
                String cell = this.getCell(sheet, results.get(0), row);
                if (cell.trim().equalsIgnoreCase("")) {
                    results.add(row - 1);
                    detectDone = true;
                    break;
                }
            }
            if (detectDone == false) {
                results.add(sheet.getRowCount() - 1);
            }

            showLog("  > DONE detecting boundary positions: start(" + results.get(0) + "," + results.get(1)
                    + "), finish(" + results.get(2) + "," + results.get(3) + ")");
            return results;
        } else {
            showLog("  > ERROR: Cannot detect boundary positions!\n");
            return null;
        }
    }

    /**
     * Read an entire sheet
     *
     * @param sheet
     * @param maxCol
     * @param maxRow
     * @return
     */
    public List<RowEntry> readSheet(Sheet sheet, int maxCol, int maxRow) {
        List<RowEntry> results = new ArrayList<>();
        ArrayList<Integer> boundaries = getBoundary(sheet, maxCol, maxRow);
        int startCol = boundaries.get(0);
        int startRow = boundaries.get(1);
        int finishCol = boundaries.get(2);
        int finishRow = boundaries.get(3);
        int columnCount = finishCol - startCol + 1;

        for (int row = startRow; row <= finishRow; row++) {
            RowEntry rowEntry = getRowEntry(sheet, startCol, row, columnCount);
            results.add(rowEntry);
        }

        showLog("- Column count: " + columnCount);
        showLog("- Number of rows: " + results.size());
        for (int i = 0; i < results.size(); i++) {
            showLog("  > Row[" + i + "]: " + results.get(i).toString());
        }
        return results;
    }

    /**
     * Read a sheet with row ranges
     *
     * @param sheet
     * @param ranges : fromRow=ranges[i].x, toRow=ranges[i].y
     * @param maxCol
     * @param maxRow
     * @return
     */
    public List<RowEntry> readSheet(Sheet sheet, ArrayList<Point> ranges, int maxCol, int maxRow) {
        List<RowEntry> results = new ArrayList<>();
        ArrayList<Integer> boundaries = getBoundary(sheet, maxCol, maxRow);
        int startCol = boundaries.get(0);
        int startRow = boundaries.get(1);
        int finishCol = boundaries.get(2);
        int finishRow = boundaries.get(3);
        int columnCount = finishCol - startCol + 1;

        results.add(getRowEntry(sheet, startCol, startRow, columnCount));

        for (int i = 0; i < ranges.size(); i++) {
            int fromRow = ranges.get(i).x;
            int toRow = ranges.get(i).y;
            if (fromRow <= toRow && fromRow <= finishRow) {
                for (int row = fromRow; row <= toRow; row++) {
                    if (row <= finishRow) {
                        RowEntry rowEntry = getRowEntry(sheet, startCol, row, columnCount);
                        results.add(rowEntry);
                    } else {
                        showLog("  > [Warning] Index out of bound: " + fromRow + ":" + toRow + "/" + finishRow);
                        setStatus("Index out of bound: " + fromRow + ":" + toRow + "/" + finishRow, true);
                        break;
                    }
                }
                setStatus("", false);
            } else {
                if (fromRow > finishRow || toRow > finishRow) {
                    showLog("  > [Warning] Index out of bound: " + fromRow + ":" + toRow + "/" + finishRow);
                    setStatus("Index out of bound: " + fromRow + ":" + toRow + "/" + finishRow, true);
                } else {
                    showLog("  > [Warning] Invalid range: " + fromRow + ":" + toRow);
                    setStatus("Invalid range: " + fromRow + ":" + toRow, true);
                }
            }
        }

        showLog("- Column count: " + columnCount);
        showLog("- Number of rows: " + results.size());
        for (int i = 0; i < ranges.size(); i++) {
            showLog("  > Range " + i + ": " + ranges.get(i).x + ", " + ranges.get(i).y);
        }
        for (int i = 0; i < results.size(); i++) {
            showLog("  > Row[" + i + "]: " + results.get(i).toString());
        }

        return results;
    }

    /**
     * Write the results to XML file
     *
     * @param list
     * @param planArc
     * @param sheetName
     * @throws IOException
     */
    public void writeXML(List<RowEntry> list, boolean planArc, String sheetName,
            boolean addHeader, boolean addFooter, boolean append) throws IOException {

        try {
            xmlFile.createNewFile();
            FileWriter fw = new FileWriter(xmlFile, append);

            String content = "";
            if (addHeader) {
                if (planArc == false) { // Windows
                    content += "<!-- WorkBook Path: " + odfFile.getPath() + " -->\n";
                    content += "<!-- Date: "
                            + new SimpleDateFormat("dd/MM/yy HH:mm:ss a").format(Calendar.getInstance().getTime()) + " -->\n";
                }
                content += "<Workbook>\n\t<Worksheet>\n\t\t<Table>\n";
                fw.write(content);
            }

            int size = list.size();
            int columnCount = (size > 0 ? list.get(0).getColumnCount() : 0);
            content = "";
            if (planArc == true) { // Linux
                content += "\t\t<!-- ";
                for (int i = 0; i < columnCount; i++) {
                    content += "'" + (size <= 0 ? "" : list.get(0).getField(i)) + "' ";
                }
                content += "-->\n\t\t\t<Row>\n";
                for (int i = 0; i < columnCount; i++) {
                    content += "\t\t\t\t<Cell>" + (size <= 1 ? "" : list.get(1).getField(i)) + "</Cell>\n";
                }
                content += "\t\t\t</Row>\n";
            } else {
                content += "\t\t\t<!-- " + sheetName + " -->\n";
                content += "\t\t\t<Row>\n";
                for (int i = 0; i < columnCount; i++) {
                    content += "\t\t\t\t<Cell>" + (size <= 1 ? "" : list.get(1).getField(i)) + "</Cell>" + "  <!-- " + (size <= 0 ? "" : list.get(0).getField(i))
                            + " -->\n";
                }
                content += "\t\t\t</Row>\n";
            }
            fw.write(content);

            for (int i = 2; i < list.size(); i++) {
                content = "\t\t\t<Row>\n";
                for (int j = 0; j < columnCount; j++) {
                    content += "\t\t\t\t<Cell>" + list.get(i).getField(j) + "</Cell>\n";
                }
                content += "\t\t\t</Row>\n";
                fw.write(content);
            }

            if (addFooter) {
                content = "\t\t</Table>\n\t</Worksheet>\n</Workbook>\n";
                fw.write(content);
            }
            fw.close();
            showLog("- WRITTEN to XML file: " + xmlFile.getPath() + "\n");
        } catch (Exception ex) {
            showLog(" > ERROR: " + ex.getMessage() + " (" + xmlFile.getPath() + ")");
            setStatus("ERROR: " + ex.getMessage() + " (" + xmlFile.getPath() + ")", true);
            throw ex;
        }
    }

    public boolean writeXML(int id) {
        try {
            Sheet sheet = spreadSheet.getSheet(id);
            showLog("- Sheet name: " + sheet.getName());
            List<RowEntry> list = this.readSheet(sheet, MAX_COL, MAX_ROW);
            writeXML(list, planArc, sheet.getName(), true, true, false);
        } catch (IOException ex) {
            Logger.getLogger(OdfToXml.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
        return true;
    }

    public boolean writeXML(int id, ArrayList<Point> ranges) {
        try {
            Sheet sheet = spreadSheet.getSheet(id);
            showLog("- Sheet name: " + sheet.getName());
            List<RowEntry> list = readSheet(sheet, ranges, MAX_COL, MAX_ROW);
            writeXML(list, planArc, sheet.getName(), true, true, false);
        } catch (IOException ex) {
            Logger.getLogger(OdfToXml.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
        return true;
    }

    public boolean writeXML(boolean isFull, String folder, boolean asc) {
        try {
            int sheetCount = this.getSheetCount();
            if (isFull) {
                xmlFile.delete();
                xmlFile.createNewFile();
                for (int i = 0; i < sheetCount; i++) {
                    Sheet sheet = spreadSheet.getSheet(asc ? i : (sheetCount - 1 - i));
                    showLog("- Sheet name: " + sheet.getName());

                    List<RowEntry> list = readSheet(sheet, MAX_COL, MAX_ROW);
                    boolean addHeader = (i == 0);
                    boolean addFooter = (i == sheetCount - 1);
                    writeXML(list, planArc, sheet.getName(), addHeader, addFooter, true);
                }
            } else {
                for (int i = 0; i < sheetCount; i++) {
                    Sheet sheet = spreadSheet.getSheet(asc ? i : (sheetCount - 1 - i));
                    showLog("- Sheet name: " + sheet.getName());

                    List<RowEntry> list = readSheet(sheet, MAX_COL, MAX_ROW);

                    xmlFile = new File(folder + "/" + sheet.getName() + ".xml");
                    xmlFile.getParentFile().mkdirs();
                    xmlFile.createNewFile();
                    writeXML(list, planArc, sheet.getName(), true, true, false);
                }
            }
        } catch (IOException ex) {
            showLog(" > ERROR: " + ex.getMessage() + " (" + xmlFile.getPath() + ")");
            setStatus("ERROR: " + ex.getMessage() + " (" + xmlFile.getPath() + ")", true);
            Logger.getLogger(OdfToXml.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
        return true;
    }

    public boolean writeXML(boolean isFull, String folder, ArrayList<Point> ranges) {
        try {
            if (isFull) {
                xmlFile.delete();
                xmlFile.createNewFile();
                for (int i = 0; i < ranges.size(); i++) {
                    for (int j = ranges.get(i).x; j <= ranges.get(i).y; j++) {
                        if (j < spreadSheet.getSheetCount()) {
                            Sheet sheet = spreadSheet.getSheet(j);
                            showLog("- Sheet name: " + sheet.getName());

                            List<RowEntry> list = readSheet(sheet, MAX_COL, MAX_ROW);
                            boolean addHeader = (i == 0 && j == ranges.get(i).x);
                            boolean addFooter = (i == ranges.size() - 1 && j == ranges.get(i).y);
                            writeXML(list, planArc, sheet.getName(), addHeader, addFooter, true);
                        } else {
                            showLog(" > ERROR: Invalid sheet ID (" + j + ")");
                            setStatus("ERROR: Invalid sheet ID (" + j + ")", true);
                        }
                    }
                }
            } else {
                for (int i = 0; i < ranges.size(); i++) {
                    for (int j = ranges.get(i).x; j <= ranges.get(i).y; j++) {
                        if (j < spreadSheet.getSheetCount()) {
                            Sheet sheet = spreadSheet.getSheet(j);
                            showLog("- Sheet name: " + sheet.getName());

                            List<RowEntry> list = readSheet(sheet, MAX_COL, MAX_ROW);

                            xmlFile = new File(folder + "/" + sheet.getName() + ".xml");
                            xmlFile.getParentFile().mkdirs();
                            xmlFile.createNewFile();
                            writeXML(list, planArc, sheet.getName(), true, true, false);
                        } else {
                            showLog(" > ERROR: Invalid sheet ID (" + j + ")");
                            setStatus("ERROR: Invalid sheet ID (" + j + ")", true);
                        }
                    }
                }
            }
        } catch (IOException ex) {
            showLog(" > ERROR: " + ex.getMessage() + " (" + xmlFile.getPath() + ")");
            setStatus("ERROR: " + ex.getMessage() + " (" + xmlFile.getPath() + ")", true);
            Logger.getLogger(OdfToXml.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
        return true;
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
}
