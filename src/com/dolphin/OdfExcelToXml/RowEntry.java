package com.dolphin.OdfExcelToXml;

import java.util.ArrayList;

public class RowEntry {

    private int columnCount;
    private ArrayList<String> contents;

    public RowEntry(int columnCount, ArrayList<String> contents) {
        this.columnCount = columnCount;
        this.contents = contents;
    }

    public String toString() {
        String content = "";
        for (int i = 0; i < columnCount - 1; i++) {
            content += contents.get(i) + " | ";
        }
        content += contents.get(columnCount - 1);
        return content;
    }

    public int getColumnCount() {
        return columnCount;
    }

    public void setColumnCount(int columnCount) {
        this.columnCount = columnCount;
    }

    public String getField(int id) {
        return contents.get(id);
    }
}
