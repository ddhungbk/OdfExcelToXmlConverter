/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dolphin.OdfExcelToXml;

import java.util.ArrayList;
import javax.swing.JComboBox;

/**
 *
 * @author hungd
 */
public class MyJCombobox extends JComboBox<SheetEntry> {

    private final JComboBox<String> cb;
    private final ArrayList<SheetEntry> arrItems; // list of Sheets

    public MyJCombobox(JComboBox cb) {
        this.cb = cb;
        arrItems = new ArrayList<>();
    }

    @Override
    public void removeAllItems() {
        super.removeAllItems();
        cb.removeAllItems();
        arrItems.clear();
    }

    @Override
    public void addItem(SheetEntry item) {
        cb.addItem(item.getName());
        arrItems.add(item);
    }

    @Override
    public SheetEntry getSelectedItem() {
        if (cb.getSelectedIndex() >= 0 && cb.getSelectedIndex() < arrItems.size()) {
            return arrItems.get(cb.getSelectedIndex());
        }
        if (arrItems.size() > 1) {
            return new SheetEntry("", -1, false);
        } else {
            return new SheetEntry("", -1, true);
        }
    }

    @Override
    public int getItemCount() {
        return arrItems.size();
    }

}
