/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dolphin.OdfExcelToXml;

/**
 *
 * @author hungd
 */
public class SheetEntry{

    private String name;
    private int id;
    private boolean sheet;

    public SheetEntry(String name, int id, boolean sheet) {
        this.name = name;
        this.id = id;
        this.sheet = sheet;
    }

    public String getName() {
        return name;
    }

    public int getId() {
        return id;
    }

    public boolean isSheet() {
        return sheet;
    }

}
