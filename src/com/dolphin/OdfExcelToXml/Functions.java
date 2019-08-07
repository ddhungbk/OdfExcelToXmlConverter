/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dolphin.OdfExcelToXml;

import java.awt.Point;
import java.util.ArrayList;

/**
 *
 * @author hungd
 */
public class Functions {
    
    public static boolean checkRanges(String s) {
        boolean result = true;
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if ((c >= '0' && c <= '9') || c == ',' || c == ';' || c == '-' || c == ':' || c == ' ') {

            } else {
                return false;
            }
        }
        return result;
    }

    public static  ArrayList<String> parseString(String src, char c) {
        ArrayList<String> results = new ArrayList<>();
        String s = src;
        if (s.indexOf(c) == -1) {
            results.add(s);
        } else {
            do {
                results.add(s.substring(0, s.indexOf(c)).trim());
                if (s.indexOf(c) < s.length()) {
                    s = s.substring(s.indexOf(c) + 1).trim();
                }
                if (s.indexOf(c) == -1) {
                    results.add(s);
                }
            } while (s.indexOf(c) != -1);
        }

        return results;
    }

    public static  int getFirstIndex(String src, char[] cs) {
        for (int i = 0; i < src.length(); i++) {
            char c = src.charAt(i);
            for (int j = 0; j < cs.length; j++) {
                if (c == cs[j]) {
                    return i;
                }
            }
        }
        return -1;
    }

    public static  ArrayList<String> parseString(String src, char[] c) {
        ArrayList<String> results = new ArrayList<>();
        String s = src;
        if (getFirstIndex(s, c) == -1) {
            results.add(s);
        } else {
            do {
                results.add(s.substring(0, getFirstIndex(s, c)).trim());
                if (getFirstIndex(s, c) < s.length()) {
                    s = s.substring(getFirstIndex(s, c) + 1).trim();
                }
                if (getFirstIndex(s, c) == -1) {
                    results.add(s);
                }
            } while (getFirstIndex(s, c) != -1);
        }
        return results;
    }

}
