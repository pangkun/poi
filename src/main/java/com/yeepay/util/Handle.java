package com.yeepay.util;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by 庞昆 on 2018-11-13.
 */
public interface Handle {
    void create(String fileName) throws IOException;
    void create(List<List<String[]>> lists, String fileName) throws IOException;
    XSSFWorkbook open(String fileName) throws IOException;
    ArrayList<ArrayList<ArrayList<String>>> read(String fileName) throws IOException;
    List<List<String[]>> readExcel(String fileName) throws IOException;

}
