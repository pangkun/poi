package com.yeepay;


import com.yeepay.util.Excel;

import java.io.IOException;
import java.util.List;


/**
 * @author pangkun
 * @date 2018/11/12 下午2:19
 */
public class Main {
    public static void main(String[] args) throws IOException {

        Excel excel = new Excel();
        excel.create("4.xlsx");
        /**
        ArrayList<ArrayList<ArrayList<String>>> file = excel.read("1.xlsx");
        for (ArrayList<ArrayList<String>> lists : file) {
            for (ArrayList<String> list : lists) {
                for (String s : list) {
                    System.out.print(s + "\t\t");
                }
                System.out.println();
            }
            System.out.println("--");
        }      */
        List<List<String[]>> lists = excel.readExcel("2.xlsx");
        for(List<String[]>strings:lists) {
            if(strings==null){
                System.out.println("sheet为空");
                continue;
            }
            for (String[] string : strings) {
                if (string == null) {
                    System.out.println("行为空");
                    continue;
                }
                for (String s : string) {
                    System.out.print(s + "  .");
                }
                System.out.println();
            }
            System.out.println("-----");
        }
        excel.create(lists,"3.xlsx");

    }


}
