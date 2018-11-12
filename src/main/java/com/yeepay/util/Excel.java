package com.yeepay.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author pangkun
 * @date 2018/11/12 下午2:28
 */
public class Excel {
    private static XSSFWorkbook xssfWorkbook = null;

    public void create(String fileName) throws IOException {
        FileOutputStream out = new FileOutputStream(
                new File(fileName));
        xssfWorkbook = new XSSFWorkbook();
        xssfWorkbook.write(out);
        out.close();
    }

    public XSSFWorkbook open(String fileName) throws IOException {
        File file = new File(fileName);
        FileInputStream fIP = new FileInputStream(file);
        xssfWorkbook = new XSSFWorkbook(fIP);
        if (!(file.isFile() && (file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx")))) {
            throw new IOException("文件不是Excel或文件不存在");
        }
        return xssfWorkbook;
    }

    public ArrayList<ArrayList<ArrayList<String>>> read(String fileName) throws IOException {
        xssfWorkbook = open(fileName);
        ArrayList<ArrayList<ArrayList<String>>> all = new ArrayList<ArrayList<ArrayList<String>>>();

        //遍历xlsx中的sheet
        for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
            if (xssfSheet == null) {
                all.add(null);
                continue;
            }
            ArrayList<ArrayList<String>> lists = new ArrayList<ArrayList<String>>();
            Iterator<Row> iterator = xssfSheet.iterator();
            while (iterator.hasNext()) {
                ArrayList<String> list = new ArrayList<String>();
                XSSFRow next = (XSSFRow) iterator.next();
                Iterator<Cell> cells = next.iterator();
                while (cells.hasNext()) {
                    XSSFCell cell = (XSSFCell) cells.next();
                    list.add(calCell(cell));
                }
                lists.add(list);
            }
            all.add(lists);
        }
        return all;
/*
            // 对于每个sheet，读取其中的每一行


            for(int i=0;i<=xssfSheet.getLastRowNum();i++){
                ArrayList<String>list=new ArrayList<String>();
                XSSFRow row = xssfSheet.getRow(i);
                for(int j=0;j<=row.getRowNum();j++){
                    XSSFCell cell = row.getCell(j);
                    String s = calCell(cell);
                    list.add(s);
                }
                lists.add(list);
            }
            all.add(lists);
*/
        /**
         for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
         //                System.out.println(xssfSheet.getLastRowNum());
         XSSFRow xssfRow = xssfSheet.getRow(rowNum);
         if (xssfRow == null) continue;
         ArrayList<String> curarr = new ArrayList<String>();
         for (int columnNum = 1; columnNum <= xssfRow.getRowNum(); columnNum++) {
         XSSFCell cell = xssfRow.getCell(columnNum);

         System.out.println(cell.getStringCellValue());
         //                    String rawValue = cell.getRawValue();
         curarr.add(Trim_str(cell.getStringCellValue()));
         }
         ans.add(curarr);
         }
         all.add(ans);*/

/**

 XSSFSheet spreadsheet = xssfWorkbook.getSheetAt(0);
 Iterator<Row> rowIterator = spreadsheet.iterator();
 XSSFRow row;
 while (rowIterator.hasNext())
 {
 row = (XSSFRow) rowIterator.next();
 Iterator < Cell > cellIterator = row.cellIterator();
 while ( cellIterator.hasNext())
 {
 Cell cell = cellIterator.next();
 switch (cell.getCellType())
 {
 case Cell.CELL_TYPE_NUMERIC:
 System.out.print(
 cell.getNumericCellValue() + " \t\t " );
 break;
 case Cell.CELL_TYPE_STRING:
 System.out.print(cell.getR
 cell.getStringCellValue() + " \t\t " );
 break;
 }
 }
 System.out.println();
 }
 fis.close();*/
    }

    private String Trim_str(String str) {
        if (str == null)
            return null;
        return str.replaceAll("[\\s\\?]", "").replace("　", "");
    }


    private String calCell(Cell cell) {
        if (cell == null)
            return null;
        switch (cell.getCellTypeEnum()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case BLANK:
                return "";
            case _NONE:
                return null;
            case FORMULA:
                // 公式
                return String.valueOf(cell.getCellFormula());
        }
        return null;
    }


    public List<List<String[]>> readExcel(String fileName) throws IOException {
        Workbook workbook = open(fileName);
        List<List<String[]>> lists = new ArrayList<List<String[]>>();
        if (workbook != null) {
            int numberOfSheets = workbook.getNumberOfSheets();
            for (int sheetNum = 0; sheetNum < numberOfSheets; sheetNum++) {
                //获得当前sheet工作表
                List<String[]> list = new ArrayList<String[]>();
                Sheet sheet = workbook.getSheetAt(sheetNum);
                if (sheet == null) {
                    System.out.println("sheet为空");
                    lists.add(null);
                    continue;
                }
                //获得当前sheet的结束行
                int lastRowNum = sheet.getLastRowNum();
                for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
                    //获得当前行
                    Row row = sheet.getRow(rowNum);
                    if (row == null) {
                        System.out.println("行为空");
                        list.add(null);
                        continue;
                    }
                    //获得当前行的列数
                    int lastCellNum = row.getLastCellNum();
                    System.out.println("该行有" + lastCellNum + "列");
                    String[] cells = new String[row.getLastCellNum()];
                    //循环当前行
                    for (int cellNum = 0; cellNum < lastCellNum; cellNum++) {
                        Cell cell = row.getCell(cellNum);
                        cells[cellNum] = calCell(cell);
                        System.out.println(cells[cellNum] + "   ");
                    }
                    list.add(cells);
                }
                lists.add(list);
            }
            workbook.close();
        }
        return lists;
    }

    public void change() {

    }


}
