package com.wj;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.LinkedList;
import java.util.List;

public class TestPOI {
    public static void main(String[] args) throws Exception {
       String filename = "C:\\Users\\Administrator\\Desktop\\商业银行数据库 - 副本 (3) - 副本.xlsx";
       Workbook wb = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(filename)));
       Sheet sheet = wb.getSheetAt(wb.getActiveSheetIndex());
       List<String> items = new LinkedList<String>();
       Row row = sheet.getRow(sheet.getFirstRowNum());
        for (int i = 3; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if(cell!=null) {
                items.add(cell.getStringCellValue());
            }
        }
        System.out.println(items);
        int count = 0;
        int pre = -11;
        for(int rowIndex = 1;rowIndex<sheet.getLastRowNum();rowIndex++){
            count++;
            if(count - pre <   10){
                continue;
            }
            Row row1 = sheet.getRow(rowIndex);
            String name = row1.getCell(0).getStringCellValue();
            String bankname = name!=null && name.endsWith("银行")? name :null;
            if(bankname !=null){
                String targetpath = "C:\\Users\\Administrator\\Desktop\\53家上市银行\\"+bankname;
                File[] files = getFiles(targetpath);
                if (files!=null){
                    System.out.println(files.length);
                    System.out.println(bankname);
                    int file_count = 0;
                    List<String> mylist;
                    int item_count = 0;
                    for (String item : items){
                        System.out.println(item);
                        while(file_count<files.length){
                            System.out.println(files[file_count].getAbsolutePath());
                            mylist = getItems(files[file_count],item);
                            int j = 0;
                            if(mylist.size()!=0){
                                pre = count;
                                System.out.println(mylist);
                                for(int i =count+10-mylist.size();i<count+10;i++){
                                        int cellIndex = item_count+3;
                                        Row row2 = sheet.getRow(i);
                                        Cell cell = row2.getCell(cellIndex);
                                        if (cell==null){
                                            cell =  row2.createCell(cellIndex);
                                        }
                                        cell.setCellValue(mylist.get(j++));
                                }
                            }
                            mylist.clear();
                            j=0;
                            file_count++;
                        }
                        item_count++;
                        file_count=0;
                    }

                }
            }
        }
        wb.write(new FileOutputStream("C:\\Users\\Administrator\\Desktop\\商业银行数据库 - 副本 (3) - 副本.xlsx"));
    }



    public static List<String> getItems(File file,String item) throws IOException {
        InputStream is = new FileInputStream(file);
        BufferedInputStream bis =  new BufferedInputStream(is);
        Workbook wb = new XSSFWorkbook(bis);
        List<String> result = new LinkedList<String>();
        for(int sheetIndex = 0 ;sheetIndex<wb.getNumberOfSheets();sheetIndex++){
            Sheet sheet = wb.getSheetAt(sheetIndex);
            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                Row row = sheet.getRow(rowIndex);
                if (row !=null){
                    Cell cell = row.getCell(0);
                    if (cell!=null){
                        String mystr = cell.getStringCellValue();
                        if (mystr != null  && item.trim().contains(mystr.trim())) {
                            for (int i = 2; i < row.getLastCellNum(); i++) {
                                result.add(String.valueOf(row.getCell(i).getCellType()==0?row.getCell(i).getNumericCellValue():row.getCell(i).getStringCellValue()));
                            }
                            break;
                        }

                    }
                }
            }
        }
        is.close();
        bis.close();
        return result;
    }


    public static File[] getFiles(String path) {
       return new File(path).listFiles();
    }
}
