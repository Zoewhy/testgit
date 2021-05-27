package com.wj;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Poi {

        public static void main(String[] args) throws Exception
        {

            String filePath = "C:\\Users\\Administrator\\Desktop\\91家非上市银行\\广东华兴银行\\广东华兴银行股份有限公司[F0200924.00]-利润表.xlsx";

            File file = new File(filePath);

            int startRowIndex = 0;//从第二行开始读取，第一行默认为列名

            String[][] content = getData(file, startRowIndex);//从excel读取数据放到“行*列”的二维数组中

            BufferedWriter writer = new BufferedWriter(new FileWriter(new File("D:\\WriteTxt.txt")));   //将生成的二维数组写入txt

            int rowLength = content.length;

            for(int i=0;i<rowLength;i++)
            {

                for(int j=0;j<content[i].length;j++)
                {

                    System.out.print(content[i][j]+"\t");

                    writer.write(content[i][j]+"\t");

                }

                writer.write("\r\n");

                System.out.println();
            }

            writer.close();
        }


        public static String[][] getData(File file, int startRowIndex)throws FileNotFoundException, IOException
        {
            // 打开文件

            Workbook wb;

            Sheet st;

            Row row;

            Cell cell;

            FileInputStream fis = new FileInputStream(file);

            BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));

            try
            {
                wb = new HSSFWorkbook(fis);    //.xls读取
            }

            catch(Exception e)
            {
                wb = new XSSFWorkbook(bis);//.xlsx读取
            }

            List<String[]> rowArray = new ArrayList<String[]>();//所有行组成的数组

            int maxColumnSize = 0;//二维数组的列最大值

            for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) //遍历sheet
            {

                st = wb.getSheetAt(sheetIndex);
                wb.setSheetName(sheetIndex,"Sheet1");
                System.out.println(st.getSheetName());
                wb.write(new FileOutputStream("C:\\Users\\Administrator\\Desktop\\91家非上市银行\\广东华兴银行\\F广东华兴银行股份有限公司[F0200924.00]-利润表.xlsx"));
                for (int rowIndex = startRowIndex; rowIndex <= st.getLastRowNum(); rowIndex++)//遍历行
                {

                    row = st.getRow(rowIndex);

                    boolean isCellValueNull = true;

                    if (row == null) //空行跳过
                    {
                        continue;
                    }

                    int columnSize = row.getLastCellNum();  //每行的列数

                    if(columnSize>maxColumnSize)//为确保数组的列容量
                    {
                        maxColumnSize = columnSize;
                    }

                    String[] rowValues = new String[columnSize];//每行的值，一维数组

                    Arrays.fill(rowValues, "");//填充默认空值

                    for (short columnIndex = 0; columnIndex < columnSize; columnIndex++) //遍历列
                    {

                        String value = "";

                        cell = row.getCell(columnIndex);

                        if (cell != null && cell.getCellType()== XSSFCell.CELL_TYPE_STRING)
                        {
                            value = cell.getStringCellValue();
                        }

                        if (cell != null && cell.getCellType()==XSSFCell.CELL_TYPE_NUMERIC)
                        {
                            value = new DecimalFormat("0").format(cell.getNumericCellValue());
                        }

                        if (value.trim().equals("")) //单元格内容为空则跳过
                        {
                            continue;
                        }

                        rowValues[columnIndex] = value;

                        isCellValueNull = false;
                    }

                    if (!isCellValueNull)//空行则跳过,包括仅有空格的行
                    {
                        rowArray.add(rowValues);
                    }

                }

            }



            bis.close();
            fis.close();

            String[][] rowColumnArray = new String[rowArray.size()][maxColumnSize];

            for (int i = 0; i < rowColumnArray.length; i++)
            {
                rowColumnArray[i] = (String[]) rowArray.get(i);
            }

            return rowColumnArray;

        }
    }

