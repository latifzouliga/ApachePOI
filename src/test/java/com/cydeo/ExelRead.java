package com.cydeo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ExelRead {

    @Test
    public void test1() throws IOException {

        FileInputStream fileInputStream = new FileInputStream("EmployeesSalary.xlsx");

        Workbook workbook = new XSSFWorkbook(fileInputStream);

       Sheet sheet = workbook.getSheet("Employees");

       // print first row first cell
        Cell cell = sheet.getRow(0).getCell(0);
        System.out.println(cell);

        // print last cell row number
        System.out.println(sheet.getLastRowNum());

        // print all rows and all cells
        for(int rowNum = 0; rowNum < sheet.getLastRowNum();rowNum++ ){

            short lastCellNum = sheet.getRow(rowNum).getLastCellNum();
            for (int cellNum = 0; cellNum < lastCellNum; cellNum++){
                System.out.print(sheet.getRow(rowNum).getCell(cellNum)+ " : ");
            }
            System.out.println();
        }

        // get the last name of Ahlam
        for(int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++){
            short lastCellNum = sheet.getRow(rowNum).getLastCellNum();
            for (int cellNum = 0; cellNum < lastCellNum; cellNum++){
                if (sheet.getRow(rowNum).getCell(cellNum).toString().equals("Ahlam")){
                    System.out.println(sheet.getRow(rowNum).getCell(1));
                }
            }
        }

        System.out.println("=======print all data with iterator");
        Iterator<Row> it = sheet.iterator();

        while (it.hasNext()){
            Row row = it.next();
            Iterator<Cell> iterator = row.iterator();
            while (iterator.hasNext())
                System.out.print(iterator.next()+" ---> ");
            System.out.println();
        }

    }
}
