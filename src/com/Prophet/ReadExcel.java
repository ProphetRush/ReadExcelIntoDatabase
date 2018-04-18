package com.Prophet;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Array;
import java.sql.Connection;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
 * Created by Prophet on 4/18/2018.
 */
public class ReadExcel {

    public static void readExcelIntoDatabase(String filePath) throws Exception{
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(filePath)));
        DatabaseOperations dbOp = new DatabaseOperations();
        Connection conn = dbOp.getDBConnection();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if(conn.getMetaData().getTables(null, null, "stu"+sheet.getSheetName(), null).next()){
                dbOp.insertDataFromExcelSheet(conn, sheet);
            }else{
                dbOp.createTable(conn, sheet);
                dbOp.insertDataFromExcelSheet(conn, sheet);
            }
        }
        conn.close();
    }


    public static List<String> readColumnName(Sheet sheet){
        List<String> list = new ArrayList<>();
        if(sheet.getLastRowNum()>=0){
            Row row = sheet.getRow(0);
            for (int i = 0; i < row.getLastCellNum(); i++) {
                list.add(row.getCell(i).toString().replace(".",""));
            }
        }
        return list;
    }

}
