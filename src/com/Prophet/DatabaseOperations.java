package com.Prophet;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.Statement;
import java.util.List;
import java.util.Properties;

/**
 * Created by Prophet on 4/18/2018.
 */
public class DatabaseOperations {

    public Connection getDBConnection() throws Exception{
        Properties p = new Properties();
        InputStream in = this.getClass().getResourceAsStream("dbSettings.properties");
        p.load(in);
        System.out.println("Connecting to database...");
        Class.forName(p.getProperty("driverName"));
        if(p.getProperty("url").contains("STUDENT")){
            return DriverManager.getConnection(p.getProperty("url"), p.getProperty("username"), p.getProperty("password"));
        }else{
            createDatabase(DriverManager.getConnection(p.getProperty("url"), p.getProperty("username"), p.getProperty("password")));
            return getDBConnection();
        }

    }

    public void createDatabase(Connection conn) throws Exception{
        Statement s = conn.createStatement();
        String sql = "CREATE DATABASE STUDENT";
        System.out.println("Creating Database STUDENT...");
        s.executeUpdate(sql);
        Properties p = new Properties();
        p.load(this.getClass().getResourceAsStream("dbSettings.properties"));
        p.setProperty("url", p.getProperty("url")+"STUDENT");
        OutputStream fos = new FileOutputStream(this.getClass().getResource("dbSettings.properties").getFile());
        p.store(fos, null);
        System.out.println("Create Database Success!");
        s.close();
    }

    public void createTable(Connection conn, Sheet sheet) throws Exception{
        List<String> columnName = ReadExcel.readColumnName(sheet);
        String sql = "CREATE TABLE stu"+sheet.getSheetName()+"(";
        for (int i = 0; i < columnName.size()-1; i++) {
            sql+=(columnName.get(i) + " DOUBLE NOT NULL,");
        }
        sql+=(columnName.get(columnName.size()-1) + " DOUBLE NOT NULL);");
        PreparedStatement ps = conn.prepareStatement(sql);
        System.out.println("Executing SQL: "+sql);
        ps.execute();
        ps.close();
    }

    public void insertDataFromExcelSheet(Connection conn, Sheet sheet) throws Exception{
        String longSQL = "INSERT INTO stu"+sheet.getSheetName()+" values";
        if(sheet.getLastRowNum()>=1){
            for (int i = 1; i < sheet.getLastRowNum(); i++) {
                String sql = "(";
                Row row = sheet.getRow(i);
                for (int j = 0; j < row.getLastCellNum()-1; j++) {
                    sql+=row.getCell(j)+", ";
                }
                sql+=row.getCell(row.getLastCellNum()-1)+"),";
                longSQL+=sql;
            }
        }
        longSQL = longSQL.substring(0, longSQL.length()-1)+";";
        PreparedStatement ps = conn.prepareStatement(longSQL);
        System.out.println("Executing SQL Statement:"+longSQL);
        ps.execute();
        ps.close();
    }
}
