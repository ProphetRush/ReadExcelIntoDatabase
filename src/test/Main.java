package test;

import com.Prophet.ReadExcel;

public class Main {

    public static void main(String[] args) {
        try{
            ReadExcel.readExcelIntoDatabase("C:\\ClassTest-DAD\\src\\test\\stu.xlsx");
        }catch (Exception e){
            e.printStackTrace();
        }
    }


}
