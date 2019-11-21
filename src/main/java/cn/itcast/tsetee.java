package cn.itcast;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class tsetee {
    public static void main(String[] args) throws Exception {
        //创建xxsf
        Workbook wb = new XSSFWorkbook();
        //创建工作簿
        Sheet sheet = wb.createSheet();
        //在哪一行
        Row row = sheet.createRow(1);
        //在那个单元格
        Cell cell = row.createCell(3);
        //在单元格中设置值
        cell.setCellValue("wdafaf张三");
        File file = new File("D:/w/a.xlsx");
        FileOutputStream pis = new FileOutputStream(file);
        wb.write(pis);
        pis.close();
    }
}
