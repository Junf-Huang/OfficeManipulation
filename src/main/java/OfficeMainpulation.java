import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;

public class OfficeMainpulation {

    public static void main(String []args) throws Exception{

        Workbook wb=new HSSFWorkbook();
//创建sheet1
        Sheet sheet = wb.createSheet("sheet1");
//创建sheet2
        wb.createSheet("hjj");
//创建行
        Row row = sheet.createRow(0);
//创建单元格
        row.createCell(0).setCellValue("省份证号");
        row.createCell(1).setCellValue("年龄");
        row.createCell(2).setCellValue("性别");
        row.createCell(3).setCellValue("身高");
        //创建一个输入流
        FileOutputStream fileOutputStream = new FileOutputStream("/home/junf/Documents/test.xlsx");
//写入
        wb.write(fileOutputStream);
    }
}
