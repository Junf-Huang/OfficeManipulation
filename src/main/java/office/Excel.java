package office;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;

public class Excel {

    public static final String XLSX_FILE_PATH = "/home/junf/Documents/";
    private String xlsxName;
    private Workbook wb;

    public Excel(String xlsxName) {

        this.xlsxName = XLSX_FILE_PATH + xlsxName + ".xlsx";
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        this.wb = new HSSFWorkbook();
        //创建sheet1
        Sheet sheet = wb.createSheet("sheet1");

        //创建行
        Row row = sheet.createRow(0);
        //创建单元格
        row.createCell(0).setCellValue("省份证号");
        row.createCell(1).setCellValue("年龄");
        row.createCell(2).setCellValue("性别");
        row.createCell(3).setCellValue("身高");
    }

    public void createSheet(String sheetName) {
        wb.createSheet(sheetName);
    }

    public Sheet getSheet(String sheetName) {
        return wb.getSheet(sheetName);
    }

    public void saveXlsx() throws Exception {

        //创建一个输入流
        FileOutputStream fileOutputStream = new FileOutputStream("/home/junf/Documents/test.xlsx");
        //写入
        wb.write(fileOutputStream);
    }
}
