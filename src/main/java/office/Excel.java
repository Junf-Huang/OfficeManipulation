package office;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class Excel {

    private static final String XLSX_FILE_PATH = "/home/junf/Documents/";
    private String xlsx;
    private Workbook wb;
    private Sheet currentSheet;


    //initalize Excel
    public Excel(String xlsxName) {

        this.xlsx = XLSX_FILE_PATH + xlsxName + ".xlsx";
        //Creating a Workbook from an Excel file (.xls or .xlsx)
        this.wb = new HSSFWorkbook();
    }

    //if exit sheet, switch it, else create new sheet
    public Sheet switchSheet(String sheetName) {
        if (getSheet(sheetName)==null)
            this.currentSheet = wb.createSheet(sheetName);
        return this.currentSheet;
    }

    //get special sheet
    public Sheet getSheet(String sheetName) {
        return wb.getSheet(sheetName);
    }

    //insert a special column and set cell value.
    public void createOneRow(int NoRow, String[] values) {

        Row row = currentSheet.createRow(NoRow);
        for(int i = 0; i < values.length; i++ ) {
            row.createCell(i).setCellValue(values[i]);
        }

    }

    //insert into a row and set values
    public void createOneRow(String[] values) {
        Row row = currentSheet.createRow(currentSheet.getLastRowNum()+1);
        for(int i = 0; i < values.length; i++ ) {
            row.createCell(i).setCellValue(values[i]);
        }
    }

    //insert into rows and set values
    public void createRows(String[] values, int columnNum) {

        int lastRow = currentSheet.getLastRowNum()+1;

        int rowNum;
        if(values.length%columnNum==0)
            rowNum = values.length/columnNum;
        else
            rowNum = values.length/columnNum+1;

        for (int i = 0; i < rowNum; i++) {
            Row row = currentSheet.createRow(i+lastRow);
            for(int j = 0; j < columnNum; j++) {
                if(j+i*columnNum >= values.length)
                    break;
                row.createCell(j).setCellValue(values[j+i*columnNum]);
            }
        }
    }

    public void createTitle(String title){
        if(this.currentSheet!=null)
            currentSheet.createRow(0).createCell(0).setCellValue(title);
        else
            System.out.println("Please choose a sheet");
    }

    //save excel to the special path
    public void save() throws Exception {
        //创建一个输入流
        FileOutputStream fileOutputStream = new FileOutputStream(xlsx);
        //写入
        wb.write(fileOutputStream);
    }

    private static void printCellValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                System.out.print(cell.getBooleanCellValue());
                break;
            case STRING:
                System.out.print(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.print(cell.getDateCellValue());
                } else {
                    System.out.print(cell.getNumericCellValue());
                }
                break;
            case FORMULA:
                System.out.print(cell.getCellFormula());
                break;
            case BLANK:
                System.out.print("");
                break;
            default:
                System.out.print("");
        }

        System.out.print("\t");
    }


    //read excel

    public void read() throws Exception{
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(xlsx));
        HSSFSheet sheet = null;
    }

}
