import office.Excel;

public class OfficeMainpulation {

    public static void main(String []args) throws Exception{

        String excelName = "test";
        String sheetName = "sheet1";
        String title = "list";
        String[] value2={"张三","李四", "小伙伴","最近","比较忙", "大家"};
        String[] value1={"张三","李四", "小伙伴","最近","比较忙"};

        Excel excel = new Excel(excelName);
        excel.switchSheet(sheetName);
        excel.createTitle(title);
        excel.createRows(value1, 2);
        excel.createOneRow(value2);
        excel.save();
    }


}
