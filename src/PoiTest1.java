import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;
import java.util.Map;

public class PoiTest1 {
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";
    public static void main(String[] args){

    }

    public static void writeExcel(List<List<Object>> dataList,
                                  String finalXlsxPath)throws IOException{
        OutputStream out = null;
        try {
            File finalXlsxFile = new File(finalXlsxPath);
            Workbook workbook = getWorkbook(finalXlsxFile);

            Sheet sheet = workbook.getSheetAt(0);
//            int rowNumber = sheet.getLastRowNum();
//            System.out.println("原始数据总行数，除属性列："+rowNumber);
//            for (int i = 1;i <= rowNumber;i++){
//                Row row = sheet.getRow(i);
//                sheet.removeRow(row);
//            }
            out = new FileOutputStream(finalXlsxPath);
            workbook.write(out);
            Cell cell39 = sheet.getRow(3).getCell(9);
            cell39.setCellValue("静液柱压力");
            Cell cell310 = sheet.getRow(3).getCell(10);
            cell310.setCellValue("沿程摩阻");
            Cell cell311 = sheet.getRow(3).getCell(11);
            cell311.setCellValue("井底压力");
            Cell cell49 = sheet.getRow(4).getCell(9);
            cell49.setCellValue("Mpa");
            Cell cell410 = sheet.getRow(4).getCell(10);
            cell410.setCellValue("Mpa");
            Cell cell411 = sheet.getRow(4).getCell(11);
            cell411.setCellValue("Mpa");
            for (int j = 0;j < dataList.size();j++){
                Row row = sheet.getRow(j+5);
                Cell cell9 = row.getCell(9);
                cell9.setCellValue(dataList.get(j).get(0).toString());
                Cell cell10 = row.getCell(10);
                cell10.setCellValue(dataList.get(j).get(1).toString());
                Cell cell11 = row.getCell(11);
                cell11.setCellValue(dataList.get(j).get(2).toString());
            }
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            try {
                if (out != null){
                    out.flush();
                    out.close();
                }
            }catch (IOException e){
                e.printStackTrace();
            }
        }
        System.out.println("数据写入成功！");
    }
    public static Workbook getWorkbook(File file)throws IOException{
        Workbook wb = null;
        FileInputStream in = new FileInputStream(file);
        if (file.getName().endsWith(EXCEL_XLS)){
            wb = new HSSFWorkbook(in);
        }else if (file.getName().endsWith(EXCEL_XLSX)){
            wb = new XSSFWorkbook();
        }
        return wb;
    }
}
