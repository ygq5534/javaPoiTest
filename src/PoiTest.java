import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

public class PoiTest {
    public static void main(String[] args){
        List<List<Object>> data ;
        List<Object> dataRow;
        try {
            data = readExcel(new File("E:\\工作文件\\2016研--敏捷开发与系统重构\\" +
                    "编程操练\\编程操练_井底压力\\数据\\输入数据.xls"));
            for (int i = 4;i<data.size();i++){
                System.out.print(i+1+"\t");
                System.out.print(data.get(i).get(0)+"\t");
                System.out.print(data.get(i).get(1)+"\t");
                System.out.print(data.get(i).get(2)+"\t");
                System.out.println(data.get(i).get(3));
            }
        }catch (IOException e){
            e.printStackTrace();
        }

    }
    public static List<List<Object>> readExcel(File file)throws IOException{
        String fileName = file.getName();
        String extension = fileName.lastIndexOf(".") == -1 ? "" :fileName.
                substring(fileName.indexOf(".")+1);
        if ("xls".equals(extension)){
            return read2003Excel(file);
        }else if("xlsx".equals(extension)){
            return read2007Excel(file);
        }else {
            throw new IOException("不支持的文件类型");
        }
    }

    public static List<List<Object>> read2007Excel(File file)
                throws IOException{
        List<List<Object>> list = new ArrayList<>();
        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = xwb.getSheetAt(0);
        Object value = null;
        XSSFRow row = null;
        XSSFCell cell = null;
        int counter = 0;
        for (int i = sheet.getFirstRowNum();counter<sheet
                .getPhysicalNumberOfRows();i++){
            row = sheet.getRow(i);
            if (row == null){
                continue;
            }else {
                counter++;
            }
            List<Object> array = new ArrayList<>();
            for (int j = row.getFirstCellNum();j <= row.getLastCellNum();j++){
                cell = row.getCell(j);
                if (j == (3) && cell == null){
                    value = list.get(i-1).get(3);
                    array.add(value);
                    continue;
                }
                DecimalFormat df = new DecimalFormat("0");
                SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
                DecimalFormat nf = new DecimalFormat("0.000");
                switch (cell.getCellType()){
                    case XSSFCell.CELL_TYPE_STRING:
                        //System.out.println(i+"行"+j+"列 is String type");
                        value = cell.getStringCellValue();
                        break;
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        //System.out.println(i+"行"+j+"列is Number type; DataFormt:"
                            //+cell.getCellStyle().getDataFormatString());
                        if ("@".equals(cell.getCellStyle().getDataFormatString())){
                            value = df.format(cell.getNumericCellValue());
                        }else if ("General".equals(cell.getCellStyle().getDataFormatString())){
                            value = nf.format(cell.getNumericCellValue());
                        }else {
                            value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
                        }
                        break;
                    case XSSFCell.CELL_TYPE_BOOLEAN:
                        //System.out.println(i+"行"+j+"列 is Boolean type.");
                        value = cell.getBooleanCellValue();
                        break;
                    case XSSFCell.CELL_TYPE_BLANK:
                        //System.out.println(i+"行"+j+"列 is Blank type.");
                        value = "";
                        break;
                    default:
                        //System.out.println(i+"行"+j+"列is default type.");
                        value = cell.toString();
                }
                if (value == null||"".equals(value)){
                    continue;
                }
                array.add(value);
            }
            list.add(array);
        }
        return list;

    }

    public static List<List<Object>> read2003Excel(File file)throws IOException{
        List<List<Object>> list = new ArrayList<>();
        HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = hwb.getSheetAt(0);
        Object value = null;
        HSSFRow row = null;
        HSSFCell cell = null;
        int counter = 0;
        for (int i = sheet.getFirstRowNum();counter < sheet
                .getPhysicalNumberOfRows();i++){
            row = sheet.getRow(i);
            if (row == null){
                continue;
            }else {
                counter++;
            }
            List<Object> array = new ArrayList<>();
            for (int j = row.getFirstCellNum();j <= row.getLastCellNum();
                    j++){
                cell = row.getCell(j);
                if (cell == null) {
                    continue;
                }
                DecimalFormat df = new DecimalFormat("0");
                SimpleDateFormat sdf = new SimpleDateFormat(
                          "HH:mm:ss");
                DecimalFormat nf = new DecimalFormat("0.000");
                switch (cell.getCellType()) {
                    case HSSFCell.CELL_TYPE_STRING:
                       // System.out.println(i + "行" + j + "列 is String type.");
                        value = cell.getStringCellValue();
                        break;
                    case HSSFCell.CELL_TYPE_NUMERIC:
                       // System.out.println(i + "行" + j + "列is Number type; Dateformt:"
                          //      + cell.getCellStyle().getDataFormatString());
                        if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                            value = df.format(cell.getNumericCellValue());
                        } else if ("General".equals(cell.getCellStyle()
                                .getDataFormatString())) {
                            value = nf.format(cell.getNumericCellValue());
                        } else {
                            value = sdf.format(HSSFDateUtil.getJavaDate(cell
                                    .getNumericCellValue()));
                        }
                        break;
                    case HSSFCell.CELL_TYPE_BOOLEAN:
                        //System.out.println(i + "行" + j + "列 is Boolean type.");
                        value = cell.getBooleanCellValue();
                        break;
                    case HSSFCell.CELL_TYPE_BLANK:
                        //System.out.println(i + "行" + j + "列 is Blank type.");
                        if (j == 3 && counter > 3){
                            value = list.get(counter-1).get(3);
                        }else if (j == 4 && counter > 3){
                            value = list.get(counter-1).get(4);
                        }else if (j == 5 && counter>3){
                            value = list.get(counter-1).get(5);
                        }else if (j == 7 && counter>3){
                            value = list.get(counter-1).get(7);
                        }else {
                            value = "";
                        }
                        break;
                    default:
                        //System.out.println(i + "行" + j + "列 is default type.");
                        value = cell.toString();
                }
                if (value == null||"".equals(value)){
                    continue;
                }
                array.add(value);
            }
            list.add(array);
        }
        return list;
    }
}
