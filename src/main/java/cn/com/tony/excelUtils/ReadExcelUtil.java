package cn.com.tony.excelUtils;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;


public class ReadExcelUtil {

    static Logger logger = LoggerFactory.getLogger(ReadExcelUtil.class);
    //默认单元格内容为数字时格式
    private static DecimalFormat df = new DecimalFormat("0");
    // 默认单元格格式化日期字符串   
    private static SimpleDateFormat sdf = new SimpleDateFormat(  "yyyy-MM-dd HH:mm:ss");
    // 格式化数字  
    private static DecimalFormat nf = new DecimalFormat("0.00");

    private static DecimalFormat df1 = new DecimalFormat("#");

    public static List<Map> readExcel(File file){
        if(file == null){
            return null;
        }
        return readExcel2007(file);

    }

    public static List<Map> readExcel2007(File file){
        try{
            List<Map> rowList = new ArrayList<>();
            Map colList;
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            XSSFCell cell;
            XSSFRow keyRow = sheet.getRow(0);
            boolean flag = true;
            for(int i = sheet.getFirstRowNum()+1; i < sheet.getPhysicalNumberOfRows() ; i++ ){
                row = sheet.getRow(i);
                colList = new LinkedHashMap();
                if(!flag){
                    break;
                }
                for( int j = row.getFirstCellNum(); j < row.getLastCellNum() ;j++){
                    cell = row.getCell(j);
                    if(j==row.getFirstCellNum() && "".equals(checkType(cell).toString())){
                        flag =false;
                        break;
                    }
                    String key = keyRow.getCell(j).toString();
                    for(int k=10;k<14;k++){
                        key = key.replaceAll(String.valueOf((char)k), "");
                    }

                    if(cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
                        //当该单元格为空
                        if(j == row.getLastCellNum()){//判断是否是该行中最后一个单元格
                            break;
                        }
                        colList.put(key,"");
                        continue;
                    }

                    colList.put(key,checkType(cell));
                }//end for j
                if(colList.size()!=0){
                    rowList.add(colList);
                }
            }//end for i

            return rowList;
        }catch(Exception e){
            logger.error("导入错误",e);
            return null;
        }
    }

    public static Object checkType(XSSFCell cell){
        Object value;
        switch(cell.getCellType()){
            case XSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
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
            case XSSFCell.CELL_TYPE_BOOLEAN:
                value = Boolean.valueOf(cell.getBooleanCellValue());
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                value = "";
                break;
            default:
                value = cell.toString();
        }// end switch
        return value;
    }
    public static DecimalFormat getDf() {
        return df;
    }
    public static void setDf(DecimalFormat df) {
        ReadExcelUtil.df = df;
    }
    public static SimpleDateFormat getSdf() {
        return sdf;
    }
    public static void setSdf(SimpleDateFormat sdf) {
        ReadExcelUtil.sdf = sdf;
    }
    public static DecimalFormat getNf() {
        return nf;
    }
    public static void setNf(DecimalFormat nf) {
        ReadExcelUtil.nf = nf;
    }


    public static void main(String[] args) {
        List<Map> temp = ReadExcelUtil.readExcel(new File("C:\\Users\\lenovo\\Desktop\\项目基本信息表(1).xlsx"));
        for(Map obj:temp){
            System.out.printf(obj.toString());
        }
    }
}

	    			