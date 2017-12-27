package cn.com.tony.excelUtils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created by Tony on 2017/12/25.
 * cn.com.cvinfo.excelUtils.TestMyBatis
 * 描述：
 */
public class MyExcelUtil {
    private static void excelExport(String fileName , String sheetName, List<Map<String,Object>> dataList, Integer[][] cellRegion , String[] titleList, HttpServletResponse response){
        //创建workbook
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建sheet页
        HSSFSheet sheet = null;
        CellRangeAddress region = null;
        if (StringUtils.isNoneBlank(sheetName)){
            sheet= workbook.createSheet(sheetName);
        }else {
            sheet=workbook.createSheet("sheet0");
        }

        sheet.setColumnWidth(0, 3766); //第一个参数代表列id(从0开始),第2个参数代表宽度值  参考 ："2012-08-10"的宽度为2500
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell1 = row.createCell(0);
        cell1.setCellValue(fileName);
        cell1.setCellStyle(buildTitleCellStyle(workbook));
        for (int i = 0 ; i < dataList.size() ; i ++) {
            row = sheet.createRow(i+1);
            Map<String, Object> map = dataList.get(i);
            for (int j = 0; j < map.size(); j++){
                HSSFCell cell = row.createCell(j);
                cell.setCellValue(map.get(titleList[j]).toString());
                cell.setCellStyle(buildCellStyle(workbook));
            }
        }
        if (cellRegion.length>0){
            for (Integer[] re :cellRegion) {
                region = new CellRangeAddress(re[0],re[1],re[2],re[3]);
                sheet.addMergedRegion(region);
            }
        }
        sheet.autoSizeColumn(1,true);
        try {
            response.reset();
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename="+ new String((fileName + ".xlsx").getBytes("gb2312"), "iso-8859-1"));
            workbook.write(response.getOutputStream());
        }catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static CellStyle buildTitleCellStyle(HSSFWorkbook workbook){
        // 生成一个样式
        CellStyle style = workbook.createCellStyle();
        // 设置这些样式
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中

        // 背景色
        style.setFillForegroundColor(HSSFColor.WHITE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setFillBackgroundColor(HSSFColor.WHITE.index);

        // 设置边框
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        // 自动换行
        style.setWrapText(true);

        // 生成一个字体
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 16);
        font.setColor(HSSFColor.BLACK.index);
        font.setFontName("宋体");

        style.setFont(font);
        return style;

    }

    public static CellStyle buildCellStyle(HSSFWorkbook workbook){
        // 生成一个样式
        CellStyle style = workbook.createCellStyle();
        // 设置这些样式
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中

        // 背景色
        style.setFillForegroundColor(HSSFColor.WHITE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setFillBackgroundColor(HSSFColor.WHITE.index);
        // 设置边框
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        // 自动换行
        style.setWrapText(true);

        // 生成一个字体
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setColor(HSSFColor.BLACK.index);
        font.setFontName("宋体");

        style.setFont(font);
        return style;

    }

    public static void main(String[] args) {
        List dataList = new ArrayList();
        Map titleMap1 = new LinkedHashMap();
        titleMap1.put("xing", "姓名");
        titleMap1.put("nian", "年龄");
        titleMap1.put("yu", "成绩");
        titleMap1.put("shu", "成绩");
        titleMap1.put("ying", "成绩");
        Map titleMap = new LinkedHashMap();
        titleMap.put("xing", "姓名");
        titleMap.put("nian", "年龄");
        titleMap.put("yu", "语文");
        titleMap.put("shu", "数学");
        titleMap.put("ying", "英语");
        Map dataMap = new LinkedHashMap();
        dataMap.put("xing", "姓名");
        dataMap.put("nian", "年龄");
        dataMap.put("yu", "80");
        dataMap.put("shu", "70");
        dataMap.put("ying", "60");
        Map dataMap1 = new LinkedHashMap();
        dataMap1.put("xing", "姓名");
        dataMap1.put("nian", "年龄");
        dataMap1.put("yu", "80");
        dataMap1.put("shu", "70");
        dataMap1.put("ying", "60");
        Map dataMap2 = new LinkedHashMap();
        dataMap2.put("xing", "姓名");
        dataMap2.put("nian", "年龄");
        dataMap2.put("yu", "80");
        dataMap2.put("shu", "70");
        dataMap2.put("ying", "60");
        dataList.add(titleMap1);
        dataList.add(titleMap);
        dataList.add(dataMap);
        dataList.add(dataMap1);
        dataList.add(dataMap2);
        int i = dataList.size() - 1;
        Integer[][] integers = {{0,0,0,i},{1,2,0,0},{1,2,1,1},{1,1,2,4}};
        String[] titleList = {"xing","nian","yu","shu","ying"};
        //excelExport("要导出的文件名称（文件名称和第一行名称相同）","学生成绩单",dataList,integers,titleList);
        //excelExport()
    }
}
