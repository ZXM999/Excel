package com.example.sbtest.service;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Service;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelTestService {
    public void makeExcel(){
        // 声明一个工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 生成一个表格
        HSSFSheet sheet = workbook.createSheet();
        //设置默认行宽 表示2个字符的高度
        sheet.setDefaultRowHeight((short) (2 * 256));
        //设置默认列宽
        sheet.setDefaultColumnWidth(20);
        //合并单元格
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 6));
        //设置样式(黑体12)
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        // 水平居中
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //垂直居中
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        //设置字体
        HSSFFont font = workbook.createFont();
        font.setFontName("黑体");
        font.setFontHeightInPoints((short)12);
        cellStyle.setFont(font);
        //设置背景颜色;
        cellStyle.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);
        //solid 填充  foreground  前景色 **必要
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        //设置样式(华文行楷12)
        HSSFCellStyle cellStyle2 = workbook.createCellStyle();
        // 水平居中
        cellStyle2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //垂直居中
        cellStyle2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        //设置字体
        HSSFFont font2 = workbook.createFont();
        font2.setFontName("华文行楷");
        font2.setFontHeightInPoints((short)12);
        cellStyle2.setFont(font2);
        //自定义设置背景颜色（把原来的红色改成想要的颜色）
        HSSFPalette palette = workbook.getCustomPalette();
        palette.setColorAtIndex(HSSFColor.RED.index, (byte) 102, (byte) 102, (byte) 153);
        cellStyle2.setFillForegroundColor(HSSFColor.RED.index);
        //solid 填充  foreground  前景色 **必要
        cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        //产生表格标题行  第0行
        HSSFRow row0 = sheet.createRow(0);
        HSSFCell cell00=row0.createCell(0);
        cell00.setCellStyle(cellStyle);
        cell00.setCellValue("订单编号");
        //产生表格标题行  第1行
        HSSFRow row1 = sheet.createRow(1);
        HSSFCell cell10=row1.createCell(0);
        cell10.setCellStyle(cellStyle);
        cell10.setCellValue("下单客户");
        //产生表格标题行  第2行
        HSSFRow row2 = sheet.createRow(2);
        HSSFCell cell20=row2.createCell(0);
        cell20.setCellStyle(cellStyle);
        cell20.setCellValue("联系方式");
        //产生表格标题行  第3行
        HSSFRow row3 = sheet.createRow(3);
        HSSFCell cell30=row3.createCell(0);
        cell30.setCellStyle(cellStyle);
        cell30.setCellValue("下单时间");
        //产生表格标题行  第4行
        HSSFRow row4 = sheet.createRow(4);
        HSSFCell cell40=row4.createCell(0);
        cell40.setCellStyle(cellStyle);
        cell40.setCellValue("收货信息");
        //产生表格标题行  第5行商品信息
        HSSFRow row5 = sheet.createRow(5);
        HSSFCell cell50=row5.createCell(0);
        HSSFCell cell51=row5.createCell(1);
        HSSFCell cell52=row5.createCell(2);
        HSSFCell cell53=row5.createCell(3);
        HSSFCell cell54=row5.createCell(4);
        HSSFCell cell55=row5.createCell(5);
        HSSFCell cell56=row5.createCell(6);
        HSSFCell cell57=row5.createCell(7);
        HSSFCell cell58=row5.createCell(8);
        HSSFCell cell59=row5.createCell(9);
        cell50.setCellStyle(cellStyle2);
        cell51.setCellStyle(cellStyle2);
        cell52.setCellStyle(cellStyle2);
        cell53.setCellStyle(cellStyle2);
        cell54.setCellStyle(cellStyle2);
        cell55.setCellStyle(cellStyle2);
        cell56.setCellStyle(cellStyle2);
        cell57.setCellStyle(cellStyle2);
        cell58.setCellStyle(cellStyle2);
        cell59.setCellStyle(cellStyle2);
        cell50.setCellValue("商品名称");
        cell51.setCellValue("商品编号");
        cell52.setCellValue("工厂号");
        cell53.setCellValue("oe号");
        cell54.setCellValue("规格型号");
        cell55.setCellValue("品牌");
        cell56.setCellValue("品质");
        cell57.setCellValue("数量");
        cell58.setCellValue("单价（￥）");
        cell59.setCellValue("小计（￥）");
        //测试数据
        List<Map> list = new ArrayList<>();
        Map map1 = new HashMap();
        map1.put("sl",2);
        map1.put("xj",1);
        Map map2 = new HashMap();
        map2.put("sl",2);
        map2.put("xj",1);
        list.add(map1);
        list.add(map2);
        int index = 6;
        //设置行
        HSSFRow rows = null;
        for(Map m:list) {
            rows = sheet.createRow(index++);
            HSSFCell c1=rows.createCell(7);
            c1.setCellValue((int) m.get("sl"));
            HSSFCell c2=rows.createCell(9);
            c2.setCellValue((int) m.get("xj"));
        }
        HSSFRow rowLast = sheet.createRow(index);
        HSSFCell c1=rowLast.createCell(7);
        c1.setCellValue("合计");
        HSSFCell c2=rowLast.createCell(9);
        c2.setCellValue("合计");
        //输出
        FileOutputStream out = null;
        try {
            out = new FileOutputStream("D:\\aaa.xls");
            workbook.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
