package com.excel.util;


import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import java.io.*;
import java.util.Calendar;
import java.util.Date;

/**
 * Quick Guide for manipulate MS-EXCEL
 * More Example see https://poi.apache.org/components/spreadsheet/quick-guide.html
 */
public class POIExcelDemo {

    //创建工作簿workbook
    @Test
    public void createWorkBook() {
        Workbook wb = new HSSFWorkbook();
        //03-version
        try (OutputStream fileOut = new FileOutputStream("测试.xls")) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //07-version
        Workbook wb1 = new XSSFWorkbook();
        try (OutputStream fileOu1 = new FileOutputStream("测试.xlsx")) {
            wb1.write(fileOu1);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //创建一个工作表sheet
    @Test
    public void createSheet() {
        Workbook wb = new HSSFWorkbook();
        Sheet s1 = wb.createSheet("sheet1");
        Sheet s2 = wb.createSheet("sheet2");
        //注意：工作表的名称不能超过31个字符
//        WorkbookUtil.validateSheetName("sheetName must not exceed 31 characters");
        try (OutputStream fileOut = new FileOutputStream("测试.xls")) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //创建单元格
    @Test
    public void createCells() {
        Workbook wb = new HSSFWorkbook();
        CreationHelper helper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("sheet1");
        //创建第一行行
        Row row = sheet.createRow(0);
        //创建第一列
        Cell cell = row.createCell(0);
        cell.setCellValue(100);

        //一次性创建
        row.createCell(1).setCellValue(101);
        row.createCell(2).setCellValue(helper.createRichTextString("这是一段富文本"));
        row.createCell(3).setCellValue(102);
        //创建日期类型单元格
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(helper.createDataFormat().getFormat("yyyy-mm-dd"));
        Cell dateCell = row.createCell(4);
        dateCell.setCellValue(new Date());
        dateCell.setCellStyle(cellStyle);
        //创建其他类型
        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue(100.1);
        row2.createCell(1).setCellValue(Calendar.getInstance());
        row2.createCell(2).setCellValue("字符串");
        row2.createCell(3).setCellValue(true);
        try (OutputStream fileOut = new FileOutputStream("测试.xls")) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    //打开一个workbook可以通过File或InputStream的方式,使用File对象相比于InputStream消耗的内存要小
    @Test
    public void usingWorkBookFactory() {
        //using a file
        try {
            Workbook wb = WorkbookFactory.create(new File("测试.xls"));
            Sheet sheet = wb.getSheetAt(0);
        } catch (Exception e) {
            e.printStackTrace();
        }
        //using an InputStream,needs more memory
        try {
            Workbook workbook = WorkbookFactory.create(new FileInputStream("测试.xlsx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void createCell(Workbook wb,Row row,int column,short h,short v){
        Cell cell=row.createCell(column);
        cell.setCellValue("对齐方式");
        CellStyle style=wb.createCellStyle();
        style.setAlignment(h);
        style.setVerticalAlignment(v);
        cell.setCellStyle(style);
    }

    //单元格对齐方式
    @Test
    public void testCellAlignment(){
        Workbook wb=new XSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");
        Row row=sheet.createRow(2);
        row.setHeightInPoints(30);

        createCell(wb,row,0,CellStyle.ALIGN_CENTER,CellStyle.VERTICAL_BOTTOM);
        createCell(wb,row,1,CellStyle.ALIGN_CENTER_SELECTION,CellStyle.VERTICAL_BOTTOM);
        createCell(wb,row,2,CellStyle.ALIGN_FILL,CellStyle.VERTICAL_CENTER);
        createCell(wb,row,3,CellStyle.ALIGN_GENERAL,CellStyle.VERTICAL_CENTER);
        createCell(wb,row,4,CellStyle.ALIGN_JUSTIFY,CellStyle.VERTICAL_JUSTIFY);
        createCell(wb,row,5,CellStyle.ALIGN_LEFT,CellStyle.VERTICAL_TOP);
        createCell(wb,row,6,CellStyle.ALIGN_RIGHT,CellStyle.VERTICAL_TOP);
        try(OutputStream fileOut=new FileOutputStream("对齐方式例子.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    //单元格边界border
    @Test
    public void testCellBorders() throws IOException {
        Workbook wb=new XSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");
        Row row=sheet.createRow(1);
        row.setHeightInPoints(30);
        Cell cell=row.createCell(1);
        cell.setCellValue(100);

        CellStyle style=wb.createCellStyle();
        //下
        style.setBorderBottom(CellStyle.BORDER_MEDIUM_DASHED);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        //左
        style.setBorderLeft(CellStyle.BORDER_DASH_DOT_DOT);
        style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        //右
        style.setBorderRight(CellStyle.BORDER_DASH_DOT_DOT);
        style.setRightBorderColor(IndexedColors.BLUE.getIndex());
        //上
        style.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(style);
        try(OutputStream fileOut=new FileOutputStream("border设置例子.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //填充和设置单元格颜色
    @Test
    public void cellsFillAndColors() throws IOException {
        Workbook wb=new XSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");
        Row row=sheet.createRow(1);

        CellStyle style=wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.PINK.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        Cell cell1=row.createCell(1);
        cell1.setCellValue("背景颜色为AQUA");
        cell1.setCellStyle(style);

        style=wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cell1 = row.createCell(2);
        cell1.setCellValue("背景颜色为ORANGE");
        cell1.setCellStyle(style);
        try(OutputStream fileOut=new FileOutputStream("设置单元格颜色例子.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //单元格合并
    @Test
    public void mergeCells() throws Exception{
        Workbook wb=new XSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");
        Row row=sheet.createRow(1);
        CellStyle style=wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);

        Cell cell=row.createCell(1);
        cell.setCellValue("测试单元格合并");
        cell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(
                1,
                1,
                1,
                3
        ));//合并第一行的1-3列
        try(OutputStream fileOut=new FileOutputStream("单元格合并例子.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //字体设置,注意重复利用Font而不是给每个Cell都new一个Font,Font最多可以32767
    @Test
    public void setCellFonts() throws Exception {
        Workbook wb=new XSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");
        Row row=sheet.createRow(1);

        //创建字体
        Font font=wb.createFont();
        font.setFontHeightInPoints((short) 30);
        font.setFontName("宋体");
        font.setItalic(true);
        font.setStrikeout(true);//删除线
        //创建样式
        CellStyle style=wb.createCellStyle();
        style.setFont(font);

        Cell cell=row.createCell(1);
        cell.setCellValue("设置字体样式");
        cell.setCellStyle(style);
        try(OutputStream fileOut=new FileOutputStream("字体设置例子.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //数据格式化
    @Test
    public void dataFormats() throws Exception{
        Workbook wb=new XSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");

        CellStyle style;
        DataFormat format=wb.createDataFormat();
        Row row;
        Cell cell;
        int rowNum=0;
        int colNum=0;

        row=sheet.createRow(rowNum++);
        cell=row.createCell(colNum);
        cell.setCellValue(123932.23);
        style=wb.createCellStyle();
        style.setDataFormat(format.getFormat("0.00"));
        cell.setCellStyle(style);

        row=sheet.createRow(rowNum++);
        cell=row.createCell(colNum);
        cell.setCellValue(123932.23);
        style=wb.createCellStyle();
        style.setDataFormat(format.getFormat("#,##0.00"));
        cell.setCellStyle(style);

        try(OutputStream fileOut=new FileOutputStream("数据格式化例子.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //将工作表调整为一页
    @Test
    public void fitSheetToOnePage() throws Exception{
        Workbook wb=new HSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");

        PrintSetup ps=sheet.getPrintSetup();
        sheet.setAutobreaks(true);
        ps.setFitHeight((short)1);
        ps.setFitWidth((short)1);
        Row row;
        Cell cell;
        for(int i=0;i<1000;i++){
            row=sheet.createRow(i);
            for(int j=0;j<5;j++){
                cell=row.createCell(j);
                cell.setCellValue(i+"-"+j);
            }
        }
        try(OutputStream fileOut=new FileOutputStream("将工作表调整为一页.xls")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //拆分和冻结窗格
    @Test
    public void splitsAndFreezePanes() throws Exception{
        Workbook wb=new HSSFWorkbook();
        Sheet sheet1=wb.createSheet("shee1");
        Sheet sheet2=wb.createSheet("shee2");
        Sheet sheet3=wb.createSheet("shee3");
        Sheet sheet4=wb.createSheet("shee4");

        //冻结一行
        sheet1.createFreezePane(0,1,0,1);
        //冻结一列
        sheet2.createFreezePane(1,0,1,0);
        //冻结前2行和2列
        sheet3.createFreezePane(2,2);
        sheet4.createSplitPane(2000,2000,0,0,Sheet.PANE_LOWER_LEFT);
        try(OutputStream fileOut=new FileOutputStream("拆分和冻结窗格.xls")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //创建头部和尾部
    @Test
    public void createHeaderAndFooter() throws Exception{
        Workbook wb=new HSSFWorkbook();
        Sheet sheet=wb.createSheet("shee1");

        Header header=sheet.getHeader();
        header.setCenter("中间");
        header.setLeft("左边");
        header.setRight(HSSFHeader.font("Stencil-Normal", "Italic") +
                HSSFHeader.fontSize((short) 16) + "Right w/ Stencil-Normal Italic font and size 16");
        try(OutputStream fileOut=new FileOutputStream("创建头部和尾部.xls")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }
}