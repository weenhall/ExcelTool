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

    //����������workbook
    @Test
    public void createWorkBook() {
        Workbook wb = new HSSFWorkbook();
        //03-version
        try (OutputStream fileOut = new FileOutputStream("����.xls")) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //07-version
        Workbook wb1 = new XSSFWorkbook();
        try (OutputStream fileOu1 = new FileOutputStream("����.xlsx")) {
            wb1.write(fileOu1);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //����һ��������sheet
    @Test
    public void createSheet() {
        Workbook wb = new HSSFWorkbook();
        Sheet s1 = wb.createSheet("sheet1");
        Sheet s2 = wb.createSheet("sheet2");
        //ע�⣺����������Ʋ��ܳ���31���ַ�
//        WorkbookUtil.validateSheetName("sheetName must not exceed 31 characters");
        try (OutputStream fileOut = new FileOutputStream("����.xls")) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //������Ԫ��
    @Test
    public void createCells() {
        Workbook wb = new HSSFWorkbook();
        CreationHelper helper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("sheet1");
        //������һ����
        Row row = sheet.createRow(0);
        //������һ��
        Cell cell = row.createCell(0);
        cell.setCellValue(100);

        //һ���Դ���
        row.createCell(1).setCellValue(101);
        row.createCell(2).setCellValue(helper.createRichTextString("����һ�θ��ı�"));
        row.createCell(3).setCellValue(102);
        //�����������͵�Ԫ��
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(helper.createDataFormat().getFormat("yyyy-mm-dd"));
        Cell dateCell = row.createCell(4);
        dateCell.setCellValue(new Date());
        dateCell.setCellStyle(cellStyle);
        //������������
        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue(100.1);
        row2.createCell(1).setCellValue(Calendar.getInstance());
        row2.createCell(2).setCellValue("�ַ���");
        row2.createCell(3).setCellValue(true);
        try (OutputStream fileOut = new FileOutputStream("����.xls")) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    //��һ��workbook����ͨ��File��InputStream�ķ�ʽ,ʹ��File���������InputStream���ĵ��ڴ�ҪС
    @Test
    public void usingWorkBookFactory() {
        //using a file
        try {
            Workbook wb = WorkbookFactory.create(new File("����.xls"));
            Sheet sheet = wb.getSheetAt(0);
        } catch (Exception e) {
            e.printStackTrace();
        }
        //using an InputStream,needs more memory
        try {
            Workbook workbook = WorkbookFactory.create(new FileInputStream("����.xlsx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void createCell(Workbook wb,Row row,int column,short h,short v){
        Cell cell=row.createCell(column);
        cell.setCellValue("���뷽ʽ");
        CellStyle style=wb.createCellStyle();
        style.setAlignment(h);
        style.setVerticalAlignment(v);
        cell.setCellStyle(style);
    }

    //��Ԫ����뷽ʽ
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
        try(OutputStream fileOut=new FileOutputStream("���뷽ʽ����.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    //��Ԫ��߽�border
    @Test
    public void testCellBorders() throws IOException {
        Workbook wb=new XSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");
        Row row=sheet.createRow(1);
        row.setHeightInPoints(30);
        Cell cell=row.createCell(1);
        cell.setCellValue(100);

        CellStyle style=wb.createCellStyle();
        //��
        style.setBorderBottom(CellStyle.BORDER_MEDIUM_DASHED);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        //��
        style.setBorderLeft(CellStyle.BORDER_DASH_DOT_DOT);
        style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        //��
        style.setBorderRight(CellStyle.BORDER_DASH_DOT_DOT);
        style.setRightBorderColor(IndexedColors.BLUE.getIndex());
        //��
        style.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(style);
        try(OutputStream fileOut=new FileOutputStream("border��������.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //�������õ�Ԫ����ɫ
    @Test
    public void cellsFillAndColors() throws IOException {
        Workbook wb=new XSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");
        Row row=sheet.createRow(1);

        CellStyle style=wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.PINK.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        Cell cell1=row.createCell(1);
        cell1.setCellValue("������ɫΪAQUA");
        cell1.setCellStyle(style);

        style=wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cell1 = row.createCell(2);
        cell1.setCellValue("������ɫΪORANGE");
        cell1.setCellStyle(style);
        try(OutputStream fileOut=new FileOutputStream("���õ�Ԫ����ɫ����.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //��Ԫ��ϲ�
    @Test
    public void mergeCells() throws Exception{
        Workbook wb=new XSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");
        Row row=sheet.createRow(1);
        CellStyle style=wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);

        Cell cell=row.createCell(1);
        cell.setCellValue("���Ե�Ԫ��ϲ�");
        cell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(
                1,
                1,
                1,
                3
        ));//�ϲ���һ�е�1-3��
        try(OutputStream fileOut=new FileOutputStream("��Ԫ��ϲ�����.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //��������,ע���ظ�����Font�����Ǹ�ÿ��Cell��newһ��Font,Font������32767
    @Test
    public void setCellFonts() throws Exception {
        Workbook wb=new XSSFWorkbook();
        Sheet sheet=wb.createSheet("sheet1");
        Row row=sheet.createRow(1);

        //��������
        Font font=wb.createFont();
        font.setFontHeightInPoints((short) 30);
        font.setFontName("����");
        font.setItalic(true);
        font.setStrikeout(true);//ɾ����
        //������ʽ
        CellStyle style=wb.createCellStyle();
        style.setFont(font);

        Cell cell=row.createCell(1);
        cell.setCellValue("����������ʽ");
        cell.setCellStyle(style);
        try(OutputStream fileOut=new FileOutputStream("������������.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //���ݸ�ʽ��
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

        try(OutputStream fileOut=new FileOutputStream("���ݸ�ʽ������.xlsx")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //�����������Ϊһҳ
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
        try(OutputStream fileOut=new FileOutputStream("�����������Ϊһҳ.xls")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //��ֺͶ��ᴰ��
    @Test
    public void splitsAndFreezePanes() throws Exception{
        Workbook wb=new HSSFWorkbook();
        Sheet sheet1=wb.createSheet("shee1");
        Sheet sheet2=wb.createSheet("shee2");
        Sheet sheet3=wb.createSheet("shee3");
        Sheet sheet4=wb.createSheet("shee4");

        //����һ��
        sheet1.createFreezePane(0,1,0,1);
        //����һ��
        sheet2.createFreezePane(1,0,1,0);
        //����ǰ2�к�2��
        sheet3.createFreezePane(2,2);
        sheet4.createSplitPane(2000,2000,0,0,Sheet.PANE_LOWER_LEFT);
        try(OutputStream fileOut=new FileOutputStream("��ֺͶ��ᴰ��.xls")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }

    //����ͷ����β��
    @Test
    public void createHeaderAndFooter() throws Exception{
        Workbook wb=new HSSFWorkbook();
        Sheet sheet=wb.createSheet("shee1");

        Header header=sheet.getHeader();
        header.setCenter("�м�");
        header.setLeft("���");
        header.setRight(HSSFHeader.font("Stencil-Normal", "Italic") +
                HSSFHeader.fontSize((short) 16) + "Right w/ Stencil-Normal Italic font and size 16");
        try(OutputStream fileOut=new FileOutputStream("����ͷ����β��.xls")){
            wb.write(fileOut);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            wb.close();
        }
    }
}