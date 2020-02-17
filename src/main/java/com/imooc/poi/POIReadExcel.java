package com.imooc.poi;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class POIReadExcel {
	public static void main(String[] args) {
        List<Person> person=new ArrayList<Person>();
		//定义需要解析的Excel文件
		File file=new File("d:/poi_test.xls");
		try {
			//创建Excel,读取文件内容
			HSSFWorkbook workbook=new HSSFWorkbook(FileUtils.openInputStream(file));
			//获取第一个工作表workbook,workbook.getSheet("Sheet0"),也可以采用下面的读取工作表
			//HSSFSheet sheet=workbook.getSheet("Sheet0");
			//读取默认第一个工作表sheet
			HSSFSheet sheet=workbook.getSheetAt(0);
			//获取sheet中最后一行行号
			int lastRowNum=sheet.getLastRowNum();
			for (int i = 1; i <= lastRowNum; i++) {//遍历行
				HSSFRow row=sheet.getRow(i);
				//获取当前行最后单元格列号
				Person pp=new Person();
				pp.setId(row.getCell(0).getStringCellValue());
				pp.setName(row.getCell(1).getStringCellValue());
				pp.setSex(row.getCell(2).getStringCellValue());
				person.add(pp);
			}
			System.out.println(person.get(1));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
