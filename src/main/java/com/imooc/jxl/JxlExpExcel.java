package com.imooc.jxl;

import java.io.File;
import java.text.Collator;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.time.DateUtils;

import jxl.Cell;
import jxl.Range;
import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class JxlExpExcel {

	/*
	 * JXL����Excel�ļ�
	 */
	public static void main(String[] args) {
		String[] title = { "���", "����", "����ʱ��", "��������", "��ע" };
		// ����Excel�ļ�
		File file = new File("D:/jxl_test.xls");
		WritableWorkbook workbook = null;
		try {
			file.createNewFile();
			// ����������
			// OutputStream os = response.getOutputStream();
			// WritableWorkbook workbook = Workbook.createWorkbook(os);
			workbook = Workbook.createWorkbook(file);
			WritableSheet sheet = workbook.createSheet("sheet1", 0);
			// ���ñ��ָ���е��п�
			sheet.setColumnView(0, 10);
			sheet.setColumnView(1, 15);
			sheet.setColumnView(2, 30);
			sheet.setColumnView(3, 50);
			sheet.setColumnView(4, 50);

			// �����������
			WritableFont titleFont = new WritableFont(WritableFont.createFont("΢���ź�"), 15, WritableFont.NO_BOLD);
			WritableFont contentFont = new WritableFont(WritableFont.createFont("���� _GB2312"), 12,
					WritableFont.NO_BOLD);
			WritableCellFormat titleFormat = new WritableCellFormat(titleFont);
			WritableCellFormat contentFormat = new WritableCellFormat(contentFont);

			contentFormat.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLACK);
			// ���ø�ʽ���ж���
			titleFormat.setAlignment(jxl.format.Alignment.CENTRE);
			titleFormat.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
			contentFormat.setAlignment(jxl.format.Alignment.CENTRE);
			contentFormat.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
	

			Label label = null;
			sheet.mergeCells(0, 0, 4, 0);
			sheet.addCell(new Label(0, 0, "xxͳ�Ʊ�", titleFormat));
			sheet.mergeCells(1, 1, 2, 1);
			sheet.addCell(new Label(1, 1, "���ţ�xx��"));

			sheet.mergeCells(3, 1, 4, 1);
			sheet.addCell(new Label(3, 1, "�ʱ�䣺xxxx��xx��xx��"));

			// ����һ����������
			for (int i = 0; i < title.length; i++) {
				label = new Label(i, 2, title[i], contentFormat);
				sheet.addCell(label);
			}

			List<Map<String, String>> list = initExportDatas();
			Map<String,String> pMap=new HashMap<String, String>();
			// ��ÿ���������
			Map<String, String> tMap = null;
			for (int j = 3; j < list.size() + 3; j++) {
				tMap = list.get(j - 3);
				String personName=tMap.get("����");
				if(pMap.containsKey(personName)){
					pMap.put(personName, pMap.get(personName).substring(0,1)+"-"+j);
				}else{
					pMap.put(personName,String.valueOf(j));
				}
				label = new Label(0, j, String.valueOf(j-2), contentFormat);
				sheet.addCell(label);
				label = new Label(1, j, tMap.get("����"), contentFormat);
				sheet.addCell(label);
				label = new Label(2, j, tMap.get("����ʱ��"), contentFormat);
				sheet.addCell(label);
				label = new Label(3, j, tMap.get("��������"), contentFormat);
				sheet.addCell(label);
				label = new Label(4, j, tMap.get("��ע"), contentFormat);
				sheet.addCell(label);
			}
			int sufixRow = list.size() + 4;
			sheet.mergeCells(0, sufixRow, 1, sufixRow);
			sheet.addCell(new Label(0, sufixRow, "�Ʊ���:"));
			sheet.mergeCells(3, sufixRow, 4, sufixRow);
			sheet.addCell(new Label(3, sufixRow, "���Ÿ�����ǩ�֣�"));

			for(String key:pMap.keySet()){
				String arr[]=pMap.get(key).split("-");
				if(arr.length==2){
					int starter=Integer.parseInt(arr[0]);
					int ender=Integer.parseInt(arr[1]);
					sheet.mergeCells(0, starter, 0, ender);//�ϲ����
					sheet.mergeCells(1, starter, 1, ender);//�ϲ�����
				
					//��������
					for(int k=ender+1;k<list.size()+3;k++){
						Label currentLabel = (Label) sheet.getWritableCell(0, k);
						int currentOrder=Integer.parseInt(currentLabel.getString());
						currentLabel.setString(String.valueOf(--currentOrder));
					}
				}
			}
			// д������
			workbook.write();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				workbook.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}

	public static List<Map<String, String>> initExportDatas() {
		List<Map<String, String>> list = new ArrayList<Map<String, String>>();
		Map<String, String> map1 = new HashMap<>();
		map1.put("���", "1");
		map1.put("����", "����");
		map1.put("����ʱ��", "2018.03.01-2018.03.03");
		map1.put("��������", "aa");
		map1.put("��ע", "bb");
		list.add(map1);
		Map<String, String> map = null;
		for (int i = 0; i < 3; i++) {
			map = new HashMap<String, String>();
			map.put("���", String.valueOf(i+2));
			map.put("����", i + "zz");
			map.put("����ʱ��", "2018.03.01-2018.03.08");
			map.put("��������", String.valueOf(i));
			map.put("��ע", String.valueOf(i));
			list.add(map);
		}
		Map<String, String> map2 = new HashMap<String, String>();
		map2.put("���", "5");
		map2.put("����", "0zz");
		map2.put("����ʱ��", "2018.03.9-2018.03.10");
		map2.put("��������", "aaaaaa");
		map2.put("��ע", "bbbbb");
		list.add(map2);
		
		Map<String, String> map3 = new HashMap<String, String>();
		map3.put("���", "6");
		map3.put("����", "����");
		map3.put("����ʱ��", "2018.03.10-2018.03.14");
		map3.put("��������", "asfasf");
		map3.put("��ע", "bbbbasdfadb");
//		list.add(map3);

		Collections.sort(list, new Comparator<Map<String, String>>() {

			@Override
			public int compare(Map<String, String> o1, Map<String, String> o2) {
				String name1 = o1.get("����");
				String name2 = o2.get("����");
				Collator cmp = Collator.getInstance(java.util.Locale.CHINA);
				if (cmp.compare(name1, name2) > 0) {
					return 1;
				} else if (cmp.compare(name1, name2) < 0) {
					return -1;
				}
				return 0;
			}

		});
		return list;
	}

}
