package main;

import java.io.File;
import java.util.List;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class CreateExcel {
	public static final int ColNum = 9;// 去掉没有姓名的行数
	public static final int classNum = 10;// 课程数
	public static void main(String args[]) {
		GetExcelInfo getExcelInfo = new GetExcelInfo();
		File file = new File("E:/1.xls");
		List<List<String>> list = getExcelInfo.readExcel(file);
		try {
			// 打开文件
			WritableWorkbook book = Workbook.createWorkbook(new File("E:/test.xls"));
			// 生成名为“sheet1”的工作表，参数0表示这是第一页
			WritableSheet sheet = book.createSheet("sheet1", 0);
			// 在Label对象的构造子中指名单元格位置是第一列第一行(0,0),单元格内容为string
			Label label = new Label(0, 0, "班级");
			// 将定义好的单元格添加到工作表中
			sheet.addCell(label);
			label = new Label(1, 0, "专业");
			sheet.addCell(label);
			label = new Label(2, 0, "学期");
			sheet.addCell(label);
			label = new Label(3, 0, "学号");
			sheet.addCell(label);
			label = new Label(4, 0, "姓名");
			sheet.addCell(label);
			label = new Label(5, 0, "学科");
			sheet.addCell(label);
			label = new Label(6, 0, "学科类型");
			sheet.addCell(label);
			label = new Label(7, 0, "学分");
			sheet.addCell(label);
			label = new Label(8, 0, "成绩");
			sheet.addCell(label);
			String[] strings = list.get(0).get(0).split(" ");
			String grade = strings[3];
			String major = strings[3].substring(0, 4);
			String term = strings[2];
			int sum=Integer.parseInt(list.get(list.size()-3).get(0));//获取最后一个学生的序号为总人数
			for (int i = 0; i < sum; i++) {// 一共有多少个人
				List<String> cellist=list.get(i+7);
				for (int k = 0; k < ColNum; k++) {// 一个人共有多少---》列
					for (int j = classNum * i + 1; j < classNum * (i + 1) + 1; j++) {// 一个人有多少门课程--->行
						if (k==0) {
							label = new Label(k, j, grade);
							sheet.addCell(label);
						}else if (k==1) {
							label = new Label(k, j, major);
							sheet.addCell(label);
						}else if (k==2) {
							label = new Label(k, j, term);
							sheet.addCell(label);
						}else if (k==3) {
							label = new Label(k, j, cellist.get(1));
							sheet.addCell(label);
						}else if (k==4) {
							label = new Label(k, j, cellist.get(2));
							sheet.addCell(label);
						}else if (k==5) {
							if ((j%classNum)==0) {
								label = new Label(k, j, list.get(1).get(classNum+3));
								sheet.addCell(label);
							}else {
								label = new Label(k, j, list.get(1).get((j%classNum)+3));
								sheet.addCell(label);
							}
						}else if (k==6) {
							if ((j%classNum)==0) {
								label = new Label(k, j, list.get(3).get(classNum+3));
								sheet.addCell(label);
							}else {
								label = new Label(k, j, list.get(3).get((j%classNum)+3));
								sheet.addCell(label);
							}
						}else if (k==7) {
							if ((j%classNum)==0) {
								label = new Label(k, j, list.get(4).get(classNum+3));
								sheet.addCell(label);
							}else {
								label = new Label(k, j, list.get(4).get((j%classNum)+3));
								sheet.addCell(label);
							}
						}else if (k==8) {
							if ((j%classNum)==0) {
								label = new Label(k, j, cellist.get(classNum+3));
								sheet.addCell(label);
							}else {
								label = new Label(k, j, cellist.get(j%classNum+3));
								sheet.addCell(label);
							}
							
						}
						
					}
				}
				
			}
			book.write();
			book.close();
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}