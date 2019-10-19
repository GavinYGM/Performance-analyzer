package main;

import java.io.File;
import java.util.List;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class CreateExcel {
	public static final int ColNum = 9;// ȥ��û������������
	public static final int classNum = 10;// �γ���
	public static void main(String args[]) {
		GetExcelInfo getExcelInfo = new GetExcelInfo();
		File file = new File("E:/1.xls");
		List<List<String>> list = getExcelInfo.readExcel(file);
		try {
			// ���ļ�
			WritableWorkbook book = Workbook.createWorkbook(new File("E:/test.xls"));
			// ������Ϊ��sheet1���Ĺ���������0��ʾ���ǵ�һҳ
			WritableSheet sheet = book.createSheet("sheet1", 0);
			// ��Label����Ĺ�������ָ����Ԫ��λ���ǵ�һ�е�һ��(0,0),��Ԫ������Ϊstring
			Label label = new Label(0, 0, "�༶");
			// ������õĵ�Ԫ����ӵ���������
			sheet.addCell(label);
			label = new Label(1, 0, "רҵ");
			sheet.addCell(label);
			label = new Label(2, 0, "ѧ��");
			sheet.addCell(label);
			label = new Label(3, 0, "ѧ��");
			sheet.addCell(label);
			label = new Label(4, 0, "����");
			sheet.addCell(label);
			label = new Label(5, 0, "ѧ��");
			sheet.addCell(label);
			label = new Label(6, 0, "ѧ������");
			sheet.addCell(label);
			label = new Label(7, 0, "ѧ��");
			sheet.addCell(label);
			label = new Label(8, 0, "�ɼ�");
			sheet.addCell(label);
			String[] strings = list.get(0).get(0).split(" ");
			String grade = strings[3];
			String major = strings[3].substring(0, 4);
			String term = strings[2];
			int sum=Integer.parseInt(list.get(list.size()-3).get(0));//��ȡ���һ��ѧ�������Ϊ������
			for (int i = 0; i < sum; i++) {// һ���ж��ٸ���
				List<String> cellist=list.get(i+7);
				for (int k = 0; k < ColNum; k++) {// һ���˹��ж���---����
					for (int j = classNum * i + 1; j < classNum * (i + 1) + 1; j++) {// һ�����ж����ſγ�--->��
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