package main;

import java.io.File;  
import java.io.FileInputStream;  
import java.io.FileNotFoundException;  
import java.io.IOException;  
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Sheet;  
import jxl.Workbook;  
import jxl.read.biff.BiffException;  
public class GetExcelInfo {  
    public static void main(String[] args) {  
        GetExcelInfo obj = new GetExcelInfo();  
        File file = new File("H:/1.xls");  
        List<List<String>> list=obj.readExcel(file);
        for (int i = 0; i < list.size(); i++) {
			List<String> list2=list.get(i);
        	System.out.println(list2.get(1)+"  ��"+list2.get(2));
        	System.out.println();
		}
    }  
    // ȥ��Excel�ķ���readExcel���÷�������ڲ���Ϊһ��File����  
    public List<List<String>> readExcel(File file) { 
    	List<List<String>> list=new ArrayList<>();
        try {  
            // ��������������ȡExcel  
            InputStream is = new FileInputStream(file.getAbsolutePath());  
            // jxl�ṩ��Workbook��  
            Workbook wb = Workbook.getWorkbook(is);  
            // Excel��ҳǩ����  
            int sheet_size = wb.getNumberOfSheets();  
            for (int index = 0; index < sheet_size; index++) {  
                // ÿ��ҳǩ����һ��Sheet����  
                Sheet sheet = wb.getSheet(index);  
                // sheet.getRows()���ظ�ҳ��������  
                for (int i = 0; i < sheet.getRows(); i++) { 
                	List<String> cell=new ArrayList<String>();
                    // sheet.getColumns()���ظ�ҳ��������  
                    for (int j = 0; j < sheet.getColumns(); j++) {  
                        String cellinfo = sheet.getCell(j, i).getContents(); 
                        cell.add(cellinfo);
                        //System.out.print(cellinfo);  
                    }  
                    list.add(cell);
                }  
            }  
        } catch (FileNotFoundException e) {  
            e.printStackTrace();  
        } catch (BiffException e) {  
            e.printStackTrace();  
        } catch (IOException e) {  
            e.printStackTrace();  
        }  
        return list;
    }  
}  