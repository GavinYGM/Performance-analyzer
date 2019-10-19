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
        	System.out.println(list2.get(1)+"  答："+list2.get(2));
        	System.out.println();
		}
    }  
    // 去读Excel的方法readExcel，该方法的入口参数为一个File对象  
    public List<List<String>> readExcel(File file) { 
    	List<List<String>> list=new ArrayList<>();
        try {  
            // 创建输入流，读取Excel  
            InputStream is = new FileInputStream(file.getAbsolutePath());  
            // jxl提供的Workbook类  
            Workbook wb = Workbook.getWorkbook(is);  
            // Excel的页签数量  
            int sheet_size = wb.getNumberOfSheets();  
            for (int index = 0; index < sheet_size; index++) {  
                // 每个页签创建一个Sheet对象  
                Sheet sheet = wb.getSheet(index);  
                // sheet.getRows()返回该页的总行数  
                for (int i = 0; i < sheet.getRows(); i++) { 
                	List<String> cell=new ArrayList<String>();
                    // sheet.getColumns()返回该页的总列数  
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