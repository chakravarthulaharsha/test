package write_to_excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class exe_code {
	static String filename="H:\\cognizant\\meeting files\\KWDFW_2.xlsx";
	static String sheetname="KEYWORD";
	public static String write_excel(int row,int col,String data) {
		
		String s = null;
		
		try {
		File f= new File(filename);
		FileInputStream fis = new FileInputStream(f);
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sh = wb.getSheet(sheetname);
		XSSFRow r = sh.getRow(row);
		XSSFCell c = r.getCell(col);
		
		
	 c.setCellValue(data);
		
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
	}catch(IOException e) {
		e.printStackTrace();
	}
		return data;
	}
	
	public static void main(String[] args) {	
		write_excel(1, 0,"Nameste");
	
	}

}
