package assmnt_1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class exe_code {
	static String filename="H:\\cognizant\\meeting files\\KWDFW_FILE.xlsx";
	static String sheetname="KEYWORD";
	
	public static String read_excel(int row ,int col) {
		String s=null;
		try {
			File f=new File(filename);
			FileInputStream fis=new FileInputStream(f);
			XSSFWorkbook wb=new XSSFWorkbook(fis);
			XSSFSheet sh = wb.getSheet(sheetname);
			XSSFRow r=sh.getRow(row);
			XSSFCell c=r.getCell(col);
			s=c.getStringCellValue();
		}catch(IOException e) {
			e.printStackTrace();
		}
		return (s);
		}
	
	public static void main(String[] args) {
		String kw, loc, td,ab;
		for(int r=1;r<=7;r++) {
			kw=read_excel(2,3);
			loc=read_excel(1,4);
			td=read_excel(3,5);
			System.out.println(kw+"\n"+loc+"\n"+td);
			}
		}
	}