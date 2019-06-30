package test.org.exc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class rcrear {
public static void main(String[] args) throws IOException {
	File f=new File("C:\\Users\\RAGHU\\eclipse-workspace\\exc\\ex\\sample.xlsx");
	Workbook w= new XSSFWorkbook();
	Sheet s=w.createSheet("raghu");
	Row r=s.createRow(0);
	Cell c=r.createCell(0);
	c.setCellValue(10);
		Cell c1=r.createCell(1);
	c1.setCellValue("30-6-19");
	Cell c2=r.createCell(2);
	c2.setCellValue("hher");
	FileOutputStream st=new FileOutputStream(f);
	w.write(st);
System.out.println("done");}
}
