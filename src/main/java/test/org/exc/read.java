package test.org.exc;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;


public class read
 {
	public static void main(String[] args) 
			throws IOException
	{
		
	
	File f=new File("C:\\Users\\RAGHU\\eclipse-workspace\\exc\\ex\\sample.xlsx");
FileInputStream st=new  FileInputStream(f);
Workbook w=new XSSFWorkbook(st);
Sheet s=w.getSheet("raghu");
for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
	Row r=s.getRow(i);
	for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
		Cell c=r.getCell(j);
		int type=c.getCellType();
if(type==1)
{
	String n=c.getStringCellValue();
	System.out.println(n);
}
if(type==0)
{
	if(DateUtil.isCellDateFormatted(c))
	{
		System.out.println("checking date");
		Date d=c.getDateCellValue();
SimpleDateFormat q=new SimpleDateFormat("DD-MMM-YYYY");
String format = q.format(c);
System.out.println(d);
	}
	else
	{
		double k=c.getNumericCellValue();
	long l=(long)k;
	String a=String.valueOf(l);
	System.out.println(a);
	System.out.println("end of programme");
	}}}}}}
