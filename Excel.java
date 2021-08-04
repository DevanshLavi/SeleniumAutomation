package DataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public ArrayList<String> getData(String testcasename) throws IOException {
		
		ArrayList<String> a= new ArrayList<String>();
		FileInputStream fis = new FileInputStream("C://Users//ah82925//Documents//Datademo1.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		//get the number of sheer
		int sheets=wb.getNumberOfSheets();
		System.out.println(sheets);
		//to get the exact sheet from where we do fetch the data
		for(int i=0;i<sheets;i++) 
		{
			if(wb.getSheetName(i).equalsIgnoreCase("testdata"))
			{
				XSSFSheet sheet=wb.getSheetAt(i);//we have got our sheet where 
				//identify Testcases(first colum) column by scaning the entire first row
				Iterator<Row> row=sheet.iterator();
				Row firstrow=row.next();//it will set the iterator to first row--which is header
				Iterator<Cell> ce=firstrow.cellIterator();
				int k =0;
				int column=0;
				while(ce.hasNext()) {
					Cell value= ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("Testcases")) {
						
						column=k;
						System.out.println(k);
					}
					
					k++;
				}
				//so column is identified. lets scan entire Testcase column.....
				
				while(row.hasNext()) {
					 Row r=row.next();
					 if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcasename)) 
					 {
						 
						 Iterator<Cell> cv= r.cellIterator();
						 while(cv.hasNext()) {
							 
							a.add(cv.next().getStringCellValue());
						 }
						
					 }
				}
				
				
			}
			
		}
		return a;
		
	}
	
	public static void main(String[] args) throws IOException {
		
		Excel e=new Excel();
		ArrayList<String> data=e.getData("AddProfile");
		
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
		System.out.println(data.get(4));
		
	}

}
