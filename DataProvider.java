package DataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataProvider {

	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
	
		DataFormatter formatter=new DataFormatter();
	
	
		FileInputStream fis = new FileInputStream("C://Users//ah82925//Documents//Datademo2.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheets=wb.getSheetAt(1);//got the sheet
		int rowCount =sheets.getPhysicalNumberOfRows();//get all row in the sheet# of
		//below 2 lines are use to get the # of column
		XSSFRow row=sheets.getRow(0);
		int colCount=row.getLastCellNum();
		//create the array for total row and colum--we dont need Header so taking rowcount-1
		Object data[][] =new Object[rowCount-1][colCount];
		for (int i=0;i<=rowCount-1;i++) {
			row=sheets.getRow(i+1);
			for(int j=0;j<=colCount-1;j++)
			{
				
			//XSSFCell cell=row.getCell(j);
			
			//data[i][j]=formatter.formatCellValue(cell);
			System.out.println(row.getCell(j));
			}
			
		}
			
		
	
	
	
	}
}
