package File;

import java.io.FileNotFoundException;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writeoperation {

	public static void main(String[] args) throws FileNotFoundException, IOException{
		// TODO Auto-generated method stub

		
		// TODO Auto-generated method stub

				 XSSFWorkbook book=new XSSFWorkbook(); //creating a book
				XSSFSheet sheet = book.createSheet();	//creating a sheet
				
				Object[] [] data = {	 //creating a data
						
						{"NAME","AGE","PLACE"},
						{"BOOPATHI","27","ERODE"},
						{"RAJ","27","ERODE"},
						{"MANI","25","CHENNAI"}
			};
				int rowCount=0; // initializing at row count
				
				for(Object[] row : data) { 
					
					XSSFRow createRow = sheet.createRow(rowCount++);
					
				int columnCount=0;  // initializing at Column count
				
				for(Object column: row) {                   
					
					XSSFCell cell = createRow.createCell(columnCount++);
					
					if(column instanceof String)
					{ 
						cell.setCellValue((String) column);
					}
					else if(column instanceof Integer)
					{
						cell.setCellValue((Integer) column);
					} 
					
					try(                                                 
							FileOutputStream output = new FileOutputStream("D:\\Guvi\\Task 13\\Filewrite.xlsx");){
							book.write(output);           
						}

					}
				}
	}

}
