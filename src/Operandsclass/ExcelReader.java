package Operandsclass;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
public class ExcelReader {

static String operand;
static int valA;
static int valB;
		// TODO Auto-generated method stub
		public static XSSFWorkbook wb;
		public static XSSFSheet wbsheet;
		public static XSSFRow row;
		public static XSSFCell cell;

		public static FileInputStream fis;
		public static FileOutputStream fileout;
		public String path;
		
		
		public ExcelReader(String path) {
			
			this.path=path;
			try {
				fis = new FileInputStream(path);
				 wb = new XSSFWorkbook(fis);
				 wbsheet = wb.getSheetAt(0);
				fis.close();
			} catch (Exception e) {
				
				e.printStackTrace();
			} 
			
		}
		
		public static void main(String[] args) throws IOException {
			// TODO Auto-generated method stub
			
	              
//			fis =new FileInputStream("D:\\automationXpath\\Cal.xlsx");
//
//		      
//		operand=getspecificCelldata(operand);
//		System.out.println();
		
//	fis.close();
//	fileout.close();
		
}
	
		//function to call specific operand 
		 public String getspecificCelldata(String operand) throws IOException
			{
				
//				wb=new XSSFWorkbook(fis);
//				
//				 wbsheet=wb.getSheetAt(0);
			 
				 int row=wbsheet.getLastRowNum();
				 System.out.println("Rows " +row);
			int col=wbsheet.getRow(0).getLastCellNum();	
			
			System.out.println("Column " + col);
			
			for(int i=1;i<=row;i++)
			{
				XSSFCell opcell=wbsheet.getRow(i).getCell(0);
				XSSFCell Acell=wbsheet.getRow(i).getCell(1);//Cells under ColA
				XSSFCell Acel2=wbsheet.getRow(i).getCell(2);//Cells under ColB
				
				 operand =opcell.getStringCellValue();
				
				 valA=(int) Acell.getNumericCellValue();
				 valB=(int) Acel2.getNumericCellValue();
				 
				
				 Row rowcal = wbsheet.getRow(i);
				 Cell cell = rowcal.getCell(3);
				 if (cell == null)
				     cell = rowcal.createCell(3);
				 cell.setCellType(Cell.CELL_TYPE_NUMERIC);

				 
				
				System.out.println("Operand " + operand);
				if(operand.equalsIgnoreCase("+"))
				{
					System.out.println("Call plus operand");
					Add add = new Add();
				
					 int c = add.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
				
				cell.setCellValue(c);
				 int valC =(int) cell.getNumericCellValue();
				
				System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +valC);
				 fis.close();
				 fileOutStream1();

		            }
				
				
				
				else if(operand.equalsIgnoreCase("-"))
				{
					System.out.println("Call minus operand");
					
					Minus min = new Minus();
					
					 int c = min.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					 cell.setCellValue(c);
				
					 int valC =(int) cell.getNumericCellValue();

						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +valC);
						 fis.close();
						 fileOutStream1();
				}
				
				else if(operand.equalsIgnoreCase("*"))
				{
					System.out.println("Call multiplication operand");
					
					
					Multiplication mul = new Multiplication();
					
					 int c = mul.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					cell.setCellValue(c);
					
					 int valC =(int) cell.getNumericCellValue();

						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +valC);
						 fis.close();
						 fileOutStream1();



                            }
				
				else if(operand.equalsIgnoreCase("/"))
				{
					System.out.println("Call division operand");
					
					Division div = new Division();
					
					 int c = div.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					cell.setCellValue(c);
					
					 int valC =(int) cell.getNumericCellValue();

						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +valC);
						 fis.close();
						 fileOutStream1();
					


				}
				
				else if(operand.equalsIgnoreCase("%"))
				{
					System.out.println("Call  percentage operand");

					
					Percentage per = new Percentage();
					
					 int c = per.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					cell.setCellValue(c);
					
					 int valC =(int) cell.getNumericCellValue();
					 fis.close();
					 fileOutStream1();


					            
				}
				
				else {
					try {
					
						System.out.println("Not a valid operand ");
					}
					catch(Exception e)
					{
						e.printStackTrace();
					}
				}
				

			}
			

			return operand;
			
			
			}
		
		 public static void fileOutStream1() throws IOException
		 {
			 fileout = new FileOutputStream("path");
	            wb.write(fileout);
fileout.close();
		 }
		
	
	}