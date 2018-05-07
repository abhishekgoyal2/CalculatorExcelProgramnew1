package CallCalcualte;

import java.io.IOException;

import Operandsclass.Add;
import Operandsclass.ExcelReader;

public class CallCal {
	public static String operand  ;


	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		ExcelReader excel =new ExcelReader("D:\\automationXpath\\Cal.xlsx");
	operand=excel.getspecificCelldata(operand);
	}

}
