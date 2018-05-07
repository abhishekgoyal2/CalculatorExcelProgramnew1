package CallCalcualte;

import java.io.IOException;

import org.apache.log4j.Logger;

import Operandsclass.Add;
import Operandsclass.ExcelReader;

public class CallCal {
	public static String operand  ;


	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		Logger log =Logger.getLogger("sdevpinbyLoger");
		ExcelReader excel =new ExcelReader("D:\\automationXpath\\Cal.xlsx");
		log.debug("calling constructor of the method");
		operand=excel.getspecificCelldata(operand);
	System.out.println(operand);
	log.debug("calling excel to get and update the data for operand");

	}

}
