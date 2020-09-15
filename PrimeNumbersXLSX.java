package primeNumbers;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

public class PrimeNumbersXLSX {

	private Workbook workBook;
	
	public PrimeNumbersXLSX(File file) throws IOException, InvalidFormatException {
		this.workBook = WorkbookFactory.create(file);
	}
	
	public List<Integer> getPrimeNumbers() {
		Sheet sheet  = this.workBook.getSheetAt(0);
		DataFormatter formatter = new DataFormatter();
		List<Integer> primeNumbers = new ArrayList<>();

		// For each Row.
		for (Row row : sheet) { 
			// Get the Cell at the Index / Column.
		    Cell cell = row.getCell(1); 
		    int temp = 0;
		    // Try convert content of Cell to integer. (Test if it's a number.)
		    try{
		    	temp = Integer.parseInt(formatter.formatCellValue(cell).trim());
		    } catch(Exception e) {
		    	// Otherwise set temp to -1, so it is not a prime number.
		    	temp = -1;
		    }
		    // Test if number is prime number and add it to list.
		    if(isPrimeNumber(temp)) {
		    	primeNumbers.add(temp);
		    }
		}
		return primeNumbers;
	}
	
	private boolean isPrimeNumber(int number) {
		if(number < 1) return false;
		for(int i=2; i <= number/2; i++) {
		   if(number%i == 0) {
			   return false;
		   }
		}
		return true;
	}
}
