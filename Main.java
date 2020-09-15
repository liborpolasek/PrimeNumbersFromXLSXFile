package primeNumbers;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Main {

	public static void main(String[] args) throws IOException, InvalidFormatException {
        // Enter data using BufferReader 
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in)); 
        String input = "";
        File f;
        String filename;
        PrimeNumbersXLSX primeNumbers;
        
        System.out.println("This program will return all prime numbers from second column in excel file(xlsx).");
        System.out.println("As input insert path to .xlsx file, to quit write 'end'.");
        // Cycle to infinite until it's ended.
        while(true) {
        	System.out.print("Input: ");
        	// Get input from user.
            input = reader.readLine(); 
            f = new File(input);
            filename = f.getName().toLowerCase();
            
            // Test if file exists, it is a file(not directory) and ends with .xlsx
            if (f.exists() && f.isFile() && filename.endsWith(".xlsx")) {
            	// Get all prime numbers from file and put them on output.
            	primeNumbers = new PrimeNumbersXLSX(f);
        		for (int i : primeNumbers.getPrimeNumbers()) {
        			System.out.println(i);
        		}
        		System.out.println();
        		
        	// If user put "end" to input, program will close.
            } else if (input.equals("end")) {
        		break;
        	// If condition isn't true(wrong file), tell user.
            } else {
            	System.out.println("File doesn't exist or it is not .xlsx file.");
            }
        }		
	}
}
