package utils;

import java.io.File;
import java.util.Scanner;

public class Read_File {
	
	private Scanner sc;
	
	public Read_File(File file) {
		// TODO Auto-generated constructor stub

		 try {
	        	this.sc = new Scanner(file);

	        } catch (Exception ex) {
	            ex.printStackTrace();
	        }
	}

	
	
	public String getValueFromFile()
	{
		
		return this.sc.nextLine();
		
	}
	
	public String skipGetValueFromFile()
	{
		this.sc.nextLine();
		return this.sc.nextLine();
	}
	
	public void closeFile()
	{
		this.sc.close();
	}

}
