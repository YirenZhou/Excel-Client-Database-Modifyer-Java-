import java.util.*;
//	The following packages are for the loading bar
import java.awt.BorderLayout;
import java.awt.Container;
import javax.swing.BorderFactory;
import javax.swing.JFrame;
import javax.swing.JProgressBar;
import javax.swing.border.Border;
//	The following packages are for file operations
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;
import java.text.DateFormat;
import java.text.DateFormatSymbols;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.FieldPosition;
import java.text.Format;
import java.text.ParseException;
import java.text.SimpleDateFormat;
public class Ultimate {
	public static void textCentering(String s)
	{
		int length = s.length();
		int pos = (int)((80 - length) / 2);
		for (int i = 0; i < pos; ++i)
			System.out.print(" ");

		System.out.println(s);
	}

	public static void systemTitle(String s)
	{
		System.out.print("\n\n\n");
		int length = s.length();
		int pos = (int)((80 - length - 2) / 2);

		for (int i = 0; i < pos; ++i)
			System.out.print(" ");

		System.out.print("*");
		for (int j = 0; j < length; ++j)
			System.out.print("*");
			System.out.println("*");

		for (int k = 0; k < pos + 1; ++k)
			System.out.print(" ");

		System.out.println(s);

		for (int i = 0; i < pos; ++i)
			System.out.print(" ");

		System.out.print("*");
		for (int j = 0; j < length; ++j)
			System.out.print("*");
		System.out.println("*");

		System.out.print("\n\n");
	}

	public static boolean emptyInput(String s)
	{
		int count = s.length();
		char[] convertStr = s.toCharArray();
		if (count == 0)
			return false;
		else if (convertStr[0] == ' ')
			return false;
		else 
			return true;
	}

	public static int textCenteringReturn(String s)
	{
		return (int)((80 - s.length()) / 2);
	}

	public static void support ()
	{
		//	Listing all of the options we have
		System.out.println("Please select one of the following: ");
		System.out.println("a. What is the program for");
		System.out.println("b. Report a problem (no more than 200 words)");
		System.out.println("c. Suggest a feature (no more than 200 words)");
		System.out.println("d. About us");
		System.out.println("q. Quit");
		//	Switch which does operations
		String alpha = "a";
		Scanner inputChar = new Scanner (System.in);
		char[] inputAlpha = alpha.toCharArray();
		char letter = inputAlpha[0];
		while (letter != 'q' || letter != 'Q')
		{
			switch (letter)
			{
				case 'a':
				case 'A':
		    	 	System.out.println("This program is designed to facilitate members in the team of EEUS with adding / removing client ");
					System.out.println("and searching clients' work areas. Team could use this program to directly complete their daily ");
			 	    System.out.println("routines without opening the workbook. However, they still need to run the macros after doing any actions.");
			 		System.out.println("The program is not responsible for any wrong names or information.");
			  		break;

				case 'b':
				case 'B':
	    		{
		   			String problem;
		    		System.out.println("Plase state your problem below: ");
		    		problem = inputChar.nextLine();
		    		if (emptyInput(problem))
		    		{
				   	 	System.out.println("Thanks for your valuable feedback. Our team would work on it. If you still have other questions, ");
				    	System.out.println("please contact Yiren Zhou: Yiren_Zhou@manulife.com.");
		    		}
		    		else 
			    		System.err.println("Please at least say something!");
		    		break;
	    		}

				case 'c':
				case 'C':
	    		{
		    		String feature;
		    		System.out.print("What feature would you suggest: ");
		    		feature = inputChar.nextLine();
		    		if (emptyInput(feature))
		    		{
		    			System.out.println("Thanks for your valuable feedback. Our team would work on it. If you still have other questions ");
		    			System.out.println("please contact Yiren Zhou: Yiren_Zhou@manulife.com.");
		    			break;
		    		}
		    		else
		   	   	 		System.out.println("Please at least say something!" );
		    		break;
	    		}

				case 'd':
				case 'D':
		    		System.out.println("EEUS stands for ENHANCED END USER SERVICE, and we are a team that helps employees at Manulife have an easier life ");
			    	System.out.println("with their electronics including laptops, mobiles, iPads, and printers, etc. We deal with customers, and fix");
		   	    	System.out.println(" their technical problems. Need a hand at work? Feel free to buzz us!");
		    		break;

		    	case 'q':
		   	 	case 'Q': 
		    		return;

		   		default:
		    		System.err.println("Wrong input. Type again.");;
		    		break;
			}
		}
	}
	
	public static void clientSearch(String name)throws FileNotFoundException, IOException, ParseException
	{
		//	Open the workbook
		File xlsxfile = new File("/Users/Yiren Zhou/Documents/workspace/ClientListMaster.xlsx");
		//	Initialize / declare a Workbook object
		Workbook wb = null;
		//	Initialize / declare a Sheet object
		Sheet sheet = null;
		//	Initialize / declare a Row object
		Row row = null;
		//	Test if the file is opened successfully
		//	If it is successful, continue to do series of actions
		FileInputStream fileInputStream = new FileInputStream(xlsxfile);
		if (xlsxfile.isFile() & xlsxfile.exists())
		{
			System.out.println("Database Opened Successfully.");
			wb = new XSSFWorkbook(fileInputStream);
			//	Navigate to the Master tab
			sheet = wb.getSheet("Master");
			boolean signal = false;
			int rowSearch = 0;
			Row rowSearching;
			Cell cellSearching;
			for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++)
			{
				rowSearching = sheet.getRow(rowIndex);
				for (int colIndex = 0; colIndex < 24; colIndex++)
				{
					if (colIndex == 2)
					{	
						cellSearching = rowSearching.getCell(colIndex);
						cellSearching.setCellType(Cell.CELL_TYPE_STRING);
						if (cellSearching.getStringCellValue().equals(name))
						{
							rowSearch = rowIndex;
							signal = true;
							break;
						}
					}
					if (signal == true)
						break;
				}
				if (signal == true)
					break;
			}

			if (signal)
			{	
				//	Navigate to the row which contains the required info
				row = sheet.getRow(rowSearch);
				//	Display the subscription date
				Cell column = row.getCell(1);
				String subscriptionDate = null;
				SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
				Date theDate = column.getDateCellValue();
				subscriptionDate = df.format(theDate);
				System.out.println("Subscription date: " + subscriptionDate);
				//	Display the corresponding reps
				column = row.getCell(3);
				if (column.getStringCellValue().equals("AT"))
					System.out.println("Designated Resources: Alex & Tracey");
				else if (column.getStringCellValue().equals("MA"))
					System.out.println("Designated Resources: Alket, Amin & Moe");
				else if (column.getStringCellValue().equals("Mike"))
					System.out.println("Designated Resources: Mike");
				else if (column.getStringCellValue().equals("RJ"))
					System.out.println("Designated Resources: Richard & Joel");
				else if (column.getStringCellValue().equals("RR"))
					System.out.println("Designated Resources: Robert C. & Robert L.");
				else if (column.getStringCellValue().equals("SD"))
					System.out.println("Designated Resources: Simon & Dean");
				else if (column.getStringCellValue().equals("SK"))
					System.out.println("Designated Resources: Steve & Kevin");
				//	Display the cost centre
				column = row.getCell(11);
				System.out.println("Cost Centre: " + column.getStringCellValue());

				//	Display address
				column = row.getCell(13);
				System.out.println("Work Address; " + column.getStringCellValue());
				//	Display city
				column = row.getCell(14);
				System.out.println("City: " + column.getStringCellValue());
				//	Display country
				column = row.getCell(15);
				System.out.println("Country: " + column.getStringCellValue());
				//	Display group
				column = row.getCell(16);
				System.out.println("Group: " + column.getStringCellValue());
				//	Display Email
				column = row.getCell(21);
				System.out.println("Email: " + column.getStringCellValue());
				//	Display user name
				column = row.getCell(22);
				System.out.println("User name: " + column.getStringCellValue());
				//	Display employee ID
				column = row.getCell(23);
				if (column.getCellType() == Cell.CELL_TYPE_NUMERIC)
				{
					String employeeID = String.format("%.0f", column.getNumericCellValue());
					System.out.println("Employee ID: " + employeeID);
				}
				else 
					System.out.println("Employee ID: " + column.getStringCellValue());
			}
			else 
			{
				System.out.println("This client is not in our list.");
				System.out.println("Going through the Un-subscription list.");
				//	Navigate to the un-subscription sheet
				wb = new XSSFWorkbook(fileInputStream);
				sheet = wb.getSheet("Unsubscribed");
				boolean unSignal = false;
				int unRowSearch = 0;
				Row unRowSearching;
				Cell unCellSearching;
				for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++)
				{
					unRowSearching = sheet.getRow(rowIndex);
					for (int colIndex = 0; colIndex < 24; colIndex++)
					{
						if (colIndex == 2)
						{	
							unCellSearching = unRowSearching.getCell(colIndex);
							unCellSearching.setCellType(Cell.CELL_TYPE_STRING);
							if (unCellSearching.getStringCellValue().equals(name))
							{
								unRowSearch = rowIndex;
								unSignal = true;
								break;
							}
						}
						if (unSignal == true)
							break;
					}
					if (unSignal == true)
						break;
				}

				if (unSignal)
				{
					System.out.println("Un-subscribed client found.");
					row = sheet.getRow(rowSearch);
					//	Display un-sub date
					Cell column = row.getCell(0);
					String unSubDate = null;
					SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
					Date theDate = column.getDateCellValue();
					unSubDate = df.format(theDate);
					System.out.println("Un-subscription date: " + unSubDate);

					//	Display subscription date
					column = row.getCell(1);
					String subDate;
					theDate = column.getDateCellValue();
					subDate = df.format(theDate);
					System.out.println("Subscription date: " + subDate);

					//	Display reps
					column = row.getCell(3);
					if (column.getStringCellValue() == "AT")
						System.out.println("Designated Resources: Alex & Tracey");
					else if (column.getStringCellValue() == "MA")
						System.out.println("Designated Resources: Alket, Amin & Moe");
					else if (column.getStringCellValue() == "Mike")
						System.out.println("Designated Resources: Mike");
					else if (column.getStringCellValue() == "RJ")
						System.out.println("Designated Resources: Richard & Joel");
					else if (column.getStringCellValue() == "RR")
						System.out.println("Designated Resources: Robert C. & Robert L.");
					else if (column.getStringCellValue() == "SD")
						System.out.println("Designated Resources: Simon & Dean");
					else if (column.getStringCellValue() == "SK")
						System.out.println("Designated Resources: Steve & Kevin");
					//	Display cost centre
					column = row.getCell(11);
					System.out.println("Cost Centre: " + column.getStringCellValue());
					//	Display address
					column = row.getCell(13);
					System.out.println("Work Address; " + column.getStringCellValue());
					//	Display city
					column = row.getCell(14);
					System.out.println("City: " + column.getStringCellValue());
					//	Display country
					column = row.getCell(15);
					System.out.println("Country: " + column.getStringCellValue());
					//	Display group
					column = row.getCell(16);
					System.out.println("Group: " + column.getStringCellValue());
				}
			}
		}
	}
	
	public static void clientRemove(String name) throws FileNotFoundException, IOException, ParseException
	{
		//	Open the workbook
		File xlsxfile = new File("/Users/Yiren Zhou/Documents/workspace/ClientListMaster.xlsx");
		int rowSearch = 0;
		int rowCount = 0;
		//	Initialize / declare a Workbook object
		Workbook wb = null;
		//	Initialize / declare a Sheet object
		Sheet sheet = null;
		//	Initialize / declare a Row object
		Row row = null;
		//	Test if the file is opened successfully
		//	If it is successful, continue to do series of actions
		FileInputStream fileInputStream = new FileInputStream(xlsxfile);
		if (xlsxfile.isFile() & xlsxfile.exists())
		{
			System.out.println("Database Opened Successfully.");
			wb = new XSSFWorkbook(fileInputStream);
			//	Navigate to the Master tab
			sheet = wb.getSheet("Master");
			//	Also navigate to the un-subscribed tab
			Sheet sheetUnsub = wb.getSheet("Unsubscribed");
			rowCount = sheetUnsub.getPhysicalNumberOfRows();
			boolean signal = false;
			Row rowSearching;
			Cell cellSearching;
			for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++)
			{
				rowSearching = sheet.getRow(rowIndex);
				for (int colIndex = 0; colIndex < 24; colIndex++)
				{
					if (colIndex == 2)
					{	
						cellSearching = rowSearching.getCell(colIndex);
						cellSearching.setCellType(Cell.CELL_TYPE_STRING);
						if (cellSearching.getStringCellValue().equals(name))
						{
							rowSearch = rowIndex;
							signal = true;
							break;
						}
					}
					if (signal == true)
						break;
				}
				if (signal == true)
					break;
			}


			if (signal)
			{
				//	Set the font first
				Font font = wb.createFont();
				font.setFontHeightInPoints((short)12);
				CellStyle fontStyle = wb.createCellStyle();
				fontStyle.setFont(font);
				
				//	Navigate to the row which contains the required info
				row = sheet.getRow(rowSearch);
				Cell column = row.getCell(1);
				//	Navigate to the last row in the unsubscribed tab
				rowCount = sheetUnsub.getPhysicalNumberOfRows();
				Row rowUnsub = sheetUnsub.getRow(rowCount);
				if (rowUnsub == null)
					rowUnsub = sheetUnsub.createRow(rowCount);
				
				//	Copy the subscription date
				Cell cell = rowUnsub.createCell((short) 1);
				String subDate = null;
				SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
				Date theDate = column.getDateCellValue();
				subDate = df.format(theDate);
				cell.setCellValue(new Date());
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				Date actualSubDate = df.parse(subDate);
				cell.setCellValue(actualSubDate);
				CellStyle subDateCellStyle = wb.createCellStyle();
				short subdf = wb.createDataFormat().getFormat("d-MMM-yy");
				CellStyle combined = wb.createCellStyle();
				combined.cloneStyleFrom(fontStyle);
				combined.setDataFormat(subdf);
				cell.setCellStyle(combined);

				//	Copy the location details
				for (int j = 2; j <= 16; ++j)
				{
					column = row.getCell(j);
					cell = rowUnsub.createCell(j);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(column.getStringCellValue());
					cell.setCellStyle(fontStyle);
				}
				//	Asking user for unsubscription date
				cell = rowUnsub.createCell((short) 0);
				String unsubDate = null;
				System.out.println("Please enter the unsubscription date (yyyy-MM-dd): ");
				//	Take user's input
				Scanner keyboard = new Scanner(System.in);
				unsubDate = keyboard.nextLine();
				cell.setCellValue(new Date());
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				SimpleDateFormat unsubFM = new SimpleDateFormat("yyyy-MM-dd");
				Date actualUnsubDate = unsubFM.parse(unsubDate);
				unsubFM = new SimpleDateFormat("d-MMM-yy");
				//	System.out.println(datetemp.format(actualDate));
				cell.setCellValue(actualUnsubDate);
				cell.setCellStyle(combined);

				//	Completely delete the row in the "Master" tab
				removeRow(sheet, rowSearch);
				FileOutputStream fileOut = new FileOutputStream("D:\\Documents\\ClientListMaster.xlsx");
				wb.write(fileOut);
				fileOut.close();
				System.out.println("Client " + name + " has been successfully removed.");
			}
		}
		else
			System.err.println("Client " + name + " not found. Please double check your request.");
	}
	
	public static void removeRow(Sheet sheet, int rowIndex)
	{
		int lastRowNum = sheet.getLastRowNum();
		if (rowIndex >= 0 && rowIndex < lastRowNum)
		{
			sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
		}
		if (rowIndex == lastRowNum)
		{
			Row removingRow = sheet.getRow(rowIndex);
			if (removingRow != null)
			{
				sheet.removeRow(removingRow);
			}
		}
	}
	
	public static void clientAdd(String full)throws FileNotFoundException, IOException, ParseException
	{
		char[] name = full.toCharArray();
		//	Open the workbook named "ClientListMaster" at a specific location
		File xlsxfile = new File("/Users/Yiren Zhou/Documents/workspace/ClientListMaster.xlsx");
		int rowCount = 0;
		//	Initialize / declare a Workbook object
		Workbook wb = null;
		//	Initialize / declare a Sheet object
		Sheet sheet = null;
		//	Declare / initialize a Row object
		Row row = null;
		/*
		 * Create a FileInputStream by opening a connection to an actual file,
		 * the file named by the File object file in the file system. A new FileDescriptor
		 * object is created to represent this file connection.
		 */
		//	Test if the file is opened successfully
		//	If it is successful, continue to do a series of actions
		FileInputStream fileInputStream = new FileInputStream(xlsxfile);
		if (xlsxfile.isFile() & xlsxfile.exists())
		{
			String legalName;
			String subscribed;
			String date;
			String rep;
			String costCentre;
			String email;
			String username;
			long employeeID;

			System.out.println("Database Opened Successfully.");
			wb = new XSSFWorkbook(fileInputStream);
			//	Navigate to the Master tab
			sheet = wb.getSheet("Master");
			//	getPhysicalNumberOfRows() returns an integer that is the number of the physically defined
			//	Not the number of rows in the sheet because that could be infinite
			//	It should only be the number of rows which actually have something
			rowCount = sheet.getPhysicalNumberOfRows();
			//	Get to the desired row. 
			sheet.createRow(rowCount);
			row = sheet.getRow(rowCount);
			if (row == null)
				row = sheet.createRow(rowCount);
			//	Now, add individual elements into the corresponding cell. 
			//	Ask for this client's information
			//	System.out.println("Please enter the name of the client (Last, First)");
			Scanner keyboard = new Scanner(System.in);
			//	name = keyboard.nextLine();

			//	Extract the first and last name from the full name
			int countFirst = 0;
			int countTotal = name.length;
			while (name[countFirst] != ' ')
				countFirst++;

			/*while (name[countTotal] != '\0')
				countTotal++;*/

			int dataLen = countFirst;
			int lastRef = countTotal - countFirst - 1;
			int reference = countTotal - countFirst + 1;
			int ref = 0;
			
			// System.out.println("datalen: " + dataLen + " lastRef: " + lastRef + " reference: " + reference);

			char[] firstName = new char[dataLen];
			char[] lastName = new char [lastRef];

			//	Extract the first name from the string
			for (int i = 0; i < countFirst; ++i)
				firstName[i] = name[i];

			//	Extract the last name from the string
			while (ref < lastRef)
			{	
				if (name[reference] == ' ')
					reference++;

				lastName[ref] = name[reference];
				ref++;
				reference++;
			}
			String firstString = new String(firstName);
			String lastString = new String (lastName);
			legalName = lastString + ", " + firstString;
			//	Convert char arrays back to Strings
			//	Fill the cells with the name
			String first = String.copyValueOf(firstName);
			String last = String.copyValueOf(lastName);
			String fullName = String.copyValueOf(name);
			
			Cell column = row.createCell(2);
			column.setCellType(Cell.CELL_TYPE_STRING);
			column.setCellValue(fullName);
			column = row.createCell(4);
			column.setCellType(Cell.CELL_TYPE_STRING);
			column.setCellValue(legalName);
			column = row.createCell(5);
			column.setCellType(Cell.CELL_TYPE_STRING);
			column.setCellValue(last);
			column.setCellType(Cell.CELL_TYPE_STRING);
			column = row.createCell(6);
			column.setCellType(Cell.CELL_TYPE_STRING);
			column.setCellValue(first);
			//	Subscription Status
			System.out.println("Please enter the client's subscription status (Yes, CDN, SLI, US or UK):");
			subscribed = keyboard.nextLine();
			column = row.createCell(0);
			column.setCellValue(subscribed);
			//	Subscription Date
			System.out.println("Please enter the client's subscription date (Example: 2013-06-30):");
			date = keyboard.nextLine();
			column = row.createCell((short)1);
			column.setCellValue(new Date());
			column.setCellType(Cell.CELL_TYPE_NUMERIC);
			SimpleDateFormat datetemp = new SimpleDateFormat("yyyy-MM-dd");
			Date actualDate = datetemp.parse(date);
			column.setCellValue(actualDate);
			CellStyle dateCellStyle = wb.createCellStyle();
			short df = wb.createDataFormat().getFormat("d-MMM-yy");
			dateCellStyle.setDataFormat(df);
			column.setCellStyle(dateCellStyle);
			
			//	Rep
			System.out.println("Please enter the client's designated resource: ");
			rep = keyboard.nextLine();
			column = row.createCell(3);
			column.setCellValue(rep);
			//	Cost centre & Location
			System.out.println("Please enter the 4-digit cost centre number: ");
			costCentre = keyboard.nextLine();
			column = row.createCell(11);
			column.setCellValue(costCentre);
			//	Search if the cost centre originally exists
			int rowNumber = 0;
			boolean signal = false;
			Row rowSearching;
			Cell cellSearching;
			
			for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++)
			{
				rowSearching = sheet.getRow(rowIndex);
				for (int colIndex = 0; colIndex < 24; colIndex++)
				{
					if (colIndex == 11)
					{	
						cellSearching = rowSearching.getCell(colIndex);
						cellSearching.setCellType(Cell.CELL_TYPE_STRING);
						if (cellSearching.getStringCellValue().equals(costCentre))
						{
							rowNumber = rowIndex;
							signal = true;
							break;
						}
					}
					if (signal == true)
						break;
				}
				if (signal == true)
					break;
			}
		
			if (rowNumber != 0)
			{	
				String division;
				String subDivision;
				String businessUnit;
				String departmentName;
				String ccName;
				String locationName;
				String city;
				String country;
				String group;
				Row row2 = sheet.getRow(rowNumber);
				Cell column2 = null;
				//	Division
				for (int index = 7; index < 11; ++index)
				{
					column2 = row2.getCell(index);
					String info = column2.getStringCellValue();
					column = row.createCell(index);
					column.setCellValue(info);
				}
				column2 = row2.getCell(7);
				division = column2.getStringCellValue();
				column = row.createCell(7);
				column.setCellValue(division);
				//	Sub Division
				column2 = row2.getCell(8);
				subDivision = column2.getStringCellValue();
				column = row.createCell(8);
				column.setCellValue(subDivision);
				//	Business Unit Name
				column2 = row2.getCell(9);
				businessUnit = column2.getStringCellValue();
				column = row.createCell(9);
				column.setCellValue(businessUnit);
				//	Department Name
				column2 = row2.getCell(10);
				departmentName = column2.getStringCellValue();
				column = row.createCell(10);
				column.setCellValue(departmentName);
				
				//	Cost Centre Name
				for (int index = 12; index < 17; ++index)
				{
					column2 = row2.getCell(index);
					String info = column2.getStringCellValue();
					column = row.createCell(index);
					column.setCellValue(info);
				}
				System.out.println("Cost centre found. Location information has been automatically filled.");
			}
			else
			{
				String division;
				String subDivision;
				String businessUnit;
				String departmentName;
				String ccName;
				String locationName;
				String city;
				String country;
				String group;

				System.out.println("Cost centre does not exist in our system. Please enter the info manually.");
				//	Division Name
				System.out.println("Please enter the division name: ");
				division = keyboard.nextLine();
				column = row.getCell(7);
				column.setCellValue(division);
				//	Sub Division Name
				System.out.println("Please enter the sub-division name: ");
				subDivision = keyboard.nextLine();
				column = row.getCell(8);
				column.setCellValue(subDivision);
				//	Business Unit Name
				System.out.println("Please enter the business unit name: ");
				businessUnit = keyboard.nextLine();
				column = row.getCell(9);
				column.setCellValue(businessUnit);
				//	Department Name
				System.out.println("Please enter the department name: ");
				departmentName = keyboard.nextLine();
				column = row.getCell(10);
				column.setCellValue(departmentName);
				//	Cost Centre Name
				System.out.println("Please enter the cost centre name: ");
				ccName = keyboard.nextLine();
				column = row.getCell(12);
				column.setCellValue(ccName);
				//	Location Name
				System.out.println("Please enter the location name: ");
				locationName = keyboard.nextLine();
				column = row.getCell(13);
				column.setCellValue(locationName);
				//	City
				System.out.println("Please enter the city name: ");
				city = keyboard.nextLine();
				column = row.getCell(14);
				column.setCellValue(city);
				//	Country
				System.out.println("Please enter the country name: ");
				country = keyboard.nextLine();
				column = row.getCell(15);
				column.setCellValue(country);
				//	Group
				System.out.println("Please enter the group name: ");
				group = keyboard.nextLine();
				column = row.getCell(16);
				column.setCellValue(group);
			}
			//	Email
			System.out.println("Please enter the email address: ");
			email = keyboard.nextLine();
			column = row.createCell(21);
			column.setCellValue(email);
			//	User name
			System.out.println("Please enter the username: ");
			username = keyboard.nextLine();
			column = row.createCell(22);
			column.setCellValue(username);
			//	Employee ID
			System.out.println("Please enter the employee ID: ");
			employeeID = keyboard.nextLong();
			column = row.createCell(23);
			column.setCellValue(employeeID);

			//	Write the output to a file
			FileOutputStream fileOut = new FileOutputStream("/Users/Yiren Zhou/Documents/workspace/ClientListMaster.xlsx");
			wb.write(fileOut);
			fileOut.close();
			System.out.println("Client has been added successfully!");
		}
	}
	
	public static void fileModify(String name)throws FileNotFoundException, IOException, ParseException
	{
		//	Open the workbook
		File xlsxfile = new File("/Users/Yiren Zhou/Documents/workspace/ClientListMaster.xlsx");
		//	Initialize / declare a Workbook object
		Workbook wb = null;
		//	Initialize / declare a Sheet object
		Sheet sheet = null;
		//	Initialize / declare a Row object
		Row row = null;
		//	Test if the file is opened successfully
		//	If it is successful, continue to do series of actions

		FileInputStream fileInputStream = new FileInputStream(xlsxfile);
		if (xlsxfile.isFile() & xlsxfile.exists())
		{
			System.out.println("Database Opened Succesfully.");
			wb = new XSSFWorkbook(fileInputStream);
			//	Navigate to the Master tab
			sheet = wb.getSheet("Master");
			//	Navigate to the correct row
			//	Search for the client's name
			int rowNumber = 0;
			boolean signal = false;
			Row rowSearching;
			Cell cellSearching;
			for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++)
			{
				rowSearching = sheet.getRow(rowIndex);
				for (int colIndex = 0; colIndex < 24; colIndex++)
				{
					if (colIndex == 2)
					{	
						cellSearching = rowSearching.getCell(colIndex);
						cellSearching.setCellType(Cell.CELL_TYPE_STRING);
						if (cellSearching.getStringCellValue().equals(name))
						{
							rowNumber = rowIndex;
							signal = true;
							break;
						}
					}
					if (signal == true)
						break;
				}
				if (signal == true)
					break;
			}
			//	If the client is found in the system
			if (signal)
			{	
				System.out.println("Client found. Processing...");
				System.out.println("Please select from below: ");
				//	Navigate to the row which contains the required info
				row = sheet.getRow(rowNumber);
				//	Let the user choose the aspect needed to be changed
				Scanner inputNumber = new Scanner(System.in);
				int input = inputNumber.nextInt();
				switch(input)
				{	
					//	Change the cost centre
					case 1:
						{	
							//	Ask for new cost centre number
							System.out.println("Please enter the new cost centre (4-digit): ");
							String ccNew;	
							Scanner keyboardInt = new Scanner(System.in);
							ccNew = keyboardInt.nextLine();
							int ccFind = 0;
							boolean signalCC = false;
							Row ccRow;
							Cell ccCell;
							for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++)
							{
								ccRow = sheet.getRow(rowIndex);
								for (int colIndex = 0; colIndex < 24; colIndex++)
								{
									if (colIndex == 11)
									{	
										ccCell = ccRow.getCell(colIndex);
										ccCell.setCellType(Cell.CELL_TYPE_STRING);
										if (ccCell.getStringCellValue().equals(ccNew))
										{
											ccFind = rowIndex;
											signalCC = true;
											break;
										}
									}
									if (signalCC == true)
										break;
								}
								if (signalCC == true)
									break;
							}
							//	If the cost center is found in the system	
							if (signalCC)
							{
								Row ccSearch = sheet.getRow(ccFind);
								//	Get Division Name
								Cell cell = ccSearch.getCell(7);
								String divisionName = cell.getStringCellValue();
								//	Get the Sub Division Name
								cell = ccSearch.getCell(8);
								String subDivName = cell.getStringCellValue();
								//	Get Business Unit Name
								cell = ccSearch.getCell(9);
								String buName = cell.getStringCellValue();
								//	Get Department Name
								cell = ccSearch.getCell(10);
								String deptName = cell.getStringCellValue();
								//	Get Cost Centre Name
								cell = ccSearch.getCell(12);
								String ccName = cell.getStringCellValue();
								//	Get Location / Address
								cell = ccSearch.getCell(13);
								String address = cell.getStringCellValue();
								//	Get City
								cell = ccSearch.getCell(14);
								String city = cell.getStringCellValue();
								//	Get Country
								cell = ccSearch.getCell(15);
								String country = cell.getStringCellValue();
								//	Get Group
								cell = ccSearch.getCell(16);
								String group = cell.getStringCellValue();

								//	Make the changes
								Cell column = row.getCell(7);
								if (column == null)
									column = row.createCell(7);
								column.setCellType(Cell.CELL_TYPE_STRING);
								column.setCellValue(divisionName);

								column = row.getCell(8);
								if (column == null)
									column = row.createCell(8);
								column.setCellType(Cell.CELL_TYPE_STRING);
								column.setCellValue(subDivName);

								column = row.getCell(9);
								if (column == null)
									column = row.createCell(9);
								column.setCellType(Cell.CELL_TYPE_STRING);
								column.setCellValue(buName);

								column = row.getCell(10);
								if (column == null)
									column = row.createCell(10);
								column.setCellType(Cell.CELL_TYPE_STRING);
								column.setCellValue(deptName);

								column = row.getCell(11);
								if (column == null)
									column = row.createCell(11);
								column.setCellType(Cell.CELL_TYPE_STRING);
								column.setCellValue(ccNew);

								column = row.getCell(12);
								if (column == null)
									column = row.createCell(12);
								column.setCellType(Cell.CELL_TYPE_STRING);
								column.setCellValue(ccName);

								column = row.getCell(13);
								if (column == null)
									column = row.createCell(13);
								column.setCellType(Cell.CELL_TYPE_STRING);
								column.setCellValue(address);

								column = row.getCell(14);
								if (column == null)
									column = row.createCell(14);
								column.setCellType(Cell.CELL_TYPE_STRING);
								column.setCellValue(city);

								column = row.getCell(15);
								if (column == null)
									column = row.createCell(15);
								column.setCellType(Cell.CELL_TYPE_STRING);
								column.setCellValue(country);

								column = row.getCell(16);
								if (column == null)
									column = row.createCell(16);
								column.setCellType(Cell.CELL_TYPE_STRING);
								column.setCellValue(group);
								//	Write the file out
								FileOutputStream fileOut = new FileOutputStream("D:\\Documents\\ClientListMaster.xlsx");
								wb.write(fileOut);
								fileOut.close();
								System.out.println("The work area information has been modified successfully.");
								System.out.println("File modified successfully.");
								break;
							}
							//	If the cost center is not found
							else
							{
								char signalUser = 'F';
								Scanner keyboard = new Scanner(System.in);
								System.err.println("Cost center not found. Please change manually.");
								Cell column = row.getCell(7);
								if (column == null)
									column = row.createCell(7);
								column.setCellType(Cell.CELL_TYPE_STRING);

								String divisionName = null;
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the division name: ");
									divisionName = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(divisionName);
								signalUser = 'F';

								column = row.getCell(8);
								String subDivName = null;
								if (column == null)
									column = row.createCell(8);
								column.setCellType(Cell.CELL_TYPE_STRING);
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the sub-division name: ");
									subDivName = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(subDivName);
								signalUser = 'F';

								column = row.getCell(9);
								String buName = null;
								if (column == null)
									column = row.createCell(9);
								column.setCellType(Cell.CELL_TYPE_STRING);
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the business unit name: ");
									buName = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(buName);
								signalUser = 'F';


								column = row.getCell(10);
								String deptName = null;
								if (column == null)
									column = row.createCell(10);
								column.setCellType(Cell.CELL_TYPE_STRING);
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the department name: ");
									deptName = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(deptName);
								signalUser = 'F';

								column = row.getCell(11);
								if (column == null)
									column = row.createCell(11);
								column.setCellType(Cell.CELL_TYPE_STRING);
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the cost center number: ");
									ccNew = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(ccNew);
								signalUser = 'F';

								column = row.getCell(12);
								String ccName = null;
								if (column == null)
									column = row.createCell(12);
								column.setCellType(Cell.CELL_TYPE_STRING);
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the name of the cost center: ");
									ccName = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(ccName);
								signalUser = 'F';

								column = row.getCell(13);
								String address = null;
								if (column == null)
									column = row.createCell(13);
								column.setCellType(Cell.CELL_TYPE_STRING);
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the name of the location / address: ");
									address = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(address);
								signalUser = 'F';


								column = row.getCell(14);
								String city;
								if (column == null)
									column = row.createCell(14);
								column.setCellType(Cell.CELL_TYPE_STRING);
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the city name: ");
									city = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(address);
								signalUser = 'F';

								column = row.getCell(15);
								String country = null;
								if (column == null)
									column = row.createCell(15);
								column.setCellType(Cell.CELL_TYPE_STRING);
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the country name: ");
									country = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(country);
								signalUser = 'F';

								column = row.getCell(16);
								String group = null;
								if (column == null)
									column = row.createCell(16);
								column.setCellType(Cell.CELL_TYPE_STRING);
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the group name: ");
									group = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(group);
								signalUser = 'F';

								//	Write the file out
								FileOutputStream fileOut = new FileOutputStream("D:\\Documents\\Workspace\\ClientListMaster.xlsx");
								wb.write(fileOut);
								fileOut.close();
								System.out.println("The work area information has been modified successfully");
								System.out.println("File modified successfully.");
								break;
							}
						}
						//	Change the corresponding rep
						case 2:
							{	
								Cell column = row.getCell(3);
								String repInput = null;
								if (column == null)
									column = row.createCell(3);
								column.setCellType(Cell.CELL_TYPE_STRING);
								Scanner keyboard = new Scanner(System.in);
								char signalUser = 'F';
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the new rep: ");
									repInput = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(repInput);
								signalUser = 'F';
								//	Write the file out
								FileOutputStream fileOut = new FileOutputStream("D:\\Documents\\ClientListMaster.xlsx");
								wb.write(fileOut);
								fileOut.close();
								System.out.println("The rep is now " + repInput + ".");
								System.out.println("File modified successfully.");
								break;
							}
						//	Change the email address
						case 3:
							{
								Cell column = row.getCell(21);
								String emailNew = null;
								if (column == null)
									column = row.createCell(21);
								column.setCellType(Cell.CELL_TYPE_STRING);
								Scanner keyboard = new Scanner(System.in);
								char signalUser = 'F';
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the new email address: ");
									emailNew = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(emailNew);
								signalUser = 'F';
								//	Write the file out
								FileOutputStream fileOut = new FileOutputStream("D:\\Documents\\ClientListMaster.xlsx");
								wb.write(fileOut);
								fileOut.close();
								System.out.println("The email address is now: " + emailNew + ".");
								System.out.println("File modified successfully.");
								break;
							}
						//	Change the user name
						case 4:
							{
								Cell column = row.getCell(22);
								String usernameNew = null;
								if (column == null)
									column = row.createCell(22);
								column.setCellType(Cell.CELL_TYPE_STRING);
								Scanner keyboard = new Scanner(System.in);
								char signalUser = 'F';
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the new user name: ");
									usernameNew = keyboard.nextLine();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.nextLine().charAt(0);
								}
								column.setCellValue(usernameNew);
								signalUser = 'F';
								//	Write the file out
								FileOutputStream fileOut = new FileOutputStream("D:\\Documents\\ClientListMaster.xlsx");
								wb.write(fileOut);
								fileOut.close();
								System.out.println("The user name is now: " + usernameNew + ".");
								System.out.println("File modified successfully.");
								break;
							}
						//	Change the employee ID
						case 5:
							{
								Cell column = row.getCell(23);
								long employeeID = 0;
								if (column == null)
									column = row.createCell(23);
								column.setCellType(Cell.CELL_TYPE_NUMERIC);
								String usernameNew;
								Scanner keyboard = new Scanner(System.in);
								char signalUser = 'F';
								while (Character.toLowerCase(signalUser) == 'f' || signalUser == 'f')
								{
									System.out.println("Please enter the new employee ID: ");
									employeeID = keyboard.nextLong();
									System.out.println("Enter T (confirm) or F (deny) to continue: ");
									signalUser = keyboard.next().charAt(0);
								}
								column.setCellValue(employeeID);
								signalUser = 'F';
								//	Write the file out
								FileOutputStream fileOut = new FileOutputStream("D:\\Documents\\ClientListMaster.xlsx");
								wb.write(fileOut);
								fileOut.close();
								System.out.println("The employee ID is now: " + employeeID + ".");
								System.out.println("File modified successfully.");
								break;
							}
						default:
							{
								System.err.println("Please enter a valid input: 1, 2, 3, 4, or 5!");
							}
					}
				}
			}
			//	If the client cannot be found in the system
			else
				System.err.println("Sorry. The client you are looking for is not in our system.");
	}
	
	

	public static void main(String[] args) throws FileNotFoundException, IOException, ParseException
	{
		// TODO Auto-generated method stub
		String title = "EEUS Client Master Beta v1.6";
		systemTitle(title);
	    
		// The following strings are lines of the interface
		String organization = "Enhanced End User Support: WELCOME!";
		String instructionTitle = "PLease choose one of the following: ";
		String optionOne = "1. Add a client";
		String optionTwo = "2. Remove a client";
		String optionThree = "3. Edit / Modify";
		String optionFour = "4. Search a client";
		String optionFive = "5. Support";
		String optionSix = "6. Quit";
		int numSpaces = textCenteringReturn (optionOne);

		// The invokings of the functions only show the interface content in a specifically "beautiful" way
		textCentering (organization);
		textCentering (instructionTitle);
		System.out.print("\n\n");
		textCentering (optionOne);
		for (int i = 0; i < numSpaces; ++i)
			System.out.print(" ");
		System.out.println(optionTwo);
		for (int j = 0; j <numSpaces; ++j)
			System.out.print(" ");
		System.out.println(optionThree);
		for (int k = 0; k < numSpaces; ++k)
			System.out.print(" ");
		System.out.println(optionFour);
		for (int l = 0; l < numSpaces; ++l)
			System.out.print(" ");
		System.out.println(optionFive);
		for (int m = 0; m < numSpaces; ++m)
			System.out.print(" ");
		System.out.println(optionSix);
		
		int input = 0;
		Scanner keyboard = new Scanner(System.in);
		input = keyboard.nextInt();
		
		switch(input)
	    {	
	    	case 1:
	    		{
	    			String name;
	    			System.out.print("Please enter the name of the client that you would like to add:");
	    			Scanner nameInput = new Scanner(System.in);
	    			name = nameInput.nextLine();
	    			clientAdd(name);
	    			break;
				}
			case 2:
				{
					String name;
					System.out.print("Please enter the name of the client that you would like to remove:");
					Scanner nameInput = new Scanner(System.in);
	    			name = nameInput.nextLine();
					clientRemove(name);
					break;
				}
				//	Edit / Modify
			case 3:
				{
					String name;
					System.out.print("Please enter the full name of the client: ");
					Scanner nameInput = new Scanner(System.in);
	    			name = nameInput.nextLine();
					fileModify(name);
					System.exit(0);
				}

			case 4:
				{
					String name;
					System.out.print("Please enter the name of the client that you would like to search:");
					Scanner nameInput = new Scanner(System.in);
	    			name = nameInput.nextLine();
					clientSearch(name);
					System.exit(0);
				}

	        case 5:
	            support();
	            break;
	        		//	Quit
			case 6:
		    	{
			    	String signalSuccess = "Thanks for using EEUS Client Master!";
			    	String signalBye = "See you next time!";
			    	textCentering(signalSuccess);
			    	textCentering(signalBye);
			    	System.exit(0);
		    	}

			//	Output when there is no correct input
			default: 
				System.err.println("Please enter 1, 2, 3, 4, 5, or 6!");;
				break;
	    }
		System.exit(0);
	}

}
