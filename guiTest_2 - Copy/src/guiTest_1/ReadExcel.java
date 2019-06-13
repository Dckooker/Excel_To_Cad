package guiTest_1;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcel {
	
	FileInputStream fis = null;
	XSSFWorkbook wb = null;    
	XSSFSheet sh = null;
	XSSFRow row;
	XSSFCell cell;
	
	//constructor for getting a file to read
	public ReadExcel(String filepath) {
		
		try {
			fis = new FileInputStream(filepath);
			wb = new XSSFWorkbook(fis);
		}
		catch(IOException e) {
			e.printStackTrace();
		}
	}
	
	
	
	//gets number of rows in file
	public int getRowCount(String filename) {
		sh = wb.getSheet(filename);
		int rows = 0;
		try {
		for(int i=0; i< sh.getLastRowNum();i++) {
			if(sh.getRow(i).getCell(0).toString() != "") {
				rows++;
			}
		  }
		}
        catch(NullPointerException f) {}
		
		return rows;
		}
		
	
	
	//gets number of cells in file
		public int getCellCount(String filename) {
			sh = wb.getSheet(filename);
			row = sh.getRow(0);
		    int columns = row.getLastCellNum();
		    return columns;
		}
		// uses cell coordinates////////////////////////////////////////////////////////////////
		public String getData(String filename, int columns, int rows) {
			sh = wb.getSheet(filename);
			row = sh.getRow(rows);
			cell = row.getCell(columns);
			
			
			
			if(cell.getCellType() == CellType.STRING) {
					return cell.getStringCellValue();
			}
			else if(cell.getCellType() == CellType.NUMERIC) {
				return String.valueOf(cell.getNumericCellValue());
			}
			else if(cell.getCellType() == CellType.BOOLEAN) {
				return String.valueOf(cell.getBooleanCellValue());
			}
			else {
				return "";
			}
		}
	
	
        //uses column name///////////////////////////////////////////////////////////////////////
	    public String getData(String filename, String colName, int rows) {
	    	int colNum = -1;
	    	sh = wb.getSheet(filename);
	    	row = sh.getRow(0);
		
	    	for(int i = 0; i< row.getLastCellNum(); i++) {
	    		if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName.trim())) {
	    			colNum = i;
	    		}
	    	}
		
	    	row = sh.getRow(rows-1);
	    	cell = row.getCell(colNum);
		
			
			if(cell.getCellType() == CellType.STRING) {
				return cell.getStringCellValue();
			}
			 if(cell.getCellType() == CellType.NUMERIC ) {
				return String.valueOf(cell.getNumericCellValue());
			}
			 if(cell.getCellType() == CellType.BOOLEAN) {
				return String.valueOf(cell.getBooleanCellValue());
			}
			else {
				return "";
			}
			
			
		}
	
	//////////////////////////////////////////////////////////////////////////////////
	 
	    
	 public void printData(ReadExcel exl) {
		 
		 int row = exl.getRowCount("Nameplates");
		 int cell = exl.getCellCount("Nameplates");
		 System.out.println(" Total rows: " +row + " Total columns: " +cell);
		 System.out.println(" ");
		 
		 System.out.println("Panel:   NP #:    Line 1:                   Line 2:             "
		         + "      Line 3:                   Line 4:              Size L1:  Size L2:  Size L3:  Size L4:  Height:  Width:  ");
		 for(int i =2; i <= row;i++) {
	
			try {
				System.out.printf("%-8s %-8s %-25s %-25s %-25s %-20s %-9s %-9s %-9s %-9s %-7s  %s",
						exl.getData("Nameplates", "Panel", i), 
						exl.getData("Nameplates", "NP #", i),
				        exl.getData("Nameplates", "Line 1", i),
				        exl.getData("Nameplates", "Line 2", i),
				        exl.getData("Nameplates", "Line 3", i),
				        exl.getData("Nameplates", "Line 4", i),
			            exl.getData("Nameplates", "Size L1", i),
				        exl.getData("Nameplates", "Size L2", i),
				        exl.getData("Nameplates", "Size L3", i), 
				        exl.getData("Nameplates", "Size L4", i),
				        exl.getData("Nameplates", "Height", i),
				        exl.getData("Nameplates", "Width", i));
						System.out.println("");
			}
			catch(NullPointerException e) { }
		
		 }
	 }
	 
	 public void exitProgram() {
		 System.exit(0);
	 }
	 
	 public void makeScriptFile() {
			File textFile = new File("NPtext.txt");
			File scriptFile = new File("X:\\DKooker\\Programs\\eclipseEXE's\\NPtext.scr");
			textFile.renameTo(scriptFile);
		}
	 
	 public ArrayList<Integer>  spacingAUTO(ReadExcel exl){
		 int row = exl.getRowCount("Nameplates");
		 ArrayList<Integer> linespacing = new ArrayList<Integer>();
		 for(int i =2; i <= row;i++) {
			  double lineNum = Integer.parseInt(exl.getData("Nameplates", "Height", i)) * 1.5;
			  		if(lineNum != 0) {
				   linespacing.add(i*10000);
			  		}
			 }
			return linespacing;
		 }
	 
	 public ArrayList<Integer> spacingMANUAL(ReadExcel exl){
		 int row = exl.getRowCount("LineSpacing");
	      ArrayList<Integer> linespacing = new ArrayList<Integer>();
		 for(int i =2; i <= row;i++) {
		  double lineNum = Integer.parseInt(exl.getData("LineSpacing", "Line #", i));
		  		if(lineNum != 0) {
		  			linespacing.add(i*10000);
		  		}
		 }
		return linespacing;
	 }
	 
	 ////////////////////////////////////////////////////////////////////////////////
	 public int getLineCount(ReadExcel exl) {
		 int totalCount = 0;
		 int col = exl.getCellCount("Nameplates");
		 int row = exl.getRowCount("Nameplates");
		 for(int i =0; i < col ;i++) {
			 for(int j =1; j< row; j++) {
		 	 if(exl.getData("Nameplates", i, 0).contains("Line") && exl.getData("Nameplates", i, j) != "") {
		 		 totalCount++;
		 	 	}
		 	 } 
		 }
		return totalCount;
	 }
	 
	 
	 
	 public ArrayList<String> getCadStrings(int hor, int vert, double height, boolean spacing,ReadExcel exl ) {
		 ArrayList<String> finalList = new ArrayList<String>();
		 String finalString = "";
		 int vert2 = 0;
		 int row = exl.getRowCount("Nameplates"); 
		 double v = 0;
		 int lineSize;
		 
		 int lineCount = exl.getLineCount(exl);
		 for(int k =2; k < row+1;k++) {
		  v = Double.parseDouble( exl.getData("Nameplates", "Size L1", k))*10000;
		 }
		 
		 if(lineCount % 2 == 0) {
			 v = v *0.75;
		 } 
		 if(lineCount % 2 == 1) {
			 v = v *1.5;
		 }
		 
		 
		 
		 //single line only
		 if( lineCount == 1) {
			 if (spacing == true) {
			     vert2 = (int) (5000 * height - v + vert)+1;
			 }
			 else{
				 vert2 = (int) (exl.spacingMANUAL(exl).get(1) * height + vert)+1; 
			 }
			 
		     lineSize  = (int)(Double.parseDouble(exl.getData("Nameplates", "Size L1" ,2))*10000);
			 String currentLine = exl.getData("Nameplates", "Line 1" , 2);
			 finalString = "text j mc " + hor + ",-" + Integer.toString(vert2) + " " + lineSize + " 0 " + currentLine;     
			 finalList.add(finalString);
			
		 }
		 
	     //2 lines
		 if(lineCount == 2) {
			 for(int i = 1; i <= lineCount; i++) {
				
				  if(spacing == true && i == 1 ) {
					  vert2 = (int) (5000 * height - v + vert)+1;
				 }
				  else if( spacing == true && i == 2) {
					  vert2 = (int)(5000 * height + v + vert);
				  }
				 else { 
					  vert2 = (int) (exl.spacingMANUAL(exl).get(i) * height + vert)+100000;
				 }
				  
				  if(exl.getData("Nameplates", "Size L" + i ,2)== ""){
						 lineSize = (int)(Double.parseDouble(exl.getData("Nameplates", "Size L1" ,2))*10000); 
					 }
					 else {
						 lineSize  = (int)(Double.parseDouble(exl.getData("Nameplates", "Size L" + i ,2))*10000);
					 }
					String currentLine = exl.getData("Nameplates", "Line " +i, 2);
					 finalString = "text j mc " + hor + ",-" + Integer.toString(vert2) + " " + lineSize + " 0 " + currentLine;     
					 finalList.add(finalString);
					
				 	}
		 }
		 //3rd line	 
		 if(lineCount == 3) {
			 for(int i = 1; i <= lineCount; i++) {
				
				 if(spacing == true && i == 1 ) {
					  vert2 = (int) (5000 * height - v + vert)+1;
				 }
				 else if( spacing == true && i == 2) {
					  vert2 = (int)(5000 * height + v + vert)+1;
				 }
				 else if( spacing == true && i == 2) {
					  vert2 = (int)(5000 * height + v + vert)+1;
				 }
				 else { 
					  vert2 = (int) (exl.spacingMANUAL(exl).get(i) * height + vert)+1;
				 }
				 
				 if(exl.getData("Nameplates", "Size L" + i ,2)== ""){
					 lineSize = (int)(Double.parseDouble(exl.getData("Nameplates", "Size L1" ,2))*10000); 
				 }
				 else {
					 lineSize  = (int)(Double.parseDouble(exl.getData("Nameplates", "Size L" + i ,2))*10000);
				 }
				String currentLine = exl.getData("Nameplates", "Line " +i, 2);
				 finalString = "text j mc " + hor + ",-" + Integer.toString(vert2) + " " + lineSize + " 0 " + currentLine;     
				 finalList.add(finalString);
			 }
		 }
		
		//4th line 
		if(lineCount == 4) {
			for(int i = 1; i <= lineCount; i++) {
					
				if(spacing == true && i == 1 ) {
					  vert2 = (int) (5000 * height - v + vert)+1;
			    }
				else if( spacing == true && i == 2) {
					  vert2 = (int)(5000 * height + v + vert)+1;
				}
				else if( spacing == true && i == 2) {
					  vert2 = (int)(5000 * height + v + vert)+1;
				}
				else if( spacing == true && i == 2) {
					  vert2 = (int)(5000 * height + v + vert)+1;
				}
				else { 
					  vert2 = (int) (exl.spacingMANUAL(exl).get(i) * height + vert)+1;
				}
				
				if(exl.getData("Nameplates", "Size L" + i ,2)== ""){
					 lineSize = (int)(Double.parseDouble(exl.getData("Nameplates", "Size L1" ,2))*10000); 
				 }
				 else {
					 lineSize  = (int)(Double.parseDouble(exl.getData("Nameplates", "Size L" + i ,2))*10000);
				 }
				String currentLine = exl.getData("Nameplates", "Line " +i, 2);
				 finalString = "text j mc " + hor + ",-" + Integer.toString(vert2) + " " + lineSize + " 0 " + currentLine;     
				 finalList.add(finalString);
			}
		}
		
		//5th line
		if(lineCount == 5) {
			for(int i = 1; i <= lineCount; i++) {
					
				if(spacing == true && i == 1 ) {
					  vert2 = (int) (5000 * height - v + vert)+1;
			    }
				else if( spacing == true && i == 2) {
					  vert2 = (int)(5000 * height + v + vert)+1;
				}
				else if( spacing == true && i == 2) {
					  vert2 = (int)(5000 * height + v + vert)+1;
				}
				else if( spacing == true && i == 2) {
					  vert2 = (int)(5000 * height + v + vert)+1;
				}
				else if( spacing == true && i == 2) {
					  vert2 = (int)(5000 * height + v + vert)+1;
				}
				else { 
					  vert2 = (int) (exl.spacingMANUAL(exl).get(i) * height + vert)+1;
				}
				
				if(exl.getData("Nameplates", "Size L" + i ,2)== ""){
					 lineSize = (int)(Double.parseDouble(exl.getData("Nameplates", "Size L1" ,2))*10000); 
				 }
				 else {
					 lineSize  = (int)(Double.parseDouble(exl.getData("Nameplates", "Size L" + i ,2))*10000);
				 }
				String currentLine = exl.getData("Nameplates", "Line " +i, 2);
				 finalString = "text j mc " + hor + ",-" + Integer.toString(vert2) + " " + lineSize + " 0 " + currentLine;     
				 finalList.add(finalString);
			}
		}
		 
		 
		 
		return finalList;
	 }
	 
	 
	 public ArrayList<String> getBorderLines( ReadExcel exl ) {
		 ArrayList<String> borderList = new ArrayList<String>();
		 String line = "";
		 int row = exl.getRowCount("Nameplates");
		 
		 for(int i =2; i < row+1;i++) {
	        	
				String Height=  exl.getData("Nameplates", "Height" ,i);
				String Width=  exl.getData("Nameplates", "Width"   ,i);
				
				double NPwidth = Double.parseDouble(Width);
				double NPheight = Double.parseDouble(Height);
				int panelLengthX= (int)(11/NPheight);
				int panelLengthY= (int)(23/NPwidth);
				int x1 = 0;
		        int y1 = 0;
		        int startHeight = 0;
		        
		        double lastHor = (startHeight + Math.ceil((i - 1) / panelLengthY) * NPheight * 10000);
		        //get cut 
 		    borderList.add("-layer s CUT");
 		    borderList.add(" ");
 		    
 		 //Vertical Lines
            while( x1 < (panelLengthY + 1))
            {
            	//             "line " +      x1 * NPwidth * 10000  + ",-" +      startHeight + " " +       x1 * NPwidth * 10000  + ",-" +       startHeight +   RoundUp((j - 1) / y)            * NPheight  * 10000 
               borderList.add( "line " +(int)((x1 * NPwidth) * 10000) + ",-" + (int)startHeight + " " + (int)((x1 * NPwidth) * 10000) + ",-" + ((startHeight + Math.round((i - 1) / panelLengthY) * NPheight) * 10000));                                                                                                                                                                                     
               x1 = x1 + 1;                                                                                                                                                                                                      
	        }  
       

         //Horizontal Lines
            while( y1 < (Math.ceil((i - 2) / panelLengthY)) + 2)
	        {
            //                 "line 0,-" +       startHeight + y1 * NPheight * 10000  + " " +       NPwidth *      y       * 10000  + ",-" +       startHeight + y1 * NPheight * 10000 
                borderList.add("line 0,-" + (int)((startHeight + y1 * NPheight) * 10000) + " " + (int)((NPwidth * panelLengthY) * 10000) + ",-" + (int)((startHeight + y1 * NPheight) * 10000));
                y1 = y1 + 1;
	        }

            


            startHeight = (int) ((((startHeight + (y1 - 1)) * NPheight) * 10000) + 10000);
    
            //Zoom out at end and scale everything
            borderList.add("zoom all");
            borderList.add("scale all");
            borderList.add("0,0 .0001");  
	        }
		return borderList;
	 } 
		
}
	 
	 
	 
	 
	
	 
	 
	 

    
    
   
    	

