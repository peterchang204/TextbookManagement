import java.util.Scanner;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import java.io.*;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.FileInputStream;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.FillPatternType;
import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.*;
public class RemoveStudent {
    public void removal(String year1, String year2, String gradeInput) {
        
        Scanner input = new Scanner (System.in); 
        int cellCount=0;
        int cellNumber=0;
        int q = 0;
        
        System.out.println("-----------------------------------------------");
        System.out.println("DELETING STUDENT");
        System.out.println("This action will not be physically deleting the student from the database, instead it will highlight them red.");
        System.out.println("-----------------------------------------------");
        System.out.println("Do you wish to continue?");
        System.out.println("1. Yes");
        System.out.println("");
        System.out.println("2. No");
        int choice = input.nextInt();
        try{
            
            
            
            String filename = "bookstore"+year1+year2+".xlsx";
            FileInputStream inputStream = new FileInputStream(new File(filename));
            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = wb.getSheet(gradeInput);
            int rowCount = sheet.getPhysicalNumberOfRows();
            
            switch(choice){
                
                        //CONFIRMS
                        case 1:
                        System.out.println("Thank you for confirming");
                        System.out.println("How many students do you wish to remove from the system?");
                        int numstudentremovalCA = input.nextInt();
                        System.out.println("You have selected " + numstudentremovalCA+ " students");
                        
                    
                    
                        for(int i =0; i<numstudentremovalCA; i++){
                            
                            //PERSONAL INFO
                            System.out.println("");
                            System.out.println("---------------------------------");
                            System.out.println("");
                            System.out.println("What is the student's last name?");
                            String lastName = input.next();
                            System.out.println("");
                            System.out.println("What is the student's first name?");
                            String firstName = input.next();
                            System.out.println("");    
                            
                                  
                            boolean found = false;
                            //CHECKS ALL EXCEL ROWS AND CELLS
                             
                                        for (Row row : sheet) {
                                        
                                        // Loop through each cell of the row
                                        for (Cell cell : row) {
                                            
                                            // Check if the cell contains the last name
                                            if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(lastName)) {
                                                
                                                // Check if the cell next to it contains the first name
                                                Cell firstNameCell = row.getCell(cell.getColumnIndex() + 1);
                                                if (firstNameCell != null && firstNameCell.getCellType() == CellType.STRING && firstNameCell.getStringCellValue().equalsIgnoreCase(firstName)) {
                                                    
                                                    // Set the cell background color to red
                                                    CellStyle cellStyle = wb.createCellStyle();
                                                    cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                                                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                                    cell.setCellStyle(cellStyle);
                                                    
                                                    
                                                    CellStyle firstNameCellStyle = wb.createCellStyle();
                                                    firstNameCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                                                    firstNameCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                                    firstNameCell.setCellStyle(firstNameCellStyle);
                                                    found = true ;
                                                    
                                                    System.out.println("");
                                                    System.out.println("You have just removed " +firstName + " " + lastName);
                                                    System.out.println("");
                                                }
                                            }
                                        }
                                    }
                                    
                                    if (!found) {
                                        System.out.println("");
                                        System.out.println("Invalid name!");
                                        System.out.println("");
                                    }
                                    
                                    //Write the modified workbook back to the file
                                    FileOutputStream outputStream = new FileOutputStream("bookstore"+year1+year2+".xlsx");
                                    wb.write(outputStream);                                    
                                    
                                }      
                        System.out.println("");
                        System.out.println("");
                        System.out.println("");
                        
                        case 2:
                        break;
                        
                        
                    
            }
               }catch(Exception e){
            e.printStackTrace();
        }
        }
}
        
        