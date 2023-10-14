import org.apache.logging.log4j.core.config.LockingReliabilityStrategy;
import org.apache.poi.ss.usermodel.Cell;
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
import java.util.Arrays;
import java.awt.Color;
import org.apache.poi.ss.usermodel.*;
public class StudentStatus {
    String cellVal;
    int choice;
    XSSFColor bgColor;
    boolean color;
    public void status(String year1, String year2, String gradeInput){
        Scanner s1 = new Scanner(System.in);
        System.out.println("");
        System.out.println("---------------------------------------------------");
        System.out.println("1. Display all students who have not returned their books.");
        System.out.println("2. Go back.");
        System.out.println("---------------------------------------------------");
        choice = s1.nextInt();
         System.out.println("");
                switch(choice){
                case 1:
                        try{
                            String filename = "bookstore"+year1+year2+".xlsx";
                            FileInputStream inputStream = new FileInputStream(new File(filename));
                            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
                            XSSFSheet sheet = wb.getSheet(gradeInput);
                            int rowCount = sheet.getPhysicalNumberOfRows();
                            
                                    for(int i=1;i<=rowCount;i++){
                                    int p=0;
                                    color=false;
                                    XSSFRow row = sheet.getRow(i);
                                    int cellCount = row.getPhysicalNumberOfCells();
                                    
                                        for(int k=4;k<=cellCount;k++){
                                            color = false;
                                            XSSFCell cell = row.getCell(k);
                                                    if(cell!=null){
                                                        String cellVal = cell.getStringCellValue();
                                                        CellStyle cellStyle = cell.getCellStyle();
                                                        boolean isColored = (cellStyle.getFillForegroundColor() != IndexedColors.AUTOMATIC.getIndex());
                                                        
                                                        if (isColored) {
                                                            color = true;
                                                        }
                                                        
                                                        if(cellVal.length()==8&&color==false){
                                                            XSSFRow firstRow = sheet.getRow(0);
                                                            XSSFCell lastname = row.getCell(1);
                                                            XSSFCell firstname = row.getCell(2);
                                                            XSSFCell lockerNum = row.getCell(3);
                                                            XSSFCell subject = firstRow.getCell(k);
                                                            String subjectGet = subject.getStringCellValue();
                                                            String lastGet = lastname.getStringCellValue();
                                                            String firstGet = firstname.getStringCellValue();
                                                            double lockerGet = lockerNum.getNumericCellValue();
                                                            if(p==0&&cellVal!=""){
                                                            System.out.println();
                                                             System.out.println("");
                                                            System.out.println("Student: "+firstGet+" "+lastGet+" Locker Number: "+lockerGet);
                                                            p++;
                                                             System.out.println("");
                                                            }
                                                            if(cellVal!=""){
                                                             System.out.println("");
                                                                System.out.println("The student hasn't returned "+subjectGet+" "+cellVal);
                                                                 System.out.println("");
                                                            }
                                                        }
                                                }
                                                
                                        }
                                    
                                }
                                
                                
                                }catch(Exception e){
                    
                                }
                break;
                case 2:
                break;
                default:
                     System.out.println("");
                System.out.println("Invalid input.");
            }
    }
}

