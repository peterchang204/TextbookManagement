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
public class RemoveTextbook {
    int rowCount;
    int numTextbooks;
    String[] textbooks;
    int cellCount;
    int l = 0;
    public void removeBook(String year1, String year2){
        Scanner s1 = new Scanner(System.in);
        System.out.println("");
        System.out.println("------------------------------");
        System.out.println("1. Continue");
        System.out.println("2. Go back");
        int choice = s1.nextInt();
        System.out.println("------------------------------");
        switch(choice){
        case 1:
            //WILL MAKE ORANGE IN EXCEL
                try{
                    String filename = "bookstore"+year1+year2+".xlsx";
                    FileInputStream inputStream = new FileInputStream(new File(filename));
                    XSSFWorkbook wb = new XSSFWorkbook(inputStream);
                    XSSFSheet sheet = wb.getSheetAt(0);
                    int numSheets = wb.getNumberOfSheets();
                    System.out.println("");
                    System.out.println("How many textbooks would you like to remove? (NOTE: This is not grade specific, as it will search through the entire system for the ISBN number.)");
                    System.out.println("By removing a textbook, you are not deleting it from the system, it is simply no longer assigned to a specifc student. It has been 'checked in'.");
                    System.out.println("");
                    numTextbooks = s1.nextInt();
                    textbooks = new String[numTextbooks];
                        for(int i=0;i<numTextbooks;i++){
                            System.out.println("");
                            System.out.println("What is the ISBN number?");
                            textbooks[i] = s1.next();
                        }
                        for(int d=0;d<numSheets;d++){
                            sheet = wb.getSheetAt(d);
                            rowCount = sheet.getPhysicalNumberOfRows();
                            for(int k=0;k<rowCount;k++){
                                XSSFRow row = sheet.getRow(k);
                                cellCount = row.getPhysicalNumberOfCells();
                                for(int m=4;m<=cellCount;m++){
                                    XSSFCell cell = row.getCell(m);
                                    if(cell!=null){
                                        String cellGet = cell.getStringCellValue();
                                        for(int j=0;j<numTextbooks;j++){
                                            if(textbooks[j].equalsIgnoreCase(cellGet)){
                                                l++;
                                                XSSFCell lastname = row.getCell(1);
                                                XSSFCell firstname = row.getCell(2);
                                                XSSFCell lockerNum = row.getCell(3);
                                                XSSFRow firstRow = sheet.getRow(0);
                                                XSSFCell subject = firstRow.getCell(m);
                                                String subjectGet = subject.getStringCellValue();
                                                String lastGet = lastname.getStringCellValue();
                                                String firstGet = firstname.getStringCellValue();
                                                double lockerGet = lockerNum.getNumericCellValue();
                                                XSSFCellStyle cellStyle = wb.createCellStyle();
                                                cellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
                                                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                                cell.setCellStyle(cellStyle);
                                                System.out.println("");
                                                System.out.println("Student: "+firstGet+" "+lastGet+" Locker Number: "+lockerGet);
                                                System.out.println("");
                                                System.out.println("You removed "+subjectGet+" "+cellGet);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(l==0){
                            System.out.println();
                            System.out.println("Invalid ISBN.");
                        }
                        
                    FileOutputStream outputStream = new FileOutputStream("bookstore"+year1+year2+".xlsx");
                    wb.write(outputStream);  
                }catch(Exception e){
                    e.printStackTrace();
                }
        
        
        
        break;
        case 2:
        break;
        default:
            System.out.println("");
            System.out.println("");
        System.out.println("Invalid input");
    }
}
}