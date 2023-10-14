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
public class CreateExcelObject{
    String firstname;
    String lastname;
    double lockerNum;
    String ISBNnum;
    String ISBNHeader;
    String[] ISBNArray;
    public void database(){
        
        //asks personal info
        Scanner s1 = new Scanner(System.in);
         System.out.println("");
        System.out.println("What is the student's lastname?");
        lastname = s1.next();
         System.out.println("");
        System.out.println("What is the student's first name?");
        firstname = s1.next();
         System.out.println("");
        System.out.println("What is the student's locker number?");
        lockerNum = s1.nextDouble();
         System.out.println("");
    }
    public void ISBN(int numBooks,String year1,String year2, String gradeInput){
        Scanner s1 = new Scanner(System.in);
        ISBNArray = new String[numBooks];
        
        //ISBN CHECKER
        try{
            String filename = "Cbookstore"+year1+year2+".xlsx";
            FileInputStream inputStream = new FileInputStream(new File(filename));
            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = wb.getSheet(gradeInput);
            XSSFRow row = sheet.getRow(0);
            for(int i=4;i<numBooks+4;i++){
                XSSFCell cell = row.getCell(i);
                String cellVal = cell.getStringCellValue();
                System.out.println(cellVal+":");
                ISBNArray[i-4] = s1.next();
            }
        }catch(Exception e){
            e.printStackTrace();
        }
    }
}