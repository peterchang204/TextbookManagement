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


public class CreateExcel {
    String lastname;
    String firstname;
    double lockerNum;
    int numStudents;
    int numBooks;
    String ISBN;
    String lastnameArray[];
    String firstnameArray[];
    double lockerNumArray[];
    public void createSheet(String year1, String year2, String gradeInput){
        Scanner s1 = new Scanner(System.in);
        try{
            String filename = "bookstore"+year1+year2+".xlsx";
            FileInputStream inputStream = new FileInputStream(new File(filename));
            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = wb.getSheet(gradeInput);
            int rowCount = sheet.getPhysicalNumberOfRows();
            XSSFRow header = sheet.getRow(0);
            
            //INPUTS FOR FOR LOOPS
            
            System.out.println("");
            System.out.println("------------------------------------------------");
            System.out.println("How many students would you like to add?");
            numStudents = s1.nextInt();
            CreateExcelObject[] ceo = new CreateExcelObject[numStudents];
            System.out.println("");
            lastnameArray = new String[numStudents];
            System.out.println("How many books would you like to add for each student?");
            numBooks = s1.nextInt();
            System.out.println("");
            System.out.println("Please type in 'quit' if you would like to stop the program.");
            System.out.println("");
            System.out.println("------------------------------------------------");
            System.out.println("");
            
            //CREATES OBJECTS AND CHECKS EXCEL
            for(int i=0;i<numStudents;i++){
                ceo[i] = new CreateExcelObject();
                ceo[i].database();
                ceo[i].ISBN(numBooks,year1,year2,gradeInput);
                if(ceo[i].lastname=="quit"){
                    break;
                }
                lastnameArray[i] = ceo[i].lastname;
            }
            Arrays.sort(lastnameArray);
            for(int i=0;i<lastnameArray.length;i++){
                for(int k=0;k<ceo.length;k++){
                    if(lastnameArray[i].equalsIgnoreCase(ceo[k].lastname)){
                        XSSFRow row = sheet.createRow(i+1);
                        row.createCell(1).setCellValue(ceo[k].lastname);
                        row.createCell(2).setCellValue(ceo[k].firstname);
                        row.createCell(3).setCellValue(ceo[k].lockerNum);
                        for(int m=4;m<numBooks+4;m++){
                            row.createCell(m).setCellValue(ceo[k].ISBNArray[m-4]);
                        }
                    }
                }
            }
            
            FileOutputStream outputStream = new FileOutputStream(filename);
            wb.write(outputStream);
        }catch(Exception e){
            e.printStackTrace(); 
        }
    }
}
