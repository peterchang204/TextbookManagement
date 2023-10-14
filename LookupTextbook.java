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
public class LookupTextbook {
    public void lookupBook(String year1, String year2, String gradeInput){
        Scanner s1 = new Scanner(System.in);
        String ISBN1 = "";
        int cellCount=0;
        int cellNumber=0;
        int numberTextbooks;
        int s = 0;
        try{
            String filename = "bookstore"+year1+year2+".xlsx";
            FileInputStream inputStream = new FileInputStream(new File(filename));
            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = wb.getSheet(gradeInput);
            int rowCount = sheet.getPhysicalNumberOfRows();
            System.out.println("");
            System.out.println("----------------------------------------------------------------");
            System.out.println("Choose an option");
            System.out.println("1. Look up textbook by ISBN number in"+" "+gradeInput);
            System.out.println("2. Lookup textbook across all grades for"+" "+year1+"-"+year2);
            System.out.println("3. Go back.");
            int decision1 = s1.nextInt();
            System.out.println("----------------------------------------------------------------");
            System.out.println("");
            switch(decision1){
                        case 1:
                            //Look up textbook by ISBN
                        
                            System.out.println("");
                            System.out.println("How many textbooks would you like to look up?");
                            numberTextbooks = s1.nextInt();
                            System.out.println("");
                            for(int i=0;i<numberTextbooks;i++){
                                System.out.println("What is the ISBN number of the textbook?");
                                ISBN1 = s1.next();
                            
                                    for(int j=1;j<rowCount;j++){
                                        XSSFRow row1 = sheet.getRow(j);
                                        cellCount = row1.getPhysicalNumberOfCells();
                                        for(int k=4;k<cellCount;k++){
                                            XSSFCell cellTest = row1.getCell(k);
                                            if(cellTest==null){
                    
                                            }else{
                                            String test = cellTest.getStringCellValue();
                                            if(test.length()==8){
                                                cellNumber=k;
                                                break;
                                            }
                                        }
                                        }
                                    }
                            
                            
                            
                                    for(int d=1;d<rowCount;d++){
                                        XSSFRow row2 = sheet.getRow(d);
                                        int cellCount2 = row2.getPhysicalNumberOfCells();
                                        for(int m=cellNumber;m<=cellCount2;m++){
                                            XSSFCell cellLookup = row2.getCell(m);
                                            if(cellLookup==null){
                        
                                            }else{
                                            String cellCheck = cellLookup.getStringCellValue();
                                                if(ISBN1.equalsIgnoreCase(cellCheck)){
                                                    s++;
                                                    XSSFCell lastName = row2.getCell(1);
                                                    XSSFCell firstName = row2.getCell(2);
                                                    XSSFCell lockerNum = row2.getCell(3);
                                                    String lastNameGet = lastName.getStringCellValue();
                                                    String firstNameGet = firstName.getStringCellValue();
                                                    double lockerNumGet = lockerNum.getNumericCellValue();
                                                    System.out.println();
                                                    System.out.println(firstNameGet+" "+lastNameGet+" Locker Number: "+lockerNumGet+" has the book with the ISBN #"+ISBN1);
                                                }
                                            }
                                        }
                                    }
                                    if(s==0){
                                        System.out.println();
                                        System.out.println("ISBN not detected. Please input a valid ISBN number.");
                                        System.out.println();
                                    }
                                    
                                }
                        break;
                        
                        //----------------
                        //----------------
                        
                        
                        case 2:
                            //Lookup textbook across all grades
                            
                            int numSheets = wb.getNumberOfSheets();
                            System.out.println("");
                            System.out.println("---------------------------------------------");
                            System.out.println("How many textbooks would you like to look up?");
                            numberTextbooks = s1.nextInt();
                            System.out.println("");
                            for(int i=0;i<numberTextbooks;i++){
                                System.out.println("");
                                System.out.println("What is the ISBN number of the textbook?");
                                ISBN1 = s1.next();
                                System.out.println("");
                            
                                            for(int d=0;d<numSheets;d++){
                                                sheet = wb.getSheetAt(d);
                                                rowCount = sheet.getPhysicalNumberOfRows();
                                                
                                                    for(int j=1;j<rowCount;j++){
                                                        XSSFRow row1 = sheet.getRow(j);
                                                        cellCount = row1.getPhysicalNumberOfCells();
                                                        for(int k=4;k<cellCount;k++){
                                                            XSSFCell cellTest = row1.getCell(k);
                                                                if(cellTest==null){
                                        
                                                                }else{
                                                                String test = cellTest.getStringCellValue();
                                                                if(test.length()==8){
                                                                    cellNumber=k;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    
                                                    for(int p=1;p<rowCount;p++){
                                                        XSSFRow row2 = sheet.getRow(p);
                                                        int cellCount2 = row2.getPhysicalNumberOfCells();
                                                        for(int m=cellNumber;m<=cellCount2;m++){
                                                            XSSFCell cellLookup = row2.getCell(m);
                                                            if(cellLookup==null){
                                        
                                                            }else{
                                                                
                                                                String cellCheck = cellLookup.getStringCellValue();
                                                            
                                                                if(ISBN1.equalsIgnoreCase(cellCheck)){
                                                                    s++;
                                                                    XSSFCell lastName = row2.getCell(1);
                                                                    XSSFCell firstName = row2.getCell(2);
                                                                    XSSFCell lockerNum = row2.getCell(3);
                                                                    String lastNameGet = lastName.getStringCellValue();
                                                                    String firstNameGet = firstName.getStringCellValue();
                                                                    double lockerNumGet = lockerNum.getNumericCellValue();
                                                                    System.out.println(firstNameGet+" "+lastNameGet+" Locker Number: "+lockerNumGet+" has the book with the ISBN #"+ISBN1);
                                                                    s++;
                                                                    break;
                                                                }else{
                                    
                                                                }
                                                            }
                                                        }
                                                    }
                                                    
                                        }
                                        if(s==0){
                                            System.out.println();
                                            System.out.println("ISBN not detected. Please input a valid ISBN number.");
                                            System.out.println();
                                        }
                            }
                        break;
                        
                        //----------------
                        //----------------
                        
                        //EXIT
                        case 3:
                        break;
                        default:
                        break;
            }
        }catch(Exception e){
            e.printStackTrace();
        }
    }
}
