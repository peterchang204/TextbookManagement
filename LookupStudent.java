
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
public class LookupStudent {
    public void lookupStud(String year1, String year2, String gradeInput){
        Scanner s1 = new Scanner(System.in);
        String lastname = "";
        String firstname = "";
        String cellVal = "";
        int lastRow = 0;
        int cellCount=0;
        int cellNumber=0;
        int p = 0;
        int j = 0;
        try{
                String filename = "bookstore"+year1+year2+".xlsx";
                FileInputStream inputStream = new FileInputStream(new File(filename));
                XSSFWorkbook wb = new XSSFWorkbook(inputStream);
                XSSFSheet sheet = wb.getSheet(gradeInput);
                int rowCount = sheet.getPhysicalNumberOfRows();
                while(true){
                        
                    //CONFIRMS TO CONTINUE
                    System.out.println("");
                        System.out.println("-------------------------");
                        System.out.println("Choose an option");
                        System.out.println("1. Look up student.");
                        System.out.println("2. Go back.");
                        int decision1 = s1.nextInt();
                         System.out.println("-------------------------");
                          System.out.println("");
                          
                        switch(decision1){
                                    
                            case 1:
                            //CHECKS EXCEL
                            
                                    for(int i=0;i<rowCount;i++){
                                        p=0;
                                        XSSFRow row = sheet.getRow(i);
                                        if(row!=null){
                                            XSSFCell cell = row.getCell(1);
                                            cellVal = cell.getStringCellValue();
                                            if(cellVal==null||cellVal==""||rowCount-1==i){
                                                p++;
                                            }
                                            if(p==1){
                                                lastRow=i;
                                                break;
                                            }
                                        }
                                    }
                             System.out.println("-------------------------");
                              System.out.println("");
                              
                              
                            System.out.println("How many students would you like to look up?");
                            int numberStudents = s1.nextInt();
                            
                            //RUNS FOR LOOP TO ASK EACH STUDENT THIER PERSONAL INFO
                            
                                for(int i=0;i<numberStudents;i++){
                                    System.out.println("");
                                    System.out.println("What is the lastname of the student?");
                                    lastname = s1.next();
                                    System.out.println("");
                                    System.out.println("What is the firstname of the student?");
                                    firstname = s1.next();
                                    
                                    //CHECKS EXCEL 
                                    for(int k=1;k<=lastRow;k++){
                                        XSSFRow row = sheet.getRow(k);
                                        XSSFCell last = row.getCell(1);
                                        String lastGet = last.getStringCellValue();
                                        XSSFCell first = row.getCell(2);
                                        String firstGet = first.getStringCellValue();
                                        XSSFCell lockerNum = row.getCell(3);
                                        double lockerNumVal = lockerNum.getNumericCellValue();
                                        int cellNum = row.getPhysicalNumberOfCells();
                                        if(lastname.equalsIgnoreCase(lastGet)&&firstname.equalsIgnoreCase(firstGet)){
                                            System.out.println(firstname+" "+lastname+" Locker Number: "+lockerNumVal+":");
                                            j++;
                                            for(int m=4;m<100;m++){
                                                XSSFCell ISBN = row.getCell(m);
                                                if(ISBN!=null){
                                                    String ISBNget = ISBN.getStringCellValue();
                                                    XSSFRow header = sheet.getRow(0);
                                                    XSSFCell headerVal = header.getCell(m);
                                                    if(headerVal!=null){
                                                        String headerGet = headerVal.getStringCellValue();
                                                        
                                                        if(ISBNget!=null||ISBNget!=""){
                                                        System.out.println();
                                                        System.out.println(headerGet+": "+ISBNget);
                                                        }
                                                        
                                                    }
                                                
                                                }
                                            }
                                        }
                                    }
                                    if(j==0){
                                        System.out.println();
                                        System.out.println("You entered a wrong first or last name. Please try again.");
                                    }
                                }
                            break;
                            case 2:
                            break;
                            default:
                            break;
                        }
                        if(decision1==2){
                            break;
                        }
                }
        }catch(Exception e){
            e.printStackTrace();
        }
    }
}
