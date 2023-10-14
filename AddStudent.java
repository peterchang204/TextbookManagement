
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
import java.util.Arrays;
public class AddStudent{
    String lastName;
    String firstName;
    int lastRow;
    int numStudents;
    int p;
    String cellVal;
    String[] firstnameArray;
    String[] lastnameArray;
    String[] allFirstname;
    String[] allLastname;
    double[] lockerNum;
    public void addStudentMethod(String year1, String year2, String gradeInput){
        
        //confirms to proceed
        
        Scanner s1 = new Scanner(System.in);
        System.out.println("------------------");
        System.out.println("1. Continue.");
        System.out.println("2. Go back.");
        int choice = s1.nextInt();
        System.out.println("------------------");
        
                switch(choice){
                //FOR LOOP FOR AMMOUNT OF STUDENTS
                case 1:
                System.out.println("");
                System.out.println("How many students would you like to add?");
                numStudents = s1.nextInt();
                lastnameArray = new String[numStudents];
                firstnameArray = new String[numStudents];
                lockerNum = new double[numStudents];
                
                        try{
                            String filename = "bookstore"+year1+year2+".xlsx";
                            FileInputStream inputStream = new FileInputStream(new File(filename));
                            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
                            XSSFSheet sheet = wb.getSheet(gradeInput);
                            int rowCount = sheet.getPhysicalNumberOfRows();
                            
                            
                            for(int i=0;i<rowCount;i++){
                                p=0;
                                XSSFRow row = sheet.getRow(i);
                                if(row!=null){
                                XSSFCell cell = row.getCell(1);
                                try{
                                cellVal = cell.getStringCellValue();
                                }catch(Exception e){
                    
                                }
                                if(cellVal==""||cellVal==null||rowCount-1==i){
                                    p++;
                                }
                                if(p==1){
                                    lastRow=i;
                                    break;
                                }
                                }
                            }
                            
                            
                            for(int i=0;i<numStudents;i++){
                                System.out.println("");
                                System.out.println("-------");
                                System.out.println("");
                                System.out.println("What is the student's first name?");
                                firstnameArray[i] = s1.next();
                                System.out.println("");
                                System.out.println("What is the student's last name?");
                                System.out.println("");
                                lastnameArray[i] = s1.next();
                                System.out.println("What is the student's locker number?");
                                System.out.println("");
                                lockerNum[i] = s1.nextDouble();
                                System.out.println("-------");
                                System.out.println("");
                            }
                            FileOutputStream outputStream = new FileOutputStream(filename);
                            wb.write(outputStream);
                            
                        }catch(Exception e){
                            e.printStackTrace();
                        }
                
                break;
                
                
                //----------------------------------------
                //----------------------------------------
                
                case 2:
                break;
                default:
                System.out.println("");
                System.out.println("Invalid input.");
                System.out.println("");
            }
    }
    
    
    //SHORTING THE SHEET
    
    public void sortSheet(String year1, String year2, String gradeInput){
        Scanner s1 = new Scanner(System.in);
        int d = 0;
        try{
            String filename = "bookstore"+year1+year2+".xlsx";
            FileInputStream inputStream = new FileInputStream(new File(filename));
            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = wb.getSheet(gradeInput);
            allLastname = new String[lastRow+lastnameArray.length];
            allFirstname = new String[lastRow];
            for(int i=0;i<=lastRow;i++){
                XSSFRow row = sheet.getRow(i);
                XSSFCell lastName = row.getCell(1);
                String lastNameGet = lastName.getStringCellValue();
                allLastname[i] = lastNameGet;
            }
            
            
            for(int k=lastRow;k<=lastRow+lastnameArray.length;k++){
                for(int i=d;i<lastnameArray.length;i++){
                    allLastname[k] = lastnameArray[i];
                    break;
                }
                d++;
            }
            
            
            System.out.println(allLastname.length);
            Arrays.sort(allLastname);
            
            
            for(int i=0;i<lastnameArray.length;i++){
                for(int k=0;k<allLastname.length;k++){
                    if(lastnameArray[i].equalsIgnoreCase(allLastname[k])){
                        if(k!=allLastname.length-1){
                        sheet.shiftRows(k+1, allLastname.length, 1, true, true);
                        }
                        XSSFRow row = sheet.createRow(k+1);
                        row.createCell(2).setCellValue(firstnameArray[i]);
                        row.createCell(1).setCellValue(lastnameArray[i]);
                        row.createCell(3).setCellValue(lockerNum[i]);
                        System.out.println(firstnameArray[i]+" "+lastnameArray[i]+" Lockernumber: "+lockerNum[i]+" has been added.");
                        break;
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