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
public class AddTextbook{ //assigning new textbook to student 
    public void addBook(String year1, String year2, String gradeInput){
           Scanner input = new Scanner(System.in);
           String ISBN = "";
           int cellCount = 0;
           int cellNumber = 0;
           int rowCount;
            int numTextbooks;
            String[] textbooks;
            int cellCount2;
            int l = 0;
           int g = 0;
           String[] subjects;
           String[] last;
           String[] first;
           String firstNameGet="";
            String lastNameGet="";
           try{
            String filename = "bookstore"+year1+year2+".xlsx";
            FileInputStream inputStream = new FileInputStream(new File(filename));
            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = wb.getSheet(gradeInput);
            rowCount = sheet.getPhysicalNumberOfRows();
            
                //ASKS USER TO CONFIRM
                System.out.println("");
                System.out.println("-------------------------------");
                System.out.println("Choose an option");
                System.out.println("1. Assign extra textbooks to student (this book will go under 'Extra Books' Column)");
                System.out.println("2. Assign textbooks by subject.");
                System.out.println("3. Go back.");
                System.out.println("-------------------------------");
                System.out.println("");
                int choice = input.nextInt();
                
                switch(choice){
                    //NEW TEXTBOOK TO STUDENT 
                        case 1: 
                        System.out.println("");
                        System.out.println("How many students would you like to assign a textbook to?");
                        int numStudents = input.nextInt();
                        System.out.println("How many textbooks do you wish to assign?");
                        int assignNumberTextbooks = input.nextInt();
                        System.out.println("");
                        
                    //RUNS FOR LOOP FOR AMMOUNT OF STUDENTS
                        for(int i=0;i<numStudents;i++){
                            System.out.println("");
                            System.out.println("---------------------------------------------------------------");
                            System.out.println("Type the name of the student you wish to assign a new book to.");
                            System.out.println("First Name: ");
                            String firstName = input.next();
                            System.out.println("Last Name: ");
                            String lastName = input.next();
                            System.out.println("---------------------------------------------------------------");
                            System.out.println("");
                            for(int j=1;j<rowCount;j++){ //trying to find the student
                                    XSSFRow row2 = sheet.getRow(0);
                                    XSSFRow row3 = sheet.getRow(j);
                                    int cellNum = row2.getPhysicalNumberOfCells();
                                    
                                    XSSFCell lastname = row3.getCell(1);
                                    String lastnameGet = lastname.getStringCellValue();
                                        if(lastnameGet.equalsIgnoreCase(lastName)){
                                            for(int p=4;p<=cellNum;p++){
                                                XSSFCell cellVal = row2.getCell(p);
                                                String cellString = cellVal.getStringCellValue();
                                                if(cellString.equalsIgnoreCase("Extra Books")){
                                                    for(int a=p;a<=cellNum;a++){
                                                        XSSFRow row1 = sheet.getRow(j);
                                                        XSSFCell cellVal2 = row1.getCell(a);
                                                if(cellVal2==null){
                                                cellNumber = a;
                                                XSSFRow test = sheet.getRow(2);
                                                XSSFCell test1 = test.getCell(6);
                                                break;
                                                }else{
                                                    String cellString2 = cellVal2.getStringCellValue();
                                                    int length = cellString2.length();
                                                    if(length!=8){
                                                        cellNumber = a;
                                                        break;
                                                    }
                                                }
                                            }
                                            }
                                        }
                                            //assigning book to student
                                            //Write in Excel infront of the cell 
                                                for(int k=0;k<assignNumberTextbooks;k++){
                                                    
                                                    
                                                    System.out.println("");
                                                    System.out.println("Please enter the ISBN for the book");
                                                    ISBN = input.next();
                                                    row3.createCell(cellNumber+k).setCellValue(ISBN);
                                                    System.out.println("The following book has been assigned to the following student");
                                                    System.out.println(" ");
                                                    System.out.println("Book: " + ISBN);
                                                    System.out.println("Student: " + firstName + " "+ lastName); //just showing what book got assigned again to whom
                                                    System.out.println("");
                                                }
                                        }
                                        
                            }
                            
                    }
                    break;
                    
                    //QUIT
                    case 2:
                        int numSheets = wb.getNumberOfSheets();
                        System.out.println("");
                        System.out.println("How many textbooks would you like to add?");
                        System.out.println("");
                        numTextbooks = input.nextInt();
                        textbooks = new String[numTextbooks];
                        subjects = new String[numTextbooks];
                        first = new String[numTextbooks];
                        last = new String[numTextbooks];
                            for(int i=0;i<numTextbooks;i++){
                                System.out.println("What is the first name of the student?");
                                first[i] = input.next();
                                System.out.println();
                                System.out.println("What is the last name of the student?");
                                last[i] = input.next();
                                System.out.println("");
                                System.out.println("What is the ISBN number?");
                                textbooks[i] = input.next();
                                System.out.println();
                                System.out.println("What is the subject?");
                                subjects[i] = input.next();
                            }
                                rowCount = sheet.getPhysicalNumberOfRows();
                                for(int k=0;k<rowCount;k++){
                                    XSSFRow row = sheet.getRow(0);
                                    XSSFRow row2 = sheet.getRow(k);
                                    cellCount2 = row.getPhysicalNumberOfCells();
                                    for(int m=4;m<=cellCount2;m++){
                                        XSSFCell cell = row.getCell(m);
                                        XSSFCell lastname = row2.getCell(1);
                                        XSSFCell firstname = row2.getCell(2);
                                        String lastGet = lastname.getStringCellValue();
                                        String firstGet = firstname.getStringCellValue();
                                        firstNameGet = firstGet;
                                        lastNameGet = lastGet;
                                        if(cell!=null){
                                            String cellGet = cell.getStringCellValue();
                                            for(int j=0;j<numTextbooks;j++){
                                                if(subjects[j].equalsIgnoreCase(cellGet)){
                                                    l++;
                                                    for(int a=0;a<numTextbooks;a++){
                                                    if(lastNameGet.equalsIgnoreCase(last[a])&&firstNameGet.equalsIgnoreCase(first[a])){
                                                    
                                                    row2.createCell(m).setCellValue(textbooks[a]);
                                                    XSSFCell lockerNum = row2.getCell(3);
                                                    XSSFRow firstRow = sheet.getRow(0);
                                                    XSSFCell subject = firstRow.getCell(m);
                                                    String subjectGet = subject.getStringCellValue();
                                                    double lockerGet = lockerNum.getNumericCellValue();
                                                    
                                                    System.out.println("");
                                                    System.out.println("Student: "+firstNameGet+" "+lastNameGet+" Locker Number: "+lockerGet);
                                                    System.out.println("");
                                                    System.out.println("You added "+textbooks[a]+" to "+cellGet);
                                                }
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
                            break;
                            
                    case 3:
                    break;
                    default:
                    System.out.println("Invalid input.");   
                    break;    
                }
            FileOutputStream outputStream = new FileOutputStream(filename);
            wb.write(outputStream);  
        }catch(Exception e){
            e.printStackTrace();
        }
    }
}