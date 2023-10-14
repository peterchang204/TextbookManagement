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


public class MainMethod{
    public static void main(String args[]){
        Scanner input = new Scanner(System.in);
        LookupTextbook l = new LookupTextbook();
        AddStudent s = new AddStudent();
        RemoveStudent r = new RemoveStudent();
        AddTextbook at = new AddTextbook();
        CreateExcel ce = new CreateExcel();
        LookupStudent ls = new LookupStudent();
        RemoveTextbook rt = new RemoveTextbook();
        StudentStatus ss = new StudentStatus();
        AssignLockernumber al = new AssignLockernumber();
        String year1 = "";
        String year2 = "";
        String gradeInput = "";
        int choice3 = 0;
        int emergencyPin = 2005;
        int key = 0;
        //Make these changeable
        while (key!=1){
        try{
            
            //PASSWORD SYSTEM
            //------------------------------------------------------------------------
            String filename = "password.xlsx";
            FileInputStream inputStream = new FileInputStream(new File(filename));
            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row = sheet.getRow(0);
            XSSFCell usernameCell = row.getCell(0);
            XSSFCell passwordCell = row.getCell(1);
            String username = usernameCell.getStringCellValue();
            String password = passwordCell.getStringCellValue();
            //------------------------------------------------------------------------
            
            //BEGINING OF THE CODE WILL ASK TO LOGIN OR RECOVERY ACCOUNT
            System.out.println("Hello, welcome to the textbook management system!");
            System.out.println("-------------------------------------------------");
            System.out.println("What would you like to do? (enter in the number corespnding to your choice ex: 1 or 2 etc.)");
            System.out.println("Disclaimer: This program is case sensitive, so please be careful with what you type in");
            System.out.println("Discalimer: The excel sheet can not be open as the program runs, so please close all excel files before launching the program.");  
            System.out.println(" ");
            System.out.println("1. Login.");
            System.out.println("");
            System.out.println("2. Forgot username or password?");
            int choice = input.nextInt();
            
            
                switch(choice){
                    // THIS SECTION OF THE CODE WILL ASK THEW USER FOR A PASSWORD AS WELL AS A USERNAME 
                    case 1:
                            while(key!=1){
                                System.out.println("");
                                System.out.println("Username:");
                                String usernameInput = input.next();
                                System.out.println("Password:");
                                String passwordInput = input.next();
                                if(usernameInput.equalsIgnoreCase(username)&&passwordInput.equalsIgnoreCase(password)){
                                    key=1;
                                }else{
                                    System.out.println("1. Wrong username or password was typed in, try again.");
                                    System.out.println("2. Forgot password?");
                                    choice3 = input.nextInt();
                                    switch(choice3){
                                        case 1:
                                        break;
                                        case 2:
                                        break;
                                        default:
                                        System.out.println("Invalid input.");
                                    }
                                }
                                if(choice3==2){
                                    break;
                                }
                            }
                         break;
                         
                    //-------------------------------------------------------
                    //-------------------------------------------------------
                         
                    //THIS SECTION WILL ACTIVATE IF THE USER REQUESTS TO RECOVER THIER ACCOUNT    
                    case 2:
                        System.out.println("Please type in the pin.");
                        int pinInput = input.nextInt();
                        if(pinInput==emergencyPin){
                                System.out.println("1. Reset username.");
                                System.out.println("");
                                System.out.println("2. Reset password.");
                                int choice2 = input.nextInt();
                            switch(choice2){
                                case 1:
                                System.out.println("Type in the new username.");
                                row.createCell(0).setCellValue(input.next());
                                System.out.println("");
                                System.out.println("The username has been reset. Try logging in again.");
                                System.out.println("");
                                break;
                                case 2:
                                System.out.println("Type in the new password.");
                                System.out.println("");
                                row.createCell(1).setCellValue(input.next());
                                System.out.println("The password has been reset. Try logging in again.");
                                System.out.println("");
                                break;
                                default:
                                System.out.println("Invalid input.");
                                break;
                            }
                            FileOutputStream outputStream = new FileOutputStream(filename);
                            wb.write(outputStream); 
                        }
                        
                    //-------------------------------------------------------
                    //-------------------------------------------------------
                    
                    // IF USER INPUT IS NOT WHAT THE COMPUTER REQUEST IT WILL BE PUT HERE
                    
                    break;
                    default:
                    System.out.println("Invalid input.");
                    break;
                }
                
                }catch(Exception e){
                e.printStackTrace();
            }
    }
    
    
        while(true){
            
            //ASKS THE USER HOW THEY WISH TO CHANGE/INPUT DATA
            System.out.println("");
            System.out.println("--------------------------------------");
            System.out.println("WHAT WOULD YOU LIKE TO DO?");
            System.out.println("Please enter in the number corespnding to your choice ex: 1 or 2 etc.");
            System.out.println("--------------------------------------");
            System.out.println("1. CONTINUE TO EXISTING TEXTBOOK DATABASE.");
            System.out.println("2. CREATE A NEW TEXTBOOK DATABASE.");
            System.out.println("3. QUIT");
            System.out.println("");
            int decision4 = input.nextInt();
            System.out.println("-------------------------------------");
            
           
            switch(decision4){
            
            //SCHOOL YEAR (THIS IS CRITICAL AS IT WILL BE NESECARY TO ACESS A EXCEL FILE
            case 1: // CASE 2 SCOLL ALL THE WAY DOWN
            
            System.out.println("");
            System.out.println("ENTER SCHOOL YEAR YOU WISH TO WORK ON: EXAMPLE (2020-2021)");
            System.out.println("From:");
            year1 = input.next();
            System.out.println("");
            System.out.println("To:");
            year2 = input.next();
            
            //ASKS WHAT THE GRADE IS SO THAT IT CAN ACESSS THE SPEFICI SHEET BASED OFF THE LAST INPUT SECTION ^^^
            System.out.println("");
            System.out.println("----------------------------------------------------------------------------");
            System.out.println("TYPE IN WHICH GRADE YOU WISH TO WORK ON (MUST BE FORMATED AS EX: 6th or 12th)");
            gradeInput = input.next();
            System.out.println("");
            System.out.println("----------------------------------------------------------------------------");
            System.out.println("");
                                
                                while(true){
                                //STUDENTS AND BOOK MANAGEMENT SYSTEMS
                                
                                
                                System.out.println("Welcome to Student and Book(s) management.");
                                System.out.println("What would you like to do? (enter in the number corespnding to your choice ex: 1 or 2 etc.)");
                                System.out.println("1. Manage textbooks.");
                                System.out.println("2. Manage students.");
                                System.out.println("3. Check status of returns.");
                                System.out.println("4. Go back.");
                                System.out.println("");
                                int decision1 = input.nextInt();
                                System.out.println("--------------------------------------------------------------------------");
                                
                                
                                    switch(decision1){
                                            
                                            //MANAGE TEXTBOOKS
                                            case 1:
                                                while(true){
                                                System.out.println("");
                                                System.out.println("WELCOME TO: MANAGE TEXTBOOK SECTION, SELECT BELOW");
                                                System.out.println("-------------------------------------");
                                                System.out.println("1. Lookup textbooks by ISBN.");
                                                System.out.println("2. Assign textbooks to students.");
                                                System.out.println("3. Check-in Textbooks."); //could add (back to school)
                                                System.out.println("4. Go back to previous page.");
                                                System.out.println("-------------------------------------");
                                                System.out.println("");
                                                int decision2 = input.nextInt();
                                                
                                                    //vvvvvvvvvv
                                                    switch(decision2){
                                                        //TEXTBOOK ISBN LOOKUP
                                                        case 1:
                                                        l.lookupBook(year1,year2,gradeInput);
                                                        
                                                        break;
                                                        
                                                        //ADD TEXTBOOK(S)
                                                        case 2:
                                                        at.addBook(year1, year2, gradeInput);
                                                        
                                                        break;
                                                        
                                                        //DELETE TEXTBOOKS
                                                        case 3:
                                                        rt.removeBook(year1, year2);
                                                        
                                                        
                                                        break;
                                                        
                                                        //PREVIOUS PAGE (RETURN FXN)
                                                        case 4:
                                                        
                                                        break;
                                                        default:
                                                        System.out.println("Invalid input.");
                                                        break;
                                                    }
                                                    
                                                    if(decision2==4){
                                                        break;
                                                    }
                                                    
                                            }
                                            
                                        //-------------------------------- 
                                        //--------------------------------  
                                            
                                            
                                        break;
                                        
                                            //MANAGE STUDENTS SECTION
                                            case 2:
                                                while(true){
                                                System.out.println("");
                                                System.out.println("----------------------------");
                                                System.out.println("1. Lookup student.");
                                                System.out.println("2. Add student.");
                                                System.out.println("3. Delete student.");
                                                System.out.println("4. Go back to previous page.");
                                                System.out.println("----------------------------");
                                                System.out.println("");
                                                int decision3 = input.nextInt();
                                                switch(decision3){
                                                    //STUDENT LOOKUP
                                                    case 1:
                                                    ls.lookupStud(year1, year2, gradeInput);
                                                    break;
                                                    //ADD STUDENT
                                                    case 2:
                                                    s.addStudentMethod(year1,year2,gradeInput);
                                                    s.sortSheet(year1,year2,gradeInput);
                                                    break;
                                                    //DELETE STUDENT
                                                    case 3:
                                                    r.removal(year1, year2, gradeInput);
                                                    break;
                                                    //PREVIOUS
                                                    case 4:
                                                    
                                                    break;
                                                    
                                                    
                                                    default:
                                                    System.out.println("Invalid input.");
                                                    break;
                                                }
                                                if(decision3==4){
                                                    break;
                                                }
                                            }
                                            
                                            
                                        //QUIT
                                        break;
                                        case 3:
                                            ss.status(year1,year2,gradeInput);
                                        break;
                                        case 4:
                                        break;
                                        default:
                                        System.out.println("Invalid input.");
                                        break;
                                    }
                                    
                                    if(decision1==4){
                                        break;
                                    }
                                    
                            }
                            
                            //-----------------------------------------------
                            //-----------------------------------------------
                            
            //^^^^^^^^^^^                
            break;
            case 2:
                            System.out.println("");
                            System.out.println("EXCEL CREATION INSTRUCTIONS");
                            System.out.println("----------------------------");
                            System.out.println("");
                            System.out.println("In order to create a new textbook database, you would have to manually create a new excel file in the "+'\n'+"SAME folder where the other files are for the existing textbook databases.");
                            System.out.println("Please be sure to properly format the new file in the following ways:");
                            System.out.println();
                            System.out.println("Open up Google Sheets, create the sheet as outlined and then download it as an excel file. DO NOT OPEN the excel file."+'\n'+"Make sure to store it in the SAME folder where the other excel files are for the existing textbook databases.");
                            System.out.println("You can create the Google sheet while the program is running, so no need to close the program and start again");
                            System.out.println("");
                            System.out.println("");
                            System.out.println("1. Please name the file 'bookstore20xx20yy' 20xx being the first year of the school year" +'\n'+"and 20yy being the second year of the school year (e.g. 20222023)");
                            System.out.println("Please make sure there are no spaces in the file name. AT ALL. Not even after 20yy."+'\n'+"The spaces do interfere with the codes ability to work. Thank you.");
                            System.out.println("");
                            System.out.println("2. The first row of the sheet should be the categories. (e.g. first name, last name) just like "+'\n'+"how the other excel spreadsheets are formatted.");
                            System.out.println("");
                            System.out.println("3. The first column should be BLANK, the second should be the student's last name, the third column should be the "+'\n'+"student's first name, and the fourth column should be the student's locker number.");
                            System.out.println("");
                            System.out.println("4. Please DO NOT manually type in the student's first and last name or the ISBN #, that is what this program is for :)");
                            System.out.println("Only the heading is required to be manually typed in.");
                            System.out.println("");
                            System.out.println("5. The next columns should hold the classes.");
                            System.out.println("");
                            System.out.println("Please press '1' if you have read this, understood it,have created the file and are ready to write in it."+'\n'+"As students are not assigned a locker till later, please fill in the name blank with a placeholder name for the time being.");
                            System.out.println("After the locker assignments are fixed, you can go in and assign students their lockers by pressing '2'."+'\n'+"Please note that when it comes to assigning students to lockers, the locker numbers show up in the order they appear in the excel file.");
                            System.out.println("If you need any help, please contact one of the developers."); 
                            System.out.println();
                            System.out.println("Template: https://docs.google.com/spreadsheets/d/13CDcSfPoSR_2Ml7-p4JeFM_hXszi-yrah0BXofN0Tf4/edit?usp=sharing");
                            System.out.println();
                            System.out.println("");
                            System.out.println("-----------------------------");
                            System.out.println("1. UNDERSTOOD, CONTINUE");
                            System.out.println("2. ASSIGN STUDENTS TO LOCKER NUMBER.");
                            System.out.println("3. GO BACK TO PREVIOUS PAGE.");
                            int decision5 = input.nextInt();
                            System.out.println("");
                            System.out.println("----------------------------");
            
            
            
                switch(decision5){
                    case 1:
                    System.out.println("");
                    System.out.println("ENTER SCHOOL YEAR YOU WISH TO WORK ON: EXAMPLE (2020-2021)");
                    System.out.println("From:");
                    year1 = input.next();
                    System.out.println("");
                    System.out.println("To:");
                    year2 = input.next();
                    System.out.println("");
                    System.out.println("--------------------------------------------------------------------------");
                    System.out.println("TYPE IN WHICH GRADE YOU WISH TO WORK ON (MUST BE FORMATED AS EX: 6th or 12th)");
                    gradeInput = input.next();
                    System.out.println("");
                    System.out.println("--------------------------------------------------------------------------");
                    System.out.println("");
                    ce.createSheet(year1, year2, gradeInput);
                    break;
                    case 2:
                    System.out.println("");
                    System.out.println("ENTER SCHOOL YEAR YOU WISH TO WORK ON: EXAMPLE (2020-2021)");
                    System.out.println("From:");
                    year1 = input.next();
                    System.out.println("");
                    System.out.println("To:");
                    year2 = input.next();
                    System.out.println("");
                    System.out.println("--------------------------------------------------------------------------");
                    System.out.println("TYPE IN WHICH GRADE YOU WISH TO WORK ON (MUST BE FORMATED AS EX: 6th or 12th)");
                    gradeInput = input.next();
                    System.out.println("");
                    System.out.println("--------------------------------------------------------------------------");
                    System.out.println("");
                    al.assignLock(year1, year2, gradeInput);
                    break;
                    case 3:
                    break;
                    default:
                    System.out.println("Invalid Input.");
                    break;
                }
                break;
                case 3:
                System.out.println("Thank you and have a good day!");
                break;
                default:
                System.out.println("Invalid Input.");
                break;
        }
    
    
        if(decision4==3){
            break;
        }
    }
    }
}