
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
public class AssignLockernumber{

    public void assignLock(String year1, String year2, String gradeInput){
        Scanner s1 = new Scanner(System.in);
        try{
            String filename = "bookstore"+year1+year2+".xlsx";
            FileInputStream inputStream = new FileInputStream(new File(filename));
            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = wb.getSheet(gradeInput);
            int rowCount = sheet.getPhysicalNumberOfRows();
            for(int i=1;i<rowCount;i++){
                XSSFRow row = sheet.getRow(i);
                XSSFCell lockerNum = row.getCell(3);
                double lockerVal = lockerNum.getNumericCellValue();
                System.out.println("First name of student for locker: "+lockerVal);
                String first = s1.next();
                System.out.println("Last name of student for locker: "+lockerVal);
                String last = s1.next();
                row.createCell(1).setCellValue(last);
                row.createCell(2).setCellValue(first);
            }
            FileOutputStream outputStream = new FileOutputStream(filename);
            wb.write(outputStream);
        }catch(Exception e){
            e.printStackTrace();
        }
    }
}