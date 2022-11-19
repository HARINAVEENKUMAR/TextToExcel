package texttoexcel;

/* Copyright (C) 2022 HariNaveenKumar TamilSelvan - All Rights Reserved
 * You may use, distribute and modify this code without any restrictions for personal/Commercial use.
 * if you face any issue in this code please contact me at harinaveen984@gmail.com
 */

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Scanner;

public class TextToExcel {

    public static void textToExcel(String txtpath, String delimitter, String excelPath, String sheetName){

    // Pass the path to the file as a parameter
    File file = new File(txtpath);
    Scanner sc = null;
    try {
        sc = new Scanner(file);
    } catch (FileNotFoundException e) {
        throw new RuntimeException(e);
    }

    // Creating a list to store text file data in it
    ArrayList<String> ls = new ArrayList<>();
    // Adding data in it
    while (sc.hasNextLine()){
        ls.add(sc.nextLine());
    }


    //Creating New Workbook object
    HSSFWorkbook workbook = new HSSFWorkbook();

    // Passing the sheetName parameter for Creating New Sheet
    HSSFSheet sheet = workbook.createSheet(sheetName);




    for(int i= 0; i < ls.size() ; i++){

        // Creating new row
        HSSFRow row = sheet.createRow((short)i);

        //Splitting the row data  into columns by using delimitter parameter to Temporary String array.
        String[] stringArray = ls.get(i).split(delimitter);

        for(int j=0 ; j < stringArray.length;j++){

            //Inserting Temporary String array values into columns (Excel Cells)
            row.createCell(j).setCellValue(stringArray[j]);
        }
    }

    //Exporting output file
    FileOutputStream fileOut = null;
    try {
        fileOut = new FileOutputStream(excelPath);
    } catch (FileNotFoundException e) {
        throw new RuntimeException(e);
    }
    try {
        workbook.write(fileOut);
        fileOut.close();
        System.out.println("Excel file has been generated successfully.");
    } catch (IOException e) {
        throw new RuntimeException(e);
    }


}





}
