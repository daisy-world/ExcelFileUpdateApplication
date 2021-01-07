package com.app.ExcelFileUpdateApplication;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );
        
        ExcelUpdateService excelService = new ExcelUpdateService();
        
        excelService.updateExcelFile();
        
        System.out.println(" main method execution end.........");
       
    }
}
