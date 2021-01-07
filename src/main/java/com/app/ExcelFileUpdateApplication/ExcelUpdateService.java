package com.app.ExcelFileUpdateApplication;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUpdateService {
	
	public void updateExcelFile() {
		
		System.out.println("inside update excel file method");
		
		String excelFilePath = "C:\\Users\\lipsa\\Desktop\\name.xlsx";  // provide your excel file path
		
		try {
			
			FileInputStream fileInputStream = new FileInputStream(excelFilePath);
			
			Workbook workbook = WorkbookFactory.create(fileInputStream);		
			Sheet sheet = workbook.getSheetAt(0);
			
			int lastRowCount = sheet.getLastRowNum();
			System.out.println("lastRowCount.. "  +  lastRowCount);
			
			List<User> userList = getUserList();
			for (int i = 0; i < userList.size(); i++) {
				
				Row dataRow = sheet.createRow(++lastRowCount);
	        	dataRow.createCell(0).setCellValue(userList.get(i).getUserId());
	        	dataRow.createCell(1).setCellValue(userList.get(i).getUserName());       	        
				
			}
			
			System.out.println("lastRowCount after excel sheet modified.. "  +  lastRowCount);
			fileInputStream.close();
			
			FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath);
			workbook.write(fileOutputStream);
			fileOutputStream.close();
			System.out.println("excel sheet updated successfully........");
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		
	}
	
	
	public List<User> getUserList(){
		
		List<User> userList = new ArrayList<User>();
		
		userList.add(new User("111" ,"daisy"));
		userList.add(new User("112" ,"daisy"));
		userList.add(new User("113" ,"daisy"));
		userList.add(new User("114" ,"daisy"));
		userList.add(new User("115" ,"daisy"));
		userList.add(new User("116" ,"daisy"));
		userList.add(new User("117" ,"daisy"));
		userList.add(new User("118" ,"daisy"));
		userList.add(new User("119" ,"daisy"));
		return userList;
		
	}
	
}
