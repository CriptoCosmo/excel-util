
package com.marbola.excel.impl;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.marbola.excel.ExcelEntity;
import com.marbola.excel.ExcelField;
import com.marbola.excel.exception.EntityNotValidException;


public class ExcelReaderImpl<T>  {

	Class<T> clazz; 
	Field[] fields;
	
	public ExcelReaderImpl(Class<T> clazz) throws EntityNotValidException {
		
		this.clazz = clazz;
		
		if(!clazz.isAnnotationPresent(ExcelEntity.class)) {
			throw new EntityNotValidException("Check ["+clazz.getName()+"] class.");
		}
		
		boolean hasLeastThenOne = false ;
		
		fields = clazz.getDeclaredFields();
		
		for (Field field : fields) {
			if(field.isAnnotationPresent(ExcelField.class)){
				hasLeastThenOne = true ;
				break;
			}
		}
		
		if (!hasLeastThenOne) {
			throw new EntityNotValidException("Check hasLeastThenOne "+hasLeastThenOne+" ["+clazz.getName()+"] fields class.");
		}
		
	}
	
	public List<T> readRows(String excelFile) throws Exception {
		return this.readRows(excelFile,0,"Sheet1");
	}
	
	public List<T> readRows(String excelFile,int indexHeader,String sheetName) throws Exception {
		ArrayList<T> list = new ArrayList<T>();
		
		Workbook wb = WorkbookFactory.create(new File(excelFile));

	    Sheet sheet = wb.getSheet(sheetName);
		// RETRIVE HEADER POSITIONS
	    
	    Row header = sheet.getRow(indexHeader);
	   
	    Row row = null ;
	    int currentRow = ++indexHeader;
	    
	    while ( !checkIfRowIsEmpty(( row = sheet.getRow(currentRow++))) ) {
	    	list.add(this.readRow(row,header));
		}
	    
	    wb.close();
	    
		return list;
	}
	
	private boolean checkIfRowIsEmpty(Row row) {
	    if (row == null) {
	        return true;
	    }
	    if (row.getLastCellNum() <= 0) {
	        return true;
	    }
	    for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
	        Cell cell = row.getCell(cellNum);
	        if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK && !cell.toString().isEmpty() ) {
	            return false;
	        }
	    }
	    return true;
	}
//	TODO
//	public T readRow(int rowIndex,Row header) throws InstantiationException, IllegalAccessException {
//		this.readRow(sheet.getRow(rowIndex),header)
//	}
	
	public T readRow(Row row,Row header) throws InstantiationException, IllegalAccessException {
		T rowInstance = clazz.newInstance();
	    
		for (Cell currentCell : row) {
			Cell cell = header.getCell(currentCell.getColumnIndex());
			String headerForIndex = cell.getStringCellValue().trim();
			
			for (Field field : fields) {
				
				if(field.getName().equals(headerForIndex)){

					boolean accessible = field.isAccessible();
					
					field.setAccessible(true);

					field.set(rowInstance, currentCell.getStringCellValue());
					
					field.setAccessible(accessible);
					
					break;
				}
			}
		}
		
		return rowInstance;
	}

}
