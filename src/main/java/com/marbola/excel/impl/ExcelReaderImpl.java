
package com.marbola.excel.impl;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.marbola.excel.ExcelEntity;
import com.marbola.excel.ExcelField;
import com.marbola.excel.exception.EntityNotValidException;


public class ExcelReaderImpl<T>  {

	private Class<T> clazz; 
	private Field[] fields;
	
	private String excelFile;
	private int indexHeader;
	private String sheetName;
	
	public ExcelReaderImpl(Class<T> clazz, String excelFile,int indexHeader,String sheetName) throws Exception {
		this.excelFile = excelFile;
		this.clazz = clazz;
		this.indexHeader = indexHeader;
		this.sheetName = sheetName;
		
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
	
	public ExcelReaderImpl(Class<T> clazz, String excelFile) throws Exception {
		this(clazz, excelFile, 0, "Sheet1");
	}
	
	public List<T> readRows() throws Exception {
		ArrayList<T> list = new ArrayList<T>();
		
		Workbook wb = WorkbookFactory.create(new File(excelFile));

	    Sheet sheet = wb.getSheet(sheetName);
	    
	    Row header = sheet.getRow(indexHeader);
	   
	    Row row = null ;
	    int currentRow = ++indexHeader;
	    
	    while ( !checkIfRowIsEmpty(( row = sheet.getRow(currentRow++))) ) {
	    	list.add(this.readRow(row,header));
		}
	    
	    wb.close();
	    
		return list;
	}
	
	public T readRow(int rowIndex,Row header) throws Exception {
		Workbook wb = WorkbookFactory.create(new File(excelFile));
	    Sheet sheet = wb.getSheet(sheetName);
		return this.readRow(sheet.getRow(rowIndex),header);
	}
	
	public T readRow(int rowIndex,int headerIndex) throws Exception {
		Workbook wb = WorkbookFactory.create(new File(excelFile));
	    Sheet sheet = wb.getSheet(sheetName);
		Row header = sheet.getRow(headerIndex);
		return this.readRow(rowIndex,header);
	}
	
	public T readRow(Row row,Row header) throws Exception {
		T rowInstance = clazz.newInstance();
	    
		for (Cell currentCell : row) {
			Cell cell = header.getCell(currentCell.getColumnIndex());
			String headerForIndex = cell.getStringCellValue().trim();
			
			for (Field field : fields) {
				
				if(field.getName().equals(headerForIndex)){

					boolean accessible = field.isAccessible();
					
					field.setAccessible(true);
					
					field.set(rowInstance,  getCellGenericValue(currentCell,field.getType()));
					
					field.setAccessible(accessible);
					
					break;
				}
			}
		}
		
		return rowInstance;
	}
	
	@SuppressWarnings({ "hiding", "unchecked" })
	private <T> T getCellGenericValue(Cell cell, Class<T> clazz) {
		
		T result = null; 
		int cellType = cell.getCellType();

		switch (cellType) {
			case Cell.CELL_TYPE_STRING:{
			    	result = (T) cell.getStringCellValue();
				break;
			}
			case Cell.CELL_TYPE_BOOLEAN:{
				result = (T) new Boolean(cell.getBooleanCellValue());
				break;
			}
			case Cell.CELL_TYPE_ERROR:{
				result = (T) new Byte(cell.getErrorCellValue());
				break;
			}
			case Cell.CELL_TYPE_FORMULA:{
				result = (T) cell.getCellFormula();
				break;
			}
			case Cell.CELL_TYPE_NUMERIC:{
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					System.out.println("HSSFDateUtil.isCellDateFormatted");
//					return (T) cell.getDateCellValue();
				}
				result = (T) new Double(cell.getNumericCellValue());
				break;
			}
			default:{
				break;
			}
		}
		
		return result;
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
}
