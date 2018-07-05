
package com.marbola.excel.impl;

import java.io.File;
import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.function.Predicate;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.marbola.excel.ExcelEntity;
import com.marbola.excel.ExcelField;
import com.marbola.excel.exception.EntityNotValidException;

public class ExcelReaderImpl<T> implements ExcelReader<T>  {

	private Class<T> clazz; 
	private Field[] fields;
	
	private String excelFile;
	private int indexHeader;
	private String sheetName;
	private Predicate<Row> predicate;
	private Locale locale;
	
	public ExcelReaderImpl(Class<T> clazz, String excelFile) throws Exception {
		this(clazz, excelFile, 0, "Sheet1");
	}
	
	public ExcelReaderImpl(Class<T> clazz, String excelFile,int indexHeader,String sheetName) throws Exception {
		this(clazz, excelFile, 0, "Sheet1", row -> true , Locale.ITALIAN);
	}	
	
	public ExcelReaderImpl(Class<T> clazz, String excelFile,int indexHeader,String sheetName, Predicate<Row> predicate,Locale locale) throws Exception {
		this.excelFile = excelFile;
		this.clazz = clazz;
		this.indexHeader = indexHeader;
		this.sheetName = sheetName;
		this.predicate = predicate;
		this.locale = locale;
		
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
	
	/* (non-Javadoc)
	 * @see com.marbola.excel.impl.ExcelReader#readRows()
	 */
	@Override
	public List<T> readRows() throws Exception {
		ArrayList<T> list = new ArrayList<T>();
		
		Workbook wb = WorkbookFactory.create(new File(excelFile));

	    Sheet sheet = wb.getSheet(sheetName);
	    
	    Row header = sheet.getRow(indexHeader);
	   
	    Row row = null ;
	    int currentRow = ++indexHeader;
	    
	    while ( !checkIfRowIsEmpty(( row = sheet.getRow(currentRow++))) && predicate.test(row) ) {
	    	list.add(this.readRow(row,header));
		}
	    
	    wb.close();
	    
		return list;
	}
	
	/* (non-Javadoc)
	 * @see com.marbola.excel.impl.ExcelReader#readRow(int, org.apache.poi.ss.usermodel.Row)
	 */
	@Override
	public T readRow(int rowIndex,Row header) throws Exception {
		Workbook wb = WorkbookFactory.create(new File(excelFile));
	    Sheet sheet = wb.getSheet(sheetName);
		return this.readRow(sheet.getRow(rowIndex),header);
	}
	
	/* (non-Javadoc)
	 * @see com.marbola.excel.impl.ExcelReader#readRow(int, int)
	 */
	@Override
	public T readRow(int rowIndex,int headerIndex) throws Exception {
		Workbook wb = WorkbookFactory.create(new File(excelFile));
	    Sheet sheet = wb.getSheet(sheetName);
		Row header = sheet.getRow(headerIndex);
		return this.readRow(rowIndex,header);
	}
	
	/* (non-Javadoc)
	 * @see com.marbola.excel.impl.ExcelReader#readRow(org.apache.poi.ss.usermodel.Row, org.apache.poi.ss.usermodel.Row)
	 */
	@Override
	public T readRow(Row row,Row header) throws Exception {
		T rowInstance = clazz.newInstance();
	    
		for (Cell currentCell : row) {
			Cell cell = header.getCell(currentCell.getColumnIndex());
			String headerForIndex = cell.getStringCellValue().trim();
			
			for (Field field : fields) {
				
				if(field.getName().equals(headerForIndex)){

					boolean accessible = field.isAccessible();
					
					field.setAccessible(true);
					ExcelField annotation = field.getAnnotation(ExcelField.class);
					field.set(rowInstance,  getCellGenericValue(currentCell,field.getType(),annotation.value()));
					
					field.setAccessible(accessible);
					
					break;
				}
			}
		}
		
		return rowInstance;
	}
	
	@SuppressWarnings({ "hiding", "unchecked" })
	private <T> T getCellGenericValue(Cell cell, Class<T> clazz,String pattern) throws ParseException {
		T result = null; 
		int cellType = cell.getCellType();

		switch (cellType) {
			case Cell.CELL_TYPE_STRING:{
				if (clazz == Date.class) {
					result = (T) new SimpleDateFormat(pattern,locale).parse(cell.getStringCellValue());
			    }else {
			    	result = (T) cell.getStringCellValue();
			    }
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
				
				if(HSSFDateUtil.isCellDateFormatted(cell)) {
					result = (T) cell.getDateCellValue();
				}else {
					result = (T) new Double(cell.getNumericCellValue());
				}
				
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
