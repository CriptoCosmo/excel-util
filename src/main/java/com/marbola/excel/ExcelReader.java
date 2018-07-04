package com.marbola.excel;

import java.util.List;

public interface ExcelReader<T> {

	List<T> readRow(String excelFile);
	
}
