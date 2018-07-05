package com.marbola.excel.impl;

import java.util.List;

import org.apache.poi.ss.usermodel.Row;

public interface ExcelReader<T> {

	List<T> readRows() throws Exception;

	T readRow(int rowIndex, Row header) throws Exception;

	T readRow(int rowIndex, int headerIndex) throws Exception;

	T readRow(Row row, Row header) throws Exception;

}