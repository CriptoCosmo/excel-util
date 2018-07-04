package com.marbola.excel.exception;

public class EntityNotValidException extends Exception {

	private static final long serialVersionUID = 4373749188384099666L;

	public EntityNotValidException() {
		super();
	}

	public EntityNotValidException(String message) {
		super(message);
	}
	
	public EntityNotValidException(Throwable e) {
		super(e);
	}
	
}
