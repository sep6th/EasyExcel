package com.github.sep6th.exception;

public class ExcelException extends RuntimeException {

	private static final long serialVersionUID = 1L;
	
	private String msg;

	public ExcelException(String msg) {
		super();
		this.msg = msg;
	}

	public String getMsg() {
		return msg;
	}

	public void setMsg(String msg) {
		this.msg = msg;
	}
	
	

}
