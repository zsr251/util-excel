package com.javasoso.util.excel;

/**
 * Excel导出导出异常
 *
 * @author jasonzhu
 * @date 2018/4/23
 */
public class ExcelException extends RuntimeException{

    public ExcelException() {
    }

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }
}
