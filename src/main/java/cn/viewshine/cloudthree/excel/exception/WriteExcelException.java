package cn.viewshine.cloudthree.excel.exception;

/**
 * 表示写入Excel时候的异常
 * @author changwei[changwei@viewshine.cn]
 * @version: 1.0
 */
public class WriteExcelException extends IllegalArgumentException {

    public WriteExcelException() {
        super("写入Excel异常错误信息");
    }

    public WriteExcelException(String message) {
        super(message);
    }

    public WriteExcelException(String meesage, Throwable cause) {
        super(meesage, cause);
    }

    public WriteExcelException(Throwable cause) {
        super(cause);
    }

}
