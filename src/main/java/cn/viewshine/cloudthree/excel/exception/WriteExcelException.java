package cn.viewshine.cloudthree.excel.exception;

/**
 * @author: 常伟
 * @create: 2019/8/11 10:48
 * @email: kmustchang@qq.com
 * @version: 1.0
 * @Description: 表示写入Excel时候的异常
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
