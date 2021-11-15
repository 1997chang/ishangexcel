package cn.viewshine.cloudthree.excel.write;

import cn.viewshine.cloudthree.excel.context.WriteContext;
import cn.viewshine.cloudthree.excel.exception.WriteExcelException;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.OutputStream;
import java.util.*;
import java.util.stream.IntStream;

/**
 * 完成Excel的写入操作
 * @author changwei[changwei@viewshine.cn]
 */
public class WriteExcel {

    /**
     * 用于将data数据写入到Excel文件中，可以使用try-with-resource实现自动关闭流
     *
     * 注意：这种方式针对于使用@ExcelField注解所在的类
     * @param data 写入Excel文件中的数据
     * @param fileName 文件的名称
     * @param xssf 是否是XSSF文件格式
     */
    public static void writeExcelByFileName(Map<String,List<?>> data, String fileName, boolean xssf) {
        //验证所有Sheet表格名称是否合法
        Map<String, Class> classMap = validationWriteData(data.entrySet());

        //创建写Excel的上下文信息，包括WorkBook等
        WriteContext writeContext=new WriteContext(xssf, classMap);

        //向Excel表格中添加数据内容
        writeContext.write(data, fileName);
    }

    /**
     * 用于将Data数据写入到Excel文件流中，必须手动关闭输出流
     *
     * 注意：这种方式针对于使用@ExcelField注解所在的类
     * @param data
     * @param outputStream
     * @param xssf
     */
    public static void wirteExcelByStream(Map<String, List<?>> data, OutputStream outputStream, boolean xssf) {
        //验证所有Sheet表格名称是否合法，并且完成sheet表名与Class的对应关系
        WriteContext writeContext = new WriteContext(xssf, validationWriteData(data.entrySet()));
        writeContext.write(data, outputStream);
    }

    /**
     * 用于将数据写入到文件中，这种方式针对于不继承@ExcelField注解的类。
     * @param data 准备写入的数据内容
     * @param headName 写入的头部分数据
     * @param fileName 文件名称
     * @param xssf 是否是2007的xlsx格式数据
     */
    public static void writeExcelByFileName(Map<String, List<List<String>>> data, Map<String, List<List<String>>> headName, String fileName, boolean xssf) {
        Objects.requireNonNull(data, "表格数据不能为空");
        if (Objects.nonNull(headName)) {
            //如果headName不为空，一定要和data中的个数相同
            if (! Objects.equals(data.size(), headName.size())) {
                throw new WriteExcelException("表格数据的Sheet页数与表格头的Sheet页数不一致");
            }
            data.forEach((key, value) -> {
                List<List<String>> headNameInSheet = headName.get(key);
                Objects.requireNonNull(headNameInSheet, "表格数据的Sheet名称与表格头的Sheet名称不一致");
                boolean columnDifferent = CollectionUtils.isNotEmpty(value) && value.get(0).size() != headNameInSheet.size();
                if (columnDifferent) {
                    throw new WriteExcelException("表格列头个数与表格数据内容个数不一致");
                }
            });
        }
        //创建写Excel的上下文信息，包括WorkBook等
        WriteContext writeContext = new WriteContext(xssf, fileName);
        writeContext.write(data, headName, fileName);
    }

    /**
     * 1.用于验证所有的Sheet名称是否合法，
     * 2.用于验证每一个Sheet的数据格式是否相同
     * 不合法抛出WriteExcelException异常
     * @param entrySet
     * @return
     */
    private static Map<String, Class> validationWriteData(Set<Map.Entry<String, List<?>>> entrySet) {
        //因为装载因子默认为0.75
        Map<String, Class> result = new HashMap(entrySet.size()*4/3+1);

        //对所有内容进行验证
        entrySet.forEach(entry -> {

            //1.首先验证所有的SheetName是否合法
            try {
                WorkbookUtil.validateSheetName(entry.getKey());
            } catch (IllegalArgumentException e) {
                throw new WriteExcelException("Sheet表格名称不合法，不合法的Sheet的名称为：" + entry.getKey(), e);
            }

            //2验证每一个Sheet中List对象是不是同一个Class，一张Sheet要保持所有Row的Class一致
            if (CollectionUtils.isNotEmpty(entry.getValue())) {
                //获取第一个Class，然后与后续的所有Class进行等值比较，从而判断是不是都是同一个Class
                Class<?> zClass = entry.getValue().get(0).getClass();
                result.put(entry.getKey(), zClass);
                //如果有一个类型和第一个类型不相同，就会抛出异常
                if(! IntStream.range(1, entry.getValue().size()).mapToObj(i -> entry.getValue().get(i).getClass()).allMatch(zClass::equals)) {
                    throw new WriteExcelException("写入到同一个Sheet的Class类型必须相同");
                }
            }
        });
        return result;
    }
}
