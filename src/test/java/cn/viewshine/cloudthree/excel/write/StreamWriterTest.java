package cn.viewshine.cloudthree.excel.write;

import cn.viewshine.cloudthree.excel.StreamExcelFactory;
import cn.viewshine.cloudthree.excel.exception.WriteExcelException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Ignore;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.stream.IntStream;

/**
 * @author moxiao
 * @Email changwei@viewshine.cn
 * @date 2021/1/23
 */
@Ignore
public class StreamWriterTest {


    @Test
    public void streamWrite() throws IOException, InterruptedException {
        TimeUnit.SECONDS.sleep(20);
        //第一个参数为：如果有文件的话，则打开文件。注意：如果文件较大，会占用非常大的内存，因为他会将全部内容放入内存中，例如打开20M的Excel，会占用几个G内存，导致Heap OUTOFMEMORY
        // 第三个参数表示是否压缩临时文件，如果为true的话，在写入文件的时候，速度慢，但是因为压缩临时文件了所以，占用内存较少，
        //                          如果为false的话，速度快，但是占用更多的内存。默认为false
//        Workbook workbook = new SXSSFWorkbook(null, 300, true);
        Workbook workbook = new SXSSFWorkbook(300);
        Sheet sheet1 = workbook.createSheet("sheet1");
        List<String> data = Arrays.asList("018501006797", "商中林","031000405541", "2020072627","13582703528","130921196909303618","冯官庄1-1-商中林","民用NB物联网表（物联网平台）","43.5","2020-12-01 00:00:36","阶梯起始日上告","2020-12-01 00:00:00","398.6","-0.75","106.83","2.4","0.01","表计预付费","500.0","0","00000000","0","开","正常","6.3","22","65511","-831","1","主动上告","43.5","39.09","表具预付费","0101", "0100");
        //				13582703528	130921196909303618	冯官庄1-1-商中林	民用NB物联网表（物联网平台）	43.5	2020-12-01 00:00:36	阶梯起始日上告	2020-12-01 00:00:00	398.6	-0.75	106.83	2.4	0.01	表计预付费	500.0	0	00000000	0	开	正常	6.3	22	65511	-831	1	主动上告	43.5	39.09			表具预付费					0101  0100
        IntStream.range(0, 100000).forEach(i -> {
            Row row = sheet1.createRow(i);
            IntStream.range(0, 34).forEach(j -> {
                Cell cell = row.createCell(j, CellType.STRING);
                cell.setCellValue(data.get(j));
            });
        });
        IntStream.range(100000, 200001).forEach(i -> {
            Row row = sheet1.createRow(i);
            IntStream.range(0, 34).forEach(j -> {
                Cell cell = row.createCell(j, CellType.STRING);
                cell.setCellValue(data.get(j));
            });
        });
        Files.createDirectories(Paths.get("/Users/xiaochang/docFile/stream1.xlsx").getParent());

        try (FileOutputStream out = new FileOutputStream("/Users/xiaochang/docFile/stream1_back.xlsx")) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
            throw new WriteExcelException("写入文件出现错误，方法：saveByFile", e);
        } finally {
            try {
                //Note that SXSSF allocates temporary files that you must always clean up explicitly, by calling the dispose method.
                //注意：如果是SXSSF的话，必须显示调用workbook的dispose方法
                ((SXSSFWorkbook)workbook).dispose();
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
                throw new WriteExcelException("关闭操作出现问题", e);
            }
        }
    }

    @Test
    public void streamWriteTools() {
        long startTime = System.currentTimeMillis();
        StreamExcelFactory streamExcelFactory = new StreamExcelFactory("/Users/xiaochang/docFile/streamTools1.xlsx", null);
        List<String> data = Arrays.asList("018501006797", "商中林","031000405541", "2020072627","13582703528","130921196909303618","冯官庄1-1-商中林","民用NB物联网表（物联网平台）","43.5","2020-12-01 00:00:36","阶梯起始日上告","2020-12-01 00:00:00","398.6","-0.75","106.83","2.4","0.01","表计预付费","500.0","0","00000000","0","开","正常","6.3","22","65511","-831","1","主动上告","43.5","39.09","表具预付费","0101", "0100");
        //				13582703528	130921196909303618	冯官庄1-1-商中林	民用NB物联网表（物联网平台）	43.5	2020-12-01 00:00:36	阶梯起始日上告	2020-12-01 00:00:00	398.6	-0.75	106.83	2.4	0.01	表计预付费	500.0	0	00000000	0	开	正常	6.3	22	65511	-831	1	主动上告	43.5	39.09			表具预付费					0101  0100
        IntStream.range(0, 100000).forEach(i -> {
            streamExcelFactory.writeData(data);
        });
        IntStream.range(0, 100000).forEach(i -> {
            streamExcelFactory.writeData(data);
        });
        IntStream.range(0, 100000).forEach(i -> {
            streamExcelFactory.writeData(data);
        });
        streamExcelFactory.finish();
        System.out.println(System.currentTimeMillis() - startTime);
    }

}
