package xxx.yyy.zzz;

import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Map;

/**
 * TestExcel03Reader: Excel 2003 reader test cases.
 *
 * @author autumind
 * @since 2019-03-25
 */
@Slf4j
public class TestExcel03Reader {


    /**
     * Test reader
     */
    @Test
    public void readerTester() throws IOException {
        Excel03Reader<Map> from = Excel03Reader.open(
                new BufferedInputStream(
                        new FileInputStream(new File("D:\\test.xls"))));
        if (from == null)
            return;
        from.readAll();
    }

    /**
     * Test excel type
     */
    @Test
    public void excelTypeTester() throws IOException {
        log.info("xxx: {}", Files.probeContentType(new File("D:\\test.csv").toPath()));
//        log.info("Excel type: {}", ExcelTypeEnum.valueOf(new FileInputStream(new File("D:\\test.xls"))));
    }

}
