package xxx.yyy.zzz;

import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;
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
    public void readerTester() {
        IExcelReader<Map> from = IExcelReader.open(new File("D:\\test.xls"));
        if (from == null)
            return;
        List<Map> maps = from.readAll().get();
        System.out.println(maps);
    }

    /**
     * Test excel type
     */
    @Test
    public void excelTypeTester() throws IOException {
//        log.info("Excel type: {}", ExcelTypeEnum.valueOf(new FileInputStream(new File("D:\\test.xlsx"))));
        IExcelReader<Excel03Reader> open = IExcelReader.open(new File("D:\\test.xls"), Excel03Reader.class);

    }

}