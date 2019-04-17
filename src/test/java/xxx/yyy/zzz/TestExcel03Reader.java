package xxx.yyy.zzz;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.Map;

/**
 * TestExcel03Reader: ExcelField 2003 reader test cases.
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
    public void readerTest() {
        ExcelReader<Map> from = ExcelReader.open(new File("D:\\test.xls"));
        if (from == null)
            return;
//        List<Map<String, Object>> maps = from.readAll().get();
//        from.iterateThen(System.out::println);
        from.stream().forEach(System.out::println);
//        Iterator<Map<String, Object>> iterator = from.iterator();
//        System.out.println(iterator.next());
//        System.out.println(iterator.next());
//        System.out.println(iterator.next());
//        System.out.println(iterator.next());
//        System.out.println(iterator.next());
        from.withHead();
    }

    /**
     * Test excel type
     */
    @Test
    public void excelTypeTest() throws IOException {
//        log.info("ExcelField type: {}", ExcelTypeEnum.valueOf(new FileInputStream(new File("D:\\test.xlsx"))));
        ExcelReader<Excel03Reader> open = ExcelReader.open(new File("D:\\test.xls"), Excel03Reader.class);

    }

    /**
     * Test method of converting number to letter.
     */
    @Test
    public void number2LetterTest() {
//        log.info(Excel03Reader.convertNumber2Letter(53));
    }

    @Data
    private static class TestBean {
        private String id;
    }

}
