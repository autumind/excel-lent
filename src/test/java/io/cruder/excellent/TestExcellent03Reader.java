package io.cruder.excellent;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.IOException;

/**
 * TestExcel03Reader: ExcelField 2003 reader test cases.
 *
 * @author cruder
 * @since 2019-03-25
 */
@Slf4j
public class TestExcellent03Reader {

    /**
     * Test reader
     */
    @Test
    public void readerTest() {
        Excellent.open("D:\\test.xls").forEach(t -> {});
    }

    /**
     * Test excel type
     */
    @Test
    public void excelTypeTest() throws IOException {

//        ExcelReader<Excel03Reader> open = ExcelReader.open(new File("D:\\test.xls"), Excel03Reader.class);
        Reader<Object> open = Excellent.open("D:\\test.xls");
        open.readRow();
        open.readRow();
    }

    /**
     * Test method of converting number to letter.
     */
    @Test
    public void number2LetterTest() throws NoSuchFieldException {
//        log.info(Excel03Reader.convertNumber2Letter(53));
//        Method method = ReflectUtil.resolveSetter(TestBean.class.getDeclaredField("id"), TestBean.class);
//        log.info("{}", method);
    }

    @Data
    private static class TestBean {
        private String id;
    }

}
