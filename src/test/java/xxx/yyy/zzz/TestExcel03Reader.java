package xxx.yyy.zzz;

import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.Map;

/**
 * TestExcel03Reader: Excel 2003 reader test cases.
 *
 * @author autumind
 * @since 2019-03-25
 */
public class TestExcel03Reader {


    /**
     * Test reader
     */
    @Test
    public void testReader() throws IOException {
        Excel03Reader<Map> from = Excel03Reader.open(new File("D:\\test.xls"));
        if (from == null)
            return;
        from.readRow(3);
        from.readRow();
    }

}
