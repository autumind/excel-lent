/*
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

package io.cruder.excellent;

import lombok.Data;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.Map;

/**
 * TestExcel03Reader: ExcelField 2003 reader test cases.
 *
 * @author cruder
 * @since 2019-03-25
 */
@Slf4j
public class TestExcel03Reader {

    /**
     * Test reader
     */
    @Test
    public void readerTest() {
        Excel.lent("D:\\test.xls", TestBean.class)
                .firstRowAsHeader()
                .forEach(bean -> log.info("{}", bean));
//                .headers("A", "B", "C")
//                .readAll()
//                .ifPresent(beans -> log.info("{}", beans ));
    }

    /**
     * Test reader
     */
    @Test
    @SneakyThrows
    public void xlsxReaderTest() {
//        Excel.lent("D:\\test.xlsx", TestBean.class).readAll();
        Excel.lent(null, new FileInputStream("D:\\test.xlsx"), TestBean.class)
                .firstRowAsHeader()
                .readAll()
                .ifPresent(list -> log.info("{}", list));

    }

    /**
     * Test excel type
     */
    @Test
    public void excelTypeTest() throws IOException {

//        ExcelReader<Excel03Reader> open = ExcelReader.open(new File("D:\\test.xls"), Excel03Reader.class);
        Reader<Map<String, String>> open = Excel.lent("D:\\test.xls").headers("A", "B", "C");
        open.readAll().ifPresent(maps -> log.info("{}", maps));
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
    public static class TestBean {
        private String name;
        private double age;
        private String remark;
        @ExcelColumn(format = "yyyyMMdd HH:mm:ss")
        private LocalDateTime date;
    }

}
