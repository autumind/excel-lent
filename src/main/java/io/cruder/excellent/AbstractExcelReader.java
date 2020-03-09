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

import io.cruder.excellent.util.DefaultRowConverter;
import io.cruder.excellent.util.RowConverter;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;
import org.apache.poi.hssf.eventusermodel.HSSFListener;

import java.util.ArrayList;
import java.util.List;

/**
 * AbstractExcelReader: Abstract implement excel reader.
 *
 * @author cruder
 * @since 2019-04-11
 */
@Data
public abstract class AbstractExcelReader<T> implements HSSFListener, Reader<T> {

    /**
     * Row class.
     */
    @Setter
    @Getter
    @Accessors(chain = true)
    protected Class<T> parameterType;

    /**
     * Data head.
     */
    protected List<String> heads = new ArrayList<>();

    /**
     * If excel has head, take 1st row as head information.
     */
    protected boolean hasHead = false;

    /**
     * Row data converter.
     */
    protected RowConverter converter = DefaultRowConverter.INSTANCE;

    /**
     * Numbers of alphabetic letter.
     */
    private final int ALPHABETIC_LETTER_NUMBERS = 26;

    /**
     * Flag excel has head.
     *
     * @return reader
     */
    @Override
    public AbstractExcelReader<T> withHead() {
        this.hasHead = true;
        return this;
    }

    /**
     * Set row data converter.
     *
     * @return reader
     */
    public AbstractExcelReader<T> withConverter(RowConverter converter) {
        this.converter = converter;
        return this;
    }

    /**
     * Convert non-negative number which less than 676 to letter.
     *
     * @param i non-negative number
     * @return letter
     */
    protected String convertNumber2Letter(int i) {
        if (i < 0) {
            throw new IllegalArgumentException();
        }

        if ((i + 1) / ALPHABETIC_LETTER_NUMBERS > ALPHABETIC_LETTER_NUMBERS) {
            throw new IllegalArgumentException("Too many columns, please decrease some useless column and retry.");
        }

        if (i / ALPHABETIC_LETTER_NUMBERS == 0) {
            return String.valueOf((char) (i + 'A'));
        } else {
            return String.valueOf((char) (i / ALPHABETIC_LETTER_NUMBERS + 'A' - 1)).concat(String.valueOf((char) (i % 26 + 'A')));
        }
    }
}
