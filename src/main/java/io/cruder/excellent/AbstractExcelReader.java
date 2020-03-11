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

import io.cruder.excellent.util.Constant;
import io.cruder.excellent.util.Converter;
import io.cruder.excellent.util.DefaultConverter;
import lombok.Data;
import lombok.Getter;
import lombok.experimental.Accessors;
import org.apache.poi.hssf.eventusermodel.HSSFListener;

import java.util.*;

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
    @Getter
    @Accessors(chain = true)
    protected Class<T> parameterizedType;

    /**
     * Data head.
     */
    protected List<String> headers = new ArrayList<>();

    /**
     * If excel has head, take 1st row as head information.
     */
    protected boolean firstRowAsHeader = false;

    /**
     * Row data converter.
     */
    protected Converter converter = DefaultConverter.INSTANCE;

    @Override
    public Reader<T> firstRowAsHeader() {
        firstRowAsHeader = true;
        return this;
    }

    @Override
    public Reader<T> headers(String... header) {
        if (header != null && header.length > 0) {
            headers.addAll(Arrays.asList(header));
        }
        return this;
    }

    @Override
    public Reader<T> converter(Converter converter) {
        this.converter = converter;
        return this;
    }

    /**
     * Do read operation.
     *
     * @return a row.
     */
    public abstract T doRead();

    @Override
    public Optional<T> readRow() {
        return Optional.ofNullable(doRead());
    }

    @Override
    public Iterator<T> iterator() {
        return new Iterator<T>() {

            /**
             * current row.
             */
            private T curr;

            /**
             * If there is next record.
             */
            private boolean hasNext = true;

            @Override
            public boolean hasNext() {
                Optional<T> optional = readRow();
                hasNext = optional.isPresent();
                if (hasNext) {
                    curr = optional.get();
                }
                return hasNext;
            }

            @Override
            public T next() {
                if (!hasNext) {
                    throw new NoSuchElementException();
                } else {
                    if (curr == null && !hasNext()) {
                        throw new NoSuchElementException();
                    }
                }
                T prev = curr;
                curr = null;
                return prev;
            }
        };
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

        if ((i + 1) / Constant.ALPHABETIC_LETTER_NUMBERS > Constant.ALPHABETIC_LETTER_NUMBERS) {
            throw new IllegalArgumentException("Too many columns, please decrease some useless column and retry.");
        }

        if (i / Constant.ALPHABETIC_LETTER_NUMBERS == 0) {
            return String.valueOf((char) (i + Constant.A));
        } else {
            return String.valueOf((char) (i / Constant.ALPHABETIC_LETTER_NUMBERS + Constant.A - 1))
                    .concat(String.valueOf((char) (i % Constant.ALPHABETIC_LETTER_NUMBERS + Constant.A)));
        }
    }
}
