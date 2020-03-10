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

import io.cruder.excellent.util.Converter;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * ExcelReader: Uniform excel reader interface.
 *
 * @author cruder
 * @since 2019-03-22
 */
interface Reader<T> extends Iterable<T> {

    /**
     * Take first row as head, if this works, headers from {@link Reader#headers(String...)} will be removed.
     * <p>Only first sheet's first row will be taken as headers and other sheet's first row will be skipped.<p/>
     *
     * @return reader
     */
    Reader<T> firstRowAsHeader();

    /**
     * Flag excel has head.
     *
     * @param head head name.
     * @return reader
     */
    Reader<T> headers(String... head);

    /**
     * Set row converter.
     *
     * @param converter row converter.
     * @return reader
     */
    Reader<T> converter(Converter converter);

    /**
     * Read one row.
     *
     * @return optional one row data.
     */
    Optional<T> readRow();

    /**
     * Read all rows. If excel contains large data, may cause OOM.
     *
     * @return optional all rows.
     */
    default Optional<List<T>> readAll() {
        List<T> list = new ArrayList<>();
        Optional<T> optional;
        while ((optional = readRow()).isPresent()) {
            list.add(optional.get());
        }
        return Optional.of(list);
    }

    /**
     * Construct stream.
     *
     * @return stream.
     */
    default Stream<T> stream() {
        return StreamSupport.stream(spliterator(), false);
    }
}
