package io.cruder.excellent;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.function.Consumer;
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
     * Flag excel has head.
     *
     * @return reader
     */
    Reader<T> withHead();

    /**
     * Read one row.
     *
     * @return optional one row data.
     */
    Optional<T> readRow();

    /**
     * Read rows for specified row amount.
     *
     * @param rowNum row amount
     * @return optional multiple rows.
     */
    Optional<List<T>> readRow(int rowNum);

    /**
     * Iterate excel data one by one and do extra something.
     *
     * @param consumer extra operation.
     */
    @Deprecated
    default void iterateThen(Consumer<T> consumer) {
        Optional<T> optional = readRow();
        if (!optional.isPresent()) {
            return;
        }
        consumer.accept(optional.get());
        this.iterateThen(consumer);
    }

    /**
     * Read all rows.
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
