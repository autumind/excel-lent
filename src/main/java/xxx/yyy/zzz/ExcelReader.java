package xxx.yyy.zzz;

import org.apache.commons.logging.LogFactory;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * ExcelReader: Uniform excel reader interface.
 *
 * @author autumind
 * @since 2019-03-22
 */
interface ExcelReader<T> extends Iterable<T> {

    /**
     * Open excel reader.
     *
     * @param file file
     * @param clz  row class
     * @return ExcelField reader.
     */
    static <E> ExcelReader<E> open(File file, Class<E> clz) {
        try (BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file))) {
            ExcelTypeEnum type = ExcelTypeEnum.valueOf(bis);
            if (type == ExcelTypeEnum.XLS) {
                return new Excel03Reader<E>(bis).setFile(file).setClz(clz);
            }
        } catch (IOException e) {
            throw new FileNotSupportedException("The file is not xls or xlsx or csv, please re-select.");
        }

        return null;
    }

    /**
     * Open excel reader.
     *
     * @param file file
     * @return ExcelField reader.
     */
    static <E> ExcelReader<E> open(File file) {
        return ExcelReader.open(file, null);
    }

    /**
     * Set sheet name which needs to be parsed.
     *
     * @param sheetName sheet name
     * @return reader
     */
    ExcelReader<T> sheetName(String sheetName);

    /**
     * Flag excel has head.
     *
     * @return reader
     */
    ExcelReader<T> withHead();

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
        if (this instanceof Excel03Reader) {
            Class<T> clz = ((Excel03Reader<T>) this).getClz();
            File file = ((Excel03Reader<T>) this).getFile();
            try (BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file))) {
                Excel03Reader<T> excel03Reader = new Excel03Reader<>(bis);
                excel03Reader.setFile(file);
                excel03Reader.setClz(clz);
                return StreamSupport.stream(excel03Reader.spliterator(), false);
            } catch (Exception e) {
                LogFactory.getLog(ExcelReader.class).error("Construct excel reader stream failure.", e);
            }
        }
        return StreamSupport.stream(spliterator(), false);
    }
}
