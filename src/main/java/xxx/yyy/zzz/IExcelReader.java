package xxx.yyy.zzz;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/**
 * IExcelReader: Uniform excel reader interface.
 *
 * @author autumind
 * @since 2019-03-22
 */
interface IExcelReader<T> {
    /**
     * Open excel reader.
     *
     * @param file file
     * @param clz  row class
     * @return Excel reader.
     */
    static <E> IExcelReader<E> open(File file, Class<E> clz) {
        try (BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file))) {
            ExcelTypeEnum type = ExcelTypeEnum.valueOf(bis);
            if (type == ExcelTypeEnum.XLS) {
                Excel03Reader<E> reader = new Excel03Reader<>(bis);
                reader.setClz(clz);
                return reader;
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
     * @return Excel reader.
     */
    static <E> IExcelReader<E> open(File file) {
        return IExcelReader.open(file, null);
    }

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
}
