package xxx.yyy.zzz;

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
     * Read one row.
     *
     * @return optional one row data.
     */
    Optional<T> readRow();

    /**
     * Read rows for specified row number.
     *
     * @param rowNum row number
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
