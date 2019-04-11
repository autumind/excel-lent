package xxx.yyy.zzz;

import java.util.List;

/**
 * RowConverter: Row data converter.
 *
 * @author autumind
 * @since 2019-04-11
 */
public interface RowConverter {

    /**
     * Convert row data to specified class object
     *
     * @param rowCells value of each row cell.
     * @param clz      class to be converted
     * @return class object
     */
    <T> T convert(List<Object> rowCells, Class<T> clz);
}
