package xxx.yyy.zzz;

import java.util.List;

/**
 * DefaultRowConverter: Default implement of row data converter.
 *
 * @author shenzb@shein.com
 * @since 2019-04-11
 */
public enum DefaultRowConverter implements RowConverter {
    /**
     * Singleton Converter
     */
    INSTANCE;

    @Override
    public <T> T convert(List<Object> rowCells, Class<T> clz) {
        return null;
    }
}
