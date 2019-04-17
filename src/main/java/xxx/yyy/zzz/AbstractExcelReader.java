package xxx.yyy.zzz;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * AbstractExcelReader: Abstract implement excel reader.
 *
 * @author autumind
 * @since 2019-04-11
 */
@Data
public abstract class AbstractExcelReader<T> implements ExcelReader<T> {

    /**
     * ExcelField file
     */
    @Setter
    @Getter
    @Accessors(chain = true)
    protected File file;

    /**
     * Row class.
     */
    @Setter
    @Getter
    @Accessors(chain = true)
    protected Class<T> clz;

    /**
     * Data sheet name
     */
    protected String sheetName;

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
     * Set sheet name which needs to be parsed.
     *
     * @param sheetName sheet name
     * @return reader
     */
    public AbstractExcelReader<T> sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    /**
     * Flag excel has head.
     *
     * @return reader
     */
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

        if ((i + 1) / 26 > 26) {
            throw new IllegalArgumentException("Too many columns, please decrease some useless column and retry.");
        }

        if (i / 26 == 0) {
            return String.valueOf((char) (i + 65));
        } else {
            return String.valueOf((char) (i / 26 + 64)).concat(String.valueOf((char) (i % 26 + 65)));
        }
    }
}
