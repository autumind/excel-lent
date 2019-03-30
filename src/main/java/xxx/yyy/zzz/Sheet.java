package xxx.yyy.zzz;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * Sheet: Excel sheet.
 *
 * @author autumind
 * @since 2019-03-28
 */
@Data
@Accessors(chain = true)
public class Sheet {

    /**
     * Sheet no, starting from 1.
     */
    private int sheetNo;

    /**
     * Sheet name.
     */
    private String sheetName;

}
