package xxx.yyy.zzz;

import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.IOException;
import java.io.InputStream;

/**
 * ExcelTypeEnum: Enum of excel type.
 *
 * @author autumind
 * @since 2019-03-27
 */
public enum ExcelTypeEnum {
    XLSX, XLS;

    ExcelTypeEnum() {
    }

    public static ExcelTypeEnum valueOf(InputStream inputStream) {
        try {
            if (!inputStream.markSupported()) {
                inputStream = FileMagic.prepareToCheckMagic(inputStream);
            }
            FileMagic fileMagic = FileMagic.valueOf(inputStream);
            if (FileMagic.OLE2.equals(fileMagic)) {
                return XLS;
            }
            if (FileMagic.OOXML.equals(fileMagic)) {
                return XLSX;
            }
            throw new FileNotSupportedException("The file is not xls or xlsx or csv, please re-select.");
        } catch (IOException e) {
            throw new FileNotSupportedException(e);
        }
    }
}
