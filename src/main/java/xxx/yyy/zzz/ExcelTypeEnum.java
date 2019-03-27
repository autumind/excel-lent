package xxx.yyy.zzz;

import lombok.Getter;
import lombok.Setter;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.util.IOUtils;

import java.io.*;

/**
 * ExcelTypeEnum: Enum of excel type.
 *
 * @author autumind
 * @since 2019-03-27
 */
public enum ExcelTypeEnum {
    XLSX, XLS, CSV(",");

    @Getter
    @Setter
    private String separator;

    ExcelTypeEnum(String separator) {
        this.separator = separator;
    }

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

            // detect if csv
            try (BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(inputStream))) {
                bufferedReader.mark(1);
                String readLine = bufferedReader.readLine();
                bufferedReader.reset();
            } catch (Exception e) {

            }
            return null;
        } catch (IOException e) {
            throw new FileNotSupportedException(e);
        }
    }
}
