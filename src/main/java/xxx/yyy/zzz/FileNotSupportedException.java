package xxx.yyy.zzz;

import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * FileNotSupportedException: Exception of not supported file.
 *
 * @author autumind
 * @since 2019-03-27
 */
@Data
@EqualsAndHashCode(callSuper = true)
public class FileNotSupportedException extends RuntimeException {

    public FileNotSupportedException() {
    }

    public FileNotSupportedException(String message) {
        super(message);
    }

    public FileNotSupportedException(String message, Throwable cause) {
        super(message, cause);
    }

    public FileNotSupportedException(Throwable cause) {
        super(cause);
    }

    public FileNotSupportedException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
