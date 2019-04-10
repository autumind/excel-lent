package xxx.yyy.zzz;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RecordFactoryInputStream;
import org.apache.poi.poifs.filesystem.DocumentInputStream;

import java.lang.reflect.Method;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.stream.Stream;

/**
 * ELHSSFEventFactory: Description of this class.
 *
 * @author autumind
 * @since 2019-03-22
 */
@Slf4j
@Data
@EqualsAndHashCode(callSuper = true)
public class ELHSSFEventFactory extends HSSFEventFactory {

    /**
     * Document input stream.
     */
    private DocumentInputStream documentInputStream;

    /**
     * Record factory input stream.
     */
    private RecordFactoryInputStream recordStream;

    /**
     * HSSF request.
     */
    private HSSFRequest req;

    /**
     * Process events
     *
     * @param consumer finish consumer
     * @param ifNext   if read next record or not supplier.
     */
    public void processEvents(Consumer<DocumentInputStream> consumer, Supplier<Boolean> ifNext) {

        try {
            while (ifNext.get()) {
                Record record = recordStream.nextRecord();
                if (record == null) {
                    consumer.accept(documentInputStream);
                    break;
                }
                Optional<Method> methodOptional = Stream.of(req.getClass().getDeclaredMethods())
                        .filter(method -> "processRecord".equals(method.getName()))
                        .findAny();
                if (!methodOptional.isPresent()) {
                    break;
                }
                Method method = methodOptional.get();
                method.setAccessible(true);
                method.invoke(req, record);
            }
        } catch (Exception e) {
            log.error("Process Excel 2003 failure.", e);
        }
    }


}
