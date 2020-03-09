package io.cruder.excellent.hssf.eventusermodel;

import org.apache.poi.hssf.eventusermodel.HSSFUserException;
import org.apache.poi.hssf.record.Record;

/**
 * io.cruder.excellent.hssf.eventusermodel.HSSFRequest: HSSF request extend from {@link org.apache.poi.hssf.eventusermodel.HSSFRequest}.
 *
 * @author cruder.io
 * @since 2020-03-03
 */
public class HssfRequest extends org.apache.poi.hssf.eventusermodel.HSSFRequest {
    @Override
    protected short processRecord(Record rec) throws HSSFUserException {
        return super.processRecord(rec);
    }
}
