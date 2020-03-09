/*
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

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
