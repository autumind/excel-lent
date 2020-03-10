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


import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.eventusermodel.HSSFUserException;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RecordFactoryInputStream;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;

import static org.apache.poi.hssf.model.InternalWorkbook.WORKBOOK_DIR_ENTRY_NAMES;

/**
 * io.cruder.excellent.hssf.eventusermodel.HSSFEventFactory: Hssf event factory override from {@link org.apache.poi.hssf.eventusermodel.HSSFEventFactory}.
 *
 * @author cruder.io
 * @since 2020-03-03
 */
@Slf4j
public class HssfEventFactory {

    /**
     * Record stream.
     */
    private RecordFactoryInputStream recordStream;

    /**
     * Hssf Request.
     */
    private HssfRequest request;

    /**
     * POIFS file system.
     */
    private POIFSFileSystem poifsFileSystem;

    /**
     * Input stream.
     */
    private InputStream inputStream;

    /**
     * If there is next record.
     */
    private boolean hasNext = true;

    /**
     * Abort reading if true.
     */
    private boolean abort = false;

    /**
     * If current row is first row in current sheet.
     */
    private boolean firstSheetRow = false;

    /**
     * Creates a new instance of HSSFEventFactory
     *
     * @param request         req
     * @param poifsFileSystem fs
     */
    public HssfEventFactory(HssfRequest request, POIFSFileSystem poifsFileSystem) {
        this.request = request;
        this.poifsFileSystem = poifsFileSystem;
    }

    /**
     * Process one row.
     */
    @SneakyThrows
    public void process() {
        abort = false;
        processWorkbookEvents(request, poifsFileSystem.getRoot());
    }

    /**
     * Processes a file into essentially record events.
     *
     * @param req an Instance of HSSFRequest which has your registered listeners
     * @param dir a DirectoryNode containing your workbook
     * @throws IOException if the workbook contained errors
     */
    private void processWorkbookEvents(HssfRequest req, DirectoryNode dir) throws IOException {
        if (inputStream == null) {
            String name = Arrays.stream(WORKBOOK_DIR_ENTRY_NAMES)
                    .filter(potentialName -> dir.getEntryNames().contains(potentialName))
                    .findAny()
                    // Default entry name "Workbook".
                    .orElse(WORKBOOK_DIR_ENTRY_NAMES[0]);
            inputStream = dir.createDocumentInputStream(name);
        }
        processEvents(req, inputStream);
    }

    /**
     * Processes a DocumentInputStream into essentially Record events.
     * <p>
     * If an <code>AbortableHSSFListener</code> causes a halt to processing during this call
     * the method will return just as with <code>abortableProcessEvents</code>, but no
     * user code or <code>HSSFUserException</code> will be passed back.
     *
     * @param req an Instance of HSSFRequest which has your registered listeners
     * @param in  a DocumentInputStream obtained from POIFS's POIFSFileSystem object
     * @see org.apache.poi.poifs.filesystem.POIFSFileSystem#createDocumentInputStream(String)
     */
    private void processEvents(HssfRequest req, InputStream in) {
        try {
            genericProcessEvents(req, in);
        } catch (HSSFUserException hue) {
            /*If an HSSFUserException user exception is thrown, ignore it.*/
        }
    }

    /**
     * Processes a DocumentInputStream into essentially Record events.
     *
     * @param req an Instance of HSSFRequest which has your registered listeners
     * @param in  a DocumentInputStream obtained from POIFS's POIFSFileSystem object
     * @see org.apache.poi.poifs.filesystem.POIFSFileSystem#createDocumentInputStream(String)
     */
    private void genericProcessEvents(HssfRequest req, InputStream in)
            throws HSSFUserException {

        short userCode;

        // Create a new RecordStream and use that
        if (recordStream == null) {
            recordStream = new RecordFactoryInputStream(in, false);
        }

        // Process each record as they come in
        while (!abort) {
            Record r = recordStream.nextRecord();
            if (r == null) {
                hasNext = false;
                try {
                    in.close();
                    poifsFileSystem.close();
                } catch (Exception e) {
                    log.error(e.getMessage(), e);
                }
                break;
            }
            userCode = req.processRecord(r);
            if (userCode != 0) {
                break;
            }
        }

    }

    /**
     * @return True if there is new record.
     */
    public boolean hasNext() {
        return hasNext;
    }

    /**
     * Abort reading process of RecordFactoryInputStream.
     */
    public void abort() {
        this.abort = true;
    }

    /**
     * @return If current sheet is new work sheet.
     */
    public boolean isFirstRow() {
        return firstSheetRow;
    }

    /**
     * Entry sheet to process first row.
     */
    public void entryFirstRow() {
        this.firstSheetRow = true;
    }

    /**
     * Do something when process one row.
     */
    public void leaveFirstRow() {
        this.firstSheetRow = false;
    }
}
