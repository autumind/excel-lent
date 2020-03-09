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
import java.util.Set;

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
     * Creates a new instance of HSSFEventFactory
     */
    public HssfEventFactory() {
        // no instance fields
    }

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
     * Read next row.
     */
    @SneakyThrows
    public void nextRow() {
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
    public void processWorkbookEvents(HssfRequest req, DirectoryNode dir) throws IOException {
        // some old documents have "WORKBOOK" or "BOOK"
        String name = null;
        Set<String> entryNames = dir.getEntryNames();
        for (String potentialName : WORKBOOK_DIR_ENTRY_NAMES) {
            if (entryNames.contains(potentialName)) {
                name = potentialName;
                break;
            }
        }
        // If in doubt, go for the default
        if (name == null) {
            name = WORKBOOK_DIR_ENTRY_NAMES[0];
        }

        if (inputStream == null) {
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
    public void processEvents(HssfRequest req, InputStream in) {
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
     * @return numeric user-specified result code.
     * @see org.apache.poi.poifs.filesystem.POIFSFileSystem#createDocumentInputStream(String)
     */
    private short genericProcessEvents(HssfRequest req, InputStream in)
            throws HSSFUserException {

        short userCode = 0;

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

        // All done, return our last code
        return userCode;
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
}
