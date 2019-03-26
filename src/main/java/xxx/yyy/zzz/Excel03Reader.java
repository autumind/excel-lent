package xxx.yyy.zzz;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RecordFactoryInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Stream;

/**
 * Excel03Reader: Implement of reading excel 2003.
 *
 * @author autumind
 * @since 2019-03-22
 */
@Slf4j
public class Excel03Reader<T> implements IExcelReader<T>, HSSFListener {

    /**
     * Input file stream
     */
    private FileInputStream fileInputStream;

    /**
     * poi file system handler
     */
    private POIFSFileSystem poifsFileSystem;

    /**
     * Cached rows when read multiple rows
     */
    private List<T> cacheRows;

    /**
     * Cached cell value of one actual data row.
     */
    private List<Object> cacheRowCells;

    /**
     * Data sheet name
     */
    private String sheetName;

    /**
     * Data head.
     */
    private List<String> heads;

    /**
     * Data head row index. -1 represents no head information.
     */
    private int headRowIndex = -1;

    /**
     * Should we output the formula, or the value it has?
     */
    private boolean outputFormulaValues = true;

    /**
     * Custom HSSFEventFactory to read rows continuously.
     */
    private ELHSSFEventFactory factory;

    /**
     * If read next record or not.
     */
    private boolean ifNext = true;

    /**
     * Open file.
     *
     * @param file file
     * @param <E>  data template
     * @return Excel03Reader
     * @throws IOException IO exception
     */
    public static <E> Excel03Reader<E> open(File file) throws IOException {
        Excel03Reader<E> reader = new Excel03Reader<>();
        reader.fileInputStream = new FileInputStream(file);
        reader.poifsFileSystem = new POIFSFileSystem(reader.fileInputStream);
        reader.factory = new ELHSSFEventFactory();
        HSSFRequest request = new HSSFRequest();

        FormatTrackingHSSFListener formatListener
                = new FormatTrackingHSSFListener(new MissingRecordAwareHSSFListener(reader));
        if (reader.outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            request.addListenerForAllRecords(new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener));
        }
        reader.factory.setReq(request);

        Set<String> entryNames = reader.poifsFileSystem.getRoot().getEntryNames();
        String name = Stream.of(InternalWorkbook.WORKBOOK_DIR_ENTRY_NAMES)
                .filter(entryNames::contains)
                .findFirst()
                .orElse(InternalWorkbook.WORKBOOK_DIR_ENTRY_NAMES[0]);

        reader.factory.setDocumentInputStream(reader.poifsFileSystem.createDocumentInputStream(name));
        reader.factory.setRecordStream(new RecordFactoryInputStream(
                reader.factory.getDocumentInputStream(), false));
        return reader;
    }

    /**
     * Set sheet name which needs to be parse.
     *
     * @param sheetName sheet name
     * @return reader
     */
    public Excel03Reader<T> sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    /**
     * Set index of head row.
     *
     * @param headRowIndex head row index
     * @return reader
     */
    public Excel03Reader<T> headRow(int headRowIndex) {
        this.headRowIndex = headRowIndex;
        return this;
    }


    @Override
    public Optional<T> readRow() {
        ifNext = true;
        factory.processEvents(documentInputStream -> {
            try {
                fileInputStream.close();
                poifsFileSystem.close();
            } catch (Exception e) {
                log.error(e.getMessage(), e);
            } finally {
                documentInputStream.close();
            }
        }, () -> ifNext);
        return Optional.ofNullable(null);
    }

    @Override
    public Optional<List<T>> readRow(int rowNum) {
        if (cacheRows == null) {
            cacheRows = new ArrayList<>();
        } else {
            cacheRows.clear();
        }
        for (int i = 0; i < rowNum; i++) {
            Optional<T> optional = readRow();
            if (!optional.isPresent()) {
                break;
            }
            cacheRows.add(optional.get());
        }
        return Optional.of(cacheRows);
    }

    @Override
    public void processRecord(Record record) {
        if (record instanceof LastCellOfRowDummyRecord) {
            log.info("Wowwwwwwwwwwwwww...{}", record);
            ifNext = false;
        } else {
            log.info("Record: {}", record);
        }
    }
}
