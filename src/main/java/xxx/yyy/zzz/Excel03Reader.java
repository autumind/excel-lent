package xxx.yyy.zzz;

import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Excel03Reader: Implement of reading excel 2003.
 *
 * @author autumind
 * @since 2019-03-22
 */
@Slf4j
public class Excel03Reader<T> extends AbstractExcelReader<T> implements HSSFListener, Cloneable {

    /**
     * Excel file
     */
    @Setter
    @Getter
    @Accessors(chain = true)
    private File file;

    /**
     * Row class.
     */
    @Setter
    @Getter
    @Accessors(chain = true)
    private Class<T> clz;

    /**
     * Input file stream
     */
    private InputStream fileInputStream;

    /**
     * poi file system handler
     */
    private POIFSFileSystem poifsFileSystem;

    /**
     * Cached rows when read multiple rows
     */
    private List<T> cacheRows;

    /**
     * Current data row.
     */
    private T currentRow;

    /**
     * Cached cell value of one actual data row.
     */
    private List<Object> cacheRowCells = new ArrayList<>();

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
    private boolean readNext = true;

    /**
     * If excel has next record or not.
     */
    private boolean hasNext = true;

    /**
     * Collection of brief sheet description.
     */
    private List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

    /**
     * Ordered sheet collection.
     */
    private BoundSheetRecord[] orderedBSRs;

    /**
     * Static String Table Record
     */
    private SSTRecord sstRecord;

    /**
     * For parsing Formulas.
     */
    private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;
    private HSSFWorkbook stubWorkbook;
    private FormatTrackingHSSFListener formatListener;

    /**
     * Sheet index.
     */
    private int sheetIndex = 0;

    /**
     * Parsed sheets.
     */
    private List<Sheet> sheets = new ArrayList<>();

    /**
     * For handling formulas with string results
     */
    private int nextRow;
    private int nextColumn;
    private boolean outputNextStringRecord;
    private int lastRowNumber;
    private int lastColumnNumber;

    /**
     * Construct excel 2003 reader.
     *
     * @param inputStream input stream
     * @throws IOException IO exception
     */
    public Excel03Reader(InputStream inputStream) throws IOException {
        fileInputStream = inputStream;
        poifsFileSystem = new POIFSFileSystem(fileInputStream);
        factory = new ELHSSFEventFactory();
        HSSFRequest request = new HSSFRequest();

        formatListener = new FormatTrackingHSSFListener(new MissingRecordAwareHSSFListener(this));
        if (outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }
        factory.setReq(request);

        Set<String> entryNames = poifsFileSystem.getRoot().getEntryNames();
        String name = Stream.of(InternalWorkbook.WORKBOOK_DIR_ENTRY_NAMES)
                .filter(entryNames::contains)
                .findFirst()
                .orElse(InternalWorkbook.WORKBOOK_DIR_ENTRY_NAMES[0]);

        factory.setDocumentInputStream(poifsFileSystem.createDocumentInputStream(name));
        factory.setRecordStream(
                new RecordFactoryInputStream(
                        factory.getDocumentInputStream(), false));
    }

    @Override
    public Optional<T> readRow() {
        if (!hasNext) {
            return Optional.empty();
        }

        readNext = true;
        factory.processEvents(documentInputStream -> {
            hasNext = false;
            try {
                fileInputStream.close();
                poifsFileSystem.close();
            } catch (Exception e) {
                log.error(e.getMessage(), e);
            } finally {
                documentInputStream.close();
            }
        }, () -> readNext);

        if (!hasNext) {
            return Optional.empty();
        }
        return Optional.ofNullable(currentRow);
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
    @SuppressWarnings("unchecked")
    public void processRecord(Record record) {
        int thisRow = -1;
        int thisColumn = -1;
        String thisStr = null;

        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add((BoundSheetRecord) record);
                break;
            case BOFRecord.sid:
                BOFRecord br = (BOFRecord) record;
                if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                    // Create sub workbook if required
                    if (workbookBuildingListener != null && stubWorkbook == null) {
                        stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
                    }

                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                    }
                    sheetIndex++;

                    Sheet sheet = new Sheet()
                            .setSheetNo(sheetIndex)
                            .setSheetName(orderedBSRs[sheetIndex - 1].getSheetname());
                    sheets.add(sheet);
                }
                break;

            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;

            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;

                thisRow = brec.getRow();
                thisColumn = brec.getColumn();
                thisStr = "";
                break;
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord) record;

                thisRow = berec.getRow();
                thisColumn = berec.getColumn();
                thisStr = "";
                break;

            case FormulaRecord.sid:
                FormulaRecord frec = (FormulaRecord) record;

                thisRow = frec.getRow();
                thisColumn = frec.getColumn();

                if (outputFormulaValues) {
                    if (Double.isNaN(frec.getValue())) {
                        // Formula result is a string
                        // This is stored in the next record
                        outputNextStringRecord = true;
                        nextRow = frec.getRow();
                        nextColumn = frec.getColumn();
                    } else {
                        thisStr = formatListener.formatNumberDateCell(frec);
                    }
                } else {
                    thisStr = HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression());
                }
                break;
            case StringRecord.sid:
                if (outputNextStringRecord) {
                    // String for formula
                    StringRecord srec = (StringRecord) record;
                    thisStr = srec.getString();
                    thisRow = nextRow;
                    thisColumn = nextColumn;
                    outputNextStringRecord = false;
                }
                break;

            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;

                thisRow = lrec.getRow();
                thisColumn = lrec.getColumn();
                thisStr = lrec.getValue();
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lsrec = (LabelSSTRecord) record;

                thisRow = lsrec.getRow();
                thisColumn = lsrec.getColumn();
                if (sstRecord == null) {
                    thisStr = "";
                } else {
                    thisStr = sstRecord.getString(lsrec.getSSTIndex()).toString();
                }
                break;
            case NoteRecord.sid:
                NoteRecord nrec = (NoteRecord) record;

                thisRow = nrec.getRow();
                thisColumn = nrec.getColumn();
                // TODO: Find object to match nrec.getShapeId()
                thisStr = "(TODO)";
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;

                thisRow = numrec.getRow();
                thisColumn = numrec.getColumn();

                // Format
                thisStr = formatListener.formatNumberDateCell(numrec);
                break;
            case RKRecord.sid:
                RKRecord rkrec = (RKRecord) record;

                thisRow = rkrec.getRow();
                thisColumn = rkrec.getColumn();
                thisStr = "";
                break;
            default:
                break;
        }

        // Handle new row
        if (thisRow != -1 && thisRow != lastRowNumber) {
            lastColumnNumber = -1;
        }

        // Handle missing column
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
            thisRow = mc.getRow();
            thisColumn = mc.getColumn();
            thisStr = "";
        }

        // If we got something to print out, do so
        if (thisStr != null) {
            cacheRowCells.add(thisStr);
        }

        // Update column and row count
        if (thisRow > -1) {
            lastRowNumber = thisRow;
        }
        if (thisColumn > -1) {
            lastColumnNumber = thisColumn;
        }

        if (record instanceof LastCellOfRowDummyRecord) {
            readNext = false;
            // We're onto a new row
            lastColumnNumber = -1;

            if (hasHead && heads.isEmpty()) {
                if (cacheRowCells.isEmpty()) {
                    throw new RuntimeException("Headed excel must has data in first row.");
                }
                heads.addAll(cacheRowCells.stream().map(String::valueOf).collect(Collectors.toList()));
                readNext = true;
                cacheRowCells.clear();
                return;
            }

            currentRow = null;
            if (clz == null) {
                HashMap<String, Object> rowMap = new HashMap<>();
                for (int i = 0; i < cacheRowCells.size(); i++) {
                    rowMap.put(convertNumber2Letter(i), cacheRowCells.get(i));
                }
                currentRow = (T) rowMap;
            } else {
                currentRow = converter.convert(cacheRowCells, clz);
            }

            cacheRowCells.clear();
        }
    }

    @Override
    public Iterator<T> iterator() {
        return new Iterator<T>() {

            /**
             * current row.
             */
            private T curr;

            /**
             * previous row.
             */
            private T prev;

            @Override
            public boolean hasNext() {
                Optional<T> optional = readRow();
                hasNext = optional.isPresent();
                if (hasNext) {
                    curr = optional.get();
                }
                return hasNext;
            }

            @Override
            public T next() {
                if (!hasNext) {
                    throw new NoSuchElementException();
                } else {
                    if (curr == null && !hasNext()) {
                        throw new NoSuchElementException();
                    }
                }
                prev = curr;
                curr = null;
                return prev;
            }
        };
    }
}
