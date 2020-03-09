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

package io.cruder.excellent.hssf;

import io.cruder.excellent.AbstractExcelReader;
import io.cruder.excellent.hssf.eventusermodel.HssfEventFactory;
import io.cruder.excellent.hssf.eventusermodel.HssfRequest;
import io.cruder.excellent.util.Reflects;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * io.cruder.excellent.hssf.XLS2CSVmra: Description of this class.
 *
 * @author cruder.io
 * @since 2020-03-02
 */
@Slf4j
public class XlsReader<T> extends AbstractExcelReader<T> implements Cloneable {

    /**
     * Should we output the formula, or the value it has?
     */
    private boolean outputFormulaValues = true;

    /**
     * For parsing Formulas
     */
    private SheetRecordCollectingListener workbookBuildingListener;
    private HSSFWorkbook stubWorkbook;

    /**
     * Records we pick up as we process
     */
    private SSTRecord sstRecord;
    private FormatTrackingHSSFListener formatListener;

    /**
     * So we known which sheet we're on
     */
    private int sheetIndex = -1;
    private BoundSheetRecord[] orderedBsrArray;
    private List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

    /**
     * For handling formulas with string results
     */
    private boolean outputNextStringRecord;

    /**
     * Cached one row's cell value.
     */
    private List<String> rowCells = new ArrayList<>();
    /**
     * Current row.
     */
    private T currRow;
    private HssfEventFactory hssf;

    /**
     * Creates a new XLS -> CSV converter
     *
     * @param fs The POIFSFileSystem to process
     */
    @SuppressWarnings("unchecked")
    public XlsReader(POIFSFileSystem fs, Class<T> clazz) {
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        formatListener = new FormatTrackingHSSFListener(listener);
        HssfRequest request = new HssfRequest();
        if (outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new SheetRecordCollectingListener(formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }
        hssf = new HssfEventFactory(request, fs);
        if (clazz == null) {
            parameterType = (Class<T>) Reflects.resolveGenericType(getClass());
        }
    }

    /**
     * Creates a new XLS reader.
     *
     * @param filename The file to process
     * @param clazz    parameterized type
     * @throws IOException IO exception
     */
    public XlsReader(String filename, Class<T> clazz) throws IOException {
        this(new POIFSFileSystem(new FileInputStream(filename)), clazz);
    }

    @Override
    public void processRecord(org.apache.poi.hssf.record.Record record) {
        String cellValue = null;

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

                    // Output the worksheet name
                    // Works by ordering the BSRs by the location of
                    //  their BOFRecords, and then knowing that we
                    //  process BOFRecords in byte offset order
                    sheetIndex++;
                    if (orderedBsrArray == null) {
                        orderedBsrArray = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                    }
                    log.info(orderedBsrArray[sheetIndex].getSheetname() + " [" + (sheetIndex + 1) + "]:");
                }
                break;

            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;

            case BlankRecord.sid:
            case BoolErrRecord.sid:

                cellValue = "";
                break;

            case FormulaRecord.sid:
                FormulaRecord frec = (FormulaRecord) record;

                if (outputFormulaValues) {
                    if (Double.isNaN(frec.getValue())) {
                        // Formula result is a string
                        // This is stored in the next record
                        outputNextStringRecord = true;
                    } else {
                        cellValue = formatListener.formatNumberDateCell(frec);
                    }
                } else {
                    cellValue = '"' +
                            HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression()) + '"';
                }
                break;
            case StringRecord.sid:
                if (outputNextStringRecord) {
                    // String for formula
                    StringRecord srec = (StringRecord) record;
                    cellValue = srec.getString();
                    outputNextStringRecord = false;
                }
                break;

            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;

                cellValue = '"' + lrec.getValue() + '"';
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lsrec = (LabelSSTRecord) record;

                if (sstRecord == null) {
                    cellValue = '"' + "(No SST Record, can't identify string)" + '"';
                } else {
                    cellValue = '"' + sstRecord.getString(lsrec.getSSTIndex()).toString() + '"';
                }
                break;
            case NoteRecord.sid:

                // TODO: Find object to match nrec.getShapeId()
                cellValue = '"' + "(TODO)" + '"';
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;

                // Format
                cellValue = formatListener.formatNumberDateCell(numrec);
                break;
            case RKRecord.sid:

                cellValue = '"' + "(TODO)" + '"';
                break;
            default:
                break;
        }

        // Handle missing column
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
            cellValue = "";
        }

        // If we got something to print out, do so
        if (cellValue != null) {
            rowCells.add(cellValue);
        }

        // Update column and row count

        // Handle end of row
        if (record instanceof LastCellOfRowDummyRecord) {

            // We're onto a new row
            hssf.abort();
            // End the row
            log.info("{}", rowCells);
            // Convert cell to entity.
            currRow = (T) new Object();
            rowCells.clear();
        }

    }

    @Override
    public Optional<T> readRow() {
        boolean hasNext = hssf.hasNext();
        if (hasNext) {
            hssf.nextRow();
            return Optional.ofNullable(currRow);
        }
        return Optional.empty();
    }

    @Override
    public Optional<List<T>> readRow(int rowNum) {
        List<T> cacheRows = new ArrayList<>();
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
    public Iterator<T> iterator() {
        return new Iterator<T>() {

            /**
             * current row.
             */
            private T curr;

            /**
             * If there is next record.
             */
            private boolean hasNext = true;

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
                T prev = curr;
                curr = null;
                return prev;
            }
        };
    }
}
