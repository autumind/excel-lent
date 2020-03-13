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

package io.cruder.excellent.xssf;

import io.cruder.excellent.AbstractExcelReader;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.*;

/**
 * io.cruder.excellent.xssf.XlsxReader: Xlsx reader implement.
 *
 * @author cruder.io
 * @since 2020-03-11
 */
@Slf4j
public class XlsxReader<T> extends AbstractExcelReader<T> implements XSSFSheetXMLHandler.SheetContentsHandler {

    /**
     * Parse thread pool.
     */
    private final static ThreadPoolExecutor EXECUTOR = new ThreadPoolExecutor(
            (int) (Runtime.getRuntime().availableProcessors() * 0.6),
            Runtime.getRuntime().availableProcessors() * 2,
            15L,
            TimeUnit.SECONDS,
            new LinkedBlockingDeque<>(Runtime.getRuntime().availableProcessors() * 10),
            Executors.defaultThreadFactory(),
            new ThreadPoolExecutor.CallerRunsPolicy()
    );

    /**
     * XSSF
     */
    private OPCPackage opc;
    private String tempFilePath = "/tmp/excel/lent/";
    private BlockingQueue<List<String>> cachedRows = new LinkedBlockingQueue<>(10);

    /**
     * Current
     */
    private String currentFilePath;
    private CompletableFuture<?> currentFuture;
    private List<String> rowCells = new ArrayList<>();

    public XlsxReader(String filePath, InputStream inputStream, Class<T> clazz) throws OpenXML4JException, IOException {
        super(clazz);

        if (filePath == null || filePath.isEmpty()) {
            currentFilePath = tempFilePath
                    .concat(clazz.getSimpleName())
                    .concat(String.valueOf(System.currentTimeMillis()));
            File destFile = new File(currentFilePath);
            IOUtils.copy(inputStream, destFile);
            IOUtils.closeQuietly(inputStream);
        } else {
            currentFilePath = filePath;
        }
        opc = OPCPackage.open(currentFilePath, PackageAccess.READ);


        currentFuture = CompletableFuture.runAsync(() -> {
            XSSFReader xssfReader;
            try {
                xssfReader = new XSSFReader(opc);

                XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
                XMLReader xmlReader = SAXHelper.newXMLReader();
                ContentHandler handler = new XSSFSheetXMLHandler(
                        xssfReader.getStylesTable(),
                        null,
                        new ReadOnlySharedStringsTable(opc),
                        this,
                        new DataFormatter(),
                        false);
                xmlReader.setContentHandler(handler);
                while (iter.hasNext()) {
                    try (InputStream is = iter.next()) {
                        xmlReader.parse(new InputSource(is));
                    }
                }
            } catch (IOException | OpenXML4JException | SAXException | ParserConfigurationException e) {
                e.printStackTrace();
            }
        }, EXECUTOR);
    }

    @Override
    protected void finalize() throws Throwable {
        super.finalize();
        IOUtils.closeQuietly(opc);
    }

    @Override
    @SneakyThrows
    public T doRead() {
        if (!currentFuture.isDone() || !cachedRows.isEmpty()) {
            List<String> rowCells = cachedRows.take();
            if (firstRowAsHeader && !headerConfirmed) {
                headerConfirmed = true;
                headers.addAll(rowCells);
                return doRead();
            }
            return getConverter().convert(headers, rowCells, getParameterizedType());
        } else {
            // Delete temp file.
            Files.deleteIfExists(new File(currentFilePath).toPath());
            return null;
        }
    }

    @Override
    public void startRow(int rowNum) {
        rowCells.clear();
    }

    @Override
    @SneakyThrows
    public void endRow(int rowNum) {
        cachedRows.put(new ArrayList<>(rowCells));
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        rowCells.add(formattedValue);
    }

    @Override
    public void endSheet() {
    }

}
