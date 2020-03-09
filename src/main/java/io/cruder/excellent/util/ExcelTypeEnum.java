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

package io.cruder.excellent.util;

import io.cruder.excellent.exception.FileNotSupportedException;
import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.IOException;
import java.io.InputStream;

/**
 * ExcelTypeEnum: Enum of excel type.
 *
 * @author cruder
 * @since 2019-03-27
 */
public enum ExcelTypeEnum {
    /**
     * XSSF
     */
    XLSX,

    /**
     * HSSF
     */
    XLS;

    ExcelTypeEnum() {
    }

    public static ExcelTypeEnum valueOf(InputStream inputStream) {
        try {
            if (!inputStream.markSupported()) {
                inputStream = FileMagic.prepareToCheckMagic(inputStream);
            }
            FileMagic fileMagic = FileMagic.valueOf(inputStream);
            if (FileMagic.OLE2.equals(fileMagic)) {
                return XLS;
            }
            if (FileMagic.OOXML.equals(fileMagic)) {
                return XLSX;
            }
            throw new FileNotSupportedException("The file is not xls or xlsx or csv, please re-select.");
        } catch (IOException e) {
            throw new FileNotSupportedException(e);
        }
    }
}
