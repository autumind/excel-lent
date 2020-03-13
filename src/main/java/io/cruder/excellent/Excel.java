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

package io.cruder.excellent;

import io.cruder.excellent.exception.FileNotSupportedException;
import io.cruder.excellent.hssf.XlsReader;
import io.cruder.excellent.util.ExcelTypeEnum;
import io.cruder.excellent.xssf.XlsxReader;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Map;

/**
 * io.cruder.excellent.Excel: An excel tool name.
 *
 * @author cruder.io
 * @since 2020-03-02
 */
@UtilityClass
public class Excel {
    /**
     * Open excel reader.
     *
     * @param file file
     * @return ExcelField reader.
     */
    public static Reader<Map<String, String>> lent(String file) {
        return lent(file, null);
    }

    /**
     * Open excel reader.
     *
     * @param file  file
     * @param clazz target class
     * @return ExcelField reader.
     */
    @SneakyThrows
    public static <E> Reader<E> lent(String file, Class<E> clazz) {
        return lent(file, new FileInputStream(file), clazz);
    }

    /**
     * Open excel reader.
     *
     * @param file file
     * @return ExcelField reader.
     */
    public static Reader<Map<String, String>> lent(File file) {
        return lent(file, null);
    }

    /**
     * Open excel reader.
     *
     * @param file  file
     * @param clazz target class
     * @return ExcelField reader.
     */
    @SneakyThrows
    public static <E> Reader<E> lent(File file, Class<E> clazz) {
        return lent(file.getPath(), new FileInputStream(file), clazz);
    }

    /**
     * Open excel reader.
     *
     * @param inputStream input stream
     * @return ExcelField reader.
     */
    @SneakyThrows
    public static Reader<Map<String, String>> lent(InputStream inputStream) {
        return lent(null, inputStream, null);
    }

    /**
     * Open excel reader.
     *
     * @param filePath    file path
     * @param inputStream input stream
     * @param clazz       target class
     * @param <E>
     * @return ExcelField reader.
     */
    public static <E> Reader<E> lent(String filePath, InputStream inputStream, Class<E> clazz) {
        try (BufferedInputStream bis = new BufferedInputStream(inputStream)) {
            ExcelTypeEnum type = ExcelTypeEnum.valueOf(bis);
            if (type == ExcelTypeEnum.XLS) {
                return new XlsReader<>(bis, clazz);
            } else if (type == ExcelTypeEnum.XLSX) {
                return new XlsxReader<>(filePath, bis, clazz);
            } else {
                return null;
            }
        } catch (Exception e) {
            throw new FileNotSupportedException("The file not xls or xlsx or csv, please re-select.", e);
        }
    }

}
