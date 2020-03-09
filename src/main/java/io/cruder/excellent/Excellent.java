package io.cruder.excellent;

import io.cruder.excellent.exception.FileNotSupportedException;
import io.cruder.excellent.hssf.XlsReader;
import io.cruder.excellent.util.ExcelTypeEnum;
import lombok.experimental.UtilityClass;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

/**
 * io.cruder.excellent.Excel: An excel tool name.
 *
 * @author cruder.io
 * @since 2020-03-02
 */
@UtilityClass
public class Excellent {
    /**
     * Open excel reader.
     *
     * @param file  file
     * @param clazz row class
     * @return ExcelField reader.
     */
    public static <E> Reader<E> open(String file, Class<E> clazz) {
        try (BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file))) {
            ExcelTypeEnum type = ExcelTypeEnum.valueOf(bis);
            if (type == ExcelTypeEnum.XLS) {
                XlsReader<E> reader = new XlsReader<>(file, clazz);
                reader.setParameterType(clazz);
                return reader;
            }
        } catch (IOException e) {
            throw new FileNotSupportedException("The file is not xls or xlsx or csv, please re-select.");
        }

        return null;
    }

    /**
     * Open excel reader.
     *
     * @param file file
     * @return ExcelField reader.
     */
    @SuppressWarnings("unchecked")
    public static <E> Reader<E> open(String file) {
        return (Reader<E>) open(file, Map.class);
    }

    /**
     * Open excel reader.
     *
     * @param inputStream input stream
     * @return ExcelField reader.
     */
    @SuppressWarnings("unchecked")
    public static <E> Reader<E> open(InputStream inputStream) {
        return (Reader<E>) open(inputStream, Map.class);
    }

    /**
     * Open excel reader.
     *
     * @param inputStream input stream
     * @param clazz       row class
     * @return ExcelField reader.
     */
    public static <E> Reader<E> open(InputStream inputStream, Class<E> clazz) {
        try (BufferedInputStream bis = new BufferedInputStream(inputStream)) {
            ExcelTypeEnum type = ExcelTypeEnum.valueOf(bis);
            if (type == ExcelTypeEnum.XLS) {
                XlsReader<E> reader = new XlsReader<>(new POIFSFileSystem(bis), clazz);
                reader.setParameterType(clazz);
                return reader;
            }
        } catch (IOException e) {
            throw new FileNotSupportedException("The file inputStream not xls or xlsx or csv, please re-select.");
        }

        return null;
    }

}
