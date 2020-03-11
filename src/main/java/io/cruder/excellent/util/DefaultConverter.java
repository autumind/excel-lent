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

import io.cruder.excellent.ExcelColumn;
import lombok.Data;
import lombok.experimental.Accessors;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

/**
 * DefaultConverter: Default implement of row data converter.
 *
 * @author cruder
 * @since 2019-04-11
 */
@Slf4j
public enum DefaultConverter implements Converter {
    /**
     * Singleton Converter
     */
    INSTANCE;

    /**
     * Class titled method cache.
     */
    private ConcurrentMap<String, List<TitledMethod>> classSetterCache = new ConcurrentHashMap<>();

    @Override
    @SuppressWarnings("unchecked")
    public <T> T convert(List<String> headers, List<String> rowCells, Class<T> clazz) {
        T bean = null;
        Map<String, Integer> headerIndex = IntStream.range(0, headers.size())
                .boxed()
                .collect(Collectors.toMap(headers::get, Function.identity()));
        try {

            if (clazz == null) {
                Map<String, String> map = new LinkedHashMap<>();
                headerIndex.forEach((header, index) -> map.put(header,
                        rowCells.size() > index ? rowCells.get(index) : Constant.EMPTY));
                return (T) map;
            }

            bean = clazz.newInstance();

            String cachedKey = clazz.getName();
            List<TitledMethod> titledMethods;
            if (classSetterCache.containsKey(cachedKey) && classSetterCache.get(cachedKey) != null) {
                titledMethods = classSetterCache.get(cachedKey);
            } else {
                titledMethods = Stream.of(clazz.getDeclaredFields())
                        .map(field -> {
                            String title = field.getName();
                            String format = null;
                            if (field.isAnnotationPresent(ExcelColumn.class)) {
                                ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                                if (!column.title().isEmpty()) {
                                    title = column.title();
                                }
                                if (!column.format().isEmpty()) {
                                    format = column.format();
                                }
                            }
                            return new TitledMethod()
                                    .setTitle(title)
                                    .setFormat(format)
                                    .setMethod(Reflects.resolveSetter(field, clazz));
                        })
                        .collect(Collectors.toList());
                classSetterCache.put(cachedKey, titledMethods);
            }

            if (titledMethods == null || titledMethods.isEmpty()) {
                return bean;
            }

            for (TitledMethod titledMethod : titledMethods) {
                Method method = titledMethod.getMethod();
                Class<?>[] parameterTypes = method.getParameterTypes();
                if (parameterTypes.length > 1) {
                    // setter should only have one parameter
                    throw new RuntimeException(String.format("Method %s is not a setter.", method.getName()));
                }
                Class parameterType = parameterTypes[0];
                Integer index = headerIndex.get(titledMethod.getTitle());
                String val = index < rowCells.size() ? rowCells.get(index) : null;
                if (val != null) {
                    // Invoke setter only when cell value is not null.
                    doInvocation(bean, titledMethod, parameterType, val);
                }
            }

        } catch (Exception e) {
            log.warn("Convert excel row value to bean field failure.", e);
        }
        return bean;
    }

    /**
     * Invoke bean set method to fill value.
     *
     * @param bean          bean
     * @param titledMethod  titled method
     * @param parameterType parameter type of set method
     * @param val           value
     */
    private <T> void doInvocation(T bean, TitledMethod titledMethod, Class<?> parameterType, String val)
            throws IllegalAccessException, InvocationTargetException, NoSuchMethodException, InstantiationException {
        Method method = titledMethod.getMethod();
        if (parameterType.equals(String.class) || parameterType.equals(Object.class)) {
            method.invoke(bean, val);
        } else if (Number.class.isAssignableFrom(parameterType)) {
            Constructor constructor = parameterType.getConstructor(String.class);
            if (constructor == null) {
                return;
            }
            if (parameterType.equals(Byte.class) || !val.isEmpty()) {
                method.invoke(bean, constructor.newInstance(val));
            }
        } else if (parameterType.isPrimitive()) {
            if (parameterType.equals(Void.TYPE)) {
                throw new RuntimeException("Field's type of target class must not be Void.class");
            }
            if (parameterType.equals(Character.TYPE)) {
                throw new RuntimeException("Please use String to replace field's type Character.class");
            }
            Object primitiveValue = getPrimitiveValue(parameterType, val);
            if (primitiveValue != null) {
                method.invoke(bean, primitiveValue);
            }
        } else if (parameterType.equals(BigDecimal.class)) {
            if (val.isEmpty()) {
                return;
            }
            method.invoke(bean, new BigDecimal(val));
        } else if (parameterType.equals(Date.class)) {
            Date date;
            Double excelDate = toExcelDate(val);
            if (excelDate != null) {
                date = DateUtil.getJavaDate(Double.parseDouble(val));
            } else {
                String format = titledMethod.getFormat();
                if (format == null) {
                    format = Constant.DEFAULT_DATETIME_FORMAT;
                }
                Instant instant = LocalDateTime.parse(val,
                        DateTimeFormatter.ofPattern(format)).atZone(ZoneId.systemDefault()).toInstant();
                date = Date.from(instant);
            }
            method.invoke(bean, date);
        } else if (parameterType.equals(LocalDateTime.class)) {
            LocalDateTime dateTime;
            Double excelDate = toExcelDate(val);
            if (excelDate == null) {
                String format = titledMethod.getFormat();
                if (format == null) {
                    format = Constant.DEFAULT_DATETIME_FORMAT;
                }
                dateTime = LocalDateTime.parse(val, DateTimeFormatter.ofPattern(format));
            } else {
                dateTime = DateUtil.getJavaDate(excelDate)
                        .toInstant()
                        .atZone(ZoneId.systemDefault())
                        .toLocalDateTime();
            }
            method.invoke(bean, dateTime);
        } else if (parameterType.equals(LocalDate.class)) {
            LocalDate date;
            Double excelDate = toExcelDate(val);
            if (excelDate == null) {
                String format = titledMethod.getFormat();
                if (format == null) {
                    format = Constant.DEFAULT_DATETIME_FORMAT;
                }
                date = LocalDate.parse(val, DateTimeFormatter.ofPattern(format));
            } else {
                date = DateUtil.getJavaDate(excelDate)
                        .toInstant()
                        .atZone(ZoneId.systemDefault())
                        .toLocalDate();
            }
            method.invoke(bean, date);
        }
    }

    /**
     * Get primitive value.
     *
     * @param primitiveType primitive type
     * @param val           string value
     * @return primitive value
     */
    private Object getPrimitiveValue(Class<?> primitiveType, String val) {
        if (primitiveType.equals(Byte.TYPE)) {
            return new Byte(val);
        }
        if (primitiveType.equals(Short.TYPE) && !val.isEmpty()) {
            return new BigDecimal(val).shortValue();
        }
        if (primitiveType.equals(Integer.TYPE) && !val.isEmpty()) {
            return new BigDecimal(val).intValue();
        }
        if (primitiveType.equals(Long.TYPE) && !val.isEmpty()) {
            return new BigDecimal(val).longValue();
        }
        if (primitiveType.equals(Float.TYPE) && !val.isEmpty()) {
            return new BigDecimal(val).floatValue();
        }
        if (primitiveType.equals(Double.TYPE) && !val.isEmpty()) {
            return new BigDecimal(val).doubleValue();
        }
        return null;
    }

    /**
     * String to excel date.
     *
     * @param val string value
     * @return excel date.
     */
    private Double toExcelDate(String val) {
        try {
            return Double.parseDouble(val);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * TitledMethod: Titled method which associated with excel field.
     *
     * @author cruder
     * @since 2019-04-11
     */
    @Data
    @Accessors(chain = true)
    private static class TitledMethod {
        /**
         * Method title
         */
        private String title;

        /**
         * Format
         */
        private String format;

        /**
         * Method
         */
        private Method method;
    }
}
