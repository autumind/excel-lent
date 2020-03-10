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

import io.cruder.excellent.Column;
import lombok.Data;
import lombok.experimental.Accessors;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Method;
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
                headerIndex.forEach((header, index) -> map.put(header, rowCells.size() > index ? rowCells.get(index) : ""));
                return (T) map;
            }

            bean = clazz.newInstance();

            String cachedKey = clazz.getName();
            List<TitledMethod> titledMethods;
            if (classSetterCache.containsKey(cachedKey) && classSetterCache.get(cachedKey) != null) {
                titledMethods = classSetterCache.get(cachedKey);
            } else {
                titledMethods = Stream.of(clazz.getDeclaredFields())
                        .filter(field -> field.isAnnotationPresent(Column.class))
                        .map(field -> new TitledMethod()
                                .setTitle(field.getAnnotation(Column.class).title())
                                .setMethod(Reflects.resolveSetter(field, clazz)))
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
                Object val = rowCells.get(headerIndex.get(titledMethod.getTitle()));
                if (val != null) {
                    // invoke setter only when cell value is not null.
                    if (parameterType.equals(String.class)) {
                        method.invoke(bean, String.valueOf(val));
                    }
                }
            }

        } catch (Exception e) {
            log.error("Convert excel row value to java object failure.", e);
        }
        return bean;
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
         * method title
         */
        private String title;

        /**
         * method
         */
        private Method method;
    }
}
