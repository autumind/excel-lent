package xxx.yyy.zzz;

import lombok.Data;
import lombok.experimental.Accessors;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * DefaultRowConverter: Default implement of row data converter.
 *
 * @author autumind
 * @since 2019-04-11
 */
@Slf4j
public enum DefaultRowConverter implements RowConverter {
    /**
     * Singleton Converter
     */
    INSTANCE;

    /**
     * Class titled method cache.
     */
    private ConcurrentMap<String, List<TitledMethod>> classSetterCache = new ConcurrentHashMap<>();

    @Override
    public <T> T convert(Map<String, Integer> headIndex, List<Object> rowCells, Class<T> clz) {
        T bean = null;
        try {
            bean = clz.newInstance();

            String cachedKey = clz.getName();
            List<TitledMethod> titledMethods;
            if (classSetterCache.containsKey(cachedKey) && classSetterCache.get(cachedKey) != null) {
                titledMethods = classSetterCache.get(cachedKey);
            } else {
                titledMethods = Stream.of(clz.getDeclaredFields())
                        .filter(field -> field.isAnnotationPresent(ExcelField.class))
                        .map(field -> new TitledMethod()
                                .setTitle(field.getAnnotation(ExcelField.class).title())
                                .setMethod(ReflectUtil.resolveSetter(field, clz)))
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
                Object val = rowCells.get(headIndex.get(titledMethod.getTitle()));
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
     * @author autumind
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
