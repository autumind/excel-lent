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

package io.cruder.excellent.xssf.xerces;

import lombok.extern.slf4j.Slf4j;
import org.apache.xerces.parsers.XIncludeAwareParserConfiguration;
import org.apache.xerces.xni.XNIException;
import org.apache.xerces.xni.parser.XMLInputSource;

import java.io.IOException;

/**
 * io.cruder.excellent.xssf.xerces.XMLConfiguration: Custom xml parser configuration.
 *
 * @author cruder.io
 * @since 2020-03-16
 */
@Slf4j
public class XmlParserConfiguration extends XIncludeAwareParserConfiguration {
    @Override
    public void parse(XMLInputSource source) throws XNIException, IOException {
        super.parse(source);
    }
}
