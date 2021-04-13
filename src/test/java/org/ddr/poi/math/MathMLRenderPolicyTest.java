/*
 * Copyright 2016 - 2021 Draco, https://github.com/draco1023
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ddr.poi.math;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import org.ddr.poi.FileReader;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Draco
 * @since 2021-04-11 8:33
 */
class MathMLRenderPolicyTest {

    @Test
    void doRender() throws IOException {
        MathMLRenderPolicy mathMLRenderPolicy = new MathMLRenderPolicy();
        Configure configure = Configure.builder()
                .bind("math1", mathMLRenderPolicy)
                .bind("math2", mathMLRenderPolicy)
                .bind("math3", mathMLRenderPolicy)
                .bind("math4", mathMLRenderPolicy)
                .build();
        List<String> mathList = new ArrayList<String>(4);
        for (int i = 0; i < 4; i++) {
            mathList.add(FileReader.readFile("/" + i + ".xml"));
        }
        Collections.shuffle(mathList);
        Map<String, Object> data = new HashMap<>(mathList.size());
        // FIXME 公式单独占用一个段落会被清空
        for (int i = 0; i < mathList.size(); i++) {
            data.put("math" + (i + 1), mathList.get(i));
        }
        try (InputStream inputStream = MathMLRenderPolicyTest.class.getResourceAsStream("/math.docx")) {
            XWPFTemplate.compile(inputStream, configure).render(data).writeToFile("math_out.docx");
        }
    }
}