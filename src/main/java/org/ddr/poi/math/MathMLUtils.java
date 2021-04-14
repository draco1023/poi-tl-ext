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

import net.sf.saxon.s9api.Processor;
import net.sf.saxon.s9api.SaxonApiException;
import net.sf.saxon.s9api.Serializer;
import net.sf.saxon.s9api.Xslt30Transformer;
import net.sf.saxon.s9api.XsltCompiler;
import net.sf.saxon.s9api.XsltExecutable;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.transform.stream.StreamSource;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.io.StringWriter;

/**
 * MathML工具类
 *
 * @author Draco
 * @since 2021-04-10 22:56
 */
public class MathMLUtils {
    private static final Logger log = LoggerFactory.getLogger(MathMLUtils.class);
    private static final String MATH_FONT = "Cambria Math";

    /**
     * 将MathML渲染到段落中
     *
     * @param paragraph 段落
     * @param math MathML字符串
     */
    public static void renderTo(XWPFParagraph paragraph, String math) {
        if (log.isDebugEnabled()) {
            log.info("Start rendering MathML: {}", math);
        }
        try (StringReader sr = new StringReader(math);
             StringWriter sw = new StringWriter()) {
            Serializer out = newSerializer(sw);

            Initializer.TRANSFORMER.transform(new StreamSource(sr), out);

            String omath = sw.toString();
            if (log.isDebugEnabled()) {
                log.info("Output OMath: {}", omath);
            }
            addMath(paragraph, omath);
        } catch (IOException | SaxonApiException | XmlException e) {
            log.warn("Failed to render math: {}", math, e);
        }
    }

    /**
     * 用于惰性加载XSL转换器
     */
    private static class Initializer {
        static final Processor PROCESSOR = new Processor(false);
        static final Xslt30Transformer TRANSFORMER = createTransformer();

        private static Xslt30Transformer createTransformer() {
            XsltCompiler compiler = PROCESSOR.newXsltCompiler();
            XsltExecutable stylesheet;
            try (InputStream inputStream = MathMLUtils.class.getResourceAsStream("/MML2OMML.XSL")) {
                stylesheet = compiler.compile(new StreamSource(inputStream));
            } catch (IOException | SaxonApiException e) {
                throw new IllegalStateException("Failed to load MML2OMML.XSL", e);
            }
            return stylesheet.load30();
        }
    }

    private static Serializer newSerializer(StringWriter sw) {
        Serializer out = Initializer.PROCESSOR.newSerializer(sw);
        out.setOutputProperty(Serializer.Property.METHOD, "xml");
        out.setOutputProperty(Serializer.Property.INDENT, "no");
        out.setOutputProperty(Serializer.Property.OMIT_XML_DECLARATION, "yes");
        return out;
    }

    /**
     * 添加公式到Word
     *
     * @param paragraph 段落
     * @param omath 由mathml转换得到的omath字符串
     */
    private static void addMath(XWPFParagraph paragraph, String omath) throws XmlException {
        CTOMath ctoMath = CTOMath.Factory.parse(omath);
        // 老版本Office可能无法正常显示，强制设置公式字体
        XmlCursor xmlCursor = ctoMath.newCursor();
        while (xmlCursor.hasNextToken()) {
            XmlCursor.TokenType tokenType = xmlCursor.toNextToken();
            if (tokenType == XmlCursor.TokenType.START) {
                if (xmlCursor.getObject() instanceof CTR) {
                    CTFonts ctFonts = ((CTR) xmlCursor.getObject()).addNewRPr2().addNewRFonts();
                    ctFonts.setAscii(MATH_FONT);
                    ctFonts.setEastAsia(MATH_FONT);
                    ctFonts.setHAnsi(MATH_FONT);
                    ctFonts.setCs(MATH_FONT);
                }
            }
        }
        xmlCursor.dispose();

        CTP ctp = paragraph.getCTP();
        ctp.addNewOMath();
        ctp.setOMathArray(ctp.sizeOfOMathArray() - 1, ctoMath.getOMathArray(0));
    }
}
