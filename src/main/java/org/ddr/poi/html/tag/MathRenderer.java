package org.ddr.poi.html.tag;

import net.sf.saxon.s9api.Processor;
import net.sf.saxon.s9api.SaxonApiException;
import net.sf.saxon.s9api.Serializer;
import net.sf.saxon.s9api.Xslt30Transformer;
import net.sf.saxon.s9api.XsltCompiler;
import net.sf.saxon.s9api.XsltExecutable;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.jsoup.nodes.Element;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.transform.stream.StreamSource;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.io.StringWriter;
import java.util.List;

/**
 * math标签渲染器
 *
 * @author Draco
 * @since 2021-02-18
 */
public class MathRenderer implements ElementRenderer {
    private static final Logger log = LoggerFactory.getLogger(MathRenderer.class);

    private static final String[] TAGS = {HtmlConstants.TAG_MATH};
    private static final String MATH_FONT = "Cambria Math";

    /**
     * 用于惰性加载XSL转换器
     */
    private static class Initializer {
        public static final Processor PROCESSOR = new Processor(false);
        public static final Xslt30Transformer TRANSFORMER = createTransformer();

        private static Xslt30Transformer createTransformer() {
            XsltCompiler compiler = PROCESSOR.newXsltCompiler();
            XsltExecutable stylesheet;
            try (InputStream inputStream = MathRenderer.class.getResourceAsStream("/MML2OMML.XSL")) {
                stylesheet = compiler.compile(new StreamSource(inputStream));
            } catch (IOException | SaxonApiException e) {
                throw new IllegalStateException("Failed to load MML2OMML.XSL", e);
            }
            return stylesheet.load30();
        }
    }

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        String math = element.outerHtml();

        try (StringReader sr = new StringReader(math);
             StringWriter sw = new StringWriter()) {
            Serializer out = newSerializer(sw);

            Initializer.TRANSFORMER.transform(new StreamSource(sr), out);

            addMath(context.getClosestParagraph(), sw.toString());
        } catch (IOException | SaxonApiException | XmlException e) {
            log.warn("Failed to render math: {}", math, e);
        }

        return false;
    }

    @Override
    public String[] supportedTags() {
        return TAGS;
    }

    @Override
    public boolean renderAsBlock() {
        return false;
    }

    private Serializer newSerializer(StringWriter sw) {
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
    private void addMath(XWPFParagraph paragraph, String omath) throws XmlException {
        CTOMath omathPara = CTOMath.Factory.parse(omath);
        // 老版本Office可能无法正常显示，强制设置公式字体
        XmlCursor xmlCursor = omathPara.newCursor();
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

        List<CTOMath> oMathList = paragraph.getCTP().getOMathList();
        oMathList.add(omathPara.getOMathArray(0));
    }
}
