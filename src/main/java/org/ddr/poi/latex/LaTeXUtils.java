package org.ddr.poi.latex;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.ddr.poi.html.util.Colors;
import org.ddr.poi.math.MathMLUtils;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.Text;
import uk.ac.ed.ph.snuggletex.InputError;
import uk.ac.ed.ph.snuggletex.SnuggleEngine;
import uk.ac.ed.ph.snuggletex.SnuggleInput;
import uk.ac.ed.ph.snuggletex.SnuggleSession;
import uk.ac.ed.ph.snuggletex.definitions.CorePackageDefinitions;
import uk.ac.ed.ph.snuggletex.definitions.Globals;
import uk.ac.ed.ph.snuggletex.definitions.LaTeXMode;
import uk.ac.ed.ph.snuggletex.dombuilding.MathComplexCommandHandler;
import uk.ac.ed.ph.snuggletex.internal.util.XMLUtilities;
import uk.ac.ed.ph.snuggletex.utilities.DefaultTransformerFactoryChooser;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.ConcurrentHashMap;

/**
 * LaTeX工具类
 *
 * @author Draco
 * @since 2023-07-17
 */
public class LaTeXUtils {
    private static final Logger log = LoggerFactory.getLogger(LaTeXUtils.class);

    public static final String TAG_MATH = "math";
    public static final String TAG_TAG = "tag";

    static final ConcurrentHashMap<String, String> textCircledMap = new ConcurrentHashMap<>();

    /**
     * 创建Snuggle会话
     */
    public static SnuggleSession createSession() {
        return Initializer.SNUGGLE_ENGINE.createSession();
    }

    /**
     * 解析字符串
     *
     * @param session Snuggle会话
     * @param data    LaTeX字符串
     * @return 是否为有效的内容
     */
    public static boolean parse(SnuggleSession session, String data) {
        SnuggleInput input = new SnuggleInput(data);
        boolean valid = false;
        try {
            valid = session.parseInput(input);
        } catch (IOException ignored) {
            // Will never throw an exception since input is raw string
        }
        if (CollectionUtils.isNotEmpty(session.getErrors())) {
            log.warn("Invalid LaTex: {}", data);
            for (InputError error : session.getErrors()) {
                log.warn("LaTeX parse error: {}", error);
            }
        }
        return valid;
    }

    /**
     * 将LaTeX渲染到段落中
     *
     * @param paragraph 段落
     * @param ctr       目标run，如果总是在末尾渲染可传null
     * @param session   Snuggle会话
     */
    public static void renderTo(XWPFParagraph paragraph, CTR ctr, SnuggleSession session) {
        CTR target = ctr;
        NodeList nodeList = session.buildDOMSubtree();
        int length = nodeList.getLength();
        for (int i = 0; i < length; i++) {
            Node node = nodeList.item(i);
            if (node instanceof Text) {
                target = paragraph.getCTP().addNewR();
                target.addNewT().setStringValue(node.getTextContent());
            } else if (TAG_MATH.equals(node.getLocalName())) {
                String math = XMLUtilities.serializeNode(node,
                        Initializer.SNUGGLE_ENGINE.getDefaultXMLStringOutputOptions());

                MathMLUtils.renderTo(paragraph, target, math);
            } else if (TAG_TAG.equals(node.getLocalName())) {
                renderTag(paragraph, ctr, node);
            }
        }
    }

    private static void renderTag(XWPFParagraph paragraph, CTR ctr, Node node) {
        XmlCursor pCursor = paragraph.getCTP().newCursor();
        pCursor.push();

        XWPFTable xwpfTable = paragraph.getBody().insertNewTbl(pCursor);
        // 100% width
        xwpfTable.setWidth(5000);
        xwpfTable.setWidthType(TableWidthType.PCT);
        // no borders
        xwpfTable.setLeftBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, Colors.BLACK);
        xwpfTable.setTopBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, Colors.BLACK);
        xwpfTable.setRightBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, Colors.BLACK);
        xwpfTable.setBottomBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, Colors.BLACK);
        xwpfTable.setInsideHBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, Colors.BLACK);
        xwpfTable.setInsideVBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, Colors.BLACK);
        // access to the first row of the new table which created in a table cell will lead to a weird exception
        xwpfTable.removeRow(0);
        xwpfTable.addNewCol();
        xwpfTable.addNewCol();
        XWPFTableRow row = xwpfTable.getRow(0);
        XWPFTableCell mathCell = row.getCell(0);
        XWPFParagraph mathParagraph = mathCell.getParagraphs().get(0);
        XmlCursor copyCursor = mathParagraph.getCTP().newCursor();
        copyCursor.toEndToken();

        pCursor.pop();
        pCursor.toFirstChild();
        boolean hasNextSibling = true;
        while (hasNextSibling) {
            XmlObject obj = pCursor.getObject();
            if (obj instanceof CTPPr) {
                pCursor.copyXml(copyCursor);
                hasNextSibling = pCursor.toNextSibling();
            } else if (obj == null || obj instanceof CTBookmark || ctr.equals(obj)) {
                hasNextSibling = pCursor.toNextSibling();
            } else {
                // moveXml附带了toNextSibling的效果
                hasNextSibling = pCursor.moveXml(copyCursor);
            }
        }
        copyCursor.dispose();
        pCursor.dispose();

        // render math
        String math = XMLUtilities.serializeNode(node.getFirstChild(),
                Initializer.SNUGGLE_ENGINE.getDefaultXMLStringOutputOptions());
        MathMLUtils.renderTo(mathParagraph, mathParagraph.createRun().getCTR(), math);

        // render tag
        XWPFTableCell tagCell = row.getCell(1);
        tagCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph tagParagraph = tagCell.getParagraphs().get(0);
        String tag = XMLUtilities.serializeNode(node.getLastChild(),
                Initializer.SNUGGLE_ENGINE.getDefaultXMLStringOutputOptions());
        MathMLUtils.renderTo(tagParagraph, tagParagraph.createRun().getCTR(), tag);
    }

    private static class Initializer {
        static final SnuggleEngine SNUGGLE_ENGINE = new SnuggleEngine(DefaultTransformerFactoryChooser.getInstance(), null);

        static {
            CorePackageDefinitions.getPackage().loadMathCharacterAliases("math-character-aliases.txt");
            CorePackageDefinitions.getPackage().addComplexCommandSameArgMode("dfrac", false, 2, Globals.MATH_MODE_ONLY, new MathComplexCommandHandler("mfrac"), null);

            try (InputStream is = Globals.class.getClassLoader().getResourceAsStream("math-character-circled.txt");
                 InputStreamReader isr = new InputStreamReader(is, StandardCharsets.UTF_8);
                 BufferedReader br = new BufferedReader(isr)) {
                String line;
                while ((line = br.readLine()) != null) {
                    if (line.startsWith("#")) {
                        continue;
                    }
                    line = line.replaceFirst("\\s+#.+$", "");
                    String[] fields = line.split("->");
                    textCircledMap.put(fields[0], fields[1].trim());
                }
            } catch (IOException e) {
                log.warn("Failed to load math-character-circled.txt", e);
            }
            CorePackageDefinitions.getPackage().addComplexCommandOneArg("textcircled", false, Globals.ALL_MODES, LaTeXMode.LR, new TextCircledHandler(), null);
            CorePackageDefinitions.getPackage().addComplexCommandOneArg("tag", false, Globals.ALL_MODES, LaTeXMode.LR, new TagHandler(), null);
        }
    }
}
