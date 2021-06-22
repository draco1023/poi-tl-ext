package org.ddr.poi.latex;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.ddr.poi.math.MathMLUtils;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.Text;
import uk.ac.ed.ph.snuggletex.SnuggleEngine;
import uk.ac.ed.ph.snuggletex.SnuggleInput;
import uk.ac.ed.ph.snuggletex.SnuggleSession;
import uk.ac.ed.ph.snuggletex.definitions.CorePackageDefinitions;
import uk.ac.ed.ph.snuggletex.internal.util.XMLUtilities;
import uk.ac.ed.ph.snuggletex.utilities.DefaultTransformerFactoryChooser;

import java.io.IOException;

/**
 * LaTeX字符串渲染策略
 *
 * @author Draco
 * @since 2021-04-14
 */
public class LaTeXRenderPolicy extends AbstractRenderPolicy<String> {
    private SnuggleSession session;

    @Override
    protected boolean validate(String data) {
        if (StringUtils.isBlank(data)) {
            return false;
        }

        // https://www2.ph.ed.ac.uk/snuggletex/documentation/overview-and-features.html
        session = Initializer.SNUGGLE_ENGINE.createSession();
        SnuggleInput input = new SnuggleInput(data);
        boolean valid = false;
        try {
            valid = session.parseInput(input);
        } catch (IOException ignored) {
            // Will never throw an exception since input is raw string
        }
        return valid;
    }

    @Override
    public void doRender(RenderContext<String> context) throws Exception {
        XWPFParagraph paragraph = (XWPFParagraph) context.getRun().getParent();
        NodeList nodeList = session.buildDOMSubtree();
        int length = nodeList.getLength();
        for (int i = 0; i < length; i++) {
            Node node = nodeList.item(i);
            if (node instanceof Text) {
                paragraph.getCTP().addNewR().addNewT().setStringValue(node.getTextContent());
            } else if ("math".equals(node.getLocalName())) {
                String math = XMLUtilities.serializeNode(node,
                        Initializer.SNUGGLE_ENGINE.getDefaultXMLStringOutputOptions());

                MathMLUtils.renderTo(paragraph, context.getRun(), math);
            }
        }
    }

    private static class Initializer {
        static final SnuggleEngine SNUGGLE_ENGINE = new SnuggleEngine(DefaultTransformerFactoryChooser.getInstance(), null);

        static {
            CorePackageDefinitions.getPackage().loadMathCharacterAliases("math-character-aliases.txt");
        }
    }

    @Override
    protected void afterRender(RenderContext<String> context) {
        clearPlaceholder(context, false);
    }
}
