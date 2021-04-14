package org.ddr.poi.latex;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.ddr.poi.math.MathMLUtils;
import org.w3c.dom.NodeList;
import uk.ac.ed.ph.snuggletex.SnuggleEngine;
import uk.ac.ed.ph.snuggletex.SnuggleInput;
import uk.ac.ed.ph.snuggletex.SnuggleSession;
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
        NodeList nodeList = session.buildDOMSubtree();
        String math = XMLUtilities.serializeNode(nodeList.item(0),
                Initializer.SNUGGLE_ENGINE.getDefaultXMLStringOutputOptions());
        MathMLUtils.renderTo((XWPFParagraph) context.getRun().getParent(), math);
    }

    private static class Initializer {
        static final SnuggleEngine SNUGGLE_ENGINE = new SnuggleEngine(DefaultTransformerFactoryChooser.getInstance(), null);
    }

    @Override
    protected void afterRender(RenderContext<String> context) {
        clearPlaceholder(context, false);
    }
}
