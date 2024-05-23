package org.ddr.poi.latex;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Element;
import uk.ac.ed.ph.snuggletex.dombuilding.CommandHandler;
import uk.ac.ed.ph.snuggletex.internal.DOMBuilder;
import uk.ac.ed.ph.snuggletex.internal.SnuggleParseException;
import uk.ac.ed.ph.snuggletex.tokens.CommandToken;

public class TextCircledHandler implements CommandHandler {

    private static final Logger log = LoggerFactory.getLogger(TextCircledHandler.class);

    @Override
    public void handleCommand(DOMBuilder builder, Element parentElement, CommandToken token) throws SnuggleParseException {
        String s = builder.extractStringValue(token.getArguments()[0]);
        String replacement = LaTeXUtils.textCircledMap.get(s);
        if (replacement != null) {
            builder.appendMathMLTextElement(parentElement, "mi", replacement, true);
        } else {
            log.warn("Text circled not found: {}", s);
        }
    }
}
