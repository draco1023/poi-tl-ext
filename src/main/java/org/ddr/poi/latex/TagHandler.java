package org.ddr.poi.latex;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import uk.ac.ed.ph.snuggletex.definitions.CorePackageDefinitions;
import uk.ac.ed.ph.snuggletex.definitions.W3CConstants;
import uk.ac.ed.ph.snuggletex.dombuilding.CommandHandler;
import uk.ac.ed.ph.snuggletex.internal.DOMBuilder;
import uk.ac.ed.ph.snuggletex.internal.SnuggleParseException;
import uk.ac.ed.ph.snuggletex.tokens.ArgumentContainerToken;
import uk.ac.ed.ph.snuggletex.tokens.CommandToken;
import uk.ac.ed.ph.snuggletex.tokens.FlowToken;
import uk.ac.ed.ph.snuggletex.tokens.MathCharacterToken;

import java.util.List;

public class TagHandler implements CommandHandler {
    @Override
    public void handleCommand(DOMBuilder builder, Element parentElement, CommandToken token) throws SnuggleParseException {
        Document document = builder.getDocument();
        Element table = document.createElementNS(W3CConstants.XHTML_NAMESPACE, LaTeXUtils.TAG_TAG);

        Node math;
        if (builder.isBuildingMathMLIsland()) {
            math = parentElement;
            while (math != null) {
                if (LaTeXUtils.TAG_MATH.equals(math.getLocalName()) && builder.getDocument().getDocumentElement().equals(math.getParentNode())) {
                    break;
                }
                math = math.getParentNode();
            }
        } else {
            math = parentElement.getLastChild();
            while (math != null) {
                if (LaTeXUtils.TAG_MATH.equals(math.getLocalName())) {
                    break;
                }
                math = math.getPreviousSibling();
            }
        }
        if (math != null) {
            Node parentNode = math.getParentNode();
            parentNode.removeChild(math);
            parentNode.appendChild(table);
            table.appendChild(math);

            ArgumentContainerToken argument = token.getArguments()[0];
            List<FlowToken> contents = argument.getContents();
            contents.add(0, new MathCharacterToken(token.getSlice(), CorePackageDefinitions.getPackage().getMathCharacter("(".codePointAt(0))));
            contents.add(new MathCharacterToken(token.getSlice(), CorePackageDefinitions.getPackage().getMathCharacter(")".codePointAt(0))));
            builder.buildMathElement(table, token, argument, true);
        }
    }

}
