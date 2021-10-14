package net.telclic.docprocessing;

import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

import javax.xml.bind.JAXBElement;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class DocxProcessorImpl implements DocxProcessor {

    private WordprocessingMLPackage  wordprocessingMLPackage;
    private InputStream              inputStream;
    private Map<String, String>      textVariables;
    private Map<String, InputStream> imageVariables;

    private DocxProcessorImpl() {
        // private constructor
    }

    public static DocxProcessorBuilder newBuilder() {
        return new DocxProcessorImpl().new DocxProcessorBuilder();
    }

    private static P addInlineImageToParagraph(Inline inline) {
        // Now add the in-line image to a paragraph
        ObjectFactory factory = new ObjectFactory();
        P paragraph = factory.createP();
        R run = factory.createR();
        paragraph.getContent().add(run);
        Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return paragraph;
    }

    @Override
    public void setTextVariables(Map<String, String> textVariables) {
        this.textVariables = textVariables;
    }

    @Override
    public void setImageVariables(Map<String, InputStream> imageVariables) {
        this.imageVariables = imageVariables;
    }

    private byte[] convertImageToByteArray(InputStream is) throws IOException {
        byte[] targetArray = new byte[is.available()];

        is.read(targetArray);
        return targetArray;
    }

    private List<Object> getAllElementFromObject(Object obj, Class toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj instanceof JAXBElement)
            obj = ((JAXBElement) obj).getValue();

        if (obj.getClass().equals(toSearch)) {
            result.add(obj);
        } else if (obj instanceof ContentAccessor) {
            List<Object> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }

        }
        return result;
    }

    private Inline createInlineImage(InputStream inputStream) throws Exception {
        byte[] bytes = convertImageToByteArray(inputStream);

        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(this.wordprocessingMLPackage, bytes);

        int docPrId = 1;
        int cNvPrId = 2;

        return imagePart.createImageInline("Filename hint", "Alternative text", docPrId, cNvPrId, false);
    }

    private WordprocessingMLPackage getTemplate() throws Docx4JException {

        return WordprocessingMLPackage.load(inputStream);
    }

    @Override
    public void processFile() throws DocxProcessingException {

        try {
            replaceImages();

            MainDocumentPart documentPart = wordprocessingMLPackage.getMainDocumentPart();
            var link = createHyperlink(documentPart, "https://google.com");
            // Add it to a paragraph
            org.docx4j.wml.P paragraph = Context.getWmlObjectFactory().createP();
            paragraph.getContent().add( link );
            documentPart.addObject(paragraph);

            VariablePrepare.prepare(wordprocessingMLPackage);
            documentPart.variableReplace(textVariables);
        } catch (Exception e) {
            throw new DocxProcessingException("Error during variable replace", e);
        }
    }

    private void replaceImages() throws Exception {
        List<Object> elements = getAllElementFromObject(wordprocessingMLPackage.getMainDocumentPart(), Tbl.class);

        for (Object obj : elements) {
            if (obj instanceof Tbl) {
                Tbl table = (Tbl) obj;
                List<Object> rows = getAllElementFromObject(table, Tr.class);
                for (Object trObj : rows) {
                    Tr tr = (Tr) trObj;
                    List<Object> cols = getAllElementFromObject(tr, Tc.class);
                    for (Object tcObj : cols) {
                        Tc tc = (Tc) tcObj;
                        List<Object> texts = getAllElementFromObject(tc, Text.class);
                        for (Object textObj : texts) {
                            Text text = (Text) textObj;
                            for (var entry : imageVariables.entrySet()) {
                                if (text.getValue().equalsIgnoreCase(String.format("${%s}", entry.getKey()))) {
                                    P paragraphWithImage = addInlineImageToParagraph(createInlineImage(entry.getValue()));
                                    tc.getContent().remove(0);

                                    tc.getContent().add(paragraphWithImage);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    public P.Hyperlink createHyperlink(MainDocumentPart mdp, String url) {
        try {
            // We need to add a relationship to word/_rels/document.xml.rels
            // but since its external, we don't use the
            // usual wordMLPackage.getMainDocumentPart().addTargetPart
            // mechanism
            org.docx4j.relationships.ObjectFactory factory =
                new org.docx4j.relationships.ObjectFactory();
            org.docx4j.relationships.Relationship rel = factory.createRelationship();
            rel.setType( Namespaces.HYPERLINK  );
            rel.setTarget(url);
            rel.setTargetMode("External");
            mdp.getRelationshipsPart().addRelationship(rel);

            // addRelationship sets the rel's @Id
            String hpl = "<w:hyperlink r:id=\"" + rel.getId() + "\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" >" +
                "<w:r>" +
                "<w:rPr>" +
                "<w:rStyle w:val=\"Hyperlink\" />" +
                "</w:rPr>" +
                "<w:t>Link</w:t>" +
                "</w:r>" +
                "</w:hyperlink>";
            return (P.Hyperlink) XmlUtils.unmarshalString(hpl);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    @Override
    public void save(OutputStream outputStream) throws DocxProcessingException {
        try {
            this.wordprocessingMLPackage.save(outputStream);
        } catch (Docx4JException e) {
            throw new DocxProcessingException("DOCX file processing error", e);
        }
    }

    public class DocxProcessorBuilder {

        private DocxProcessorBuilder() {
            // private constructor
        }

        public DocxProcessorBuilder setTemplate(InputStream inputStream) {
            DocxProcessorImpl.this.inputStream = inputStream;
            return this;
        }

        public DocxProcessorBuilder setTextPlaceholders(Map<String, String> variables) {
            DocxProcessorImpl.this.textVariables = variables;
            return this;
        }

        public DocxProcessorBuilder setImagePlaceholders(Map<String, InputStream> variables) {
            DocxProcessorImpl.this.imageVariables = variables;
            return this;
        }

        public DocxProcessor build() throws DocxProcessingException {
            try {
                DocxProcessorImpl.this.wordprocessingMLPackage = DocxProcessorImpl.this.getTemplate();
            } catch (Docx4JException e) {
                throw new DocxProcessingException("Cannot build DocxProcessor", e);
            }
            return DocxProcessorImpl.this;
        }
    }

}
