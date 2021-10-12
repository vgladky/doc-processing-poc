package net.telclic.docprocessing;

import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;

public class DocXProcessor {

    public static final String TEMPLATE_FILE = "./template.docx";
    public static final String OUTPUT_FILE = "output.docx";

    public static void main(String[] args) throws Exception {
        WordprocessingMLPackage wordMLPackage = getTemplate(TEMPLATE_FILE);

        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        VariablePrepare.prepare(wordMLPackage);

        documentPart.variableReplace(getVariables());

        wordMLPackage.save(createOutputStream(OUTPUT_FILE));
    }

    private static HashMap<String, String> getVariables() {
        HashMap<String, String> variables = new HashMap<>();
        variables.put("firstName", "John");
        variables.put("lastName", "Smith");
        variables.put("message", "Test message");
        return variables;
    }

    private static FileOutputStream createOutputStream(String filePath) throws IOException {
        File outputFile = new File(filePath);
        if (!outputFile.exists()) {
            outputFile.createNewFile();
        }
        return new FileOutputStream(outputFile, false);
    }

    private static WordprocessingMLPackage getTemplate(String templateFileName) throws FileNotFoundException, Docx4JException {
        final File initialFile = new File(templateFileName);
        final InputStream templateInputStream = new DataInputStream(new FileInputStream(initialFile));

        return WordprocessingMLPackage.load(templateInputStream);
    }
}
