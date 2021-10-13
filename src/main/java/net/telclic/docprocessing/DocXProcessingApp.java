package net.telclic.docprocessing;

import com.thedeanda.lorem.Lorem;
import com.thedeanda.lorem.LoremIpsum;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

public class DocXProcessingApp {

    private static final String TEMPLATE_FILE = "template.docx";
    private static final Lorem  lorem         = LoremIpsum.getInstance();

    public static void main(String[] args) throws Exception {
        var template = DocXProcessingApp.class.getClassLoader().getResourceAsStream(TEMPLATE_FILE);

        var docxProcessor = DocxProcessorImpl.newBuilder().setTemplate(template).build();

        for (var i = 0; i < 10; i++) {
            docxProcessor.setTextVariables(getTextVariables());
            docxProcessor.setImageVariables(getImageVariables());
            docxProcessor.processFile();
            docxProcessor.save(new FileOutputStream("temp/out" + i + ".docx", false));
        }
    }

    private static Map<String, String> getTextVariables() {
        Map<String, String> variables = new HashMap<>();
        variables.put("firstName", lorem.getFirstName());
        variables.put("lastName", lorem.getLastName());
        variables.put("message", lorem.getParagraphs(5, 10));
        return variables;
    }

    private static Map<String, InputStream> getImageVariables() {
        HashMap<String, InputStream> variables = new HashMap<>();
        variables.put("MONKEY", DocXProcessingApp.class.getClassLoader().getResourceAsStream("monkey.jpg"));
        variables.put("LION", DocXProcessingApp.class.getClassLoader().getResourceAsStream("lion.jpg"));
        return variables;
    }
}
