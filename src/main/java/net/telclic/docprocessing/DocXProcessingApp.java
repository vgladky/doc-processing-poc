package net.telclic.docprocessing;

import com.thedeanda.lorem.Lorem;
import com.thedeanda.lorem.LoremIpsum;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.management.ManagementFactory;
import java.lang.management.MemoryMXBean;
import java.lang.management.MemoryUsage;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

public class DocXProcessingApp {

    private static final String TEMPLATE_FILE = "template.docx";
    private static final Lorem  lorem         = LoremIpsum.getInstance();

    public static void main(String[] args) throws Exception {

        MemoryMXBean mbean = ManagementFactory.getMemoryMXBean();
        MemoryUsage beforeHeapMemoryUsage = mbean.getHeapMemoryUsage();

        long start2 = System.currentTimeMillis();

        runInSingleThread();
        //runMultipleThreads();

        // benchmarking results
        long end2 = System.currentTimeMillis();
        System.out.println("Elapsed Time in seconds: " + (end2 - start2) / 1000.0);

        MemoryUsage afterHeapMemoryUsage = mbean.getHeapMemoryUsage();
        long consumed = afterHeapMemoryUsage.getUsed() -
            beforeHeapMemoryUsage.getUsed();
        System.out.println("Total consumed Memory:" + consumed / 1000000.0);


    }

    private static void runInSingleThread() throws DocxProcessingException, FileNotFoundException {
        var template = DocXProcessingApp.class.getClassLoader().getResourceAsStream(TEMPLATE_FILE);

        var docxProcessor = DocxProcessorImpl.newBuilder().setTemplate(template).build();

        for (var i = 0; i < 20; i++) {
            docxProcessor.setTextVariables(getTextVariables());
            docxProcessor.setImageVariables(getImageVariables());
            docxProcessor.processFile();
            docxProcessor.save(new FileOutputStream("temp/out" + i + ".docx", false));
        }
    }

    private static void runMultipleThreads() throws InterruptedException {
        ExecutorService es = Executors.newCachedThreadPool();

        for (int i = 0; i < 20; i++) {
            final int n = i;
            es.execute(() -> {
                try {
                    var template = DocXProcessingApp.class.getClassLoader().getResourceAsStream(TEMPLATE_FILE);
                    var docxProcessor = DocxProcessorImpl.newBuilder().setTemplate(template).build();
                    docxProcessor.setTextVariables(getTextVariables());
                    docxProcessor.setImageVariables(getImageVariables());
                    docxProcessor.processFile();
                    docxProcessor.save(new FileOutputStream("temp/out" + n + ".docx", false));
                } catch (Exception e) {
                    e.printStackTrace();
                }
            });
        }
        es.shutdown();
        boolean finished = es.awaitTermination(5, TimeUnit.MINUTES);
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
