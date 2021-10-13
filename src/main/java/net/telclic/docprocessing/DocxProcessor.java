package net.telclic.docprocessing;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

public interface DocxProcessor {

    void setTextVariables(Map<String, String> textVariables);

    void setImageVariables(Map<String, InputStream> imageVariables);

    void processFile() throws DocxProcessingException;

    void save(OutputStream outputStream) throws DocxProcessingException;
}
