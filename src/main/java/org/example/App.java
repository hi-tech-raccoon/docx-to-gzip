package org.example;

import java.io.IOException;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        DocxToGzipConverter docxToGzip = new DocxToGzipConverter();
        String outputPath = "C:\\Users\\<YOUR_USER>\\Desktop\\test.zip";
        try {
            docxToGzip.createDocument(outputPath);
            System.out.println("ZIP Escrito en: " + outputPath);
        } catch (IOException e) {e.printStackTrace();}
    }
}
