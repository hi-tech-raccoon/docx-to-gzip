package org.example;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class DocxToGzipConverter
{
    public void createDocument(String outputPath) throws IOException
    {
        // Adquiriendo el fichero Word
        XWPFDocument document = prepareWordDocument();

        // Preparando el Input Stream
        InputStream docxInputStream = getWordDocumentInputStream(document);

        // Preparando el Output Stream ZIP
        OutputStream zipOutStream = prepareZipOutputStreamToFile(outputPath);

        // Escribiendo el fichero comprimido Zip
        writeZipFile(docxInputStream, zipOutStream);
    }

    private XWPFDocument prepareWordDocument()
    {
        // Usando Apache POI para generar un documento word.
        XWPFDocument document = new XWPFDocument();
        XWPFParagraph paragraph = document.createParagraph();

        // Escribiendo texto en el documento word.
        XWPFRun run = paragraph.createRun();
        run.setText("Hello World!");
        return document;
    }

    private InputStream getWordDocumentInputStream(XWPFDocument document) throws IOException
    {
        // Adquiriendo el Input Stream del documento word.
        // Preparando un OutputStream capaz de proporcionar bytes.
        ByteArrayOutputStream docxByteOutStream = new ByteArrayOutputStream();
        // Preparando para escribir el documento word en una secuencia de bytes.
        document.write(docxByteOutStream);
        document.close();
        // Preparando un Input Stream usando los bytes del stream que preparamos.
        ByteArrayInputStream docxInStream = new ByteArrayInputStream(docxByteOutStream.toByteArray());
        return docxInStream;
    }

    private OutputStream prepareZipOutputStreamToFile(String path) throws IOException
    {
        // Creando un Output Stream para fichero usando el path del parametro.
        FileOutputStream fileOutStream = new FileOutputStream(path);
        // Creando un Output Stream de tipo Zip usando el que ya tenemos sobre ficheros.
        ZipOutputStream zipOutStream = new ZipOutputStream(fileOutStream);

        // Configurando el ZIP que vamos a crear.
        ZipEntry zipEntry = new ZipEntry("test.docx");
        zipOutStream.putNextEntry(zipEntry);

        return zipOutStream;
    }

    private void writeZipFile(InputStream inputStream, OutputStream outputStream) throws IOException
    {
        // Leyendo del Input Stream y guardando en Output Stream
        byte[] readBuffer = new byte[1024];
        int bytesRead;
        while((bytesRead = inputStream.read(readBuffer)) > 0)
        {
            outputStream.write(readBuffer, 0, bytesRead);
        }
        outputStream.close();
    }
}
