package com.github.rzymek;

import org.apache.ant.compress.taskdefs.Unzip;
import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveOutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.nio.file.Files;
import java.util.Arrays;
import java.util.List;

public class XlsxTest {

    @Test
    public void test() throws IOException {
        File file = new File("target", "sample.xlsx");
        File dir = new File("target", "sample");
        File result = new File("target", "repackaged.xlsx");

        createXlsxUsingPOI(file);
        unpackTo(file, dir);
        repack(dir, result);
        System.out.println("Open " + result.toURI() + " with Excel");
    }

    private void repack(File dir, File xlsx) throws IOException {
        List<String> paths = Arrays.asList(
                "[Content_Types].xml",
                "_rels/.rels",
                "docProps/app.xml",
                "docProps/core.xml",
                "xl/styles.xml",
                "xl/workbook.xml",
                "xl/sharedStrings.xml",
                "xl/_rels/workbook.xml.rels",
                "xl/worksheets/sheet1.xml"
        );
        try (ZipArchiveOutputStream out = new ZipArchiveOutputStream(new FileOutputStream(xlsx))) {
            out.setUseZip64(Zip64Mode.Always);
            for (String path : paths) {
                out.putArchiveEntry(new ZipArchiveEntry(path));
                Files.copy(new File(dir, path).toPath(), out);
                out.closeArchiveEntry();
            }
        }
    }

    private void unpackTo(File file, File dir) {
        Unzip unzipper = new Unzip();
        unzipper.setSrc(file);
        unzipper.setDest(dir);
        unzipper.execute();
    }

    private void createXlsxUsingPOI(File file) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet();
            XSSFRow row = sheet.createRow(0);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue(42);
            try (OutputStream out = new BufferedOutputStream(new FileOutputStream(file))) {
                wb.write(out);
            }
        }
    }

}
