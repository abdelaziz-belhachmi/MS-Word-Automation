package controller;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import  java.io.*;
import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;

public class CrossWordGenerator8x5by11 {

    private String BasicPath ="/*path to working folder*/";
    private String firstpage = "1stPageCrossWord";//update each time
    private String secondpage = "2ndPageCrossWord";//update each time


    public void Start(String donePath) throws IOException {

        XWPFDocument doc=new XWPFDocument();

        // Set the page dimensions in EMUs (English Metric Units)
        int width = 12240;  // 1440*8.5
        int height = 1440*11;  // 1440*11
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageSz pageSz = sectPr.addNewPgSz();
        pageSz.setW(BigInteger.valueOf(width)); // Page width 8.5'
        pageSz.setH(BigInteger.valueOf(height)); // Page height 11'
        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(850L)); // Left margin
        pageMar.setRight(BigInteger.valueOf(600L)); // Right margin
        pageMar.setTop(BigInteger.valueOf(600L)); // Top margin
        pageMar.setBottom(BigInteger.valueOf(500L)); // Bottom margin
        //pageMar.setGutter(BigInteger.valueOf(540L)); // Gutter margin (0.375" x 1440 EMUs per inch)

        // Create a new CTBorder object and set its size and style
//         CTBorder borderTemplate = CTBorder.Factory.newInstance();
//         borderTemplate.setSz(BigInteger.valueOf(4L)); // Border size in eighths of a point (1 point = 72/8 eighths of a point)
//         borderTemplate.setVal(STBorder.SINGLE); // Border style (SINGLE, DOUBLE, DASHED, etc.)
        XWPFParagraph paragraph01 = doc.createParagraph();
        XWPFRun r1 = paragraph01.createRun();
        r1.addBreak(BreakType.TEXT_WRAPPING);
        r1.addBreak(BreakType.TEXT_WRAPPING);
        r1.addBreak(BreakType.TEXT_WRAPPING);
        r1.addBreak(BreakType.TEXT_WRAPPING);
        r1.addBreak(BreakType.TEXT_WRAPPING);
        r1.addBreak(BreakType.TEXT_WRAPPING);
        r1.addBreak(BreakType.TEXT_WRAPPING);
        r1.addBreak(BreakType.TEXT_WRAPPING);
        r1.setText(" BOOK TITLE ");
        r1.setFontFamily("Arial Rounded MT Bold");
        r1.setFontSize(28);
        r1.setCapitalized(true);
        r1.setBold(true);
//        r1.setItalic(true);
        paragraph01.setVerticalAlignment(TextAlignment.CENTER);
        paragraph01.setAlignment(ParagraphAlignment.CENTER);

        XWPFParagraph paragraph02 = doc.createParagraph();
        XWPFRun r2 = paragraph02.createRun();
        r2.setText("TITLE");
        r2.setFontSize(18);
        r1.setFontFamily("Arial Rounded MT Bold");
        r2.setCapitalized(true);
        r2.setBold(true);
        r2.setItalic(true);
        paragraph02.setVerticalAlignment(TextAlignment.CENTER);
        paragraph02.setAlignment(ParagraphAlignment.CENTER);

        XWPFParagraph paragraph03 = doc.createParagraph();
        XWPFRun r3 = paragraph03.createRun();
        paragraph03.setPageBreak(true);

        File folder = new File(BasicPath+donePath+firstpage);
        File[] files = folder.listFiles();
//        File folder2 = new File(BasicPath+donePath+secondpage);
//        String blackPath = C:/Users/unknow/Desktop/kdp/crosswordBooks+/doneX/+2ndPageCrossWord;

        for (File file : files) {
            // Create new paragraph and add page break
            XWPFParagraph paragraph1 = doc.createParagraph();
            XWPFParagraph paragraph2 = doc.createParagraph();

            paragraph1.setPageBreak(true);
            //alignement
            paragraph1.setVerticalAlignment(TextAlignment.CENTER);
            paragraph1.setAlignment(ParagraphAlignment.CENTER );
            // border
//            paragraph1.setBorderLeft(Borders.THICK_THIN_SMALL_GAP);
//            paragraph1.setBorderRight(Borders.THICK_THIN_SMALL_GAP);
//            paragraph1.setBorderTop(Borders.THICK_THIN_SMALL_GAP);
//            paragraph1.setBorderBottom(Borders.THICK_THIN_SMALL_GAP);


            // Add empty run to the paragraph
            XWPFRun run = paragraph1.createRun();
            run.setText("");

            // Load image data into byte array
            Path imagePath = file.toPath();


            paragraph1.setPageBreak(true);

            try {
                //first page
                XWPFPicture picture = run.addPicture(new FileInputStream(String.valueOf(imagePath)), XWPFDocument.PICTURE_TYPE_PNG, file.getName(), Units.toEMU(420) , Units.toEMU(730));

            } catch (InvalidFormatException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
            //2nd page
            XWPFRun run2 = paragraph2.createRun();
            run2.addBreak(BreakType.TEXT_WRAPPING);
            run2.setText("Solutions to Your Previous Puzzle");
            run2.setFontSize(25);
            run2.setBold(true);
            run2.addBreak(BreakType.TEXT_WRAPPING);
            run2.addBreak(BreakType.TEXT_WRAPPING);
            paragraph2.setVerticalAlignment(TextAlignment.CENTER);
            paragraph2.setAlignment(ParagraphAlignment.CENTER );
            //insert 2nd page here
            try {
                XWPFPicture picture = run2.addPicture(new FileInputStream(BasicPath+donePath+secondpage+"/"+file.getName()), XWPFDocument.PICTURE_TYPE_PNG, file.getName(), Units.toEMU(450) , Units.toEMU(450));
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
            //new page
            paragraph2.setPageBreak(true);

        }


        //save the file
        FileOutputStream out = new FileOutputStream(BasicPath+donePath+"/CrossWordGeneratedDoc8by11.docx");
        doc.write(out);
        out.close();


    }


}
