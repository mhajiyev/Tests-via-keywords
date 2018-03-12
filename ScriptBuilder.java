package com.example.scriptbuilder;
import java.awt.Font;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import java.lang.String;



public class ScriptBuilder {
    public static void main(String[] args) throws Exception{


        try {XWPFDocument docx = new XWPFDocument(new FileInputStream("/Users/mustafahajiyev/Documents/tablr.docx"));
            //XWPFWordExtractor we = new XWPFWordExtractor(docx);
           Iterator <IBodyElement> bodyElementIterator = docx.getBodyElementsIterator();


           // XWPFTable t = new XWPFTable(docx.getTable(IBodyElement(docx)), iface,6,4);
           // System.out.println(t.getRow(2).getCell(1).getText());

             /*   for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()){
                        System.out.println(cell.getText());
                        String sFieldValue = cell.getText();
                        if (sFieldValue.matches("Approved")){
                            System.out.println("The match as per document is true");
                        }
                        System.out.println("\t");
                    }
                    System.out.println(" ");
                }  */
             while (bodyElementIterator.hasNext()){
                 IBodyElement element = bodyElementIterator.next();
                 if ("TABLE".equalsIgnoreCase(element.getElementType().name())){
                     List<XWPFTable> tableList = element.getBody().getTables();
                     for (XWPFTable table: tableList){
                        // System.out.println("# of Rows"+table.getNumberOfRows());
                       //  System.out.println(table.getText());
                         for (XWPFTableRow row : table.getRows()) {
                             if (row.getCell(1).getText().contains("Drive communicating train")) {
                                 System.out.println(row.getCell(1).getText());}
                           XWPFParagraph par = row.getCell(1).getParagraphs().get(0);

                             XWPFRun run = par.createRun();
                             //System.out.println(run.isBold());



                         }

                         for (XWPFTableRow row : table.getRows()) {
                             for (XWPFTableCell cell : row.getTableCells()) {
                                 for (XWPFParagraph p : cell.getParagraphs()) {
                                     for (XWPFRun r : p.getRuns()) {
                                         String text = r.getText(0);
                                         if (r.isBold()) {
                                             for (String boldWord : text){
                                                if (boldWord.isBold()) System.out.println();
                                             }
                                         }
                                     }
                                 }}}
                     }
                 }
             }

        }
        catch(Exception e){
            System.out.println(e);
        }
    }
}
