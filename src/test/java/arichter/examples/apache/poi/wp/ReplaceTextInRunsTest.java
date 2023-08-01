package arichter.examples.apache.poi.wp;

import java.io.InputStream;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.*;
import java.util.List;

import static org.junit.Assert.assertTrue;
import org.junit.Test;

public class ReplaceTextInRunsTest {
    
    @Test
    public void testReplace() throws Exception {
        
        ReplaceTextInRuns replaceTextInRuns = new ReplaceTextInRuns();
        
        try (
            InputStream is = getClass().getResourceAsStream("/wp/WordDocument.docx");
            XWPFDocument document = new XWPFDocument(is);
            FileOutputStream out = new FileOutputStream("./WordDocumentNew.docx");
            ) {

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                List<XWPFRun> runs = replaceTextInRuns.replace(paragraph, "${Firma}", "Sample Corporation LTD");
                System.out.println(runs);
            }


            document.write(out);
        }
        
        assertTrue(true);

    }

}
