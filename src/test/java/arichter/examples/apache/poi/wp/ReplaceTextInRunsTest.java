package arichter.examples.apache.poi.wp;

import java.io.InputStream;

import org.apache.poi.xwpf.usermodel.*;
import java.util.List;

import static org.junit.Assert.assertTrue;
import org.junit.Test;

/**
* Test the methods of class {@link arichter.examples.apache.poi.wp.ReplaceTextInRuns}  
*
* @author Axel Richter
* 
* @see arichter.examples.apache.poi.wp.ReplaceTextInRuns
* */
public class ReplaceTextInRunsTest {
    
    /**
    * Constructor not used
    */
    public ReplaceTextInRunsTest() {
    }
    
    /**
    * Test method for method replace of {@link arichter.examples.apache.poi.wp.ReplaceTextInRuns}.
    * Test is passed if no exceptions thrown.
    * @throws Exception at any Exception
    */
    @Test
    public void testReplace() throws Exception {
        
        ReplaceTextInRuns replaceTextInRuns = new ReplaceTextInRuns();
        
        try (
            InputStream is = getClass().getResourceAsStream("/wp/WordDocumentHavingTextToReplace.docx");
            XWPFDocument document = new XWPFDocument(is);
            ) {

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                List<XWPFRun> runs = replaceTextInRuns.replace(paragraph, "${placeholder}", "text to replace the placeholder");
            }
        }
        
        // test is passed if no exceptions thrown
        assertTrue(true);

    }

}
