package arichter.examples.apache.poi.sl;

import java.io.InputStream;
import java.io.FileOutputStream;

import org.apache.poi.xslf.usermodel.*;
import java.util.List;

import static org.junit.Assert.assertTrue;
import org.junit.Test;

/**
* Test the methods of class {@link arichter.examples.apache.poi.sl.ReplaceTextInRuns}  
*
* @author Axel Richter
* 
* @see arichter.examples.apache.poi.sl.ReplaceTextInRuns
* */
public class ReplaceTextInRunsTest {
    
    /**
    * Test method for method replace of {@link arichter.examples.apache.poi.sl.ReplaceTextInRuns}.
    * Test is passed if no exceptions thrown.
    * @throws Exception at any Exception
    */
    @Test
    public void testReplace() throws Exception {
        
        ReplaceTextInRuns replaceTextInRuns = new ReplaceTextInRuns();
        
        try (
            InputStream is = getClass().getResourceAsStream("/sl/PPTHavingTextToReplace.pptx");
            XMLSlideShow slideShow = new XMLSlideShow(is);
            FileOutputStream out = new FileOutputStream("./PPTHavingTextToReplaceResult.pptx");
            ) {

            for (XSLFSlide slide : slideShow.getSlides()) {
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape textShape = (XSLFTextShape)shape;
                        for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                            List<XSLFTextRun> runs = replaceTextInRuns.replace(paragraph, "${placeholder}", "text to replace the placeholder");
                        }
                    }
                }
                
            }
            slideShow.write(out);
        }
        
        // test is passed if no exceptions thrown
        assertTrue(true);

    }

}
