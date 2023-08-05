package arichter.examples.apache.poi.sl;

import java.io.InputStream;
import java.io.FileOutputStream;

import org.apache.poi.xslf.usermodel.*;
import java.util.List;

import static org.junit.Assert.assertTrue;
import org.junit.Test;

/**
* Test the methods of class ReplaceTextInRuns  
*
* @author Axel Richter
* 
* @see arichter.examples.apache.poi.sl.ReplaceTextInRuns
* */
public class ReplaceTextInRunsTest {
    
    /**
    * Text method for method replace of {@link arichter.examples.apache.poi.sl.ReplaceTextInRuns}.
    * Test is passed if no exceptions thrown.
    */
    @Test
    public void testReplace() throws Exception {
        
        ReplaceTextInRuns replaceTextInRuns = new ReplaceTextInRuns();
        
        try (
            InputStream is = getClass().getResourceAsStream("/sl/PPTHavingTextShape.pptx");
            XMLSlideShow slideShow = new XMLSlideShow(is);
            FileOutputStream out = new FileOutputStream("./PPTHavingTextShapeResult.pptx");
            ) {

            for (XSLFSlide slide : slideShow.getSlides()) {
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape textShape = (XSLFTextShape)shape;
                        for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                            List<XSLFTextRun> runs = replaceTextInRuns.replace(paragraph, "${placeholder}", "text to replace the placeholder");
                            System.out.println(runs);
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
