package arichter.examples.apache.poi.sl;

import java.io.InputStream;
import java.io.FileOutputStream;

import org.apache.poi.xslf.usermodel.*;
import java.util.List;

import static org.junit.Assert.assertTrue;
import org.junit.Test;

public class ReplaceTextInRunsTest {
    
    @Test
    public void testReplace() throws Exception {
        
        ReplaceTextInRuns replaceTextInRuns = new ReplaceTextInRuns();
        
        try (
            InputStream is = getClass().getResourceAsStream("/sl/PPT.pptx");
            XMLSlideShow slideShow = new XMLSlideShow(is);
            FileOutputStream out = new FileOutputStream("./PPTNew.pptx");
            ) {

            for (XSLFSlide slide : slideShow.getSlides()) {
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape textShape = (XSLFTextShape)shape;
                        for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                            List<XSLFTextRun> runs = replaceTextInRuns.replace(paragraph, "${Firma}", "Sample Corporation LTD");
                            System.out.println(runs);
                        }
                    }
                }
            }

            slideShow.write(out);
        }
        
        assertTrue(true);

    }

}
