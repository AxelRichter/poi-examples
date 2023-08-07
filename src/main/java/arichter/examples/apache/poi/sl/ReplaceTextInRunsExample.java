package arichter.examples.apache.poi.sl;

import java.io.InputStream;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.File;

import org.apache.poi.xslf.usermodel.*;
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
* Example usage of {@link arichter.examples.apache.poi.sl.ReplaceTextInRuns}
*
* @author Axel Richter
* 
* @see arichter.examples.apache.poi.sl.ReplaceTextInRuns
* */
public class ReplaceTextInRunsExample {
    
    private static final Logger LOG = LogManager.getLogger(ReplaceTextInRunsExample.class);
    
    /**
    * Constructor method to show the usage of {@link arichter.examples.apache.poi.sl.ReplaceTextInRuns}
    */
    public ReplaceTextInRunsExample() {
    }
    
    /**
    * Rub method to show the usage of {@link arichter.examples.apache.poi.sl.ReplaceTextInRuns}
    * @param sourceFilePath sourceFilePath
    * @param targetFilePath targetFilePath
    * @param placeholderText placeholderText
    * @param replacementText replacementText
    */
    public void run(String sourceFilePath, String targetFilePath, String placeholderText, String replacementText) {
        try {            
            File sourceFile = null;
            if(sourceFilePath != null) {
                sourceFile = new File(sourceFilePath);
                if(!sourceFile.exists() || !sourceFile.isFile()) {
                    LOG.atWarn().log("Could not open file {}. Default source file used.", sourceFile);
                    sourceFile = null;
                }
            }          
            File targetFile = null;
            if(targetFilePath != null) {
                targetFile = new File(targetFilePath);  
            } else {
                targetFile = new File("./PPTHavingTextToReplaceNew.pptx");
            }
           
            ReplaceTextInRuns replaceTextInRuns = new ReplaceTextInRuns();
                    
            XMLSlideShow slideShow = null;
            if (sourceFile == null) {
                InputStream is = ReplaceTextInRunsExample.class.getResourceAsStream("/sl/PPTHavingTextToReplace.pptx");
                slideShow = new XMLSlideShow(is);
            } else {
                try {
                    slideShow = new XMLSlideShow(new FileInputStream(sourceFile));    
                } catch (Exception ex) {
                    LOG.atError().log("Could not open file {} as XMLSlideShow.", sourceFile);
                }
            }
            
            if (slideShow != null) {
                for (XSLFSlide slide : slideShow.getSlides()) {
                    for (XSLFShape shape : slide.getShapes()) {
                        if (shape instanceof XSLFTextShape) {
                            XSLFTextShape textShape = (XSLFTextShape)shape;
                            for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                                List<XSLFTextRun> runs = replaceTextInRuns.replace(paragraph, placeholderText, replacementText);
                            }
                        }
                    }
                }
               
                FileOutputStream out = new FileOutputStream(targetFile);
                slideShow.write(out);
                out.close();
                slideShow.close();
            }
        
        } catch (Exception ex) {
            LOG.atError().withThrowable(ex).log("Exception thrown.");
        }
    }

    /**
    * Main method 
    * @param args array of parameters <br/>
    *             args[0] = sourceFilePath <br/>
    *             args[1] = targetFilePath <br/>
    *             args[2] = placeholderText, default "${placeholder}" <br/>
    *             args[3] = replacementText, default "text to replace the placeholder"
    */
    public static void main( String[] args ) {
                
        String sourceFilePath = null;
        String targetFilePath = null;
        String placeholderText = "${placeholder}";
        String replacementText = "text to replace the placeholder";
        
        if (args.length > 0) {
            sourceFilePath = args[0];
        } 
        if (args.length > 1) {
            targetFilePath = args[1];
        }
        if (args.length > 2) {
            placeholderText = args[2];
        }
        if (args.length > 3) {
            replacementText = args[3];
        }

        ReplaceTextInRunsExample example = new ReplaceTextInRunsExample();
        example.run(sourceFilePath, targetFilePath, placeholderText, replacementText);

    }
}
