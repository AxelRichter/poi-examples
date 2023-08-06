package arichter.examples.apache.poi;

import arichter.examples.apache.poi.sl.ReplaceTextInRuns;

import java.io.InputStream;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.File;

import org.apache.poi.xslf.usermodel.*;
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

/**
* Example usage of ...
*
* @author Axel Richter
* 
* @see arichter.examples.apache.poi.sl.ReplaceTextInRuns
* */
public class ReplaceTextInRunsSlExample {
    
    private static final Logger LOG = LogManager.getLogger(ReplaceTextInRunsSlExample.class);
    
    /**
    * Constructor not used
    */
    public ReplaceTextInRunsSlExample() {
    }

    /**
    * Simple static method to show the usage of {@link arichter.examples.apache.poi.sl.ReplaceTextInRuns}
    * @param args default arguments
    */
    public static void main( String[] args ) {
        try {
        
            ReplaceTextInRuns replaceTextInRuns = new ReplaceTextInRuns();
        
            File selectedFile = null;
            JFileChooser chooser = new JFileChooser();
            FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "PowerPoint Persentaton (*.pptx)", "pptx");
            chooser.setFileFilter(filter);
            chooser.setAcceptAllFileFilterUsed(false);
            int returnVal = chooser.showOpenDialog(null);
            if(returnVal == JFileChooser.APPROVE_OPTION) {
                selectedFile = chooser.getSelectedFile();
            }
            
            String placeholderText = "${placeholder}";
            String replacementText = "text to replace the placeholder";
            XMLSlideShow slideShow = null;
            if (selectedFile == null) {
                InputStream is = ReplaceTextInRunsSlExample.class.getResourceAsStream("/sl/PPTHavingTextToReplace.pptx");
                slideShow = new XMLSlideShow(is);
            } else {
                try {
                    slideShow = new XMLSlideShow(new FileInputStream(selectedFile));    
                    JTextField placeholderField = new JTextField();
                    placeholderField.setText(placeholderText);
                    JTextField replacementTextField = new JTextField();
                    replacementTextField.setText(replacementText);
                    Object[] message = {
                        "PLacehoder:", placeholderField,
                        "Replacement text:", replacementTextField
                    };
                    int option = JOptionPane.showConfirmDialog(null, message, "Input parameters", JOptionPane.OK_CANCEL_OPTION);
                    if (option == JOptionPane.OK_OPTION) {
                        placeholderText = placeholderField.getText();
                        replacementText = replacementTextField.getText();
                    }
                } catch (Exception ex) {
                    System.out.println("could not open " + selectedFile);
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

                returnVal = chooser.showSaveDialog(null);
                if(returnVal == JFileChooser.APPROVE_OPTION) {
                    selectedFile = chooser.getSelectedFile();
                }

                if (selectedFile != null) {
                    FileOutputStream out = new FileOutputStream(selectedFile);
                    slideShow.write(out);
                    out.close();
                    slideShow.close();
                }
            }
        
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}