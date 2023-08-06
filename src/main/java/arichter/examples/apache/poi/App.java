package arichter.examples.apache.poi;

import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * The main app for this package
 *
 */
public class App {

    private static final Logger LOG = LogManager.getLogger(App.class);

    /**
    * Constructor not used
    */
    public App() {
    }
    
    /**
    * main method for this app
    * @param args default arguments
    */
    public static void main( String[] args ) {
        System.out.println( "Hello World!" );
    }
    
    private void codeDump() {

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


                returnVal = chooser.showSaveDialog(null);
                if(returnVal == JFileChooser.APPROVE_OPTION) {
                    selectedFile = chooser.getSelectedFile();
                }
        
        
    }
}
