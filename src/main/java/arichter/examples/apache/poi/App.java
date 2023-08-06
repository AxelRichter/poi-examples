package arichter.examples.apache.poi;

import java.io.File;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JFileChooser;
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
        
    private static void createAndShowGUI() {
        //Create and set up the window.
        JFrame frame = new JFrame("HelloWorldSwing");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
 
        //Add the ubiquitous "Hello World" label.
        JLabel label = new JLabel("Hello World");
        frame.getContentPane().add(label);
 
        //Display the window.
        frame.pack();
        frame.setVisible(true);
    }
 
    /**
    * main method for this app
    * @param args default arguments
    */
    public static void main(String[] args) {
        //Schedule a job for the event-dispatching thread:
        //creating and showing this application's GUI.
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });
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
