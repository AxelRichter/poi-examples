package arichter.examples.apache.poi;

import java.io.File;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.*;

import java.util.Map; 
import java.util.HashMap; 

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * The main app for this package
 *
 */
public class App extends JPanel implements ActionListener {

    private static final Logger LOG = LogManager.getLogger(App.class);
    
    private static final Map<String, String> examples = new HashMap<String, String>();
    static {
        examples.put("Example replace text in PowertPoint", "arichter.examples.apache.poi.sl.ReplaceTextInRunsExample");    
        examples.put("Example replace text in Word", "arichter.examples.apache.poi.wp.ReplaceTextInRunsExample");    
        examples.put("Example 3", null);    
        examples.put("Example 4", null);    
        examples.put("Example 5", null);    
    }

    /**
    * Constructor for main app
    */
    public App() {
        super(new GridLayout(0,1));
        JPanel panel = new JPanel(new GridLayout(0,1));
        JButton button = null;
        for (String example : examples.keySet()) {
            button = new JButton(example);
            button.setActionCommand(example);
            button.addActionListener(this);
            panel.add(button);
        }
        JScrollPane scrollPane = new JScrollPane(panel);
        add(scrollPane);
    }
    
    /**
    * Action listener method
    * @param e {@link java.awt.event.ActionEvent}
    */
    public void actionPerformed(ActionEvent e) {
        String example = e.getActionCommand();
        String className = examples.get(example);
        if (className != null) {
            try {
                Class classObject = Class.forName(className);
                runExample(classObject);
            } catch (Exception ex) {
                LOG.atWarn().log("Could not open class for name {}.", className);
            }
        } 
    }   

    private void runExample(Class classObject) {
        if (classObject.getName().equals("arichter.examples.apache.poi.sl.ReplaceTextInRunsExample")) {
   
            String sourceFilePath = null;
            File selectedFile = null;
            JFileChooser chooser = new JFileChooser();
            FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "PowerPoint Persentaton (*.pptx)", "pptx");
            chooser.setFileFilter(filter);
            chooser.setAcceptAllFileFilterUsed(false);
            int returnVal = chooser.showOpenDialog(this);
            if(returnVal == JFileChooser.APPROVE_OPTION) {
                selectedFile = chooser.getSelectedFile();
                sourceFilePath = selectedFile.getAbsolutePath();
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
            int option = JOptionPane.showConfirmDialog(this, message, "Input parameters", JOptionPane.OK_CANCEL_OPTION);
            if (option == JOptionPane.OK_OPTION) {
                placeholderText = placeholderField.getText();
                replacementText = replacementTextField.getText();
            }

            String targetFilePath = null;
            returnVal = chooser.showSaveDialog(this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                selectedFile = chooser.getSelectedFile();
                targetFilePath = selectedFile.getAbsolutePath();
                if (!targetFilePath.endsWith(".pptx")) {
                    targetFilePath = targetFilePath + ".pptx";   
                }
            }

            try {
                arichter.examples.apache.poi.sl.ReplaceTextInRunsExample example = 
                    (arichter.examples.apache.poi.sl.ReplaceTextInRunsExample)classObject.newInstance();
                example.run(sourceFilePath, targetFilePath, placeholderText, replacementText);
            } catch (Exception ex) {
                LOG.atError().withThrowable(ex).log("Exception thrown.");
            }
            
        } else if (classObject.getName().equals("arichter.examples.apache.poi.wp.ReplaceTextInRunsExample")) {
            
            String sourceFilePath = null;
            File selectedFile = null;
            JFileChooser chooser = new JFileChooser();
            FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "PowerPoint Persentaton (*.docx)", "docx");
            chooser.setFileFilter(filter);
            chooser.setAcceptAllFileFilterUsed(false);
            int returnVal = chooser.showOpenDialog(this);
            if(returnVal == JFileChooser.APPROVE_OPTION) {
                selectedFile = chooser.getSelectedFile();
                sourceFilePath = selectedFile.getAbsolutePath();
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
            int option = JOptionPane.showConfirmDialog(this, message, "Input parameters", JOptionPane.OK_CANCEL_OPTION);
            if (option == JOptionPane.OK_OPTION) {
                placeholderText = placeholderField.getText();
                replacementText = replacementTextField.getText();
            }

            String targetFilePath = null;
            returnVal = chooser.showSaveDialog(this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                selectedFile = chooser.getSelectedFile();
                targetFilePath = selectedFile.getAbsolutePath();
                if (!targetFilePath.endsWith(".docx")) {
                    targetFilePath = targetFilePath + ".docx";   
                }
            }

            try {
                arichter.examples.apache.poi.wp.ReplaceTextInRunsExample example = 
                    (arichter.examples.apache.poi.wp.ReplaceTextInRunsExample)classObject.newInstance();
                example.run(sourceFilePath, targetFilePath, placeholderText, replacementText);
            } catch (Exception ex) {
                LOG.atError().withThrowable(ex).log("Exception thrown.");
            }
            
        }
    }        
    
    private static void createAndShowGUI() {
        JFrame frame = new JFrame("Examples for usage of Apache POI");
        frame.setLayout(new BorderLayout());
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
 
        // JMenuBar menuBar = new JMenuBar();
        // menuBar.setOpaque(true);
        // menuBar.setPreferredSize(new Dimension(0, 20));
        // JMenu menu = new JMenu("A Menu");
        // menu.setMnemonic(KeyEvent.VK_A);
        // menu.getAccessibleContext().setAccessibleDescription("The only menu in this program that has menu items");
        // menuBar.add(menu);
        // frame.setJMenuBar(menuBar);
        
        JComponent newContentPane = new App();
        newContentPane.setOpaque(true);
        frame.setContentPane(newContentPane); 
        frame.setPreferredSize(new Dimension(400, 300));
        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }
 
    /**
    * main method for this app
    * @param args default arguments
    */
    public static void main(String[] args) {
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });
    }
}
