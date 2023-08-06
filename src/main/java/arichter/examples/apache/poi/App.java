package arichter.examples.apache.poi;

import java.io.File;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.*;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * The main app for this package
 *
 */
public class App extends JPanel implements ActionListener {

    private static final Logger LOG = LogManager.getLogger(App.class);

    /**
    * Constructor not used
    */
    public App() {
        super(new GridLayout(0,1));
        JPanel panel = new JPanel(new GridLayout(0,1));
        JButton button = new JButton("Example 1");
        button.setActionCommand("Example 1");
        button.addActionListener(this);
        panel.add(button);
        button = new JButton("Example 2");
        button.setActionCommand("Example 2");
        button.addActionListener(this);
        panel.add(button);
        JScrollPane scrollPane = new JScrollPane(panel);
        add(scrollPane);
    }
    
    public void actionPerformed(ActionEvent e) {
        System.out.println(e.getActionCommand());
    }      
    
    private static void createAndShowGUI() {
        //Create and set up the window.
        JFrame frame = new JFrame("HelloWorldSwing");
        frame.setLayout(new BorderLayout());
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
 
        //Create the menu bar.  Make it have a green background.
        JMenuBar menuBar = new JMenuBar();
        menuBar.setOpaque(true);
        menuBar.setPreferredSize(new Dimension(0, 20));
        JMenu menu = new JMenu("A Menu");
        menu.setMnemonic(KeyEvent.VK_A);
        menu.getAccessibleContext().setAccessibleDescription("The only menu in this program that has menu items");
        menuBar.add(menu);
 
        //Set the menu bar and add the label to the content pane.
        frame.setJMenuBar(menuBar);
        
        //Create and set up the content pane.
        JComponent newContentPane = new App();
        newContentPane.setOpaque(true); //content panes must be opaque
        frame.setContentPane(newContentPane); 
        
        frame.setLocationRelativeTo(null);
       
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
