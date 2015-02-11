package watsondb;

import java.awt.Desktop;
import java.awt.EventQueue;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JButton;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.IOException;

import javax.swing.JTextField;
import javax.swing.JSeparator;

public class GUI {

	private File inputFile;
	
	private JFrame frmBuWatsonDashboard;
	private JTextField fileNameTextField;
	private JSeparator separator;
	private JButton browseButton;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					GUI window = new GUI();
					window.frmBuWatsonDashboard.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public GUI() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmBuWatsonDashboard = new JFrame();
		frmBuWatsonDashboard.setTitle("BU Watson Dashboard - Version 1.0");
		frmBuWatsonDashboard.setBounds(100, 100, 400, 150);
		frmBuWatsonDashboard.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmBuWatsonDashboard.getContentPane().setLayout(null);
		
		final JButton createButton = new JButton("Create Dashboard");
		createButton.setEnabled(false);
		createButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				// Create a new SpreadsheetManager, passing in desired input/output file.
				SpreadsheetManager sm = new SpreadsheetManager(inputFile, new File("WatsonDashboard.xlsx"));
				
				// Perform the initial read of the input file.
				sm.read();
				
				// Perform the final write to the output file.
				sm.write();
				
				try {
					Desktop.getDesktop().open(new File("WatsonDashboard.xlsx"));
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		});
		createButton.setBounds(10, 58, 364, 37);
		frmBuWatsonDashboard.getContentPane().add(createButton);
		
		fileNameTextField = new JTextField();
		fileNameTextField.setEditable(false);
		fileNameTextField.setBounds(10, 11, 244, 23);
		frmBuWatsonDashboard.getContentPane().add(fileNameTextField);
		fileNameTextField.setColumns(10);
		
		separator = new JSeparator();
		separator.setBounds(10, 45, 364, 2);
		frmBuWatsonDashboard.getContentPane().add(separator);
		
		browseButton = new JButton("Browse");
		browseButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				
				JFileChooser fileChooser = new JFileChooser();
				int returnVal = fileChooser.showOpenDialog(null);
				
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					inputFile = fileChooser.getSelectedFile();
					fileNameTextField.setText(inputFile.getAbsolutePath());
					createButton.setEnabled(true);
				}
				
			}
		});
		browseButton.setBounds(264, 11, 110, 23);
		frmBuWatsonDashboard.getContentPane().add(browseButton);
	}
	
}
