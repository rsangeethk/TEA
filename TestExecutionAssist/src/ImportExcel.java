package src;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Properties;
import java.util.prefs.Preferences;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.JTree;
import javax.swing.tree.TreeModel;

public class ImportExcel {
	JFrame frmImportExcel = new JFrame();
	private static JTextField importXLPath;
	private static JTextField projectPathField = new JTextField();
	Properties props = new Properties();
	Preferences prefs = Preferences.userRoot().node(this.getClass().getName());
	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
	
	public ImportExcel(String projectPath, JTree tcTree){
		/*Import from Excel frame*/
		frmImportExcel.setIconImage(Toolkit.getDefaultToolkit().getImage(ImportExcel.class.getResource("/res/logo48x48.png")));
		frmImportExcel.setTitle("Import Excel");
		frmImportExcel.getContentPane().setLayout(null);
		frmImportExcel.setResizable(false);
		frmImportExcel.setSize( 500, 217);
		frmImportExcel.setLocation(dim.width/2- frmImportExcel.getSize().width/2, dim.height/2-frmImportExcel.getSize().height/2);
		frmImportExcel.setVisible(false);
		JPanel importXLPanel = new JPanel();
		importXLPanel.setBounds(0, 0, 485, 187);
		frmImportExcel.getContentPane().add(importXLPanel);
		importXLPanel.setLayout(null);
		
		importXLPath = new JTextField();
		importXLPath.setBounds(120, 39, 300, 25);
		importXLPanel.add(importXLPath);
		importXLPath.setColumns(10);
		
		JButton btnBrowse = new JButton("Browse");
		btnBrowse.setBounds(428, 39, 40, 25);
		importXLPanel.add(btnBrowse);
		btnBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				int rVal = fileChooser.showOpenDialog(null);
				if (rVal == JFileChooser.APPROVE_OPTION) {
					importXLPath.setText(fileChooser.getSelectedFile().toString());
				}
			}
		});
		JButton btnImport = new JButton("Import");
		btnImport.setFont(new Font("Tahoma", Font.BOLD, 11));
		btnImport.setBounds(350, 146, 105, 25);
		importXLPanel.add(btnImport);
		btnImport.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				createProject importXLCls = new createProject();
				if(!(importXLPath.getText().trim().equals("") && projectPathField.getText().trim().equals(""))){
					importXLCls.createProjectExcel(projectPath, importXLPath.getText(), projectPathField.getText());
					TreeModel model = new FileTreeModel(new File(projectPath));
					tcTree.setModel(model);
					tcTree.repaint();
					tcTree.setVisible(true);
					frmImportExcel.setVisible(false);
				}
				else {
					JOptionPane.showMessageDialog(null, "All fields are manadatory. Mandatory fields should not be empty.");
				}
			}
		});
		
		JLabel lblExcel = new JLabel("Import File");
		lblExcel.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblExcel.setBounds(21, 44, 100, 14);
		importXLPanel.add(lblExcel);
		
		JLabel lblProjectPath = new JLabel("Test Suite Name");
		lblProjectPath.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblProjectPath.setBounds(21, 95, 100, 14);
		importXLPanel.add(lblProjectPath);
		
		projectPathField.setColumns(10);
		projectPathField.setBounds(120, 90, 300, 25);
		importXLPanel.add(projectPathField);
		frmImportExcel.setVisible(true);
	}
}
