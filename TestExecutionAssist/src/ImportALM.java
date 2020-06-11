package src;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Properties;
import java.util.prefs.Preferences;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JTextField;
import javax.swing.JTree;
import javax.swing.tree.TreeModel;

import com.mercury.qualitycenter.otaclient.ClassFactory;
import com.mercury.qualitycenter.otaclient.ICustomization;
import com.mercury.qualitycenter.otaclient.ICustomizationList;
import com.mercury.qualitycenter.otaclient.ICustomizationListNode;
import com.mercury.qualitycenter.otaclient.ICustomizationLists;
import com.mercury.qualitycenter.otaclient.IList;
import com.mercury.qualitycenter.otaclient.ITDConnection;

import com4j.Com4jObject;

public class ImportALM {
	
	JFrame frmImportALM = new JFrame();
	private JTextField qcURLField;
	private JTextField qcUserIDfield;
	private JPasswordField passwordField;
	private JComboBox<String> domainCombo = new JComboBox<String>();
	private JComboBox<String> projectCombo = new JComboBox<String>();
	private static JTextField projectPathFieldALM = new JTextField();
	Properties props = new Properties();
	Preferences prefs = Preferences.userRoot().node(this.getClass().getName());
	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
	
	public ImportALM(String projectPath, JTree tcTree) {
		try {
			props.load(new FileInputStream(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg"));
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		/*Import from ALM frame*/
		frmImportALM.setIconImage(Toolkit.getDefaultToolkit().getImage(ImportALM.class.getResource("/res/logo48x48.png")));
		frmImportALM.setTitle("Import From ALM");
		frmImportALM.getContentPane().setLayout(null);
		frmImportALM.setResizable(false);
		frmImportALM.setSize( 500, 335);
		frmImportALM.setLocation(dim.width/2- frmImportALM.getSize().width/2, dim.height/2-frmImportALM.getSize().height/2);
		JPanel importALMPanel = new JPanel();
		importALMPanel.setBounds(0, 0, 494, 304);
		frmImportALM.getContentPane().add(importALMPanel);
		importALMPanel.setLayout(null);
		
		JCheckBox runImport = new JCheckBox("Import Last Run Details & Attachments");
		runImport.setFont(new Font("Tahoma", Font.BOLD, 11));
		runImport.setBounds(26, 218, 400, 20);
		importALMPanel.add(runImport);
		
		JButton btnALMImport = new JButton("Import ALM");
		btnALMImport.setFont(new Font("Tahoma", Font.BOLD, 11));
		btnALMImport.setBounds(351, 257, 105, 25);
		importALMPanel.add(btnALMImport);
		btnALMImport.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				createProject importXLCls = new createProject();
				if(!projectPathFieldALM.getText().trim().equals("")){
					if(checkALMconnection(domainCombo.getSelectedItem().toString(), projectCombo.getSelectedItem().toString(), qcUserIDfield.getText(), passwordField.getText())){
						setStatuses(props.getProperty("ALMURL"), domainCombo.getSelectedItem().toString(), projectCombo.getSelectedItem().toString(), qcUserIDfield.getText(), passwordField.getText());
						importXLCls.createProjectALM(projectPath, projectPathFieldALM.getText(), domainCombo.getSelectedItem().toString(), projectCombo.getSelectedItem().toString(), qcUserIDfield.getText(), passwordField.getText(),runImport.isSelected() );
						TreeModel model = new FileTreeModel(new File(projectPath));
						tcTree.setModel(model);
						tcTree.repaint();
						tcTree.setVisible(true);
						frmImportALM.dispose();
					}
					else {
						JOptionPane.showMessageDialog(null, "Unable to connect to ALM with give credentials!. Please enter correct ALM credentials");
					}
				}
				else {
					JOptionPane.showMessageDialog(null, "All fields are manadatory. Mandatory fields should not be empty.");
				}
			}
		});
		
		
		JLabel lblALMProjectPath = new JLabel("Test Suite Path");
		lblALMProjectPath.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblALMProjectPath.setBounds(26, 186, 100, 14);
		importALMPanel.add(lblALMProjectPath);
		
		
		
		projectPathFieldALM.setColumns(10);
		projectPathFieldALM.setBounds(136, 181, 320, 25);
		importALMPanel.add(projectPathFieldALM);
		
		JLabel qcURLlbl = new JLabel("ALM URL:");
		qcURLlbl.setFont(new Font("Tahoma", Font.BOLD, 11));
		qcURLlbl.setBounds(26, 27, 100, 14);
		importALMPanel.add(qcURLlbl);
		
		qcURLField = new JTextField();
		qcURLField.setEnabled(false);
		qcURLField.setColumns(10);
		qcURLField.setBounds(136, 23, 233, 25);
		qcURLField.setText(props.getProperty("ALMURL"));
		importALMPanel.add(qcURLField);
		
		qcUserIDfield = new JTextField();
		qcUserIDfield.setColumns(10);
		qcUserIDfield.setBounds(136, 54, 151, 25);
		qcUserIDfield.setText(prefs.get("UserID", ""));
		importALMPanel.add(qcUserIDfield);
		
		JLabel qcUserIDlbl = new JLabel("UserID:");
		qcUserIDlbl.setFont(new Font("Tahoma", Font.BOLD, 11));
		qcUserIDlbl.setBounds(26, 58, 100, 14);
		importALMPanel.add(qcUserIDlbl);
		
		JLabel Passwordlbl = new JLabel("Password:");
		Passwordlbl.setFont(new Font("Tahoma", Font.BOLD, 11));
		Passwordlbl.setBounds(26, 89, 100, 14);
		importALMPanel.add(Passwordlbl);
		
		JLabel lblDomain = new JLabel("Domain:");
		lblDomain.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblDomain.setBounds(26, 120, 100, 14);
		importALMPanel.add(lblDomain);
		
		JLabel lblProject = new JLabel("Project:");
		lblProject.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblProject.setBounds(26, 152, 100, 14);
		importALMPanel.add(lblProject);
		
		
		domainCombo.setBounds(136, 115, 151, 25);
		importALMPanel.add(domainCombo);
		ITDConnection itdc = ClassFactory.createTDConnection();
		if(!props.getProperty("ALMURL").equals("")) {
	        itdc.initConnectionEx(props.getProperty("ALMURL"));
	        IList domainsL = itdc.domainsList().queryInterface(IList.class);
	        for(int dCount = 1; dCount<=domainsL.count(); dCount++) {
	        	domainCombo.addItem(domainsL.item(dCount).toString());
	        }
		}
		
		
		projectCombo.setBounds(136, 147, 151, 25);
		importALMPanel.add(projectCombo);
		
		domainCombo.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				if(!domainCombo.getSelectedItem().toString().equals("")) {
					projectCombo.removeAllItems();
					IList projsL = itdc.projectsListEx(domainCombo.getSelectedItem().toString()).queryInterface(IList.class);
			        for(int dCount = 1; dCount<=projsL.count(); dCount++) {
			        	projectCombo.addItem(projsL.item(dCount).toString());
			        }
				}
			}
		});
		
		passwordField = new JPasswordField();
		passwordField.setBounds(136, 85, 151, 25);
		passwordField.setText(prefs.get("pass", ""));
		importALMPanel.add(passwordField);
		frmImportALM.setVisible(true);
	}
	
	public boolean checkALMconnection(String domainName, String projectname, String userID, String pass) {
		boolean connected = false;
		try {
			ITDConnection itdc= ClassFactory.createTDConnection();
	        itdc.initConnectionEx(props.getProperty("ALMURL"));
	        itdc.connectProjectEx(domainName, projectname, userID, pass);
	        if(itdc.connected()) {
	        	connected = true;
	        }
	        else {
	        	connected = false;
	        }
		}
		catch(Exception e) {
			connected = false;
		}
		return connected;
	}
	
	public void setStatuses(String URL, String domain, String project, String userID, String pass) {
		String statuslist = "";
		try{
			ITDConnection itdc= ClassFactory.createTDConnection();
	        itdc.initConnectionEx(URL);
	        itdc.connectProjectEx(domain, project, userID, pass);
	        itdc.ignoreHtmlFormat(true);
	        
	        ICustomization cust = itdc.customization().queryInterface(ICustomization.class);
	        ICustomizationLists fields = cust.lists().queryInterface(ICustomizationLists.class);
	        for(int eachCount = 1; eachCount <= fields.count(); eachCount++) {
	        	ICustomizationList eachList = fields.listByCount(eachCount).queryInterface(ICustomizationList.class);
	        	if(eachList.rootNode().queryInterface(ICustomizationListNode.class).name().equals("Status")){
	        		ICustomizationListNode eachNode = eachList.rootNode().queryInterface(ICustomizationListNode.class);
	        		for(int eachChildNode = 1; eachChildNode <= eachNode.childrenCount() ; eachChildNode++) {
	        			ICustomizationListNode eachCild = ((Com4jObject)eachNode.children(eachChildNode)).queryInterface(ICustomizationListNode.class);
	        			System.out.println(eachCild.name());
	        			statuslist = statuslist.equals("") ? eachCild.name() : statuslist + ";" + eachCild.name();
	        		}
	        		break;
	        	}
	        }
	        props.setProperty("Status", statuslist);
			props.store(new FileOutputStream(new File(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg")), "");
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
}
