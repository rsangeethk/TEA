package src;

import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Enumeration;
import java.util.Properties;
import java.util.prefs.Preferences;

import javax.swing.AbstractButton;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JRadioButton;
import javax.swing.JTabbedPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;

import com.mercury.qualitycenter.otaclient.ClassFactory;
import com.mercury.qualitycenter.otaclient.ICustomization;
import com.mercury.qualitycenter.otaclient.ICustomizationList;
import com.mercury.qualitycenter.otaclient.ICustomizationListNode;
import com.mercury.qualitycenter.otaclient.ICustomizationLists;
import com.mercury.qualitycenter.otaclient.IList;
import com.mercury.qualitycenter.otaclient.ITDConnection;

import com4j.Com4jObject;

public class Prefs {
	private JTextField almurlField;
	private JTextField tcNameCol;
	private JTextField tDescCol;
	private JTextField stepNCol;
	private JTextField stepDescCol;
	private JTextField stepExpCol;
	private JTextField stepActCol;
	private JPasswordField passField;
	private JTextField userIDField;
	private JRadioButton rdbtnNone;
	private JRadioButton rdbtnAll;
	private JCheckBox chckbxExpectedResult;
	private JCheckBox chckbxActualResult;
	private JCheckBox chckbxStepDescription;
	private JCheckBox chckbxStepName;
	private JRadioButton rdbtnMaximized;
	private JRadioButton rdbtnMinimized;
	private JRadioButton rdbtnAttachInTest = new JRadioButton("Attach in Test Run");
	private JRadioButton rdbtnAttachInTest_1 = new JRadioButton("Attach in Test Case");
	private JRadioButton rdbtnNoAttach = new JRadioButton("No Attachment");
	JTextArea textAreaPassActual = new JTextArea();
	JTextArea textAreaFailActual = new JTextArea();
	
	private JButton btnsetApply;
	private JButton btnsetCancel;
	private ButtonGroup scrnshtOption = new ButtonGroup();
	private JFrame frmPreferences = new JFrame();
	private JLabel lblDelayTime;
	private JTextField delayField;
	JComboBox<String> screenType = new JComboBox<String>();
	JButton btnArrange = new JButton("");
	String expandedModule = "";
	JTextField lclNameField = new JTextField();
	Properties props = new Properties();
	Preferences prefs = Preferences.userRoot().node(this.getClass().getName());
	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
	
	public Prefs() {
		/* Settings Panel */
		frmPreferences.setIconImage(Toolkit.getDefaultToolkit().getImage(Prefs.class.getResource("/res/logo48x48.png")));
		frmPreferences.setTitle("Preferences");
		frmPreferences.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		frmPreferences.setResizable(false);
		frmPreferences.getContentPane().setLayout(null);
		frmPreferences.setSize(400, 385);
		frmPreferences.setLocation(dim.width/2- frmPreferences.getSize().width/2, dim.height/2-frmPreferences.getSize().height/2);
		JPanel settingsPanel = new JPanel();
		settingsPanel.setBounds(0, 0, 394, 363);
		frmPreferences.getContentPane().add(settingsPanel);
		settingsPanel.setLayout(null);
		
		JTabbedPane settingTabbedPane = new JTabbedPane(JTabbedPane.TOP);
		settingTabbedPane.setBorder(null);
		settingTabbedPane.setBounds(0, 0, 394, 305);
		settingsPanel.add(settingTabbedPane);
		
		JPanel reportPanel = new JPanel();
		settingTabbedPane.addTab("Reports", null, reportPanel, null);
		reportPanel.setLayout(null);
		
		JLabel lblReportContents = new JLabel("Report contents");
		lblReportContents.setFont(new Font("Tahoma", Font.PLAIN, 11));
		lblReportContents.setBounds(22, 11, 135, 14);
		reportPanel.add(lblReportContents);
		
		chckbxStepName = new JCheckBox("Step name");
		chckbxStepName.setBounds(22, 32, 97, 23);
		reportPanel.add(chckbxStepName);
		
		chckbxStepDescription = new JCheckBox("Step description");
		chckbxStepDescription.setBounds(22, 58, 135, 23);
		reportPanel.add(chckbxStepDescription);
		
		chckbxActualResult = new JCheckBox("Actual Result");
		chckbxActualResult.setBounds(22, 84, 135, 23);
		reportPanel.add(chckbxActualResult);
		
		chckbxExpectedResult = new JCheckBox("Expected Result");
		chckbxExpectedResult.setBounds(22, 110, 135, 23);
		reportPanel.add(chckbxExpectedResult);
		
		JLabel lblScreenshots = new JLabel("Screenshots");
		lblScreenshots.setBounds(22, 140, 73, 14);
		reportPanel.add(lblScreenshots);
		
		
		
		rdbtnAll = new JRadioButton("All");
		rdbtnAll.setActionCommand("All");
		rdbtnAll.setBounds(46, 161, 47, 23);
		scrnshtOption.add(rdbtnAll);
		reportPanel.add(rdbtnAll);
		
		rdbtnNone = new JRadioButton("None");
		rdbtnNone.setActionCommand("None");
		rdbtnNone.setBounds(95, 161, 73, 23);
		scrnshtOption.add(rdbtnNone);
		reportPanel.add(rdbtnNone);
		
		JPanel importsPanel = new JPanel();
		settingTabbedPane.addTab("Excel", null, importsPanel, null);
		importsPanel.setLayout(null);
		
		JLabel lblImportFromExcel = new JLabel("<html><body>Import from Excel Configuration <i>(Column Names in Import sheet)</i></body></html>");
		lblImportFromExcel.setBounds(10, 11, 364, 14);
		importsPanel.add(lblImportFromExcel);
		
		JLabel lblTestCaseName = new JLabel("Test Case Name:");
		lblTestCaseName.setBounds(20, 40, 114, 14);
		importsPanel.add(lblTestCaseName);
		
		JLabel lblTestDescription = new JLabel("Test Description:");
		lblTestDescription.setBounds(20, 76, 114, 14);
		importsPanel.add(lblTestDescription);
		
		JLabel lblStepName = new JLabel("Step Name:");
		lblStepName.setBounds(20, 112, 92, 14);
		importsPanel.add(lblStepName);
		
		JLabel lblStepDescription = new JLabel("Step Description:");
		lblStepDescription.setBounds(20, 149, 114, 14);
		importsPanel.add(lblStepDescription);
		
		JLabel lblExpectedResult = new JLabel("Expected Result:");
		lblExpectedResult.setBounds(20, 185, 92, 14);
		importsPanel.add(lblExpectedResult);
		
		JLabel lblActualResult = new JLabel("Actual Result:");
		lblActualResult.setBounds(20, 222, 92, 14);
		importsPanel.add(lblActualResult);
		
		tcNameCol = new JTextField();
		tcNameCol.setBounds(133, 36, 169, 28);
		importsPanel.add(tcNameCol);
		tcNameCol.setColumns(10);
		
		tDescCol = new JTextField();
		tDescCol.setColumns(10);
		tDescCol.setBounds(133, 72, 169, 28);
		importsPanel.add(tDescCol);
		
		stepNCol = new JTextField();
		stepNCol.setColumns(10);
		stepNCol.setBounds(133, 108, 169, 28);
		importsPanel.add(stepNCol);
		
		stepDescCol = new JTextField();
		stepDescCol.setColumns(10);
		stepDescCol.setBounds(133, 144, 169, 28);
		importsPanel.add(stepDescCol);
		
		stepExpCol = new JTextField();
		stepExpCol.setColumns(10);
		stepExpCol.setBounds(133, 180, 169, 28);
		importsPanel.add(stepExpCol);
		
		stepActCol = new JTextField();
		stepActCol.setColumns(10);
		stepActCol.setBounds(133, 216, 169, 28);
		importsPanel.add(stepActCol);
		
		JPanel almPanel = new JPanel();
		settingTabbedPane.addTab("ALM", null, almPanel, null);
		almPanel.setLayout(null);
		
		JLabel lblAlmUrl = new JLabel("ALM URL: ");
		lblAlmUrl.setBounds(10, 14, 66, 14);
		almPanel.add(lblAlmUrl);
		
		JLabel lblUserId = new JLabel("User ID:");
		lblUserId.setBounds(10, 41, 64, 25);
		almPanel.add(lblUserId);
		
		JLabel lblPassword = new JLabel("Password:");
		lblPassword.setBounds(10, 75, 64, 14);
		almPanel.add(lblPassword);
		
		almurlField = new JTextField();
		almurlField.setBounds(76, 8, 201, 28);
		almPanel.add(almurlField);
		almurlField.setColumns(10);
		
		userIDField = new JTextField();
		userIDField.setColumns(10);
		userIDField.setBounds(76, 38, 120, 28);
		almPanel.add(userIDField);
		
		passField = new JPasswordField();
		passField.setBounds(76, 69, 120, 27);
		almPanel.add(passField);
		
		JButton btnConnect = new JButton("Test");
		btnConnect.setBounds(221, 75, 91, 23);
		almPanel.add(btnConnect);
		btnConnect.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				if(!almurlField.getText().equals("") && !userIDField.getText().equals("") && !passField.getText().equals("")) {
					ITDConnection itdc= ClassFactory.createTDConnection();
			        itdc.initConnectionEx(almurlField.getText());
			        
					JDialog domainDialog = new JDialog();
					domainDialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
					domainDialog.setIconImage(Toolkit.getDefaultToolkit().getImage(Try.class.getResource("/res/logo48x48.png")));
					domainDialog.setTitle("Domain and Project");
					domainDialog.setSize(300, 164);
					FlowLayout flowLayout = new FlowLayout();
					flowLayout.setVgap(15);
					flowLayout.setHgap(10);
					domainDialog.getContentPane().setLayout(flowLayout);
					
					JLabel lblDomain = new JLabel("Domain :");
					domainDialog.getContentPane().add(lblDomain);
					
					JComboBox<String> comboDomain = new JComboBox<String>();
					comboDomain.setPreferredSize(new Dimension(200, 26));
					domainDialog.getContentPane().add(comboDomain);
					IList domainsL = itdc.domainsList().queryInterface(IList.class);
			        for(int dCount = 1; dCount<=domainsL.count(); dCount++) {
			        	comboDomain.addItem(domainsL.item(dCount).toString());
			        }
					
					JLabel lblProject = new JLabel("Project   :");
					domainDialog.getContentPane().add(lblProject);
					
					JComboBox<String> comboProject = new JComboBox<String>();
					comboProject.setPreferredSize(new Dimension(200, 26));
					domainDialog.getContentPane().add(comboProject);
					
					comboDomain.addActionListener(new ActionListener() {
						@Override
						public void actionPerformed(ActionEvent e) {
							if(!comboDomain.getSelectedItem().toString().equals("")) {
								comboProject.removeAllItems();
								IList projsL = itdc.projectsListEx(comboDomain.getSelectedItem().toString()).queryInterface(IList.class);
						        for(int dCount = 1; dCount<=projsL.count(); dCount++) {
						        	comboProject.addItem(projsL.item(dCount).toString());
						        }
							}
						}
					});
					
					JButton btnTest = new JButton("Test");
					btnTest.setPreferredSize(new Dimension(100, 28));
					domainDialog.getContentPane().add(btnTest);
					btnTest.addActionListener(new ActionListener() {
						@Override
						public void actionPerformed(ActionEvent e) {
							if(!comboDomain.getSelectedItem().toString().trim().equals("") && !comboProject.getSelectedItem().toString().trim().equals("")) {
							try {
						        itdc.connectProjectEx(comboDomain.getSelectedItem().toString(), comboProject.getSelectedItem().toString(), userIDField.getText(), passField.getText());
							        if(itdc.connected()) {
							        	prefs.put("UserID", userIDField.getText());
										prefs.put("pass", passField.getText());
							        	JOptionPane.showMessageDialog(null, "ALM Connection Test : Success");
							        }
							        else {
							        	JOptionPane.showMessageDialog(null, "ALM Connection Test : Failed. Please check your credentials");
							        }
						        }
						        catch(Exception ex) {
						        	JOptionPane.showMessageDialog(null, "ALM Connection Test : Failed. Please check your credentials");
						        }
							}
							else {
								JOptionPane.showMessageDialog(null, "Please select a domain and project");
							}
							domainDialog.dispose();
						}
					});
					domainDialog.setVisible(true);
					
			        
				}
				else {
					JOptionPane.showMessageDialog(null, "Please enter all mandatory ALM details!");
				}
			}
		});
		
		JLabel lblPasswordStoredSecurely = new JLabel("(Password Stored securely)");
		lblPasswordStoredSecurely.setFont(new Font("Tahoma", Font.ITALIC, 11));
		lblPasswordStoredSecurely.setBounds(76, 97, 136, 14);
		almPanel.add(lblPasswordStoredSecurely);
		
		ButtonGroup attachLocation = new ButtonGroup();
		JLabel lblReportAttachment = new JLabel("Report Attachment");
		lblReportAttachment.setBounds(10, 118, 114, 14);
		almPanel.add(lblReportAttachment);
		rdbtnAttachInTest.setBounds(15, 139, 120, 23);
		attachLocation.add(rdbtnAttachInTest);
		almPanel.add(rdbtnAttachInTest);
		
		rdbtnAttachInTest_1.setBounds(137, 139, 140, 23);
		attachLocation.add(rdbtnAttachInTest_1);
		almPanel.add(rdbtnAttachInTest_1);
		
		rdbtnNoAttach.setBounds(270, 139, 120, 23);
		attachLocation.add(rdbtnNoAttach);
		almPanel.add(rdbtnNoAttach);
		
		JLabel lblStepRun = new JLabel("Step Run Actuals");
		lblStepRun.setBounds(10, 169, 114, 14);
		almPanel.add(lblStepRun);
		
		JLabel lblActualResultFor = new JLabel("Actual Result for Passed: ");
		lblActualResultFor.setBounds(20, 194, 149, 14);
		almPanel.add(lblActualResultFor);
		textAreaPassActual.setBounds(159, 190, 201, 37);
		almPanel.add(textAreaPassActual);
		textAreaFailActual.setBounds(159, 238, 201, 37);
		almPanel.add(textAreaFailActual);
		
		JLabel lblActualResultFor_1 = new JLabel("Actual Result for Failed: ");
		lblActualResultFor_1.setBounds(20, 242, 149, 14);
		almPanel.add(lblActualResultFor_1);
		
		JPanel runPanel = new JPanel();
		settingTabbedPane.addTab("Run Dialog", null, runPanel, null);
		runPanel.setLayout(null);
		
		JLabel lblOpenRunDialog = new JLabel("Open Run Dialog in view");
		lblOpenRunDialog.setBounds(11, 11, 148, 14);
		runPanel.add(lblOpenRunDialog);
		
		ButtonGroup runView = new ButtonGroup();
		rdbtnMaximized = new JRadioButton("Maximized");
		rdbtnMaximized.setBounds(21, 29, 95, 23);
		runView.add(rdbtnMaximized);
		runPanel.add(rdbtnMaximized);
		
		rdbtnMinimized = new JRadioButton("Minimized");
		rdbtnMinimized.setBounds(117, 29, 109, 23);
		runView.add(rdbtnMinimized);
		runPanel.add(rdbtnMinimized);
		
		lblDelayTime = new JLabel("Time Delay for screenshot capture: ");
		lblDelayTime.setBounds(4, 74, 194, 14);
		runPanel.add(lblDelayTime);
		
		delayField = new JTextField();
		delayField.setBounds(196, 71, 65, 23);
		runPanel.add(delayField);
		delayField.setColumns(10);
		
		JLabel lblSeconds = new JLabel("milliseconds");
		lblSeconds.setBounds(267, 74, 79, 14);
		runPanel.add(lblSeconds);
		ButtonGroup autoGenRadio = new ButtonGroup();
		
		
		
		btnsetApply = new JButton("Apply");
		btnsetApply.setBounds(194, 316, 91, 28);
		settingsPanel.add(btnsetApply);
		btnsetApply.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				setPreferencestoFile();
			}
		});
		
		btnsetCancel = new JButton("Cancel");
		btnsetCancel.setBounds(295, 316, 91, 28);
		btnsetCancel.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				frmPreferences.dispose();
			}
		});
		settingsPanel.add(btnsetCancel);
		frmPreferences.setVisible(true);	
		setPreferencestoUI();
	}
	
	public void setPreferencestoFile() {
		try {
			
			//Run Dialog
			props.load(new FileInputStream(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg"));
			
			/*if(chckbxTakeScreenshotOn.isSelected()) props.setProperty("PassScreenshot", "Yes"); else props.setProperty("PassScreenshot", "No");
			if(chckbxTakeScreenshotOn_1.isSelected()) props.setProperty("FailScreenshot", "Yes"); else props.setProperty("FailScreenshot", "No");
			if(chckbxAutoNavigationTo.isSelected()) props.setProperty("AutoNextStep", "Yes"); else props.setProperty("AutoNextStep", "No");*/	
			
			if(rdbtnMaximized.isSelected()) props.setProperty("DialogView", "Maximized"); else props.setProperty("DialogView", "Minimized");
			props.setProperty("CaptureDelay", delayField.getText());
			
			//Reports
			/*if(rdbtnAutoGenerateReport.isSelected()) props.setProperty("AutoReport", "Yes"); else props.setProperty("AutoReport", "No");*/
			if(chckbxStepName.isSelected()) props.setProperty("StepName", "Yes"); else props.setProperty("StepName", "No");
			if(chckbxStepDescription.isSelected()) props.setProperty("StepDescription", "Yes"); else props.setProperty("StepDescription", "No");
			if(chckbxActualResult.isSelected()) props.setProperty("ActualResult", "Yes"); else props.setProperty("ActualResult", "No");
			if(chckbxExpectedResult.isSelected()) props.setProperty("ExpectedResult", "Yes"); else props.setProperty("ExpectedResult", "No");
			props.setProperty("Screenshots", scrnshtOption.getSelection().getActionCommand());
			
			//ALM
			props.setProperty("ALMURL", almurlField.getText());
			if(rdbtnAttachInTest.isSelected()){ props.setProperty("AttachTestRun", "Yes");} 
			else if(rdbtnAttachInTest_1.isSelected()){ props.setProperty("AttachTestRun", "No");}
			else if(rdbtnNoAttach.isSelected()){ props.setProperty("AttachTestRun", "None");}
			props.setProperty("PassedActual", textAreaPassActual.getText());
			props.setProperty("FailedActual", textAreaFailActual.getText());
			
			//Import
			props.setProperty("TCName", tcNameCol.getText());
			props.setProperty("TCDesc", tDescCol.getText());
			props.setProperty("SName", stepNCol.getText());
			props.setProperty("SDesc", stepDescCol.getText());
			props.setProperty("ActRes", stepActCol.getText());
			props.setProperty("ExpRes", stepExpCol.getText());
			prefs.put("UserID", userIDField.getText());
			prefs.put("pass", passField.getText());
			
			props.store(new FileOutputStream(new File(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg")), "");
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public void setPreferencestoUI() {
		try {
			
			//Run Dialog
			props.load(new FileInputStream(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg"));
			
			/*if(props.getProperty("PassScreenshot").equalsIgnoreCase("Yes")) chckbxTakeScreenshotOn.setSelected(true) ; else chckbxTakeScreenshotOn.setSelected(false);
			if(props.getProperty("FailScreenshot").equalsIgnoreCase("Yes")) chckbxTakeScreenshotOn_1.setSelected(true) ; else chckbxTakeScreenshotOn_1.setSelected(false);
			if(props.getProperty("AutoNextStep").equalsIgnoreCase("Yes")) chckbxAutoNavigationTo.setSelected(true) ; else chckbxAutoNavigationTo.setSelected(false);*/
			
			if(props.getProperty("DialogView").equalsIgnoreCase("Maximized")) rdbtnMaximized.setSelected(true) ; else rdbtnMinimized.setSelected(true);
			delayField.setText(props.getProperty("CaptureDelay"));
			
			//Reports
			/*if(props.getProperty("AutoReport").equalsIgnoreCase("Yes")) rdbtnAutoGenerateReport.setSelected(true) ; else rdbtnAutoGenerateReport.setSelected(false);*/
			if(props.getProperty("StepName").equalsIgnoreCase("Yes")) chckbxStepName.setSelected(true) ; else chckbxStepName.setSelected(false);
			if(props.getProperty("StepDescription").equalsIgnoreCase("Yes")) chckbxStepDescription.setSelected(true) ; else chckbxStepDescription.setSelected(false);
			if(props.getProperty("ActualResult").equalsIgnoreCase("Yes")) chckbxActualResult.setSelected(true) ; else chckbxActualResult.setSelected(false);
			if(props.getProperty("ExpectedResult").equalsIgnoreCase("Yes")) chckbxExpectedResult.setSelected(true) ; else chckbxExpectedResult.setSelected(false);
			
			scrnshtOption.clearSelection();
			Enumeration<AbstractButton> btns = scrnshtOption.getElements();
			while(btns.hasMoreElements()) {
				AbstractButton radBtn = btns.nextElement();
				if(radBtn.getActionCommand().equalsIgnoreCase(props.getProperty("Screenshots"))) {
					radBtn.setSelected(true);
					break;
				}
			}
			
			//ALM
			almurlField.setText(props.getProperty("ALMURL"));
			userIDField.setText(prefs.get("UserID", ""));
			passField.setText(prefs.get("pass", ""));
			textAreaPassActual.setText(props.getProperty("PassedActual"));
			textAreaFailActual.setText(props.getProperty("FailedActual"));
			if(props.getProperty("AttachTestRun").equalsIgnoreCase("Yes")){ rdbtnAttachInTest.setSelected(true) ;} 
			else if(props.getProperty("AttachTestRun").equalsIgnoreCase("No")){ rdbtnAttachInTest_1.setSelected(true);}
			else if(props.getProperty("AttachTestRun").equalsIgnoreCase("None")){ rdbtnNoAttach.setSelected(true);}
			
			
			//Import
			tcNameCol.setText(props.getProperty("TCName"));
			tDescCol.setText(props.getProperty("TCDesc"));
			stepNCol.setText(props.getProperty("SName"));
			stepDescCol.setText(props.getProperty("SDesc"));
			stepActCol.setText(props.getProperty("ActRes"));
			stepExpCol.setText(props.getProperty("ExpRes"));
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
}

