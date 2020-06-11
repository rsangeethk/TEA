package src;

import java.awt.Component;
import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JDialog;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.border.EmptyBorder;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Branch;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

import com.mercury.qualitycenter.otaclient.ClassFactory;
import com.mercury.qualitycenter.otaclient.IAttachment;
import com.mercury.qualitycenter.otaclient.IAttachmentFactory;
import com.mercury.qualitycenter.otaclient.IBaseFactory;
import com.mercury.qualitycenter.otaclient.IDesignStep;
import com.mercury.qualitycenter.otaclient.IExtendedStorage;
import com.mercury.qualitycenter.otaclient.IList;
import com.mercury.qualitycenter.otaclient.IRun;
import com.mercury.qualitycenter.otaclient.IRunFactory;
import com.mercury.qualitycenter.otaclient.IStep;
import com.mercury.qualitycenter.otaclient.ITDConnection;
import com.mercury.qualitycenter.otaclient.ITSTest;
import com.mercury.qualitycenter.otaclient.ITest;
import com.mercury.qualitycenter.otaclient.ITestSet;
import com.mercury.qualitycenter.otaclient.ITestSetFolder;
import com.mercury.qualitycenter.otaclient.ITestSetTreeManager;

import com4j.Com4jObject;

public class createProject {
	Properties props = new Properties();
	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();

	public void createProjectExcel(String projectPath, String importFile, String moduleName) {
		try {
			try {
				props.load(new FileInputStream(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg"));
			}
			catch(Exception e) {
				
			}
			if(!props.getProperty("TCName").trim().equals("") && !props.getProperty("TCDesc").trim().equals("") && !props.getProperty("SDesc").trim().equals("") && !props.getProperty("ExpRes").trim().equals("")) {
				FileInputStream tcStream = new FileInputStream(new File(importFile));
				XSSFWorkbook tcBook = new XSSFWorkbook(tcStream);
				XSSFSheet tcSheet = tcBook.getSheetAt(0);
				int testNameCol = getColumnIndex(props.getProperty("TCName"), 0, tcSheet);
				int testDescriptionCol = getColumnIndex(props.getProperty("TCDesc"), 0, tcSheet);
				int testStepCol = getColumnIndex(props.getProperty("SDesc"), 0, tcSheet);
				int testExpectedResult = getColumnIndex(props.getProperty("ExpRes"), 0, tcSheet);
				String curTCName = "";
				Document document = null;
				String curtestDet = "";
				String modulePath = projectPath + "\\" + moduleName;
				if(!new File(modulePath).exists()) {
					new File(modulePath).mkdir();
					Document moduleDoc = DocumentFactory.getInstance().createDocument();
			        Element moduleRoot =  (Element) moduleDoc.addElement("module");
			        moduleRoot.addElement("project").setText(new File(projectPath).getName());
			        moduleRoot.addElement("pulled").setText("Excel");
			        moduleRoot.addElement("TestLab").setText("");
			        moduleRoot.addElement("ALMDomain").setText("");
			        moduleRoot.addElement("ALMProject").setText("");
			        moduleRoot.addElement("pulledon").setText(new SimpleDateFormat("MM/dd/yyyy").format(new Date()));
			        moduleRoot.addElement("plannedstart").setText("");
			        moduleRoot.addElement("actualstart");
			        moduleRoot.addElement("plannedend").setText("");
			        moduleRoot.addElement("actualend");
			        FileOutputStream modFOS = new FileOutputStream(modulePath + "\\moduleDetails.ts");
					OutputFormat format = OutputFormat.createPrettyPrint();
			        XMLWriter modWriter = new XMLWriter(modFOS, format);
			        modWriter.write(moduleDoc);
			        modWriter.close();
					for(int rowIndex = 1; rowIndex<=tcSheet.getLastRowNum(); rowIndex++) {
						if(tcSheet.getRow(rowIndex).getCell(testNameCol)!=null){
							if(document!=null) {
								FileOutputStream fos = new FileOutputStream(curtestDet);
						        XMLWriter writer = new XMLWriter(fos, format);
						        writer.write(document);
						        writer.close();
							}
							curTCName = tcSheet.getRow(rowIndex).getCell(testNameCol).getStringCellValue();
							if(!new File(modulePath + "\\" + curTCName).exists()) {
								new File(modulePath + "\\" + curTCName).mkdir();
								new File(modulePath + "\\" + curTCName + "\\Screenshots").mkdir();
								new File(modulePath + "\\" + curTCName + "\\Reports").mkdir();
								curtestDet = modulePath + "\\" + curTCName + "\\testDetails.tc";
						        document = DocumentFactory.getInstance().createDocument();
						        Element root =  (Element) document.addElement("testcase");
						        root.addElement("testcasename").setText(curTCName);
						        root.addElement("testdescription").setText(tcSheet.getRow(rowIndex).getCell(testDescriptionCol).getStringCellValue());
						        root.addElement("testdata");
						        root.addElement("executiontime");
						        root.addElement("lastexecutedstep");
						        root.addElement("executedby");
						        root.addElement("executedon");
						        root.addElement("overallstatus");
						        root.addElement("report");
						        root.addElement("runid").setText("");
						        Element tsNode = root.addElement("teststep");
								tsNode.addElement("stepdescription").setText(tcSheet.getRow(rowIndex).getCell(testStepCol).getStringCellValue());
								tsNode.addElement("expectedresult").setText(tcSheet.getRow(rowIndex).getCell(testExpectedResult).getStringCellValue());
								tsNode.addElement("actualresult").setText("");
								tsNode.addElement("stepstatus").setText("");
								tsNode.addElement("screenshot").setText("");
							}
						}
						else {
							if(document!=null) {
								Element tcNode =  (Element) document.selectNodes("/testcase").get(0);
								Element tsNode = tcNode.addElement("teststep");
								tsNode.addElement("stepdescription").setText(tcSheet.getRow(rowIndex).getCell(testStepCol).getStringCellValue());
								tsNode.addElement("expectedresult").setText(tcSheet.getRow(rowIndex).getCell(testExpectedResult).getStringCellValue());
								tsNode.addElement("actualresult").setText("");
								tsNode.addElement("stepstatus").setText("");
								tsNode.addElement("screenshot").setText("");
							}
						}
					}
					FileOutputStream fos = new FileOutputStream(curtestDet);
			        XMLWriter writer = new XMLWriter(fos, format);
			        writer.write(document);
			        writer.close();
				}
			}
			else{
				JOptionPane.showMessageDialog(null, "Test case columns should not be empty. Please setup Import Column details in Option->Preferences");
			}
		}
        catch(Exception e) {
        	e.printStackTrace();
        }
	}
	
	public void createProjectALM(String projectPath, String testLabPath, String domainName, String projectname, String userID, String Pass, boolean download) {
		try {
			ITDConnection itdc= ClassFactory.createTDConnection();
	        itdc.initConnectionEx(props.getProperty("ALMURL"));
	        itdc.connectProjectEx(domainName, projectname, userID, Pass);
	        itdc.ignoreHtmlFormat(true);	        
	        ITestSetTreeManager sTestSetTree = itdc.testSetTreeManager().queryInterface(ITestSetTreeManager.class);
	        ITestSetFolder sTestFolder = null;
	        try {
	        	sTestFolder = sTestSetTree.nodeByPath(testLabPath).queryInterface(ITestSetFolder.class);
	        }
	        catch(Exception exe) {
	        	JOptionPane.showMessageDialog(null, "The given ALM folder path doesn't exist. Please enter correct ALM folder path");
	        	return;
	        }
	        IList foundList = sTestFolder.findTestSets(null, true, null).queryInterface(IList.class);
			JDialog dialog = new JDialog();
			dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
	        dialog.setTitle("Select Test Suites");
			dialog.getContentPane().setLayout(null);
			dialog.setSize(300, 300);
			dialog.setLocation(dim.width/2- dialog.getSize().width/2, dim.height/2-dialog.getSize().height/2);
			JPanel suitesPanel = new JPanel();
			WrapFlowLayout wfl_suitesPanel = new WrapFlowLayout(WrapFlowLayout.LEFT);
			wfl_suitesPanel.setVgap(10);
			suitesPanel.setLayout(wfl_suitesPanel);
			JScrollPane suitesScroll = new JScrollPane(suitesPanel);
			suitesScroll.setBounds(0, 0, 290, 239);
			dialog.getContentPane().add(suitesScroll);
			
			JButton btnImport = new JButton("Import");
			btnImport.setBounds(47, 244, 91, 23);
			dialog.getContentPane().add(btnImport);
			
			JButton btnCancel = new JButton("Cancel");
			btnCancel.setBounds(156, 244, 91, 23);
			dialog.getContentPane().add(btnCancel);
			btnCancel.addActionListener(new ActionListener() {
				@Override
				public void actionPerformed(ActionEvent e) {
					dialog.dispose();
				}
			});
	        for(int eachTLcnt = 1; eachTLcnt<=foundList.count(); eachTLcnt++) {
	        	ITestSet tsfac = ((Com4jObject) foundList.item(eachTLcnt)).queryInterface(ITestSet.class);
	        	if(((ITestSetFolder)tsfac.testSetFolder().queryInterface(ITestSetFolder.class)).path().equals(testLabPath)) {
		    		JCheckBox testChkBox = new JCheckBox(tsfac.name());
		    		if(testChkBox.getPreferredSize().width<=280) {
		    			testChkBox.setBorder(new EmptyBorder(0, 0, 0, 280-testChkBox.getPreferredSize().width));
		    		}
		    		System.out.println(testChkBox.getText());
		    		suitesPanel.add(testChkBox);
	        	}
	        }
	        List<String> checkboxes = new ArrayList<String>();
	        btnImport.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent arg0) {
					try {
						for( Component comp : suitesPanel.getComponents() ) {
							   if(comp instanceof JCheckBox){
								   if(((JCheckBox)comp).isSelected()){
									   checkboxes.add(((JCheckBox)comp).getText());
							   }
						   }
						}
						String curtestDet = "";
				        for(int eachTLcnt = 1; eachTLcnt<=foundList.count(); eachTLcnt++) {
				        	ITestSet tsfac = ((Com4jObject) foundList.item(eachTLcnt)).queryInterface(ITestSet.class);
				        	if(checkboxes.contains(tsfac.name())) {
				        		String modulePath = projectPath + "\\" + tsfac.name();
				        		int foldcount = 1;
				        		while(new File(modulePath).exists()){
				        			modulePath = projectPath + "\\" + tsfac.name() + "_" + foldcount;
				        			foldcount++;
				        		}
				        		new File(modulePath).mkdir();
								Document moduleDoc = DocumentFactory.getInstance().createDocument();
						        Element moduleRoot =  (Element) moduleDoc.addElement("module");
						        moduleRoot.addElement("project").setText(new File(projectPath).getName());
						        moduleRoot.addElement("pulled").setText("ALM");
						        moduleRoot.addElement("TestLab").setText(testLabPath + "\\" + tsfac.name());
						        moduleRoot.addElement("ALMDomain").setText(domainName);
						        moduleRoot.addElement("ALMProject").setText(projectname);
						        moduleRoot.addElement("pulledon").setText(new SimpleDateFormat("MM/dd/yyyy").format(new Date()));
						        moduleRoot.addElement("plannedstart").setText("");
						        moduleRoot.addElement("actualstart");
						        moduleRoot.addElement("plannedend").setText("");
						        moduleRoot.addElement("actualend");
						        FileOutputStream modFOS = new FileOutputStream(modulePath + "\\moduleDetails.ts");
								OutputFormat format = OutputFormat.createPrettyPrint();
						        XMLWriter modWriter = new XMLWriter(modFOS, format);
						        modWriter.write(moduleDoc);
						        modWriter.close();
					            IBaseFactory tsfc = tsfac.tsTestFactory().queryInterface(IBaseFactory.class);
					            IList testList = tsfc.newList("");
					            //IRunFactory runFac = itdc.runFactory().queryInterface(IRunFactory.class);
					            for(int eachTest = 1; eachTest<=testList.count(); eachTest++) {
					    	        ITSTest tsTest = ((Com4jObject)testList.item(eachTest)).queryInterface(ITSTest.class);
					    	        ITest actualTest = tsTest.test().queryInterface(ITest.class);
					    	        String curTCName = actualTest.name();
					    	        IBaseFactory desgFac = actualTest.designStepFactory().queryInterface(IBaseFactory.class);
					    	        IList steps = desgFac.newList("");
					    	        if(!new File(modulePath + "\\" + curTCName).exists()) {
										new File(modulePath + "\\" + curTCName).mkdir();
										new File(modulePath + "\\" + curTCName + "\\Screenshots").mkdir();
										new File(modulePath + "\\" + curTCName + "\\Reports").mkdir();
										curtestDet = modulePath + "\\" + curTCName + "\\testDetails.tc";
										IRunFactory runFac = tsTest.runFactory().queryInterface(IRunFactory.class);
										IRun lastrun = null;
										if(runFac.newList("").count() > 0 && download ) {
											lastrun = tsTest.lastRun().queryInterface(IRun.class);
										}
								        Document document = DocumentFactory.getInstance().createDocument();
								        Element root =  (Element) document.addElement("testcase");
								        root.addElement("testcasename").setText(curTCName);
								        root.addElement("testdescription").setText(((String) actualTest.field("TS_DESCRIPTION")).replace("<br />", " "));
								        root.addElement("testdata");
								        root.addElement("executiontime");
								        root.addElement("lastexecutedstep");
								        if(runFac.newList("").count() > 0 && download){
								        	root.addElement("executedby").setText(lastrun.field("RN_TESTER_NAME").toString());
								        	String executedDate = new SimpleDateFormat("MM/dd/yyyy hh:mm").format(new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy").parse(lastrun.field("RN_EXECUTION_DATE").toString()));
								        	root.addElement("executedon").setText(executedDate);
								        	root.addElement("overallstatus").setText(lastrun.status());
								        }
								        else {
								        	root.addElement("executedby");
								        	root.addElement("executedon");
								        	root.addElement("overallstatus");
								        }
								        root.addElement("report");
								        if(download) {
									        IAttachmentFactory attchFac = tsTest.attachments().queryInterface(IAttachmentFactory.class);
									        IList attachList = attchFac.newList("");
									        for(int eachAttach = 1; eachAttach<=attachList.count();eachAttach++) {
									        	IAttachment eachAttachment = ((Com4jObject)attachList.item(eachAttach)).queryInterface(IAttachment.class);
									        	IExtendedStorage storage = eachAttachment.attachmentStorage().queryInterface(IExtendedStorage.class);
									        	storage.clientPath(modulePath + "\\" + curTCName + "\\Reports");
									        	storage.load(eachAttachment.name(eachAttachment.type()), false);
									        }
								        }
								        
							        	IBaseFactory runStepFac = null;
							        	IList runStepList = null;
								        if(runFac.newList("").count() > 0 && download ) {
								        	runStepFac = lastrun.stepFactory().queryInterface(IBaseFactory.class);
								        	runStepList = runStepFac.newList("");
									        root.addElement("runid").setText(lastrun.name());
									        IAttachmentFactory attchRunFac = lastrun.attachments().queryInterface(IAttachmentFactory.class);
									        IList attachRunList = attchRunFac.newList("");
									        for(int eachAttach = 1; eachAttach<=attachRunList.count();eachAttach++) {
									        	IAttachment eachAttachment = ((Com4jObject)attachRunList.item(eachAttach)).queryInterface(IAttachment.class);
									        	IExtendedStorage storage = eachAttachment.attachmentStorage().queryInterface(IExtendedStorage.class);
									        	storage.clientPath(modulePath + "\\" + curTCName + "\\Reports");
									        	storage.load(eachAttachment.name(eachAttachment.type()), false);
									        }
								        }
								        else {
								        	root.addElement("runid").setText("");
								        }
						    	        for(int eachStep = 1; eachStep<=steps.count();eachStep++) {
						    		        IDesignStep step = ((Com4jObject)steps.item(eachStep)).queryInterface(IDesignStep.class);
									        Element tsNode = root.addElement("teststep");
											tsNode.addElement("stepdescription").setText(step.stepDescription().replace("<br />", " "));
											tsNode.addElement("expectedresult").setText(step.stepExpectedResult().replace("<br />", " "));
											if(runFac.newList("").count() > 0 && download ) {
												IStep runStep = ((Com4jObject)runStepList.item(eachStep)).queryInterface(IStep.class);
												tsNode.addElement("actualresult").setText(runStep.field("ST_ACTUAL").toString());
												tsNode.addElement("stepstatus").setText(runStep.status());
											}
											else {
												tsNode.addElement("actualresult").setText("");
												tsNode.addElement("stepstatus").setText("");
											}
											
											tsNode.addElement("screenshot").setText("");
										}
						    	        FileOutputStream fos = new FileOutputStream(curtestDet);
								        XMLWriter writer = new XMLWriter(fos, format);
								        writer.write(document);
								        writer.close();
					    	        }
				        		}
				        	}
				        }
				        dialog.dispose();
				        return;
					}
					catch(Exception ex ) {
						ex.printStackTrace();
					}
				}
			});
	        dialog.setVisible(true);
			
		}
        catch(Exception e) {
        	e.printStackTrace();
        }
	}
	
	/*TO get the column index of header in excel sheet*/
	public static int getColumnIndex(String ColumnText, int Row, XSSFSheet tcSheet) {
		int colIndex = 0;
		try {
			for (int columnIndex = 0; columnIndex<=100; columnIndex++){
		        if(tcSheet.getRow(0).getCell(columnIndex).getStringCellValue().equalsIgnoreCase(ColumnText)) {
		        	colIndex = columnIndex;
		        	break;
		        }
			}
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return colIndex;
	}
	
	public void refreshProjectALM(String projectPath, String userID, String Pass) {
		try {
			File inputFile = new File(projectPath + "\\moduleDetails.ts");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectNodes("/module").get(0);
        	Element element = (Element) node;
        	String testLabPath = element.selectSingleNode("TestLab").getText();
        	String domainName = element.selectSingleNode("ALMDomain").getText();
        	String projectname = element.selectSingleNode("ALMProject").getText();
			ITDConnection itdc= ClassFactory.createTDConnection();
	        itdc.initConnectionEx(props.getProperty("ALMURL"));
	        itdc.connectProjectEx(domainName, projectname, userID, Pass);
	        itdc.ignoreHtmlFormat(true);
	        ITestSetTreeManager sTestSetTree = itdc.testSetTreeManager().queryInterface(ITestSetTreeManager.class);
	        //String testLabPath = "Root\\Overall_PCMH_Regression\\01.27.2017_Release\\Manual_Regression\\Cognizant_Scope\\SRH and Error Handling";
	        String testLabName = testLabPath.substring(testLabPath.lastIndexOf("\\") + 1);
	        String testLabFolder = testLabPath.substring(0, testLabPath.lastIndexOf("\\"));
	        ITestSetFolder sTestFolder = sTestSetTree.nodeByPath(testLabFolder).queryInterface(ITestSetFolder.class);
	        IList foundList = sTestFolder.findTestSets(null, true, null).queryInterface(IList.class);
			String curtestDet = "";
			String modulePath = projectPath;
			for(int eachTLcnt = 1; eachTLcnt<=foundList.count(); eachTLcnt++) {
	        	ITestSet tsfac = ((Com4jObject) foundList.item(eachTLcnt)).queryInterface(ITestSet.class);
	        	if(tsfac.name().equals(testLabName)) {
		            IBaseFactory tsfc = tsfac.tsTestFactory().queryInterface(IBaseFactory.class);
		            IList testList = tsfc.newList("");
		            //IRunFactory runFac = itdc.runFactory().queryInterface(IRunFactory.class);
		            for(int eachTest = 1; eachTest<=testList.count(); eachTest++) {
		    	        ITSTest tsTest = ((Com4jObject)testList.item(eachTest)).queryInterface(ITSTest.class);
		    	        ITest actualTest = tsTest.test().queryInterface(ITest.class);
		    	        String curTCName = actualTest.name();
		    	        IBaseFactory desgFac = actualTest.designStepFactory().queryInterface(IBaseFactory.class);
		    	        IList steps = desgFac.newList("");
		    	        if(!new File(modulePath + "\\" + curTCName).exists()) {
							new File(modulePath + "\\" + curTCName).mkdir();
							new File(modulePath + "\\" + curTCName + "\\Screenshots").mkdir();
							new File(modulePath + "\\" + curTCName + "\\Reports").mkdir();
							curtestDet = modulePath + "\\" + curTCName + "\\testDetails.tc";
							IRunFactory runFac = tsTest.runFactory().queryInterface(IRunFactory.class);
							IRun lastrun = null;
							if(runFac.newList("").count() > 0) {
								lastrun = tsTest.lastRun().queryInterface(IRun.class);
							}
					        Document tcdocument = DocumentFactory.getInstance().createDocument();
					        Element root =  (Element) tcdocument.addElement("testcase");
					        root.addElement("testcasename").setText(curTCName);
					        root.addElement("testdescription").setText(((String) actualTest.field("TS_DESCRIPTION")).replace("<br />", " "));
					        root.addElement("testdata");
					        root.addElement("executiontime");
					        root.addElement("lastexecutedstep");
					        if(runFac.newList("").count() > 0){
					        	root.addElement("executedby").setText(lastrun.field("RN_TESTER_NAME").toString());
					        	String executedDate = new SimpleDateFormat("MM/dd/yyyy hh:mm").format(new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy").parse(lastrun.field("RN_EXECUTION_DATE").toString()));
					        	root.addElement("executedon").setText(executedDate);
					        	root.addElement("overallstatus").setText(lastrun.status());
					        }
					        else {
					        	root.addElement("executedby");
					        	root.addElement("executedon");
					        	root.addElement("overallstatus");
					        }
					        root.addElement("report");
					        IAttachmentFactory attchFac = tsTest.attachments().queryInterface(IAttachmentFactory.class);
					        IList attachList = attchFac.newList("");
					        for(int eachAttach = 1; eachAttach<=attachList.count();eachAttach++) {
					        	IAttachment eachAttachment = ((Com4jObject)attachList.item(eachAttach)).queryInterface(IAttachment.class);
					        	IExtendedStorage storage = eachAttachment.attachmentStorage().queryInterface(IExtendedStorage.class);
					        	storage.clientPath(modulePath + "\\" + curTCName + "\\Reports");
					        	storage.load(eachAttachment.name(eachAttachment.type()), false);
					        }
					        
				        	IBaseFactory runStepFac = null;
				        	IList runStepList = null;
					        if(runFac.newList("").count() > 0) {
					        	runStepFac = lastrun.stepFactory().queryInterface(IBaseFactory.class);
					        	runStepList = runStepFac.newList("");
						        root.addElement("runid").setText(lastrun.name());
						        IAttachmentFactory attchRunFac = lastrun.attachments().queryInterface(IAttachmentFactory.class);
						        IList attachRunList = attchRunFac.newList("");
						        for(int eachAttach = 1; eachAttach<=attachRunList.count();eachAttach++) {
						        	IAttachment eachAttachment = ((Com4jObject)attachRunList.item(eachAttach)).queryInterface(IAttachment.class);
						        	IExtendedStorage storage = eachAttachment.attachmentStorage().queryInterface(IExtendedStorage.class);
						        	storage.clientPath(modulePath + "\\" + curTCName + "\\Reports");
						        	storage.load(eachAttachment.name(eachAttachment.type()), false);
						        }
					        }
					        else {
					        	root.addElement("runid").setText("");
					        }
			    	        for(int eachStep = 1; eachStep<=steps.count();eachStep++) {
			    		        IDesignStep step = ((Com4jObject)steps.item(eachStep)).queryInterface(IDesignStep.class);
						        Element tsNode = root.addElement("teststep");
								tsNode.addElement("stepdescription").setText(step.stepDescription().replace("<br />", " "));
								tsNode.addElement("expectedresult").setText(step.stepExpectedResult().replace("<br />", " "));
								if(runFac.newList("").count() > 0) {
									IStep runStep = ((Com4jObject)runStepList.item(eachStep)).queryInterface(IStep.class);
									tsNode.addElement("actualresult").setText(runStep.field("ST_ACTUAL").toString());
									tsNode.addElement("stepstatus").setText(runStep.status());
								}
								else {
									tsNode.addElement("actualresult").setText("");
									tsNode.addElement("stepstatus").setText("");
								}
								
								tsNode.addElement("screenshot").setText("");
							}
			    	        FileOutputStream fos = new FileOutputStream(curtestDet);
			    	        OutputFormat format = OutputFormat.createPrettyPrint();
					        XMLWriter writer = new XMLWriter(fos, format);
					        writer.write(tcdocument);
					        writer.close();
		    	        }
		    	        else {
		    	        	curtestDet = modulePath + "\\" + curTCName + "\\testDetails.tc";
		    	        	Document doc = reader.read(curtestDet);
		    	        	Element root =  (Element) doc.selectSingleNode("/testcase");
		    	        	root.selectSingleNode("testdescription").setText(((String) actualTest.field("TS_DESCRIPTION")).replace("<br />", " "));
		    	        	IRunFactory runFac = tsTest.runFactory().queryInterface(IRunFactory.class);
					        if(runFac.newList("").count() > 0) {
					        	IRun lastrun = tsTest.lastRun().queryInterface(IRun.class);
						        root.selectSingleNode("runid").setText(lastrun.name());
					        }
					        else {
					        	root.selectSingleNode("runid").setText("");
					        }
					        for(int eachStep = 1; eachStep<=steps.count();eachStep++) {
			    		        IDesignStep step = ((Com4jObject)steps.item(eachStep)).queryInterface(IDesignStep.class);
						        List<Node> tsNodes = doc.selectNodes("/testcase/teststep");
						        if(eachStep <= tsNodes.size()) {
						        	Element tsNode = (Element) tsNodes.get(eachStep - 1);
									tsNode.selectSingleNode("stepdescription").setText(step.stepDescription().toString());
									tsNode.selectSingleNode("expectedresult").setText(step.stepExpectedResult().toString());
						        }
						        else {
						        	Element tsNode = root.addElement("teststep");
									tsNode.addElement("stepdescription").setText(step.stepDescription().replace("<br />", " "));
									tsNode.addElement("expectedresult").setText(step.stepExpectedResult().replace("<br />", " "));
									tsNode.addElement("actualresult").setText("");
									tsNode.addElement("stepstatus").setText("");
									tsNode.addElement("screenshot").setText("");
						        }
							}
					        FileOutputStream fos = new FileOutputStream(curtestDet);
			    	        OutputFormat format = OutputFormat.createPrettyPrint();
					        XMLWriter writer = new XMLWriter(fos, format);
					        writer.write(doc);
					        writer.close();
		    	        }
		    	        
		            }
	            }
				
			}
		}
        catch(Exception e) {
        	e.printStackTrace();
        }
	}

}
