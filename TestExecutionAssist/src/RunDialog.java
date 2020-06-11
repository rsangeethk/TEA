package src;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.SwingWorker;
import javax.swing.Timer;
import javax.imageio.ImageIO;
import javax.imageio.ImageWriter;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import java.awt.Font;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.GraphicsDevice;
import java.awt.GraphicsEnvironment;
import java.awt.Image;
import java.awt.Point;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.Dialog.ModalExclusionType;
import java.awt.Dialog.ModalityType;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Window.Type;
import java.awt.AWTException;
import java.awt.AlphaComposite;
import java.awt.BorderLayout;
import java.awt.Desktop;
import java.awt.event.ActionListener;
import java.awt.event.ComponentAdapter;
import java.awt.event.ComponentEvent;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.ExecutionException;
import java.awt.event.ActionEvent;
import java.awt.Color;
import java.awt.SystemColor;

import javax.swing.border.EtchedBorder;
import javax.swing.border.CompoundBorder;
import javax.swing.border.SoftBevelBorder;
import javax.swing.border.BevelBorder;
import javax.swing.border.Border;
import javax.swing.border.TitledBorder;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

import javax.swing.UIManager;
import javax.swing.border.LineBorder;
import javax.swing.JSeparator;
import javax.swing.SwingConstants;
import javax.swing.JTextArea;
import javax.swing.JTextPane;
import javax.swing.KeyStroke;
import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.BorderFactory;
import javax.swing.ImageIcon;

public class RunDialog {
/*	private static JTextField textField;
	private static JTextField textField_1;
	private static Timer timer;*/
	JButton btnDetails = new JButton("...");
	JTextArea txtpnDescpane = new JTextArea();
	JTextArea txtpnExpected = new JTextArea();
	JTextArea txtrActual = new JTextArea();
	JButton btnP = new JButton("");
	JButton btnNext = new JButton("");
	JLabel lblName = new JLabel();
	JLabel labelStep = new JLabel("");
	JLabel lblStatus = new JLabel();
	String tcPath = "";
	private static int curStepNum = 0;
	private static int lastStepNum = 0;
	Border emptyBorder = BorderFactory.createEmptyBorder();
	Properties props = new Properties();
	public RunDialog(String testPath){
		
		try {
			props.load(new FileInputStream(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg"));
		}
		catch(Exception e) {
			
		}
		Date startime = new Date();
		tcPath = testPath; 
		lastStepNum = getlastStepNum();
		JDialog exeHandle = new JDialog();
		exeHandle.setModal(true);
        GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
        Rectangle rect = ge.getMaximumWindowBounds();
		//final JFrame exeHandle = new JFrame();
		exeHandle.setBackground(Color.LIGHT_GRAY);
		exeHandle.setUndecorated(true);
		exeHandle.setResizable(false);   
		exeHandle.setBounds(0, 0, 315, 430);
		exeHandle.getContentPane().setLayout(null);
		exeHandle.setAlwaysOnTop(true);
		exeHandle.setModalityType(ModalityType.APPLICATION_MODAL);
		
		JPanel executionPanel = new JPanel();
		executionPanel.setBorder(new LineBorder(new Color(160, 160, 160), 2, true));
		executionPanel.setBackground(SystemColor.menu);
		executionPanel.setBounds(0, 0, 315, 76);
		exeHandle.getContentPane().add(executionPanel);
		executionPanel.setLayout(null);
		
		JButton btnStop = new JButton("");
		btnStop.setIcon(new ImageIcon(RunDialog.class.getResource("/res/Stop30x30.png")));
		btnStop.setBackground(SystemColor.menu);
		btnStop.setBorder(emptyBorder);
		btnStop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(getTestDetails("executiontime").equals("")) {
					setTestDetails("executiontime",  Long.toString((long)(new Date().getTime() - startime.getTime())/(60 * 1000) % 60));
				}
				else{
					setTestDetails("executiontime",  Long.toString(Long.parseLong(getTestDetails("executiontime")) + (long)(new Date().getTime() - startime.getTime())/(60 * 1000) % 60));
				}
				setTestDetails("lastexecutedstep",  Integer.toString(curStepNum + 1));
				if(getTestDetails("executedon").equals("")) {
					setTestDetails("executedon", new SimpleDateFormat("MM/dd/yyyy HH:mm").format(new Date()));
				}
				else{
					setTestDetails("executedon", getTestDetails("executedon") + ";" + new SimpleDateFormat("MM/dd/yyyy HH:mm").format(new Date()));
				}
				if(getTestDetails("executedby").equals("")) {
					setTestDetails("executedby", System.getProperty("user.name"));
				}
				else{
					setTestDetails("executedby", getTestDetails("executedby") + ";" + System.getProperty("user.name"));
				}
				setTestDetails("overallstatus", getOverAllStatus());
				try {
					if(props.getProperty("AutoReport").equalsIgnoreCase("Yes")) {
						generateReportWord();
					}
				}
				catch(Throwable exp) {
					JOptionPane.showMessageDialog(null, "Unable to generate word report.");
				}
				exeHandle.dispose();
			}
		});
		
		JLabel lblPass = new JLabel("PASS");
		lblPass.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblPass.setHorizontalAlignment(SwingConstants.CENTER);
		lblPass.setForeground(new Color(0, 128, 0));
		lblPass.setBounds(65, 30, 46, 14);
		executionPanel.add(lblPass);
		
		JLabel lblFail = new JLabel("FAIL");
		lblFail.setHorizontalAlignment(SwingConstants.CENTER);
		lblFail.setForeground(new Color(255, 0, 0));
		lblFail.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblFail.setBounds(136, 30, 46, 14);
		executionPanel.add(lblFail);
		btnStop.setBounds(203, 23, 30, 30);
		executionPanel.add(btnStop);
		
		JButton btnDetails = new JButton("");
		btnDetails.setIcon(new ImageIcon(RunDialog.class.getResource("/res/arrowUp30x30.png")));
		btnDetails.setBackground(SystemColor.menu);
		btnDetails.setBounds(10, 40, 30, 30);
		btnDetails.setBorder(emptyBorder);
		executionPanel.add(btnDetails);
		
		  final Action alPass = new AbstractAction("") {
	          public void actionPerformed(ActionEvent ae) {
	        	  Robot robot;
				try {
					robot = new Robot();
					Rectangle screenSize = new Rectangle(
		                  Toolkit.getDefaultToolkit().getScreenSize());
		              Image image = robot.createScreenCapture(screenSize);
		              // construct the buffered image
		              BufferedImage bImage = new BufferedImage(image.getWidth(null), image.getHeight(null), BufferedImage.TYPE_INT_RGB);
	
		              //obtain it's graphics
		              Graphics2D bImageGraphics = bImage.createGraphics();
	
		              //draw the Image (image) into the BufferedImage (bImage)
		              bImageGraphics.drawImage(image, null, null);
	
		              // cast it to rendered image
		              RenderedImage rImage      = (RenderedImage)bImage;
		              String scrshotName = "Step_" + Integer.toString(curStepNum + 1)  + "_" + new SimpleDateFormat("HH_mm_ss").format(new Date()) +".png";
		              ImageIO.write(rImage, "png", new File(tcPath + "\\Screenshots\\" + scrshotName));
		              setStepDetails(curStepNum, "Passed", scrshotName, txtrActual.getText());
		              if(props.getProperty("AutoNextStep").equalsIgnoreCase("Yes")) {
		                if(curStepNum<=lastStepNum) {
							curStepNum++;
							if(curStepNum>=lastStepNum) {
								btnNext.setEnabled(false);
								btnP.setEnabled(true);
							}
							else{
								getStepDetails(curStepNum);
								btnP.setEnabled(true);
								btnNext.setEnabled(true);
							}
							labelStep.setText(Integer.toString(curStepNum + 1) + "/" + Integer.toString(lastStepNum+1));
		                }
		              }
		              exeHandle.setVisible(true);
		              
				} catch (Exception e) {
					e.printStackTrace();
				}
	              
	          }
	      };
	      
	      final Action alFail = new AbstractAction("") {
	          public void actionPerformed(ActionEvent ae) {
	        	  Robot robot;
				try {
					robot = new Robot();
					Rectangle screenSize = new Rectangle(
		                  Toolkit.getDefaultToolkit().getScreenSize());
		              Image image = robot.createScreenCapture(screenSize);
		              // construct the buffered image
		              BufferedImage bImage = new BufferedImage(image.getWidth(null), image.getHeight(null), BufferedImage.TYPE_INT_RGB);
	
		              //obtain it's graphics
		              Graphics2D bImageGraphics = bImage.createGraphics();
	
		              //draw the Image (image) into the BufferedImage (bImage)
		              bImageGraphics.drawImage(image, null, null);
	
		              // cast it to rendered image
		              RenderedImage rImage      = (RenderedImage)bImage;
		              String scrshotName = "Step_" + Integer.toString(curStepNum + 1)  + "_" + new SimpleDateFormat("HH_mm_ss").format(new Date()) +".png";
		              ImageIO.write(rImage, "png", new File(tcPath + "\\Screenshots\\" + scrshotName));
		              setStepDetails(curStepNum, "Failed", scrshotName, txtrActual.getText());
		              if(props.getProperty("AutoNextStep").equalsIgnoreCase("Yes")) {
		                if(curStepNum<=lastStepNum) {
							curStepNum++;
							if(curStepNum>=lastStepNum) {
								btnNext.setEnabled(false);
								btnP.setEnabled(true);
							}
							else{
								getStepDetails(curStepNum);
								btnP.setEnabled(true);
								btnNext.setEnabled(true);
							}
							labelStep.setText(Integer.toString(curStepNum + 1) + "/" + Integer.toString(lastStepNum+1));
		                }
		              }
		              exeHandle.setVisible(true);
		              
				} catch (Exception e) {
					e.printStackTrace();
				}
	          }
	      };
		Action pAct = new AbstractAction("") {
			public void actionPerformed(ActionEvent e) {
				if(props.getProperty("PassScreenshot").equalsIgnoreCase("Yes")) {
					exeHandle.setVisible(false);
	                Timer timer = new Timer(100, alPass);
	                timer.setRepeats(false);
	                timer.start();
				}
				else{
					setStepDetails(curStepNum, "Passed", "", txtrActual.getText());
					if(props.getProperty("AutoNextStep").equalsIgnoreCase("Yes")) {
		                if(curStepNum<=lastStepNum) {
							curStepNum++;
							if(curStepNum>=lastStepNum) {
								btnNext.setEnabled(false);
								btnP.setEnabled(true);
							}
							else{
								getStepDetails(curStepNum);
								btnP.setEnabled(true);
								btnNext.setEnabled(true);
							}
							labelStep.setText(Integer.toString(curStepNum + 1) + "/" + Integer.toString(lastStepNum+1));
		                }
					}
				}
			}
		};
		Action fAct = new AbstractAction("") {
			public void actionPerformed(ActionEvent e) {
				if(props.getProperty("FailScreenshot").equalsIgnoreCase("Yes")) {
					exeHandle.setVisible(false);
					Timer timer = new Timer(100, alFail);
	                timer.setRepeats(false);
	                timer.start();
				}
				else {
					setStepDetails(curStepNum, "Failed", "", txtrActual.getText());
					if(props.getProperty("AutoNextStep").equalsIgnoreCase("Yes")) {
		                if(curStepNum<=lastStepNum) {
							curStepNum++;
							if(curStepNum>=lastStepNum) {
								btnNext.setEnabled(false);
								btnP.setEnabled(true);
							}
							else{
								getStepDetails(curStepNum);
								btnP.setEnabled(true);
								btnNext.setEnabled(true);
							}
							labelStep.setText(Integer.toString(curStepNum + 1) + "/" + Integer.toString(lastStepNum+1));
		                }
					}
				}
                
			}
		};
		
		JButton btnPass = new JButton("PASS");
		btnPass.setBackground(new Color(255, 255, 255));
		btnPass.setForeground(new Color(0, 128, 0));
		btnPass.setFont(new Font("Dialog", Font.BOLD, 11));
		btnPass.setText("PASS");
		btnPass.setAction(pAct);
		btnPass.getInputMap().put(KeyStroke.getKeyStroke(KeyEvent.VK_P,KeyEvent.CTRL_DOWN_MASK | KeyEvent.ALT_DOWN_MASK), "passAct");
		btnPass.getActionMap().put("passAct", pAct);
		btnPass.setBounds(57, 12, 64, 52);
		executionPanel.add(btnPass);
		
		JButton btnFail = new JButton("FAIL");
		btnFail.setAction(fAct);
		btnFail.getInputMap().put(KeyStroke.getKeyStroke(KeyEvent.VK_P,KeyEvent.CTRL_DOWN_MASK | KeyEvent.ALT_DOWN_MASK), "failAct");
		btnFail.getActionMap().put("failAct", fAct);
		btnFail.setForeground(new Color(255, 0, 0));
		btnFail.setFont(new Font("Dialog", Font.BOLD, 11));
		btnFail.setBackground(new Color(255, 255, 255));
		btnFail.setBounds(127, 12, 64, 52);
		executionPanel.add(btnFail);
		
		JSeparator separator = new JSeparator();
		separator.setOrientation(SwingConstants.VERTICAL);
		separator.setBackground(UIManager.getColor("ProgressBar.shadow"));
		separator.setForeground(UIManager.getColor("ProgressBar.shadow"));
		separator.setBounds(50, 0, 2, 76);
		executionPanel.add(separator);
		
		JButton btnCollapse = new JButton("");
		btnCollapse.setIcon(new ImageIcon(RunDialog.class.getResource("/res/arrowRight30x30.png")));
		btnCollapse.setBackground(SystemColor.menu);
		btnCollapse.setBorder(emptyBorder);
		btnCollapse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(exeHandle.getSize().getWidth()>150) {
					exeHandle.setBounds(0, 0, 64, 76);
					btnDetails.setEnabled(false);
					btnCollapse.setIcon(new ImageIcon(RunDialog.class.getResource("/res/arrowLeft30x30.png")));
					btnDetails.setIcon(new ImageIcon(RunDialog.class.getResource("/res/arrowDown30x30.png")));
				}
				else {
					exeHandle.setBounds(0, 0, 315, 76);
					btnDetails.setEnabled(true);
					btnCollapse.setIcon(new ImageIcon(RunDialog.class.getResource("/res/arrowRight30x30.png")));
				}
				Dimension windowSize = exeHandle.getSize();
		        int x = (int) rect.getWidth() - (int) windowSize.getWidth();
		        int y = (int) rect.getHeight()/2 - (int) windowSize.getHeight()/2;
		        exeHandle.setLocation(x, y);
			}
		});
		btnCollapse.setBounds(10, 7, 30, 30);
		executionPanel.add(btnCollapse);
		
		
		btnDetails.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(exeHandle.getSize().getHeight()>150) {
					exeHandle.setBounds(0, 0, exeHandle.getWidth(), 76);
					btnDetails.setIcon(new ImageIcon(RunDialog.class.getResource("/res/arrowDown30x30.png")));
				}
				else {
					exeHandle.setBounds(0, 0, exeHandle.getWidth(), 430);
					btnDetails.setIcon(new ImageIcon(RunDialog.class.getResource("/res/arrowUp30x30.png")));
				}
				Dimension windowSize = exeHandle.getSize();
		        int x = (int) rect.getWidth() - (int) windowSize.getWidth();
		        int y = (int) rect.getHeight()/2 - (int) windowSize.getHeight()/2;
		        exeHandle.setLocation(x, y);
			}
		});
		/*btnDetails.setBounds(7, 45, 25, 25);
		executionPanel.add(btnDetails);*/
		
		btnP.setIcon(new ImageIcon(RunDialog.class.getResource("/res/Prev30x30.png")));
		btnP.setBackground(SystemColor.menu);
		btnP.setBounds(243, 11, 30, 30);
		btnP.setBorder(emptyBorder);
		executionPanel.add(btnP);
		
		
		btnNext.setIcon(new ImageIcon(RunDialog.class.getResource("/res/Next30x30.png")));
		btnNext.setBackground(SystemColor.menu);
		btnNext.setBorder(emptyBorder);
		btnNext.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				//setStepDetails(curStepNum);
				curStepNum++;
				if(curStepNum>=lastStepNum - 1) {
					btnNext.setEnabled(false);
					btnP.setEnabled(true);
					getStepDetails(curStepNum);
				}
				else{
					getStepDetails(curStepNum);
					btnP.setEnabled(true);
					btnNext.setEnabled(true);
				}
				labelStep.setText(Integer.toString(curStepNum + 1) + "/" + Integer.toString(lastStepNum));
			}
		});
		btnNext.setBounds(273, 11, 30, 30);
		executionPanel.add(btnNext);
		
		JLabel lblStep = new JLabel("Step :");
		lblStep.setBounds(251, 50, 36, 14);
		executionPanel.add(lblStep);
		
		labelStep.setBounds(285, 50, 24, 14);
		labelStep.setText(Integer.toString(curStepNum + 1) + "/" + Integer.toString(lastStepNum));
		executionPanel.add(labelStep);
		
		
		
		btnP.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				//setStepDetails(curStepNum);
				curStepNum--;
				if(curStepNum<=0) {
					btnP.setEnabled(false);
					btnNext.setEnabled(true);
					getStepDetails(curStepNum);
				}
				else{
					getStepDetails(curStepNum);
					btnP.setEnabled(true);
					btnNext.setEnabled(true);
				}
				labelStep.setText(Integer.toString(curStepNum + 1) + "/" + Integer.toString(lastStepNum));
			}
		});
		
		
		JPanel panelDetails = new JPanel();
		panelDetails.setBackground(SystemColor.menu);
		panelDetails.setBorder(new LineBorder(new Color(160, 160, 160), 2, true));
		panelDetails.setBounds(0, 76, 315, 354);
		exeHandle.getContentPane().add(panelDetails);
		panelDetails.setLayout(null);
		
		lblName.setBounds(10, 11, 260, 14);
		panelDetails.add(lblName);
		
		
		lblStatus.setBounds(260, 11, 40, 15);
		panelDetails.add(lblStatus);
		
		JLabel lblDdescription = new JLabel("Description:");
		lblDdescription.setBounds(10, 36, 78, 14);
		panelDetails.add(lblDdescription);
		
		
		txtpnDescpane.setLineWrap(true);
		txtpnDescpane.setWrapStyleWord(true);
		txtpnDescpane.setEditable(false);
		txtpnDescpane.setBounds(10, 52, 298, 76);
		JScrollPane DescScroll = new JScrollPane(txtpnDescpane);
		DescScroll.setBounds(10, 52, 298, 76);
		panelDetails.add(DescScroll);
		
		JLabel lblExpected = new JLabel("Expected:");
		lblExpected.setBounds(10, 140, 64, 14);
		panelDetails.add(lblExpected);
		txtpnExpected.setLineWrap(true);
		txtpnExpected.setWrapStyleWord(true);
		
		txtpnExpected.setEditable(false);
		txtpnExpected.setText("expected");
		txtpnExpected.setBounds(10, 157, 298, 76);
		JScrollPane expScroll = new JScrollPane(txtpnExpected);
		expScroll.setBounds(10, 157, 298, 76);
		panelDetails.add(expScroll);
		
		JLabel lblActual = new JLabel("Actual:");
		lblActual.setBounds(10, 243, 64, 14);
		panelDetails.add(lblActual);
		
		txtrActual.setLineWrap(true);
		txtrActual.setText("actual");
		txtrActual.setBounds(10, 260, 298, 82);
		JScrollPane actScroll = new JScrollPane(txtrActual);
		actScroll.setBounds(10, 260, 298, 82);
		panelDetails.add(actScroll);
		
		if(curStepNum<=0) {
			btnP.setEnabled(false);
			btnNext.setEnabled(true);
		}
		else if(curStepNum>=lastStepNum) {
			btnNext.setEnabled(false);
			btnP.setEnabled(true);
		}
		else{
			btnP.setEnabled(true);
			btnNext.setEnabled(true);
		}
		getStepDetails(curStepNum);
		
		Dimension windowSize = exeHandle.getSize();
        int x = (int) rect.getWidth() - (int) windowSize.getWidth();
        int y = (int) rect.getHeight()/2 - (int) windowSize.getHeight()/2;
        exeHandle.setLocation(x, y);
		exeHandle.setVisible(true);
	}
	
	public void getStepDetails(int stepNum) {
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        String tcName = ((Element) document.selectNodes("/testcase/testcasename").get(0)).getText();
	        Node node = (Node) document.selectNodes("/testcase/teststep").get(stepNum);
        	Element element = (Element) node;
        	String tcDesc = element.selectSingleNode("stepdescription").getText();
        	String tcStatus = element.selectSingleNode("stepstatus").getText();
        	String tcExp = element.selectSingleNode("expectedresult").getText();
        	String tcActual = element.selectSingleNode("actualresult").getText();
        	lblName.setText(tcName);
        	if(tcStatus.equalsIgnoreCase("Passed")) {
        		lblStatus.setText(tcStatus);
        		lblStatus.setForeground(new Color(0, 128, 0));
        	}
        	else if(tcStatus.equalsIgnoreCase("Failed")) {
        		lblStatus.setText(tcStatus);
        		lblStatus.setForeground(new Color(255, 0, 0));
        	}
        	txtpnDescpane.setText(tcDesc);
        	txtpnExpected.setText(tcExp);
        	txtrActual.setText(tcActual);
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public void setStepDetails(int stepNum) {
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectNodes("/testcase/teststep").get(stepNum);
        	Element element = (Element) node;
        	element.selectSingleNode("actualresult").setText(txtrActual.getText());
        	FileOutputStream fos = new FileOutputStream(inputFile);
			OutputFormat format = OutputFormat.createPrettyPrint();
	        XMLWriter writer = new XMLWriter(fos, format);
	        writer.write(document);
	        writer.close();
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public void setStepDetails(int stepNum, String Status, String Screenshotname, String actuals) {
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectNodes("/testcase/teststep").get(stepNum);
        	Element element = (Element) node;
        	element.selectSingleNode("actualresult").setText(actuals);
        	element.selectSingleNode("stepstatus").setText(Status);
        	if(element.selectSingleNode("screenshot").getText().equals(""))  {
        		element.selectSingleNode("screenshot").setText(Screenshotname);
        	}
        	else {
        		element.addElement("screenshot").setText(Screenshotname);
        	}
        	FileOutputStream fos = new FileOutputStream(inputFile);
			OutputFormat format = OutputFormat.createPrettyPrint();
	        XMLWriter writer = new XMLWriter(fos, format);
	        writer.write(document);
	        writer.close();
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public int getlastStepNum() {
		int stepNum = 0;
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        stepNum = document.selectNodes("/testcase/teststep").size();
		}
		catch(Exception e) {
			
		}
		return stepNum-1;
	}
	public int getCurStepNum() {
		int stepNum = 0;
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        String lastExecuted = ((Element)document.selectNodes("/testcase/lastexecutedstep").get(0)).getText();
	        if(!lastExecuted.trim().equals("")) {
	        	stepNum = Integer.parseInt(lastExecuted);	        
	        }
		}
		catch(Exception e) {
			
		}
		return stepNum;
	}
	
	private void setTestDetails(String tagName, String value) {
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectNodes("/testcase").get(0);
        	Element element = (Element) node;
        	element.selectSingleNode(tagName).setText(value.trim());
        	FileOutputStream fos = new FileOutputStream(inputFile);
			OutputFormat format = OutputFormat.createPrettyPrint();
	        XMLWriter writer = new XMLWriter(fos, format);
	        writer.write(document);
	        writer.close();
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	private String getTestDetails(String tagName) {
		String value = "";
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectNodes("/testcase").get(0);
        	Element element = (Element) node;
        	value = element.selectSingleNode(tagName).getText();
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return value;
	}
	
	private String getOverAllStatus() {
		String status = "";
		try {
			String value = "";
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        List<Node> nodes = document.selectNodes("/testcase/teststep");
	        for(Node  eachNode : nodes) {
	        	Element element = (Element) eachNode;
	        	if(element.selectSingleNode("stepstatus").getText().equals("")) {
	        		value = value + ";null";
	        	}
	        	else {
	        		value = value + ";" + element.selectSingleNode("stepstatus").getText();
	        	}
	        }
	        value = value.substring(1);
	        if((value.contains("Passed") || value.contains("Failed")) && value.contains("null")) {
	        	status = "Not Completed";
	        }
	        else if(!value.contains("Passed") && !value.contains("Failed")) {
	        	status = "";
	        }
	        else if(value.contains("Passed") && !value.contains("null") && !value.contains("Failed")) {
	        	status = "Passed";
	        }
	        else if(value.contains("Failed")) {
	        	status = "Failed";
	        }
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return status;
	}
	
	public void generateReportWord() throws Throwable{
		  CustomXWPFDocument document= new CustomXWPFDocument();
		  XWPFRun header = document.createParagraph().createRun();
		  File inputFile = new File(tcPath + "\\testDetails.tc");
	      SAXReader reader = new SAXReader();
	      Document xmldocument = reader.read(inputFile);
	      Node tcnode = (Node) xmldocument.selectNodes("/testcase").get(0);
	      Element tcelement = (Element) tcnode;
	      header.setBold(true);
	      header.setFontSize(20);
	      header.setTextPosition(0);
	      header.getParagraph().setAlignment(ParagraphAlignment.CENTER);
	      String tcName = tcelement.selectSingleNode("testcasename").getText();
		  header.setText(tcelement.selectSingleNode("testcasename").getText());
		  
		  XWPFParagraph desc = document.createParagraph();
		  XWPFRun descRun = desc.createRun();
		  descRun.setText("Description: ");
		  descRun.setBold(true);
		  XWPFRun descCont = desc.createRun();
		  descCont.setText(tcelement.selectSingleNode("testdescription").getText());
		  descCont.addBreak();
		  
		  XWPFParagraph status = document.createParagraph();
		  XWPFRun statusRun = desc.createRun();
		  statusRun.setText("Status: ");
		  statusRun.setBold(true);
		  XWPFRun statusCont = desc.createRun();
		  statusCont.setText(tcelement.selectSingleNode("overallstatus").getText());
		  List<Node> stepnodes = xmldocument.selectNodes("/testcase/teststep");
		  int StepCount = 1;
		  for(Node eachnode : stepnodes){
			  Element element = (Element) eachnode;
			  XWPFParagraph exp = document.createParagraph();
			  XWPFRun stepNum = exp.createRun();
			  if(StepCount!=1)stepNum.addBreak();
			  stepNum.setText("Step: " + StepCount);
			  stepNum.setBold(true);
			  stepNum.addBreak();
			  
			  if(props.getProperty("").equalsIgnoreCase("Yes")) {
				  XWPFRun desRun = exp.createRun();
				  desRun.setText("Expected Result: ");
				  desRun.setBold(true);
				  XWPFRun desCont = exp.createRun();
				  desCont.setText(element.selectSingleNode("expectedresult").getText());
				  desCont.addBreak();
		  	  }
			  
			  if(props.getProperty("").equalsIgnoreCase("Yes")) {
				  XWPFRun expRun = exp.createRun();
				  expRun.setText("Expected Result: ");
				  expRun.setBold(true);
				  XWPFRun expCont = exp.createRun();
				  expCont.setText(element.selectSingleNode("expectedresult").getText());
				  expCont.addBreak();
			  }
			  
			  if(props.getProperty("").equalsIgnoreCase("Yes")) {
				  XWPFRun actRun = exp.createRun();
				  actRun.setText("Actual Result: ");
				  actRun.setBold(true);
				  XWPFRun actCont = exp.createRun();
				  actCont.setText(element.selectSingleNode("actualresult").getText());
				  actCont.addBreak();
			  }
			  
			  
			  XWPFRun statRun = exp.createRun();
			  statRun.setText("Step Status: ");
			  statRun.setBold(true);
			  XWPFRun statCont = exp.createRun();
			  statCont.setText(element.selectSingleNode("stepstatus").getText());
			  statCont.addBreak();
			  
			  if((props.getProperty("Screenshots").equalsIgnoreCase("Yes")) ||
					  (props.getProperty("Screenshots").equalsIgnoreCase("Only Passed Steps") && element.selectSingleNode("stepstatus").getText().equalsIgnoreCase("Passed")) || 
					  	(props.getProperty("Screenshots").equalsIgnoreCase("Only Failed Steps") && element.selectSingleNode("stepstatus").getText().equalsIgnoreCase("Failed")))  {
				  XWPFRun scrnRun = exp.createRun();
				  scrnRun.setText("Screenshots: ");
				  scrnRun.setBold(true);
				  List<Node> scrnNodes = element.selectNodes("screenshot");
				  for(Node EachScrn : scrnNodes) {
					  File scrnPath = new File(tcPath + "\\Screenshots\\" + EachScrn.getText());
					  if(scrnPath.exists()) {
						  String blipId = exp.getDocument().addPictureData(new FileInputStream(scrnPath), XWPFDocument.PICTURE_TYPE_PNG);
						  document.createPicture(blipId, document.getNextPicNameNumber(XWPFDocument.PICTURE_TYPE_PNG), 640, 360);
					  }
				  }
			  }
			  StepCount++;
		  }
	      //Write the Document in file system
		  File wordDoc= new File(tcPath + "\\Reports\\"  + tcName + "_" + new SimpleDateFormat("hh_mm_ss").format(new Date().toString()) + ".docx");
	      FileOutputStream out = new FileOutputStream(wordDoc);
	      document.write(out);
	      out.close();
	      setTestDetails("report", wordDoc.getName());
	}
}
