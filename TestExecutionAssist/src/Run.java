package src;

import java.awt.Color;
import java.awt.Dialog.ModalityType;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.GraphicsEnvironment;
import java.awt.Image;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.SystemColor;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import javax.imageio.ImageIO;
import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.BorderFactory;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSeparator;
import javax.swing.JTextArea;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingConstants;
import javax.swing.Timer;
import javax.swing.UIManager;
import javax.swing.border.Border;
import javax.swing.border.LineBorder;

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

public class Run {
/*	private static JTextField textField;
	private static JTextField textField_1;
	private static Timer timer;*/
	JButton btnDetails = new JButton("...");
	JTextArea txtpnDescpane = new JTextArea();
	JTextArea txtpnExpected = new JTextArea();
	JButton btnP = new JButton("");
	JButton btnNext = new JButton("");
	JLabel lblName = new JLabel();
	JLabel labelStep = new JLabel("");
	JTextArea scrnComments = new JTextArea();
	JScrollPane scrollPane = new JScrollPane(scrnComments);
	String tcPath = "";
	String scrshotName = "";
	private static int curStepNum = 0;
	private static int curScreenNum;
	private static int lastStepNum = 0;
	Border emptyBorder = BorderFactory.createEmptyBorder();
	Properties props = new Properties();
	boolean bolCtrl = false;
	boolean bolShift = false;
	boolean bolAlt = false;
	private JLabel imgLabel;
	JDialog exeHandle = null;
	
	public Run(String testPath, JFrame parentFrame){
		
		try {
			props.load(new FileInputStream(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg"));
		}
		catch(Exception e) {
			
		}
		/*try {
			GlobalScreen.registerNativeHook();
		} catch (NativeHookException e1) {
			e1.printStackTrace();
		}*/
		Date startime = new Date();
		tcPath = testPath; 
		lastStepNum = getlastStepNum();
		curScreenNum = getlastScreenNum() + 1;
		exeHandle = new JDialog();
		exeHandle.setModal(true);
        GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
        Rectangle rect = ge.getMaximumWindowBounds();
		//final JFrame exeHandle = new JFrame();
		exeHandle.setBackground(Color.LIGHT_GRAY);
		exeHandle.setUndecorated(true);
		exeHandle.setResizable(false);   
		exeHandle.setBounds(0, 0, 315, 162);
		exeHandle.getContentPane().setLayout(null);
		exeHandle.setAlwaysOnTop(true);
		exeHandle.setModalityType(ModalityType.APPLICATION_MODAL);
		
		JPanel executionPanel = new JPanel();
		executionPanel.setBorder(new LineBorder(new Color(160, 160, 160), 2, true));
		executionPanel.setBackground(SystemColor.menu);
		executionPanel.setBounds(0, 85, 315, 76);
		exeHandle.getContentPane().add(executionPanel);
		executionPanel.setLayout(null);
		
		JButton btnStop = new JButton("");
		btnStop.setIcon(new ImageIcon(Run.class.getResource("/res/Stop30x30.png")));
		btnStop.setBackground(SystemColor.menu);
		btnStop.setBorder(null);
		btnStop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(getTestDetails("executiontime").equals("")) {
					setTestDetails("executiontime",  Long.toString((long)(new Date().getTime() - startime.getTime())/(60 * 1000) % 60));
				}
				else{
					setTestDetails("executiontime",  Long.toString(Long.parseLong(getTestDetails("executiontime")) + (long)(new Date().getTime() - startime.getTime())/(60 * 1000) % 60));
				}
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
				parentFrame.setVisible(true);
				exeHandle.dispose();
			}
		});
		btnStop.setBounds(225, 23, 30, 30);
		executionPanel.add(btnStop);
		
		JButton btnDetails = new JButton("");
		//btnDetails.setIcon(new ImageIcon(Run.class.getResource("/res/arrowUp30x30.png")));
		btnDetails.setBackground(SystemColor.menu);
		btnDetails.setBounds(265, 23, 30, 30);
		btnDetails.setBorder(null);
		executionPanel.add(btnDetails);
		btnDetails.setIcon(new ImageIcon(Run.class.getResource("/res/arrowDown30x30.png")));
		
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
		              scrshotName = "Screenshot_" + Integer.toString(curScreenNum)  + "_" + new SimpleDateFormat("HH_mm_ss").format(new Date()) +".png";
		              ImageIO.write(rImage, "png", new File(tcPath + "\\Screenshots\\" + scrshotName));
		              addGeneralScreenshot(scrshotName, scrnComments.getText());
		              scrnComments.setText("");
		              curScreenNum++;
		              imgLabel.setVisible(true);
		              //((JFrame)exeHandle.getRootPane().getTopLevelAncestor()).setVisible(false);
		              parentFrame.setVisible(false);
		              exeHandle.setVisible(true);
		              
		              
				} catch (Exception e) {
					e.printStackTrace();
				}
	              
	          }
	      };
	    
	      
		JButton btnPass = new JButton("");
		btnPass.setBackground(Color.WHITE);
		btnPass.setToolTipText("Capture");
		/*Action pAct = new MouseEvent("") {
			public void actionPerformed(ActionEvent e) {
				
			}
		};*/
		btnPass.addMouseListener(new MouseAdapter() {
			public void mousePressed(MouseEvent ev) {
				try {
					if (ev.getClickCount() == 1) {
						exeHandle.setVisible(false);
		                Timer timer = new Timer(100, alPass);
		                timer.setRepeats(false);
		                timer.start();
					}
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
			}
		});
		//btnPass.getInputMap().put(KeyStroke.getKeyStroke(KeyEvent.VK_P,KeyEvent.CTRL_DOWN_MASK | KeyEvent.ALT_DOWN_MASK), "passAct");
		//btnPass.getActionMap().put("passAct", pAct);
		
		imgLabel = new JLabel("");
		imgLabel.setForeground(Color.WHITE);
		imgLabel.setBackground(Color.WHITE);
		imgLabel.setIcon(new ImageIcon(Run.class.getResource("/res/capture45x45.png")));
		imgLabel.setBounds(97, 13, 50, 50);
		executionPanel.add(imgLabel);
		btnPass.setBounds(69, 11, 102, 52);
		executionPanel.add(btnPass);
		
		JSeparator separator = new JSeparator();
		separator.setOrientation(SwingConstants.VERTICAL);
		separator.setBackground(UIManager.getColor("ProgressBar.shadow"));
		separator.setForeground(UIManager.getColor("ProgressBar.shadow"));
		separator.setBounds(50, 0, 2, 76);
		executionPanel.add(separator);
		
		JButton btnCollapse = new JButton("");
		btnCollapse.setIcon(new ImageIcon(Run.class.getResource("/res/arrowRight30x30.png")));
		btnCollapse.setBackground(SystemColor.menu);
		btnCollapse.setBorder(emptyBorder);
		btnCollapse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(exeHandle.getSize().getWidth()>150) {
					scrollPane.setVisible(false);
					executionPanel.setLocation(0, 0);
					exeHandle.setBounds(0, 0, 60, 76);
					btnDetails.setEnabled(false);
					btnCollapse.setIcon(new ImageIcon(Run.class.getResource("/res/arrowLeft30x30.png")));
					btnDetails.setIcon(new ImageIcon(Run.class.getResource("/res/arrowDown30x30.png")));
				}
				else {
					scrollPane.setVisible(true);
					executionPanel.setLocation(0, 85);
					exeHandle.setBounds(0, 0, 315, 160);
					btnDetails.setEnabled(true);
					btnCollapse.setIcon(new ImageIcon(Run.class.getResource("/res/arrowRight30x30.png")));
				}
				Dimension windowSize = exeHandle.getSize();
		        int x = (int) rect.getWidth() - (int) windowSize.getWidth();
		        int y = (int) rect.getHeight()/2 - (int) windowSize.getHeight()/2;
		        exeHandle.setLocation(x, y);
			}
		});
		btnCollapse.setBounds(10, 23, 30, 30);
		executionPanel.add(btnCollapse);
		
		JButton btnScrns = new JButton("");
		btnScrns.setIcon(new ImageIcon(Run.class.getResource("/res/photo-gallery20x20.png")));
		btnScrns.setBounds(185, 23, 30, 30);
		executionPanel.add(btnScrns);
		
		
		MouseListener imageEdi = new MouseListener() {
			
			@Override
			public void mouseReleased(java.awt.event.MouseEvent e) {}
			@Override
			public void mousePressed(java.awt.event.MouseEvent e) {}
			@Override
			public void mouseExited(java.awt.event.MouseEvent e) {}
			@Override
			public void mouseEntered(java.awt.event.MouseEvent e) {}
			
			@Override
			public void mouseClicked(java.awt.event.MouseEvent e) {
				try {
					File inputFile = new File(tcPath + "\\testDetails.tc");
			        SAXReader reader = new SAXReader();
			        Document document = reader.read(inputFile);
			        List<Node> nodes = document.selectNodes("/testcase/general");
			        if(nodes.size()!=0) {
				        Element lastEle = (Element) nodes.get(nodes.size() - 1);
				        String lastScrnName = lastEle.selectSingleNode("screenshot").getStringValue();
				        exeHandle.setAlwaysOnTop(false);
						exeHandle.setModalityType(ModalityType.MODELESS);
			            ImageEditor editor = new ImageEditor(lastScrnName, tcPath, "general");
			            exeHandle.setAlwaysOnTop(true);
						exeHandle.setModalityType(ModalityType.APPLICATION_MODAL);
			        }
			        else{
			        	JOptionPane.showMessageDialog(null, "No screenshots captured for this testcase yet!");
			        }
				} catch (Exception e1) {
					e1.printStackTrace();
				} catch (Throwable e1) {
					e1.printStackTrace();
				}
			}
		};
		btnScrns.addMouseListener(imageEdi);
		
		btnDetails.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(exeHandle.getSize().getHeight()>210) {
					exeHandle.setBounds(0, 0, exeHandle.getWidth(), 160);
					btnDetails.setIcon(new ImageIcon(Run.class.getResource("/res/arrowDown30x30.png")));
				}
				else {
					exeHandle.setBounds(0, 0, exeHandle.getWidth(), 402);
					btnDetails.setIcon(new ImageIcon(Run.class.getResource("/res/arrowUp30x30.png")));
				}
				Dimension windowSize = exeHandle.getSize();
		        int x = (int) rect.getWidth() - (int) windowSize.getWidth();
		        int y = (int) rect.getHeight()/2 - (int) windowSize.getHeight()/2;
		        exeHandle.setLocation(x, y);
			}
		});
		
		
		JPanel panelDetails = new JPanel();
		panelDetails.setBackground(SystemColor.menu);
		panelDetails.setBorder(new LineBorder(new Color(160, 160, 160), 2, true));
		panelDetails.setBounds(0, 161, 315, 242);
		exeHandle.getContentPane().add(panelDetails);
		panelDetails.setLayout(null);
		
		lblName.setBounds(10, 11, 235, 30);
		panelDetails.add(lblName);
		
		JLabel lblDdescription = new JLabel("Description:");
		lblDdescription.setBounds(10, 40, 78, 14);
		panelDetails.add(lblDdescription);
		
		
		txtpnDescpane.setLineWrap(true);
		txtpnDescpane.setWrapStyleWord(true);
		txtpnDescpane.setEditable(false);
		txtpnDescpane.setBounds(10, 52, 298, 76);
		JScrollPane DescScroll = new JScrollPane(txtpnDescpane);
		DescScroll.setBounds(10, 56, 298, 76);
		panelDetails.add(DescScroll);
		
		JLabel lblExpected = new JLabel("Expected:");
		lblExpected.setBounds(10, 144, 64, 14);
		panelDetails.add(lblExpected);
		txtpnExpected.setLineWrap(true);
		txtpnExpected.setWrapStyleWord(true);
		
		txtpnExpected.setEditable(false);
		txtpnExpected.setText("expected");
		txtpnExpected.setBounds(10, 157, 298, 76);
		JScrollPane expScroll = new JScrollPane(txtpnExpected);
		expScroll.setBounds(10, 161, 298, 76);
		panelDetails.add(expScroll);
		btnP.setBounds(241, 7, 30, 30);
		panelDetails.add(btnP);
		
		btnP.setIcon(new ImageIcon(Run.class.getResource("/res/Prev30x30.png")));
		btnP.setBackground(SystemColor.menu);
		btnP.setBorder(emptyBorder);
		btnNext.setBounds(271, 7, 30, 30);
		panelDetails.add(btnNext);
		
		
		btnNext.setIcon(new ImageIcon(Run.class.getResource("/res/Next30x30.png")));
		btnNext.setBackground(SystemColor.menu);
		btnNext.setBorder(emptyBorder);
		
		JLabel lblStep = new JLabel("Step :");
		lblStep.setBounds(247, 35, 36, 14);
		panelDetails.add(lblStep);
		labelStep.setBounds(281, 35, 24, 14);
		panelDetails.add(labelStep);
		labelStep.setText(Integer.toString(curStepNum + 1) + "/" + Integer.toString(lastStepNum));
		scrollPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
		
		
		scrollPane.setBounds(0, 0, 315, 85);
		exeHandle.getContentPane().add(scrollPane);
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
		
		if(curStepNum<=0) {
			btnP.setEnabled(false);
		}
		else if(curStepNum>=lastStepNum) {
			btnNext.setEnabled(false);
		}
		else{
			btnP.setEnabled(true);
			btnNext.setEnabled(true);
		}
		getStepDetails(curStepNum);
		
		Dimension windowSize = exeHandle.getSize();
        int x = (int) rect.getWidth() - (int) windowSize.getWidth();
        int y = (int) rect.getHeight()/2 - (int) windowSize.getHeight()/2;
		/*// Get the logger for "org.jnativehook" and set the level to off.
		Logger logger = Logger.getLogger(GlobalScreen.class.getPackage().getName());
		logger.setLevel(Level.OFF);

		// Change the level for all handlers attached to the default logger.
		Handler[] handlers = Logger.getLogger("").getHandlers();
		for (int i = 0; i < handlers.length; i++) {
		    handlers[i].setLevel(Level.OFF);
		}
		GlobalScreen.addNativeKeyListener(new NativeKeyListener() {
			
			@Override
			public void nativeKeyTyped(NativeKeyEvent arg0) {
				
			}
			
			@Override
			public void nativeKeyReleased(NativeKeyEvent e) {
				if(e.getKeyCode() == NativeKeyEvent.VC_CONTROL_L || e.getKeyCode() == NativeKeyEvent.VC_CONTROL_R) {
					bolCtrl = false;
				}
				else if(e.getKeyCode() == NativeKeyEvent.VC_SHIFT_L || e.getKeyCode() == NativeKeyEvent.VC_SHIFT_R) {
					bolShift = false;
				}
				else if(e.getKeyCode() == NativeKeyEvent.VC_ALT_L || e.getKeyCode() == NativeKeyEvent.VC_ALT_R) {
					bolAlt = false;
				}
			}
			
			@Override
			public void nativeKeyPressed(NativeKeyEvent e) {
				if(e.getKeyCode() == NativeKeyEvent.VC_CONTROL_L || e.getKeyCode() == NativeKeyEvent.VC_CONTROL_R) {
					bolCtrl = true;
				}
				else if(e.getKeyCode() == NativeKeyEvent.VC_SHIFT_L || e.getKeyCode() == NativeKeyEvent.VC_SHIFT_R) {
					bolShift = true;
				}
				else if(e.getKeyCode() == NativeKeyEvent.VC_ALT_L || e.getKeyCode() == NativeKeyEvent.VC_ALT_R) {
					bolAlt = true;
				}
				
				if(bolCtrl && bolShift && bolAlt) {
					// Capture screen method
					exeHandle.setVisible(false);
	                Timer timer = new Timer(100, alPass);
	                timer.setRepeats(false);
	                timer.start();
				}
			}
		});*/
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
        	txtpnDescpane.setText(tcDesc);
        	txtpnExpected.setText(tcExp);
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
        	//element.selectSingleNode("actualresult").setText(txtrActual.getText());
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
	
	public void addGeneralScreenshot(String screenshotPath, String comment) {
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectSingleNode("/testcase");
        	Element element = (Element) node;
        	Element genEle = element.addElement("general");
        	genEle.addElement("screenshot").setText(screenshotPath);
        	genEle.addElement("comment").setText(comment);
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
		return stepNum;
	}
	
	public int getlastScreenNum() {
		int stepNum = 0;
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        stepNum = document.selectNodes("/testcase/general").size();
		}
		catch(Exception e) {
			
		}
		return stepNum;
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
