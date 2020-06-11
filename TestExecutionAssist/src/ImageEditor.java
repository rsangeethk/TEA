package src;

import java.awt.Color;
import java.awt.Component;
import java.awt.Dialog;
import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Image;
import java.awt.SystemColor;
import java.awt.Toolkit;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.imageio.ImageIO;
import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.BorderFactory;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JScrollPane;
import javax.swing.JSeparator;
import javax.swing.JTextArea;
import javax.swing.JToggleButton;
import javax.swing.KeyStroke;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.border.Border;

import org.apache.commons.lang.ArrayUtils;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

public class ImageEditor {
	JFrame imgframe = new JFrame();
	JTextArea scrnCommnt = new JTextArea();
	JScrollPane cmntScroll =  new JScrollPane(scrnCommnt);
	JLabel cmtLabel = new JLabel("Comments:");
	JLabel expLabel = new JLabel("Expected Result:");
	JLabel actLabel = new JLabel("Actual Result:");
	JTextArea expRes = new JTextArea();
	JTextArea actRes = new JTextArea();
	JToggleButton redColor = new JToggleButton("");
	JToggleButton blueColor = new JToggleButton("");
	JToggleButton greenColor = new JToggleButton("");
	JToggleButton blackColor = new JToggleButton("");
	JButton prevBtn =  new JButton();
	JButton nextBtn =  new JButton();
	private int curIndex = 0;
	String curScrnSht = "";
	Border emptyBorder = BorderFactory.createEmptyBorder();
	private JLabel imgLabel = new JLabel();
	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
	String tcpath = "";
	private String[] scrns;
	int intbackX = 0;
	int intbackY = 32;
	private double zoom = 1.0;
	String screenType = "";
	
	DrawingFrame asPanel = null;
	private final JButton btnDelete = new JButton();
	public ImageEditor(String screenName, String tcPath, String screenType) throws Exception {
		this.tcpath = tcPath;
		this.screenType = screenType;
		asPanel = new DrawingFrame();
		imgframe.getContentPane().setLayout(null);
		imgframe.setResizable(false);
		imgframe.setTitle("Snapshot editor");
		imgframe.setIconImage(Toolkit.getDefaultToolkit().getImage(tea.class.getResource("/res/logo48x48.png")));
		imgframe.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		imgframe.setSize(820, 620);
		imgframe.setLocation(dim.width/2- imgframe.getSize().width/2, dim.height/2-imgframe.getSize().height/2);
		
		JPanel backPack = new JPanel();
		backPack.setLayout(null);
		
		backPack.setBounds(intbackX, intbackY, 820, 460);
		imgframe.getContentPane().add(backPack);
		/*backPack.addMouseListener(new MouseAdapter() {
	        @Override
	        public void mousePressed(MouseEvent me) {
	        	intbackX = me.getX();
	        	intbackY = me.getY();
	        }
	    });
		backPack.addMouseMotionListener(new MouseMotionAdapter() {
	        @Override
	        public void mouseDragged(MouseEvent me) {
	        	if(me.getX()<0 || me.getY()<32) {
		            me.translatePoint(me.getComponent().getLocation().x-intbackX, me.getComponent().getLocation().y-intbackY);
		            backPack.setLocation(me.getX(), me.getY());
	        	}
	        }
	    });
		
		backPack.addMouseWheelListener(new MouseWheelListener() {
			
			@Override
			public void mouseWheelMoved(MouseWheelEvent e) {
                int notches = e.getWheelRotation();
                double temp = zoom - (notches * 0.2);
                // minimum zoom factor is 1.0
                temp = Math.max(temp, 1.0);
                if (temp != zoom) {
                    zoom = temp;
                    resizeImage();
                }
                imgLabel.revalidate();
                imgLabel.repaint();
                backPack.revalidate();
                backPack.repaint();
            }

		});*/
		
		JPanel panel = asPanel;
		panel.setLayout(null);
		panel.setBounds(0, 0, 820, 460);
		imgLabel.setBounds(0, 0, 820, 460);
	    backPack.add(imgLabel);
		
		backPack.add(panel);
		backPack.setComponentZOrder(panel, 0);
		
		
		for (Component eachComponent : panel.getComponents()) {
			if (eachComponent != null) {
				if (eachComponent instanceof JToggleButton) {
					JToggleButton tog = (JToggleButton) eachComponent;
					tog.setIcon(new ImageIcon(DrawingFrame.class.getResource("/res/highlight15x15.png")));
					tog.setBackground(SystemColor.activeCaptionBorder);
					tog.setBounds(32, 5, 25, 25);
					tog.setVisible(true);
					imgframe.getContentPane().add(tog);
				}
			}
		}
		redColor.setToolTipText("");
		
		redColor.setBounds(95, 5, 25, 25);
		redColor.setIcon(new ImageIcon(DrawingFrame.class.getResource("/res/red10x10.png")));
		imgframe.getContentPane().add(redColor);
		blueColor.setToolTipText("");
		
		
		
		blueColor.setBounds(120, 5, 25, 25);
		blueColor.setIcon(new ImageIcon(DrawingFrame.class.getResource("/res/blue10x10.png")));
		imgframe.getContentPane().add(blueColor);
		greenColor.setToolTipText("");
		
		
		greenColor.setBounds(145, 5, 25, 25);
		greenColor.setIcon(new ImageIcon(DrawingFrame.class.getResource("/res/green10x10.png")));
		imgframe.getContentPane().add(greenColor);
		blackColor.setToolTipText("");
		
		
		blackColor.setBounds(170, 5, 25, 25);
		blackColor.setIcon(new ImageIcon(DrawingFrame.class.getResource("/res/black10x10.png")));
		imgframe.getContentPane().add(blackColor);
		
		blueColor.setName("blue");
		redColor.setName("red");
		greenColor.setName("green");
		blackColor.setName("black");
		
		redColor.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				if(((JToggleButton) e.getSource()).isSelected()) {
					redColor.setToolTipText("Selected");
					blueColor.setToolTipText("");
					greenColor.setToolTipText("");
					blackColor.setToolTipText("");
					blueColor.setSelected(false);
					greenColor.setSelected(false);
					blackColor.setSelected(false);
				}
				
			}
		});
		
		blueColor.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				if(((JToggleButton) e.getSource()).isSelected()) {
					blueColor.setToolTipText("Selected");
					redColor.setToolTipText("");
					greenColor.setToolTipText("");
					blackColor.setToolTipText("");
					redColor.setSelected(false);
					greenColor.setSelected(false);
					blackColor.setSelected(false);
				}
			}
		});
		
		greenColor.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				if(((JToggleButton) e.getSource()).isSelected()) {
					greenColor.setToolTipText("Selected");
				blueColor.setToolTipText("");
				redColor.setToolTipText("");
				blackColor.setToolTipText("");
				blueColor.setSelected(false);
				redColor.setSelected(false);
				blackColor.setSelected(false);
				}
			}
		});
		
		blackColor.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				if(((JToggleButton) e.getSource()).isSelected()) {
				blackColor.setToolTipText("Selected");
				blueColor.setToolTipText("");
				greenColor.setToolTipText("");
				redColor.setToolTipText("");
				blueColor.setSelected(false);
				greenColor.setSelected(false);
				redColor.setSelected(false);
				}
			}
		});
		
		JButton btnSave = new JButton();
		btnSave.setToolTipText("Save");
		btnSave.setBounds(5, 5, 25, 25);
		btnSave.setIcon(new ImageIcon(DrawingFrame.class.getResource("/res/Save15x15.png")));
		btnSave.setBackground(SystemColor.activeCaptionBorder);
		//btnSave.setBorder(emptyBorder);
		imgframe.getContentPane().add(btnSave);
		
		Action saveAction = new AbstractAction() {
			@Override
			public void actionPerformed(ActionEvent e) {
				BufferedImage panelImg = new BufferedImage(820, 460, BufferedImage.TYPE_INT_ARGB);
				Graphics g = panelImg.getGraphics();
	            backPack.paint(g);
	            try {
	                ImageIO.write(panelImg, "png", new File(curScrnSht));
	            } catch (IOException ex) {
	               	ex.printStackTrace();
	            }
	            int startINdex = scrns[curIndex].indexOf("_") + 1;
				int secondIndex = scrns[curIndex].indexOf("_", startINdex) ;
	            if(screenType.equalsIgnoreCase("general")) {
					setgeneralDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "comment", scrnCommnt.getText());
				}
				else{
					setStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "expectedresult", expRes.getText());
					setStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "actualresult",actRes.getText());
				}
			}
		};
		btnSave.addActionListener(saveAction);
		btnSave.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_S, KeyEvent.CTRL_DOWN_MASK), "saveAction");
		btnSave.getActionMap().put("saveAction", saveAction);
		
		prevBtn.setBounds(8, 530, 30, 30);
		prevBtn.setIcon(new ImageIcon(DrawingFrame.class.getResource("/res/Prev30x30.png")));
		imgframe.getContentPane().add(prevBtn);
		
		
		nextBtn.setBounds(770, 530, 30, 30);
		nextBtn.setIcon(new ImageIcon(DrawingFrame.class.getResource("/res/Next30x30.png")));
		imgframe.getContentPane().add(nextBtn);
		
		cmtLabel.setBounds(50, 490, 120, 20);
		imgframe.getContentPane().add(cmtLabel);
		
		expLabel.setBounds(50, 490, 120, 20);
		imgframe.getContentPane().add(expLabel);
		
		actLabel.setBounds(415, 490, 120, 20);
		imgframe.getContentPane().add(actLabel);
		
		scrnCommnt.setLineWrap(true);
		
		cmntScroll.setBounds(50, 510, 705, 70);
		imgframe.getContentPane().add(cmntScroll);
		
		expRes.setLineWrap(true);
		JScrollPane expScrl =  new JScrollPane(expRes);
		expScrl.setBounds(50, 510, 340, 70);
		imgframe.getContentPane().add(expScrl);
		
		
		actRes.setLineWrap(true);
		JScrollPane actScrl =  new JScrollPane(actRes);
		actScrl.setBounds(415, 510, 340, 70);
		imgframe.getContentPane().add(actScrl);
		
		JButton btnClear = new JButton();
		btnClear.setToolTipText("Clear");
		btnClear.setIcon(new ImageIcon(ImageEditor.class.getResource("/res/eraser16x16.png")));
		btnClear.setBackground(SystemColor.activeCaptionBorder);
		btnClear.setBounds(57, 5, 25, 25);
		imgframe.getContentPane().add(btnClear);
		
		
		Object[] thicknessItems = {
				new ImageIcon(ImageEditor.class.getResource("/res/small_line.png")),
				new ImageIcon(ImageEditor.class.getResource("/res/medium_line.png"))
	        };

		JComboBox<Object> thicknessCombo = new JComboBox<Object>(thicknessItems);
		thicknessCombo.setToolTipText("Select thickness");
		thicknessCombo.setBounds(196, 5, 99, 26);
		imgframe.getContentPane().add(thicknessCombo);
		
		JSeparator editSeperator = new JSeparator();
		editSeperator.setOrientation(SwingConstants.VERTICAL);
		editSeperator.setBackground(Color.DARK_GRAY);
		editSeperator.setBounds(88, 5, 2, 25);
		imgframe.getContentPane().add(editSeperator);
		
		JLabel pagLbl = new JLabel("0 / 0");
		pagLbl.setHorizontalAlignment(SwingConstants.CENTER);
		pagLbl.setBounds(745, 11, 55, 16);
		
		imgframe.getContentPane().add(pagLbl);
		
		
		/*JSlider slider = new JSlider();
		slider.setMinorTickSpacing(50);
		slider.setPaintTicks(true);
		slider.setPaintLabels(true);
		slider.setMajorTickSpacing(100);
		slider.setValue(0);
		slider.setMaximum(200);
		slider.setBounds(289, 7, 101, 39);
		imgframe.getContentPane().add(slider);
		slider.addChangeListener(new ChangeListener() {
			@Override
			public void stateChanged(ChangeEvent e) {
				JSlider theJSlider = (JSlider) e.getSource();
			      if (!theJSlider.getValueIsAdjusting()) {
			        if(theJSlider.getValue() > 25) {
			        	zoom = 25;
		                resizeImage();
		                imgLabel.revalidate();
		                imgLabel.repaint();
		                backPack.revalidate();
		                backPack.repaint();
			        }
			      }

			}
		});*/
		
		Action clearAction = new  AbstractAction() {
			@SuppressWarnings("static-access")
			@Override
			public void actionPerformed(ActionEvent e) {
				try{
					curScrnSht = tcpath + "\\Screenshots\\" + scrns[curIndex];
					Image bImg1 = ImageIO.read(new File(tcpath + "\\Screenshots\\" + scrns[curIndex])).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
		            ImageIcon newImc1 = new ImageIcon(bImg1);
					imgLabel.setIcon(newImc1);
					Graphics  g = asPanel.backgroundImg.getGraphics();
					g.drawImage(bImg1, 0, 0, null);
					g.dispose();
					int startINdex = scrns[curIndex].indexOf("_") + 1;
					int secondIndex = scrns[curIndex].indexOf("_", startINdex) ;
					if(screenType.equalsIgnoreCase("general")) {
						scrnCommnt.setText(getgeneralDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "comment"));
					}
					else{
						expRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "expectedresult"));
						actRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "actualresult"));
					}
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
			}
		};
		btnClear.addActionListener(clearAction);
		btnClear.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_DELETE, 0), "clearAction");
		btnClear.getActionMap().put("clearAction", clearAction);
		
		Action prevBtnAction = new  AbstractAction() {
			@SuppressWarnings("static-access")
			@Override
			public void actionPerformed(ActionEvent e) {
				try{
					
					if(curIndex > 0 && new File(tcpath + "\\Screenshots\\" + scrns[curIndex]).exists()) {
						curIndex--;
						curScrnSht = tcpath + "\\Screenshots\\" + scrns[curIndex];
						Image bImg1 = ImageIO.read(new File(tcpath + "\\Screenshots\\" + scrns[curIndex])).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
			            ImageIcon newImc1 = new ImageIcon(bImg1);
						imgLabel.setIcon(newImc1);
						Graphics  g = asPanel.backgroundImg.getGraphics();
						g.drawImage(bImg1, 0, 0, null);
						g.dispose();
						int startINdex = scrns[curIndex].indexOf("_") + 1;
						int secondIndex = scrns[curIndex].indexOf("_", startINdex) ;
						if(screenType.equalsIgnoreCase("general")) {
							scrnCommnt.setText(getgeneralDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "comment"));
						}
						else{
							expRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "expectedresult"));
							actRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "actualresult"));
						}
					}
					if(curIndex > 0){
						prevBtn.setEnabled(true);
						nextBtn.setEnabled(true);
					}
					else{
						prevBtn.setEnabled(false);
						nextBtn.setEnabled(true);
					}
					pagLbl.setText("" + (curIndex + 1) +" / " + scrns.length);
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
			}
		};
		prevBtn.addActionListener(prevBtnAction);
		prevBtn.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_LEFT, 0), "prevAction");
		prevBtn.getActionMap().put("prevAction", prevBtnAction);
		
		Action nextBtnAction = new  AbstractAction() {
			@SuppressWarnings("static-access")
			@Override
			public void actionPerformed(ActionEvent e) {
				try{
					if(curIndex < scrns.length - 1 && new File(tcpath + "\\Screenshots\\" + scrns[curIndex]).exists()) {
						curIndex++;
						curScrnSht = tcpath + "\\Screenshots\\" + scrns[curIndex];
						Image bImg2 = ImageIO.read(new File(tcpath + "\\Screenshots\\" + scrns[curIndex])).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
			            ImageIcon newImc2 = new ImageIcon(bImg2);
						imgLabel.setIcon(newImc2);
						Graphics  g = asPanel.backgroundImg.getGraphics();
						g.drawImage(bImg2, 0, 0, null);
						g.dispose();
						int startINdex = scrns[curIndex].indexOf("_") + 1;
						int secondIndex = scrns[curIndex].indexOf("_", startINdex) ;
						if(screenType.equalsIgnoreCase("general")) {
							scrnCommnt.setText(getgeneralDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "comment"));
						}
						else{
							
							expRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "expectedresult"));
							actRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "actualresult"));
						}
					}
					if(curIndex < scrns.length - 1){
						nextBtn.setEnabled(true);
						prevBtn.setEnabled(true);
					}
					else{
						nextBtn.setEnabled(false);
						prevBtn.setEnabled(true);
					}
					pagLbl.setText("" + (curIndex + 1) +" / " + scrns.length);
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
			}
		};
		nextBtn.addActionListener(nextBtnAction);
		nextBtn.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_RIGHT, 0), "nextAction");
		nextBtn.getActionMap().put("nextAction", nextBtnAction);
		
		imgframe.setVisible(false);
		openImageEditor(screenName, screenType);
		pagLbl.setText("" + (curIndex + 1) +" / " + scrns.length);
		
		btnDelete.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					if(JOptionPane.showConfirmDialog(null, "Are you sure you want to delete this Screenshot and its comments from the testcase?") == JOptionPane.YES_OPTION) {
						deleteNode();
						if(curIndex > 0 && scrns.length > 0) {
							curIndex--;
							curScrnSht = tcpath + "\\Screenshots\\" + scrns[curIndex];
							Image bImg1 = ImageIO.read(new File(tcpath + "\\Screenshots\\" + scrns[curIndex])).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
				            ImageIcon newImc1 = new ImageIcon(bImg1);
							imgLabel.setIcon(newImc1);
							Graphics  g = asPanel.backgroundImg.getGraphics();
							g.drawImage(bImg1, 0, 0, null);
							g.dispose();
							int startINdex = scrns[curIndex].indexOf("_") + 1;
							int secondIndex = scrns[curIndex].indexOf("_", startINdex) ;
							if(screenType.equalsIgnoreCase("general")) {
								scrnCommnt.setText(getgeneralDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "comment"));
							}
							else{
								expRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "expectedresult"));
								actRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "actualresult"));
							}
						}
						else if(curIndex == 0 && scrns.length > 0){
							curIndex++;
							curScrnSht = tcpath + "\\Screenshots\\" + scrns[curIndex];
							Image bImg2 = ImageIO.read(new File(tcpath + "\\Screenshots\\" + scrns[curIndex])).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
				            ImageIcon newImc2 = new ImageIcon(bImg2);
							imgLabel.setIcon(newImc2);
							Graphics  g = asPanel.backgroundImg.getGraphics();
							g.drawImage(bImg2, 0, 0, null);
							g.dispose();
							int startINdex = scrns[curIndex].indexOf("_") + 1;
							int secondIndex = scrns[curIndex].indexOf("_", startINdex) ;
							if(screenType.equalsIgnoreCase("general")) {
								scrnCommnt.setText(getgeneralDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "comment"));
							}
							else{
								expRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "expectedresult"));
								actRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "actualresult"));
							}
						}
						if(curIndex == 0){
							prevBtn.setEnabled(false);
							nextBtn.setEnabled(true);
						}
						else if(curIndex == scrns.length - 1){
							nextBtn.setEnabled(false);
							prevBtn.setEnabled(true);
						}
						else {
							nextBtn.setEnabled(true);
							prevBtn.setEnabled(true);
						}
						scrns = getScreenshots(screenType).split(";");
						pagLbl.setText("" + (curIndex + 1) +" / " + scrns.length);
					}
				}
				catch(Exception e) {
					e.printStackTrace();
				}
			}
		});
		btnDelete.setIcon(new ImageIcon(ImageEditor.class.getResource("/res/delete-15x15.png")));
		btnDelete.setToolTipText("Delete");
		btnDelete.setBackground(SystemColor.activeCaptionBorder);
		btnDelete.setBounds(297, 5, 25, 25);
		
		JPopupMenu imagePopup = new JPopupMenu();
		JMenuItem copyMenu = new JMenuItem("Copy Screenshot");
		imagePopup.add(copyMenu);
		copyMenu.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					curScrnSht = tcpath + "\\Screenshots\\" + scrns[curIndex];
					Image bImg = ImageIO.read(new File(tcpath + "\\Screenshots\\" + scrns[curIndex])).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
					setClipboard(bImg);
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
			}
		});
		MouseListener imageViewerListener = new MouseAdapter() {
			public void mousePressed(MouseEvent me) {
				try{
					if (me.getClickCount() == 1 && SwingUtilities.isRightMouseButton(me)) {
						imagePopup.show(me.getComponent(), me.getX(), me.getY());
				    }
				}
				catch(Exception e) {
					
				}
			}
		};
		asPanel.addMouseListener(imageViewerListener);
		
		imgframe.getContentPane().add(btnDelete);
		imgframe.setVisible(true);
    	imgframe.setEnabled(true);
    	imgframe.setModalExclusionType(Dialog.ModalExclusionType.APPLICATION_EXCLUDE);
	}
	
	private String deleteNode() {
		String value = "";
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        List<Object> allNodes;
	        if(screenType.equalsIgnoreCase("general")) {
	        	allNodes = document.selectNodes("/testcase/general");
	        }
	        else{
	        	allNodes = document.selectNodes("/testcase/teststep");
	        }
	        for(Object eachNode : allNodes) {
				Node node = (Node) eachNode;
				Element element = (Element) node;
	        	if(scrns[curIndex].equalsIgnoreCase(element.selectSingleNode("screenshot").getText())){
	        		element.getParent().remove(element);
	        		break;
	        	}
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
		return value;
	}
	
	private String getScreenshots(String type) {
		String value = "";
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        List<Node> nodes;
	        if(type.equalsIgnoreCase("general")) {
		        nodes = document.selectNodes("/testcase/general");
	        }
	        else{
	        	nodes = document.selectNodes("/testcase/teststep");
	        }
	        for(Node eachNode : nodes)  {
	        	List<Node> scrnNodes = eachNode.selectNodes("screenshot");
	        	for(Node scrnNode : scrnNodes) {
		        	Element element = (Element) scrnNode;
		        	if(!element.getText().equals("")) {
		        	value = value + ";" + element.getText();
		        	}
	        	}
	        }
	        if(!value.equals("")) {
	        	value =  value.substring(1);
	        }
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return value;
	}
	
	public String getStepDetails(int stepNum, String tagName) {
		String tcStepDetails =  "";
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        List<Object> allNodes = document.selectNodes("/testcase/teststep");
			for(Object eachNode : allNodes) {
				Node node = (Node)eachNode;
				Element element = (Element) node;
	        	if(scrns[curIndex].equalsIgnoreCase(element.selectSingleNode("screenshot").getText())){
	        		tcStepDetails = element.selectSingleNode(tagName).getText();
	        	}
			}
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return tcStepDetails;
	}
	
	public void setStepDetails(int stepNum, String tagName, String value) {
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        List<Object> allNodes = document.selectNodes("/testcase/teststep");
			for(Object eachNode : allNodes) {
				Node node = (Node)eachNode;
				Element element = (Element) node;
	        	if(scrns[curIndex].equalsIgnoreCase(element.selectSingleNode("screenshot").getText())){
	        		element.selectSingleNode(tagName).setText(value);
	        	}
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
	
	public String getgeneralDetails(int stepNum, String tagName) {
		String tcStepDetails =  "";
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        List<Object> allNodes = document.selectNodes("/testcase/general");
			for(Object eachNode : allNodes) {
				Node node = (Node)eachNode;
				Element element = (Element) node;
	        	if(scrns[curIndex].equalsIgnoreCase(element.selectSingleNode("screenshot").getText())){
	        		tcStepDetails = element.selectSingleNode(tagName).getText();
	        	}
			}
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return tcStepDetails;
	}
	
	public void setgeneralDetails(int stepNum, String tagName, String value) {
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        List<Object> allNodes = document.selectNodes("/testcase/general");
			for(Object eachNode : allNodes) {
				Node node = (Node)eachNode;
				Element element = (Element) node;
	        	if(scrns[curIndex].equalsIgnoreCase(element.selectSingleNode("screenshot").getText())){
	        		element.selectSingleNode(tagName).setText(value);
	        	}
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
	
	 public void openImageEditor(String scrnName, String type) {
			if(type.equalsIgnoreCase("general")){
				expLabel.setVisible(false);
				actLabel.setVisible(false);
				expRes.setVisible(false);
				actRes.setVisible(false);
				cmtLabel.setVisible(true);
				scrnCommnt.setVisible(true);
				cmntScroll.setVisible(true);
				scrns = getScreenshots("general").split(";");
				//String scrnName = ((JLabel) e.getComponent()).getToolTipText();
				curIndex = ArrayUtils.indexOf(scrns, scrnName);
				curScrnSht = tcpath + "\\Screenshots\\" + scrns[curIndex];
				try {
					Image bImg = ImageIO.read(new File(curScrnSht)).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
					ImageIcon newImc = new ImageIcon(bImg);
				    imgLabel.setIcon(newImc);
				    asPanel.prefW = bImg.getWidth(null);
				    asPanel.prefH = bImg.getHeight(null);
				    asPanel.backgroundImg = new BufferedImage(asPanel.prefW, asPanel.prefH, BufferedImage.TYPE_INT_ARGB);
					Graphics  g = asPanel.backgroundImg.getGraphics();
					g.drawImage(bImg, 0, 0, null);
					g.dispose();
					int startINdex = scrns[curIndex].indexOf("_") + 1;
					int secondIndex = scrns[curIndex].indexOf("_", startINdex) ;
					scrnCommnt.setText(getgeneralDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "comment"));
					if(curIndex == 0 & scrns.length > 1){
						prevBtn.setEnabled(false);
						nextBtn.setEnabled(true);
					}
					else if(curIndex == 0 && scrns.length <= 1){
						prevBtn.setEnabled(false);
						nextBtn.setEnabled(false);
					}
					else if(curIndex == scrns.length - 1){
						prevBtn.setEnabled(true);
						nextBtn.setEnabled(false);
					}
					else{
						prevBtn.setEnabled(true);
						nextBtn.setEnabled(true);
					}
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
			else{
				expLabel.setVisible(true);
				actLabel.setVisible(true);
				expRes.setVisible(true);
				actRes.setVisible(true);
				cmtLabel.setVisible(false);
				scrnCommnt.setVisible(false);
				cmntScroll.setVisible(false);
				scrns  = getScreenshots("step").split(";");
				//String scrnName = ((JLabel) e.getComponent()).getToolTipText();
				curIndex = ArrayUtils.indexOf(scrns, scrnName);
				curScrnSht = tcpath + "\\Screenshots\\" + scrns[curIndex];
				try {
					Image bImg = ImageIO.read(new File(curScrnSht)).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
					ImageIcon newImc = new ImageIcon(bImg);
				    imgLabel.setIcon(newImc);
				    asPanel.prefW = bImg.getWidth(null);
				    asPanel.prefH = bImg.getHeight(null);
				    asPanel.backgroundImg = new BufferedImage(asPanel.prefW, asPanel.prefH, BufferedImage.TYPE_INT_ARGB);
					Graphics  g = asPanel.backgroundImg.getGraphics();
					g.drawImage(bImg, 0, 0, null);
					g.dispose();
					int startINdex = scrns[curIndex].indexOf("_") + 1;
					int secondIndex = scrns[curIndex].indexOf("_", startINdex) ;
					expRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "expectedresult"));
					actRes.setText(getStepDetails(Integer.parseInt(scrns[curIndex].substring(startINdex, secondIndex)) - 1, "actualresult"));
					if(curIndex == 0 & scrns.length > 1){
						prevBtn.setEnabled(false);
						nextBtn.setEnabled(true);
					}
					else if(curIndex == 0 && scrns.length <= 1){
						prevBtn.setEnabled(false);
						nextBtn.setEnabled(false);
					}
					else if(curIndex == scrns.length - 1){
						nextBtn.setEnabled(false);
						prevBtn.setEnabled(true);
					}
					else{
						prevBtn.setEnabled(true);
						nextBtn.setEnabled(true);
					}
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
	    }
	 
	 	//This method writes a image to the system clipboard.
		//otherwise it returns null.
		public static void setClipboard(Image image)
		{
		   ImageSelection imgSel = new ImageSelection(image);
		   Toolkit.getDefaultToolkit().getSystemClipboard().setContents(imgSel, null);
		}


		// This class is used to hold an image while on the clipboard.
		static class ImageSelection implements Transferable
		{
			  private Image image;
	
			  public ImageSelection(Image image)
			  {
			    this.image = image;
			  }
	
			  // Returns supported flavors
			  public DataFlavor[] getTransferDataFlavors()
			  {
			    return new DataFlavor[] { DataFlavor.imageFlavor };
			  }
	
			  // Returns true if flavor is supported
			  public boolean isDataFlavorSupported(DataFlavor flavor)
			  {
			    return DataFlavor.imageFlavor.equals(flavor);
			  }
	
			  // Returns image
			  public Object getTransferData(DataFlavor flavor)
			      throws UnsupportedFlavorException, IOException
			  {
			    if (!DataFlavor.imageFlavor.equals(flavor))
			    {
			      throw new UnsupportedFlavorException(flavor);
			    }
			    return image;
			  }
		}
}
