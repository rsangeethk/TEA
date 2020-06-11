package src;

import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import javax.imageio.ImageIO;
import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.DefaultComboBoxModel;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFormattedTextField.AbstractFormatter;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenuItem;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JScrollPane;
import javax.swing.JToggleButton;
import javax.swing.KeyStroke;
import javax.swing.SpringLayout;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.border.LineBorder;
import javax.swing.border.TitledBorder;

import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;
import org.jdatepicker.impl.JDatePanelImpl;
import org.jdatepicker.impl.JDatePickerImpl;
import org.jdatepicker.impl.UtilDateModel;


public class ImageRepo {
	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
	static String tcPath = "";
	JFrame imgViewerfrm = new JFrame();
	JLabel imgLabel = new JLabel("");
	int curImgIndex = 0;
	int selectedImgCount = 0;
	JPanel imgPanel = new JPanel();
	JToggleButton selectImgBtn = new JToggleButton("");
	JPopupMenu imageOpenMenu = new JPopupMenu();
	JLabel lblviewerSelected = new JLabel("0 Selected");
	JLabel imageCount = new JLabel("");
	int curScreenNum = 0;
	private JComboBox<String> repoCombo;
	String repoPath = "";
	private JDatePickerImpl datePickerFrom;
	private JDatePickerImpl datePickerTo;
	private JComboBox<String> userFiltercmbo;
	String userName = System.getProperty("user.name");
	
	@SuppressWarnings("serial")
	public ImageRepo(String tcPath, String projectpath) throws Throwable {
		//UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
		this.repoPath = projectpath + "\\Snapshot Repo";
		this.tcPath = tcPath;
		curScreenNum = getlastScreenNum() + 1;
		//Add steps list in the current Test case
		List<String> steps = new ArrayList<>();
		steps.add("General");
		for(int eachStep = 1; eachStep<=getStepsCount(); eachStep++) {
			steps.add("Step " + eachStep);
		}
		
		//Add scrnsRepo List
		ArrayList<String> scrnsList = new ArrayList<String>();
		ArrayList<String> usersList = new ArrayList<String>();
		ArrayList<String> repoList = new ArrayList<String>();
		usersList.add("All");
		repoList.add("<Default>");
		for(File eachScrnSht : new File(projectpath + "\\Snapshot Repo").listFiles()) {
			if(eachScrnSht.getName().toLowerCase().endsWith(".png")) {
				scrnsList.add(eachScrnSht.getName());
				if(eachScrnSht.getName().split("_").length>0) {
					if(!usersList.contains(eachScrnSht.getName().split("_")[0])) {
						usersList.add(eachScrnSht.getName().split("_")[0]);	
					}
 				}
			}
			else if(eachScrnSht.isDirectory()) {
				repoList.add(eachScrnSht.getName());
			}
		}
		
		JFrame frmImportFromRepo = new JFrame();
		frmImportFromRepo.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		frmImportFromRepo.setIconImage(Toolkit.getDefaultToolkit().getImage(tea.class.getResource("/res/logo48x48.png")));
		frmImportFromRepo.setResizable(false);
		frmImportFromRepo.getContentPane().setLayout(null);
		frmImportFromRepo.setSize(new Dimension(950, 620));
		frmImportFromRepo.setTitle("Import from Repo");
		frmImportFromRepo.setLocation(dim.width/2 - frmImportFromRepo.getSize().width/2, dim.height/2-frmImportFromRepo.getSize().height/2);
		
		
		JPanel repoPanel = new JPanel();
		repoPanel.setBounds(0, 0, 603, 619);
		frmImportFromRepo.getContentPane().add(repoPanel);
		repoPanel.setLayout(null);
		JPanel imagesPanel = new JPanel();
		imagesPanel.setLayout(new WrapFlowLayout());
		JScrollPane scrollPane = new JScrollPane(imagesPanel);
		scrollPane.setBounds(0, 76, 600, 520);
		repoPanel.add(scrollPane);
		
		JLabel lblSelected = new JLabel("0 Selected");
		lblSelected.setBounds(531, 6, 63, 14);
		repoPanel.add(lblSelected);
		
		JPopupMenu imageMenu = new JPopupMenu();
		JMenuItem openImage = new JMenuItem("Open in Viewer");
		imageMenu.add(openImage);
		openImage.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try{
					JToggleButton selecImage = (JToggleButton)((JPopupMenu)((JMenuItem)e.getSource()).getParent()).getInvoker();
					curImgIndex = scrnsList.indexOf(selecImage.getToolTipText());
					Image bImg = ImageIO.read(new File(new File(repoPath + "\\" + selecImage.getToolTipText()).getAbsolutePath())).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
					ImageIcon newImc = new ImageIcon(bImg);
				    imgLabel.setIcon(newImc);
				    BufferedImage backgroundImg = new BufferedImage(820, 460, BufferedImage.TYPE_INT_ARGB);
					Graphics  g = backgroundImg.getGraphics();
					g.drawImage(bImg, 0, 0, null);
					g.dispose();
					if(((JToggleButton)imagesPanel.getComponent(curImgIndex)).isSelected()) {
						selectImgBtn.setSelected(true);
					}
					else{
						selectImgBtn.setSelected(false);
					}
					lblSelected.setText(selectedImgCount + " Selected");
					lblviewerSelected.setText(selectedImgCount + " Selected");
					imageCount.setText((curImgIndex + 1) + " of " + scrnsList.size() );
					imgViewerfrm.setVisible(true);
				}
				catch(Exception ex){
					ex.printStackTrace();
				}
			}
		});
		
		ItemListener imageStateListener = new ItemListener() {
			
			@Override
			public void itemStateChanged(ItemEvent e) {
				if(e.getStateChange() == ItemEvent.SELECTED) {
					selectedImgCount++;
				}
				else if(e.getStateChange() == ItemEvent.DESELECTED){
					selectedImgCount--;
				}
			lblSelected.setText(selectedImgCount + " Selected");
			}
		};
		
		MouseListener imageViewerListener = new MouseAdapter() {
			public void mousePressed(MouseEvent me) {
				try{
				if (me.getClickCount() == 1 && SwingUtilities.isRightMouseButton(me)) {
					imageMenu.show(me.getComponent(), me.getX(), me.getY());
			    }
				/*if(e.getClickCount() == 2 && e.getButton() == MouseEvent.BUTTON1) {
					try{
						JToggleButton selecImage = (JToggleButton)e.getSource();
						curImgIndex = scrnsList.indexOf(selecImage.getToolTipText());
						Image bImg = ImageIO.read(new File(new File(repoPath + "\\" + selecImage.getToolTipText()).getAbsolutePath())).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
						ImageIcon newImc = new ImageIcon(bImg);
					    imgLabel.setIcon(newImc);
					    BufferedImage backgroundImg = new BufferedImage(820, 460, BufferedImage.TYPE_INT_ARGB);
						Graphics  g = backgroundImg.getGraphics();
						g.drawImage(bImg, 0, 0, null);
						g.dispose();
						if(((JToggleButton)imagesPanel.getComponent(curImgIndex)).isSelected()) {
							selectImgBtn.setSelected(true);
						}
						else{
							selectImgBtn.setSelected(false);
						}
						lblSelected.setText(selectedImgCount + " Selected");
						lblviewerSelected.setText(selectedImgCount + " Selected");
						imageCount.setText((curImgIndex + 1) + " of " + scrnsList.size() );
						imgViewerfrm.setVisible(true);
					}
					catch(Exception ex){
						ex.printStackTrace();
					}
				}*/
				}
				catch(Exception ex) {
					
				}
			}
        };
        userFiltercmbo = new JComboBox<String>();
		DefaultComboBoxModel<String> userIDmodel = new DefaultComboBoxModel<String>(usersList.toArray(new String[usersList.size()]));
		userFiltercmbo.setModel(userIDmodel);
		userFiltercmbo.setBounds(442, 44, 115, 27);
		if(usersList.contains(userName)) {
			userFiltercmbo.getModel().setSelectedItem(userName);
		}
		/*userFiltercmbo.addActionListener(new ActionListener() {
			
			@SuppressWarnings("unchecked")
			@Override
			public void actionPerformed(ActionEvent e) {
				int x = 20;
		    	int width = 140;
		    	int height = 100;
		    	int y = 10;
				int eachNode = 0;
				scrnsList.removeAll(scrnsList);
				imagesPanel.removeAll();
				JToggleButton[] pic = new JToggleButton[new File(repoPath).listFiles().length];
		        for(File eachFile : new File(repoPath).listFiles()) {
			        	if(!eachFile.isDirectory() && eachFile.getName().endsWith(".png") && (eachFile.getName().startsWith(((JComboBox<String>)e.getSource()).getSelectedItem().toString()) || ((JComboBox<String>)e.getSource()).getSelectedItem().toString().equals("All"))) {
				        	ImageIcon icon = new ImageIcon(eachFile.getAbsolutePath());
				        	pic[eachNode] = new JToggleButton("");
				        	pic[eachNode].setBounds(x, y, width, height);
				        	pic[eachNode].setToolTipText(eachFile.getName());
				            Image img = icon.getImage();
				            Image newImg = img.getScaledInstance(pic[eachNode].getWidth(), pic[eachNode].getHeight(), Image.SCALE_SMOOTH);
				            ImageIcon newImc = new ImageIcon(newImg);
				            pic[eachNode].setIcon(newImc);
				            pic[eachNode].setVisible(true);
				            pic[eachNode].addMouseListener(imageViewerListener);
				            pic[eachNode].addItemListener(imageStateListener);
				            //pic[eachNode].addActionListener(imgSelectChange);
				            imagesPanel.add(pic[eachNode]);
				            eachNode = eachNode++;
				            scrnsList.add(eachFile.getName());
			        	}
		        }
		        imagesPanel.revalidate();
		        imagesPanel.repaint();
			}
		});*/
		
		JButton btnFilter = new JButton("");
		btnFilter.setToolTipText("Filter");
		btnFilter.setIcon(new ImageIcon(ImageRepo.class.getResource("/res/filter15x15.png")));
		btnFilter.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					datePickerFrom.getModel().setSelected(true);
					datePickerTo.getModel().setSelected(true);
					if(!datePickerFrom.getModel().getValue().toString().trim().equals("") && !datePickerTo.getModel().getValue().toString().trim().equals("")) {
						Date fromDate = new SimpleDateFormat("MM/dd/yyyy HH mm ss").parse(new SimpleDateFormat("MM/dd/yyyy").format(datePickerFrom.getModel().getValue()) + " 00 00 00");
						Date toDate = new SimpleDateFormat("MM/dd/yyyy HH mm ss").parse(new SimpleDateFormat("MM/dd/yyyy").format(datePickerTo.getModel().getValue()) + " 23 59 59");
						if(toDate.compareTo(fromDate) > 0 || toDate.compareTo(fromDate) == 0) {
							int x = 20;
					    	int width = 140;
					    	int height = 100;
					    	int y = 10;
							int eachNode = 0;
							scrnsList.removeAll(scrnsList);
							imagesPanel.removeAll();
							JToggleButton[] pic = new JToggleButton[new File(repoPath).listFiles().length];
							for(File eachFile : new File(repoPath).listFiles()) {
								if(eachFile.isFile() && eachFile.getName().endsWith(".png")){
								String fileName = eachFile.getName();
								Date curDate = new SimpleDateFormat("MM_dd_yy_HH_mm_ss").parse(fileName.substring(fileName.indexOf("_") + 1).replace(".png", ""));
								if((!curDate.before(fromDate) && !curDate.after(toDate)) && (eachFile.getName().startsWith(userFiltercmbo.getSelectedItem().toString()) || userFiltercmbo.getSelectedItem().toString().equalsIgnoreCase("All"))) {
									ImageIcon icon = new ImageIcon(eachFile.getAbsolutePath());
						        	pic[eachNode] = new JToggleButton("");
						        	pic[eachNode].setBounds(x, y, width, height);
						        	pic[eachNode].setToolTipText(eachFile.getName());
						        	pic[eachNode].setName(repoCombo.getSelectedItem().toString());
						            Image img = icon.getImage();
						            Image newImg = img.getScaledInstance(pic[eachNode].getWidth(), pic[eachNode].getHeight(), Image.SCALE_SMOOTH);
						            ImageIcon newImc = new ImageIcon(newImg);
						            pic[eachNode].setIcon(newImc);
						            pic[eachNode].setVisible(true);
						            pic[eachNode].addMouseListener(imageViewerListener);
						            pic[eachNode].addItemListener(imageStateListener);
						            //pic[eachNode].addActionListener(imgSelectChange);
						            imagesPanel.add(pic[eachNode]);
						            eachNode++;
						            scrnsList.add(eachFile.getName());
								}
								}
							}
							selectedImgCount = 0;
							lblSelected.setText("0 Selected");
							imagesPanel.revalidate();
					        imagesPanel.repaint();
						}
					}
				}
				catch(Exception e) {
					
				}
			}
		});
		btnFilter.setBounds(561, 43, 27, 27);
		repoPanel.add(btnFilter);
		repoPanel.add(userFiltercmbo);
		
		Properties p = new Properties();
		p.put("text.today", "Today");
		p.put("text.month", "Month");
		p.put("text.year", "Year");
		UtilDateModel model = new UtilDateModel();
		model.setDate(Integer.parseInt(new SimpleDateFormat("yyyy").format(new Date())), new Date().getMonth(), new Date().getDate());
		model.setSelected(true);
		JDatePanelImpl datePanel = new JDatePanelImpl(model, p);
		datePickerFrom = new JDatePickerImpl(datePanel, new AbstractFormatter () {
			private String datePattern = "MM/dd/yyyy";
		    private SimpleDateFormat dateFormatter = new SimpleDateFormat(datePattern);
			@Override
			public String valueToString(Object value) throws ParseException {
				if (value != null) {
	                Calendar cal = (Calendar) value;
	                return dateFormatter.format(cal.getTime());
	            }

	            return "";
			}
			
			@Override
			public Object stringToValue(String text) throws ParseException {
				 return dateFormatter.parseObject(text);

			}
		});
		SpringLayout sl_datePickerFrom = (SpringLayout) datePickerFrom.getLayout();
		sl_datePickerFrom.putConstraint(SpringLayout.WEST, datePickerFrom.getJFormattedTextField(), 0, SpringLayout.WEST, datePickerFrom);
		datePickerFrom.setBounds(194, 44, 120, 27);
		repoPanel.add(datePickerFrom);
		
		
		
		UtilDateModel tomodel = new UtilDateModel();
		tomodel.setDate(Integer.parseInt(new SimpleDateFormat("yyyy").format(new Date())), new Date().getMonth(), new Date().getDate());
		tomodel.setSelected(true);
		JDatePanelImpl to_datePanel = new JDatePanelImpl(tomodel, p);
		datePickerTo = new JDatePickerImpl(to_datePanel, new AbstractFormatter () {
			private String datePattern = "MM/dd/yyyy";
		    private SimpleDateFormat dateFormatter = new SimpleDateFormat(datePattern);
			@Override
			public String valueToString(Object value) throws ParseException {
				if (value != null) {
	                Calendar cal = (Calendar) value;
	                return dateFormatter.format(cal.getTime());
	            }
	            return "";
			}
			
			@Override
			public Object stringToValue(String text) throws ParseException {
				 return dateFormatter.parseObject(text);

			}
		});
		SpringLayout sl_datePickerTo = (SpringLayout) datePickerTo.getLayout();
		sl_datePickerTo.putConstraint(SpringLayout.WEST, datePickerTo.getJFormattedTextField(), 0, SpringLayout.WEST, datePickerTo);
		datePickerTo.setBounds(318, 44, 120, 27);
		repoPanel.add(datePickerTo);
		
		JLabel lblFrom = new JLabel("From:");
		lblFrom.setFont(new Font("SansSerif", Font.PLAIN, 10));
		lblFrom.setBounds(199, 30, 55, 16);
		repoPanel.add(lblFrom);
		
		JLabel lblTo = new JLabel("To:");
		lblTo.setFont(new Font("SansSerif", Font.PLAIN, 10));
		lblTo.setBounds(322, 29, 55, 16);
		repoPanel.add(lblTo);
		
		JLabel lblUserId = new JLabel("User ID:");
		lblUserId.setFont(new Font("SansSerif", Font.PLAIN, 10));
		lblUserId.setBounds(448, 29, 55, 16);
		repoPanel.add(lblUserId);
		
		JPanel titlePanel = new JPanel();
		titlePanel.setBorder(new TitledBorder(new LineBorder(new Color(192, 192, 192), 1, true), "Filter", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		titlePanel.setBounds(185, 17, 412, 61);
		repoPanel.add(titlePanel);
		
		JPanel repotitlePanel = new JPanel();
		repotitlePanel.setBorder(new TitledBorder(new LineBorder(new Color(192, 192, 192), 1, true), "Repo", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		repotitlePanel.setBounds(0, 17, 182, 61);
		repoPanel.add(repotitlePanel);
		repotitlePanel.setLayout(null);
		
		DefaultComboBoxModel<String> repoComboModel = new DefaultComboBoxModel<String>(repoList.toArray(new String[usersList.size()]));
		repoCombo = new JComboBox<String>();
		repoCombo.setModel(repoComboModel);
		repoCombo.setBounds(6, 29, 170, 27);
		repoCombo.addActionListener(new ActionListener() {
			
			@SuppressWarnings("unchecked")
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
				if(((JComboBox<String>)e.getSource()).getSelectedItem().toString().equals("<Default>")) {
					repoPath = projectpath + "\\Snapshot Repo";
				}
				else {
					repoPath = projectpath + "\\Snapshot Repo\\" + ((JComboBox<String>)e.getSource()).getSelectedItem().toString();
				}
				int x = 20;
		    	int width = 140;
		    	int height = 100;
		    	int y = 10;
				int eachNode = 0;
				scrnsList.removeAll(scrnsList);
				imagesPanel.removeAll();
				userFiltercmbo.setSelectedItem(userName);
				datePickerFrom.getModel().setSelected(true);
				datePickerTo.getModel().setSelected(true);
				if(!datePickerFrom.getModel().getValue().toString().trim().equals("") && !datePickerTo.getModel().getValue().toString().trim().equals("")) {
					Date fromDate = new SimpleDateFormat("MM/dd/yyyy HH mm ss").parse(new SimpleDateFormat("MM/dd/yyyy").format(datePickerFrom.getModel().getValue()) + " 00 00 00");
					Date toDate = new SimpleDateFormat("MM/dd/yyyy HH mm ss").parse(new SimpleDateFormat("MM/dd/yyyy").format(datePickerTo.getModel().getValue()) + " 23 59 59");
					if(toDate.compareTo(fromDate) > 0 || toDate.compareTo(fromDate) == 0) {
						
						JToggleButton[] pic = new JToggleButton[new File(repoPath).listFiles().length];
				        for(File eachFile : new File(repoPath).listFiles()) {
				        	if(!eachFile.isDirectory() && eachFile.getName().endsWith(".png")) {
				        		String fileName = eachFile.getName();
								Date curDate = new SimpleDateFormat("MM_dd_yy_HH_mm_ss").parse(fileName.substring(fileName.indexOf("_") + 1).replace(".png", ""));
								if((!curDate.before(fromDate) && !curDate.after(toDate)) && (eachFile.getName().startsWith(userFiltercmbo.getSelectedItem().toString()) || userFiltercmbo.getSelectedItem().toString().equalsIgnoreCase("All"))) {
						        	ImageIcon icon = new ImageIcon(eachFile.getAbsolutePath());
						        	pic[eachNode] = new JToggleButton("");
						        	pic[eachNode].setBounds(x, y, width, height);
						        	pic[eachNode].setToolTipText(eachFile.getName());
						        	pic[eachNode].setName(repoCombo.getSelectedItem().toString());
						            Image img = icon.getImage();
						            Image newImg = img.getScaledInstance(pic[eachNode].getWidth(), pic[eachNode].getHeight(), Image.SCALE_SMOOTH);
						            ImageIcon newImc = new ImageIcon(newImg);
						            pic[eachNode].setIcon(newImc);
						            pic[eachNode].setVisible(true);
						            pic[eachNode].addMouseListener(imageViewerListener);
						            pic[eachNode].addItemListener(imageStateListener);
						            //pic[eachNode].addActionListener(imgSelectChange);
						            imagesPanel.add(pic[eachNode]);
						            eachNode++;
						            scrnsList.add(eachFile.getName());
								}
				        	}
			        	}
					}
		        }

				selectedImgCount = 0;
				lblSelected.setText("0 Selected");
		        imagesPanel.revalidate();
		        imagesPanel.repaint();
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
			}
		});
		repotitlePanel.add(repoCombo);
		
		JLabel lblSelectRepository = new JLabel("Select Repository: ");
		lblSelectRepository.setFont(new Font("SansSerif", Font.PLAIN, 10));
		lblSelectRepository.setBounds(12, 15, 89, 16);
		repotitlePanel.add(lblSelectRepository);
		
		JPanel stepsimgPanel = new JPanel();
		stepsimgPanel.setBounds(634, 0, 297, 619);
		frmImportFromRepo.getContentPane().add(stepsimgPanel);
		stepsimgPanel.setLayout(null);
		
		JPanel stepsListPanel = new JPanel();
		stepsListPanel.setLayout(new WrapFlowLayout());
		JScrollPane stepsscrollPane = new JScrollPane(stepsListPanel);
		stepsscrollPane.setBounds(2, 75, 297, 520);
		stepsimgPanel.add(stepsscrollPane);
		
		JButton btnAddToTest = new JButton("Add to Test");
		btnAddToTest.setBounds(206, 44, 91, 27);
		stepsimgPanel.add(btnAddToTest);
		btnAddToTest.addActionListener(new ActionListener() {
			@SuppressWarnings("unchecked")
			public void actionPerformed(ActionEvent arg0) {
				Component[] arrayOfComponent;
		        int j = (arrayOfComponent = stepsListPanel.getComponents()).length;
		        for (int i = 0; i < j; i++)
		        {
		          Component eachComp = arrayOfComponent[i];
		          if (JPanel.class.isInstance(eachComp))
		          {
		            JPanel imgFinPanel = (JPanel)eachComp;
		            JLabel imgLbl = (JLabel)imgFinPanel.getComponent(1);
		            JComboBox<String> imgStepCm = (JComboBox<String>)imgFinPanel.getComponent(2);
		            String fromPath = projectpath + "\\Snapshot Repo";
		            if(!imgLbl.getName().equals("<Default>")) {
		            	fromPath = projectpath + "\\Snapshot Repo\\" + imgLbl.getName();
		            }
		            if (imgStepCm.getSelectedItem().toString().equalsIgnoreCase("general"))
		            {
		              String fileName = "Screenshot_" + ImageRepo.this.curScreenNum + "_" + new SimpleDateFormat("HH_mm_ss").format(new Date()) + ".png";
		              if (new File(fromPath + "\\" + imgLbl.getToolTipText()).exists())
		              {
		                try
		                {
		                  Files.copy(new File(fromPath + "\\" + imgLbl.getToolTipText()).toPath(), new File(tcPath + "\\Screenshots\\" + fileName).toPath());
		                }
		                catch (IOException localIOException) {}
		                ImageRepo.this.addGeneralScreenshot(fileName, "");
		                ImageRepo.this.curScreenNum += 1;
		              }
		            }
		            else
		            {
		              String fileName = imgStepCm.getSelectedItem().toString().replace(" ", "_") + "_" + new SimpleDateFormat("HH_mm_ss").format(new Date()) + ".png";
		              if (new File(fromPath + "\\" + imgLbl.getToolTipText()).exists()) {
		                try
		                {
		                  Files.copy(new File(fromPath + "\\" + imgLbl.getToolTipText()).toPath(), new File(tcPath + "\\Screenshots\\" + fileName).toPath());
		                }
		                catch (IOException localIOException1) {}
		              }
		              ImageRepo.setStepDetails(Integer.parseInt(imgStepCm.getSelectedItem().toString().split(" ")[1]) - 1, fileName);
		            }
		          }
		        }
		        frmImportFromRepo.dispose();
			}
		});
		
		JButton selectAllBtn = new JButton("");
		selectAllBtn.setToolTipText("Select All");
		selectAllBtn.setIcon(new ImageIcon(ImageRepo.class.getResource("/res/selectal16x16l.png")));
		selectAllBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				for(Component eachComp : stepsListPanel.getComponents()) {
					if(JPanel.class.isInstance(eachComp)) {
						JPanel imgFinPanel = (JPanel) eachComp;
						JCheckBox chkBox = (JCheckBox)imgFinPanel.getComponent(0);
						chkBox.setSelected(true);
					}
				}
			}
		});
		selectAllBtn.setBounds(6, 45, 27, 27);
		stepsimgPanel.add(selectAllBtn);
		
		JButton deleteBtn = new JButton("");
		deleteBtn.setToolTipText("Delete");
		deleteBtn.setIcon(new ImageIcon(ImageRepo.class.getResource("/res/delete-15x15.png")));
		deleteBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				for(Component eachComp : stepsListPanel.getComponents()) {
					if(JPanel.class.isInstance(eachComp)) {
						JPanel imgFinPanel = (JPanel) eachComp;
						JCheckBox chkBox = (JCheckBox)imgFinPanel.getComponent(0);
						if(chkBox.isSelected()) {
							imgFinPanel.getParent().remove(imgFinPanel);
						}
					}
				}
				stepsListPanel.revalidate();
				stepsListPanel.repaint();
			}
		});
		deleteBtn.setBounds(35, 45, 27, 27);
		stepsimgPanel.add(deleteBtn);
		
		
		JButton button = new JButton(">");
		button.setFont(new Font("SansSerif", Font.BOLD, 12));
		button.setBounds(601, 274, 35, 25);
		frmImportFromRepo.getContentPane().add(button);
		frmImportFromRepo.setSize(940, 620);
		
		int x = 20;
    	int width = 140;
    	int height = 100;
    	int y = 10;
		int eachNode = 0;
		datePickerFrom.getModel().setSelected(true);
		datePickerTo.getModel().setSelected(true);
		if(!datePickerFrom.getModel().getValue().toString().trim().equals("") && !datePickerTo.getModel().getValue().toString().trim().equals("")) {
			Date fromDate = new SimpleDateFormat("MM/dd/yyyy HH mm ss").parse(new SimpleDateFormat("MM/dd/yyyy").format(datePickerFrom.getModel().getValue()) + " 00 00 00");
			Date toDate = new SimpleDateFormat("MM/dd/yyyy HH mm ss").parse(new SimpleDateFormat("MM/dd/yyyy").format(datePickerTo.getModel().getValue()) + " 23 59 59");
			if(toDate.compareTo(fromDate) > 0 || toDate.compareTo(fromDate) == 0) {
				ArrayList<String> snapsList = new ArrayList<String>(); 
				JToggleButton[] pic = new JToggleButton[new File(repoPath).listFiles().length];
		        for(File eachFile : new File(repoPath).listFiles()) {
		        	if(!eachFile.isDirectory() && eachFile.getName().endsWith(".png")) {
		        		String fileName = eachFile.getName();
						Date curDate = new SimpleDateFormat("MM_dd_yy_HH_mm_ss").parse(fileName.substring(fileName.indexOf("_") + 1).replace(".png", ""));
						if((!curDate.before(fromDate) && !curDate.after(toDate)) && (eachFile.getName().startsWith(userFiltercmbo.getSelectedItem().toString()) || userFiltercmbo.getSelectedItem().toString().equalsIgnoreCase("All"))) {
				        	ImageIcon icon = new ImageIcon(eachFile.getAbsolutePath());
				        	pic[eachNode] = new JToggleButton("");
				        	pic[eachNode].setBounds(x, y, width, height);
				        	pic[eachNode].setToolTipText(eachFile.getName());
				        	pic[eachNode].setName(repoCombo.getSelectedItem().toString());
				            Image img = icon.getImage();
				            Image newImg = img.getScaledInstance(pic[eachNode].getWidth(), pic[eachNode].getHeight(), Image.SCALE_SMOOTH);
				            ImageIcon newImc = new ImageIcon(newImg);
				            pic[eachNode].setIcon(newImc);
				            pic[eachNode].setVisible(true);
				            pic[eachNode].addMouseListener(imageViewerListener);
				            pic[eachNode].addItemListener(imageStateListener);
				            imagesPanel.add(pic[eachNode]);
				            eachNode++;
				            snapsList.add(eachFile.getName());
						}
		        	}
		        }
		     }
		 }
        
        button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				List<JComponent> imagepanels = new ArrayList<JComponent>();
				for(Component eachComp :  imagesPanel.getComponents()) {
					JToggleButton eachImage = (JToggleButton) eachComp;
					if(eachImage.isSelected()) {
						if(new File(repoPath + "\\" + eachImage.getToolTipText()).exists()) {
							JPanel eachimagepanel = new JPanel();
							eachimagepanel.setSize(270, 115);
							JCheckBox checkBox = new JCheckBox();
							checkBox.setSize(20, 20);
							eachimagepanel.add(checkBox);
							ImageIcon icon = new ImageIcon(new File(repoPath + "\\" + eachImage.getToolTipText()).getAbsolutePath());
				        	JLabel eachLabelImg = new JLabel("");
				        	eachLabelImg.setName(eachImage.getName());
				        	eachLabelImg.setSize(width, height);
				        	eachLabelImg.setToolTipText(new File(repoPath + "\\" + eachImage.getToolTipText()).getName());
				            Image img = icon.getImage();
				            Image newImg = img.getScaledInstance(eachLabelImg.getWidth(), eachLabelImg.getHeight(), Image.SCALE_SMOOTH);
				            ImageIcon newImc = new ImageIcon(newImg);
				            eachLabelImg.setIcon(newImc);
				            eachLabelImg.setVisible(true);
				            eachimagepanel.add(eachLabelImg);
				            JComboBox<String> stepComb = new JComboBox<String>(steps.toArray(new String[steps.size()]));
				            stepComb.setPrototypeDisplayValue("General num");
				            stepComb.setSize(70, 20);
				            eachimagepanel.add(stepComb);
				            imagepanels.add(eachimagepanel);
				            eachImage.setSelected(false);
						}
					}
				}
				for(Component eachPanel : imagepanels) {
					stepsListPanel.add(eachPanel);
				}
				stepsListPanel.revalidate();
				stepsListPanel.repaint();
			}
		});
        
        
        /*Image Viewer*/
        imgViewerfrm.setTitle("Repo Image Viewer");
		imgViewerfrm.setResizable(false);
		imgViewerfrm.setSize(856, 599);
		imgViewerfrm.setLocation(dim.width/2- imgViewerfrm.getSize().width/2, dim.height/2-imgViewerfrm.getSize().height/2);
		imgViewerfrm.getContentPane().setLayout(null);
		
		
		imgPanel.setBounds(10, 37, 822, 502);
		imgViewerfrm.getContentPane().add(imgPanel);
		imgPanel.add(imgLabel);
		
		
		lblviewerSelected.setHorizontalAlignment(SwingConstants.RIGHT);
		lblviewerSelected.setBounds(757, 12, 75, 14);
		imgViewerfrm.getContentPane().add(lblviewerSelected);
		
		ItemListener selectImg = new ItemListener() {

			@Override
			public void itemStateChanged(ItemEvent e) {
				if(JToggleButton.class.isInstance(imagesPanel.getComponent(curImgIndex))) {
					if(e.getStateChange() == ItemEvent.SELECTED) {
						((JToggleButton)imagesPanel.getComponent(curImgIndex)).setSelected(true);
					}
					else{
						((JToggleButton)imagesPanel.getComponent(curImgIndex)).setSelected(false);
					}
					lblviewerSelected.setText(selectedImgCount + " Selected");
					lblSelected.setText(selectedImgCount + " Selected");
				}
			}
		};
		
		
		selectImgBtn.setBounds(10, 8, 27, 23);
		selectImgBtn.addItemListener(selectImg);
		selectImgBtn.setIcon(new ImageIcon(ImageRepo.class.getResource("/res/select16x16.png")));
		imgViewerfrm.getContentPane().add(selectImgBtn);
		
		
		
		imageCount.setBounds(410, 530, 100, 50);
		imgViewerfrm.getContentPane().add(imageCount);
		
		Action nextBtnAction = new AbstractAction() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try{
					if(curImgIndex < scrnsList.size() - 1) {
						curImgIndex++;
						Image bImg = ImageIO.read(new File(new File(repoPath + "\\" + scrnsList.get(curImgIndex)).getAbsolutePath())).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
						ImageIcon newImc = new ImageIcon(bImg);
					    imgLabel.setIcon(newImc);
					    BufferedImage backgroundImg = new BufferedImage(820, 460, BufferedImage.TYPE_INT_ARGB);
						Graphics  g = backgroundImg.getGraphics();
						g.drawImage(bImg, 0, 0, null);
						g.dispose();
						if(((JToggleButton)imagesPanel.getComponent(curImgIndex)).isSelected()) {
							selectImgBtn.setSelected(true);
						}
						else{
							selectImgBtn.setSelected(false);
						}
						selectImgBtn.revalidate();
						selectImgBtn.repaint();
						imageCount.setText((curImgIndex + 1) + " of " + scrnsList.size() );
					}
				}
				catch(Exception ex){
					
				}
			}
		};
		
		
		Action prevBtnAction = new AbstractAction() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					if(curImgIndex>0) {
						curImgIndex--;
						Image bImg = ImageIO.read(new File(new File(repoPath + "\\" + scrnsList.get(curImgIndex)).getAbsolutePath())).getScaledInstance(820, 460, Image.SCALE_SMOOTH);
						ImageIcon newImc = new ImageIcon(bImg);
					    imgLabel.setIcon(newImc);
					    BufferedImage backgroundImg = new BufferedImage(820, 460, BufferedImage.TYPE_INT_ARGB);
						Graphics  g = backgroundImg.getGraphics();
						g.drawImage(bImg, 0, 0, null);
						g.dispose();
						if(((JToggleButton)imagesPanel.getComponent(curImgIndex)).isSelected()) {
							selectImgBtn.setSelected(true);
						}
						else{
							selectImgBtn.setSelected(false);
						}
						selectImgBtn.revalidate();
						selectImgBtn.repaint();
						imageCount.setText((curImgIndex + 1) + " of " + scrnsList.size() );
					}
				}
				catch(Exception ex) {
					
				}
			}
		};
		
		JButton prevBtn = new JButton("<");
		prevBtn.setBounds(339, 543, 27, 23);
		prevBtn.addActionListener(prevBtnAction);
		imgViewerfrm.getContentPane().add(prevBtn);
		prevBtn.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_LEFT, 0), "prevAction");
		prevBtn.getActionMap().put("prevAction", prevBtnAction);
		
		JButton nexBtn = new JButton(">");
		nexBtn.setBounds(482, 543, 27, 23);
		nexBtn.addActionListener(nextBtnAction);
		imgViewerfrm.getContentPane().add(nexBtn);
        frmImportFromRepo.setVisible(true);
        nexBtn.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_RIGHT, 0), "nextAction");
        nexBtn.getActionMap().put("nextAction", nextBtnAction);
		
	}
	
	@SuppressWarnings("unchecked")
	private static int getStepsCount() {
		int stepCount = 0;
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        List<Node> node = document.selectNodes("/testcase/teststep");
        	stepCount = node.size();
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return stepCount;
	}
	
	public static void setStepDetails(int stepNum, String Screenshotname) {
		try {
			File inputFile = new File(tcPath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectNodes("/testcase/teststep").get(stepNum);
        	Element element = (Element) node;
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
	
	
}
