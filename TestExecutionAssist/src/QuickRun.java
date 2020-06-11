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
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
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
import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JSeparator;
import javax.swing.SwingConstants;
import javax.swing.JTextArea;
import javax.swing.JTextPane;
import javax.swing.KeyStroke;
import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.BorderFactory;
import javax.swing.ImageIcon;

public class QuickRun {
	JButton btnDetails = new JButton("...");
	String repoPath = "";
	String userName = System.getProperty("user.name");
	Border emptyBorder = BorderFactory.createEmptyBorder();
	Properties props = new Properties();
	String projectPath = "";
	String repoName = "";
	private JDialog exeHandle;
	private JLabel lblRepoName;
	public QuickRun(String projectPath, JFrame parentFrame){
		this.projectPath = projectPath;
		try {
			props.load(new FileInputStream(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg"));
			repoPath = projectPath + "\\Snapshot Repo";
			exeHandle = new JDialog();
			exeHandle.setModal(true);
	        GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
	        Rectangle rect = ge.getMaximumWindowBounds();
			//final JFrame exeHandle = new JFrame();
			exeHandle.setBackground(Color.LIGHT_GRAY);
			exeHandle.setUndecorated(true);
			exeHandle.setResizable(false);   
			exeHandle.setBounds(0, 0, 257, 76);
			exeHandle.getContentPane().setLayout(null);
			exeHandle.setAlwaysOnTop(true);
			exeHandle.setModalityType(ModalityType.APPLICATION_MODAL);
			
			JPanel executionPanel = new JPanel();
			executionPanel.setBorder(new LineBorder(new Color(160, 160, 160), 2, true));
			executionPanel.setBackground(SystemColor.menu);
			executionPanel.setBounds(0, 0, 255, 76);
			exeHandle.getContentPane().add(executionPanel);
			executionPanel.setLayout(null);
			
			JButton btnStop = new JButton("");
			btnStop.setToolTipText("Stop");
			btnStop.setIcon(new ImageIcon(QuickRun.class.getResource("/res/Stop30x30.png")));
			btnStop.setBackground(SystemColor.menu);
			btnStop.setBorder(emptyBorder);
			btnStop.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent e) {
					exeHandle.dispose();
					parentFrame.setVisible(true);
				}
			});
			btnStop.setBounds(171, 27, 30, 30);
			executionPanel.add(btnStop);
			
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
			              String scrshotName =  userName  + "_" + new SimpleDateFormat("MM_dd_yy_HH_mm_ss").format(new Date()) +".png";
			              ImageIO.write(rImage, "png", new File(repoPath + "\\" + scrshotName));
			              parentFrame.setVisible(false);
			              exeHandle.setVisible(true);
			              
					} catch (Exception e) {
						e.printStackTrace();
					}
		              
		          }
		      };
		      
	          
			JButton btnPass = new JButton("");
			btnPass.setToolTipText("Capture Screen");
			btnPass.setForeground(new Color(0, 128, 0));
			btnPass.setHorizontalAlignment(SwingConstants.LEFT);
			btnPass.setFont(new Font("Dialog", Font.BOLD, 11));
			btnPass.setBackground(Color.WHITE);
			/*Action pAct = new AbstractAction("") {
				public void actionPerformed(ActionEvent e) {
						exeHandle.setVisible(false);
		                Timer timer = new Timer(100, alPass);
		                timer.setRepeats(false);
		                timer.start();
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
			/*btnPass.getInputMap().put(KeyStroke.getKeyStroke(KeyEvent.VK_C,KeyEvent.CTRL_DOWN_MASK | KeyEvent.ALT_DOWN_MASK), "passAct");
			btnPass.getActionMap().put("passAct", pAct);*/
			
			JLabel lblCaptureImg = new JLabel("");
			lblCaptureImg.setIcon(new ImageIcon(QuickRun.class.getResource("/res/capture45x45.png")));
			lblCaptureImg.setBounds(89, 16, 50, 50);
			executionPanel.add(lblCaptureImg);
			btnPass.setBounds(67, 16, 95, 52);
			executionPanel.add(btnPass);
	
			
			JSeparator separator = new JSeparator();
			separator.setOrientation(SwingConstants.VERTICAL);
			separator.setBackground(UIManager.getColor("ProgressBar.shadow"));
			separator.setForeground(UIManager.getColor("ProgressBar.shadow"));
			separator.setBounds(50, 0, 2, 76);
			executionPanel.add(separator);
			
			JButton btnCollapse = new JButton("");
			btnCollapse.setIcon(new ImageIcon(QuickRun.class.getResource("/res/arrowRight30x30.png")));
			btnCollapse.setBackground(SystemColor.menu);
			btnCollapse.setBorder(emptyBorder);
			btnCollapse.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent e) {
					if(exeHandle.getSize().getWidth()>150) {
						exeHandle.setBounds(0, 0, 64, 76);
						btnCollapse.setIcon(new ImageIcon(QuickRun.class.getResource("/res/arrowLeft30x30.png")));
					}
					else {
						exeHandle.setBounds(0, 0, 255, 76);
						btnCollapse.setIcon(new ImageIcon(QuickRun.class.getResource("/res/arrowRight30x30.png")));
					}
					Dimension windowSize = exeHandle.getSize();
			        int x = (int) rect.getWidth() - (int) windowSize.getWidth();
			        int y = (int) rect.getHeight()/2 - (int) windowSize.getHeight()/2;
			        exeHandle.setLocation(x, y);
				}
			});
			btnCollapse.setBounds(10, 23, 30, 30);
			executionPanel.add(btnCollapse);
			
			JButton button = new JButton("<[]>");
			button.setToolTipText("Switch Repository");
			button.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent arg0) {
					repoPath = projectPath + "\\Snapshot Repo";
					triggerSelectRepo(parentFrame);
				}
			});
			button.setBorder(emptyBorder);
			button.setBackground(SystemColor.activeCaptionBorder);
			button.setBounds(213, 27, 30, 30);
			executionPanel.add(button);
			
			lblRepoName = new JLabel("");
			lblRepoName.setHorizontalAlignment(SwingConstants.LEFT);
			lblRepoName.setFont(new Font("SansSerif", Font.PLAIN, 10));
			lblRepoName.setBounds(56, 2, 145, 16);
			executionPanel.add(lblRepoName);
			
			Dimension windowSize = exeHandle.getSize();
	        int x = (int) rect.getWidth() - (int) windowSize.getWidth();
	        int y = (int) rect.getHeight()/2 - (int) windowSize.getHeight()/2;
	        exeHandle.setLocation(x, y);
			//exeHandle.setVisible(true);
	        triggerSelectRepo(parentFrame);
		}
		catch(Exception  e) {
			e.printStackTrace();
		}
	}
	
	//Function to Pop Select repository frame
	public void triggerSelectRepo(JFrame parentFrame) {
		Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
		parentFrame.setVisible(false);
		exeHandle.setVisible(false);
		JFrame selectRepoFrame = new JFrame();
		selectRepoFrame.setResizable(false);
		selectRepoFrame.setTitle("Select Repository");
		selectRepoFrame.getContentPane().setLayout(null);
		selectRepoFrame.setSize(440, 150);
		selectRepoFrame.setLocation(dim.width/2- selectRepoFrame.getSize().width/2, dim.height/2-selectRepoFrame.getSize().height/2);
		
		JPanel repoPanel = new JPanel();
		repoPanel.setSize(435, 125);
		selectRepoFrame.getContentPane().add(repoPanel);
		
		JComboBox<String> repoCmbo = new JComboBox<String>();
		repoCmbo.setBounds(140, 34, 270, 25);
		repoPanel.add(repoCmbo);
		
		DefaultComboBoxModel<String> cmbBoxModel = new DefaultComboBoxModel<String>();
		repoCmbo.setModel(cmbBoxModel);
		
		repoCmbo.addItem("<Default>");
		for(File eachFolder : new File(repoPath).listFiles()) {
			if(eachFolder.isDirectory()) {
				repoCmbo.addItem(eachFolder.getName());
			}
		}
		
		
		JLabel repoLabel = new JLabel("Select Repository");
		repoLabel.setHorizontalAlignment(SwingConstants.TRAILING);
		repoLabel.setBounds(17, 36, 100, 22);
		repoPanel.setLayout(null);
		repoPanel.add(repoLabel);
		
		JButton btnOk = new JButton("Proceed");
		btnOk.setBounds(319, 79, 90, 25);
		repoPanel.add(btnOk);
		btnOk.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				if(!repoCmbo.getSelectedItem().toString().equals("<Default>")) {
					repoName = repoCmbo.getSelectedItem().toString();
					repoPath = projectPath + "\\Snapshot Repo" + "\\" + repoCmbo.getSelectedItem().toString();
				}
				else {
					repoName = "<Default>";
					repoPath = projectPath + "\\Snapshot Repo";
				}
				selectRepoFrame.dispose();
				lblRepoName.setText(repoName);
				exeHandle.setVisible(true);
			}
		});
		
		JButton btnCreateNew = new JButton("");
		btnCreateNew.setBounds(14, 4, 25, 25);
		btnCreateNew.setIcon(new ImageIcon(QuickRun.class.getResource("/res/createRepo16x16.png")));
		repoPanel.add(btnCreateNew);
		btnCreateNew.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String newFolderName = JOptionPane.showInputDialog("Please input new repository name", null);
				if(!newFolderName.trim().equals("")) {
					try{
						new File(repoPath + "\\" + newFolderName).mkdir();
						cmbBoxModel.removeAllElements();
						repoCmbo.addItem("<Default>");
						for(File eachFolder : new File(repoPath).listFiles()) {
							if(eachFolder.isDirectory()) {
								repoCmbo.addItem(eachFolder.getName());
							}
						}
					}
					catch(Exception e) {
						JOptionPane.showMessageDialog(null, "Error in creating repo folder with the given name!");
					}
				}
			}
		});
		
		selectRepoFrame.setVisible(true);
	}
}
