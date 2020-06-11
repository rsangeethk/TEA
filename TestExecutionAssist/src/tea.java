package src;

import java.awt.Color;
import java.awt.Component;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Frame;
import java.awt.Image;
import java.awt.Rectangle;
import java.awt.SystemColor;
import java.awt.Toolkit;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.InputEvent;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.prefs.Preferences;

import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.BorderFactory;
import javax.swing.DefaultComboBoxModel;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.JTextPane;
import javax.swing.JToggleButton;
import javax.swing.JTree;
import javax.swing.KeyStroke;
import javax.swing.RepaintManager;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.Border;
import javax.swing.border.LineBorder;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.event.TreeExpansionEvent;
import javax.swing.event.TreeExpansionListener;
import javax.swing.event.TreeWillExpandListener;
import javax.swing.plaf.TreeUI;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumnModel;
import javax.swing.tree.DefaultTreeCellRenderer;
import javax.swing.tree.ExpandVetoException;
import javax.swing.tree.TreeModel;
import javax.swing.tree.TreePath;
import javax.swing.tree.TreeSelectionModel;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.ArrayUtils;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

import com.mercury.qualitycenter.otaclient.ClassFactory;
import com.mercury.qualitycenter.otaclient.IAttachment;
import com.mercury.qualitycenter.otaclient.IAttachmentFactory;
import com.mercury.qualitycenter.otaclient.IBaseFactory;
import com.mercury.qualitycenter.otaclient.ICustomization;
import com.mercury.qualitycenter.otaclient.ICustomizationList;
import com.mercury.qualitycenter.otaclient.ICustomizationListNode;
import com.mercury.qualitycenter.otaclient.ICustomizationLists;
import com.mercury.qualitycenter.otaclient.IExtendedStorage;
import com.mercury.qualitycenter.otaclient.IList;
import com.mercury.qualitycenter.otaclient.IRun;
import com.mercury.qualitycenter.otaclient.IRunFactory;
import com.mercury.qualitycenter.otaclient.IStep;
import com.mercury.qualitycenter.otaclient.ITDConnection;
import com.mercury.qualitycenter.otaclient.ITDFilter;
import com.mercury.qualitycenter.otaclient.ITSTest;
import com.mercury.qualitycenter.otaclient.ITest;
import com.mercury.qualitycenter.otaclient.ITestSet;
import com.mercury.qualitycenter.otaclient.ITestSetFolder;
import com.mercury.qualitycenter.otaclient.ITestSetTreeManager;

import com4j.Com4jObject;

@SuppressWarnings("serial")
public class tea extends JFrame {
	JPanel contentPane = new JPanel();
	JTree tcTree = new JTree();
	private static String projectPath = "";
	private static String tcpath = "";
	JTable tcTable = new JTable();
	DefaultTableModel testStepsModel = new DefaultTableModel(new String[] {"SN" , "Description", "Expected Result", "Actual Result","Status"}, 0);
	private JPanel screenShtPanel;
	private JScrollPane scrnShtScrlPane;
	private JLabel[] pic;
	private String[] scrns;
	private int curIndex = 0;
	String curScrnSht = "";
	Border emptyBorder = BorderFactory.createEmptyBorder();
	JButton wordLbl = new JButton("");
	JButton pdfLbl = new JButton("");
	String wordReport = "";
	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
	private JLabel lblTestCaseName;
	private JLabel lblDescription;
	private JTextArea lblTCName;
	private JComboBox<String> lblOverallStatus;
	private JComboBox<String> stepStatus = new JComboBox<String>();
	private JLabel lblexecutedres;
	private JLabel lblexecutedDt;
	private JTextPane descPane;
	
	
	JTextArea textAreaPassActual = new JTextArea();
	JTextArea textAreaFailActual = new JTextArea();
	
	
	JComboBox<String> screenType = new JComboBox<String>();
	JButton btnArrange = new JButton("");
	String expandedModule = "";
	JTextField lclNameField = new JTextField();
	Properties props = new Properties();
	Preferences prefs = Preferences.userRoot().node(this.getClass().getName());
	DrawingFrame asPanel = new DrawingFrame();
	ReportPanel reportSumm = null;
	JPanel testCasePanel = new JPanel();
	String copiedPath = "";
	boolean copyComments = false;
	
	JMenu mnTestSuite = new JMenu("Test Suite");
	JMenu mnTestCase = new JMenu("Test Case");
	JMenuItem uploadALMItem = new JMenuItem("Upload to ALM");
	JMenuItem genReportItem = new JMenuItem("Generate Word Report");
	JMenu changeStatusItem = new JMenu("Change Status");
	
	JMenuItem refreshALM = new JMenuItem("Sync with ALM");
	JMenu mntestData = new JMenu("Test Data");
	
	
	public static void main(String[] args) throws Throwable {
		String uilayout = "Nimbus";
		try {
			if(args.length>0) {
				if(new File(args[0]).exists()) {
					projectPath = args[0];
				}
			}
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if (uilayout.equals(info.getName())) {
                    UIManager.setLookAndFeel(info.getClassName());
                    /*if (uilayout.equals("Nimbus")) {
                        tweakNimbusUI();
                    }*/
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            ex.printStackTrace();
        }
		Icon leafico = new ImageIcon(tea.class.getResource("/res/tc-ico12x12.png"));
		UIManager.put("Tree.leafIcon", leafico);
		/*UIManager.put("Tree.collapsedIcon", new ImageIcon(tea.class.getResource("/res/expand-module.gif")));
	    UIManager.put("Tree.expandedIcon", new ImageIcon(tea.class.getResource("/res/collapse-module.gif")));*/
	    UIManager.put("Tree.openIcon", new ImageIcon(tea.class.getResource("/res/ts-ico12x12.png")));
	    UIManager.put("Tree.closedIcon", new ImageIcon(tea.class.getResource("/res/ts-ico12x12.png")));
		tea teassist = new tea();
		teassist.setVisible(true);
	}
	
	public tea() throws IOException, Throwable{
		
		setTitle("Test Execution Assist");
		setIconImage(Toolkit.getDefaultToolkit().getImage(tea.class.getResource("/res/logo48x48.png")));
		props.load(new FileInputStream(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg"));
		

		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setResizable(false);
		getContentPane().setLayout(null);
		setBounds(new Rectangle(940, 600));
		setLocation(dim.width/2- getSize().width/2, dim.height/2-getSize().height/2);
		contentPane = new JPanel();
		contentPane.setAlignmentX(Component.RIGHT_ALIGNMENT);
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JMenuBar menuBar = new JMenuBar();
		menuBar.setBounds(0, 0, 940, 25);
		contentPane.add(menuBar);
		
		JMenu mnFile = new JMenu("File");
		menuBar.add(mnFile);
		
		
		JMenuItem mntmOpenProject = new JMenuItem("Open Project");
		mnFile.add(mntmOpenProject);
		KeyStroke openKS = KeyStroke.getKeyStroke(KeyEvent.VK_O, KeyEvent.CTRL_DOWN_MASK);
		mntmOpenProject.setAccelerator(openKS);
		mntmOpenProject.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = null;
			    try {
			        fileChooser = new JFileChooser();
			    } catch (Exception ex) {}
			    String recentPath = prefs.get("TEA_RECENT_PROJECT_PATH", "");
				fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				if(!recentPath.equals("")) {
					fileChooser.setCurrentDirectory(new File(recentPath));
				}
				int rVal = fileChooser.showOpenDialog(null);
				if (rVal == JFileChooser.APPROVE_OPTION) {
					prefs.put("TEA_RECENT_PROJECT_PATH", fileChooser.getSelectedFile().getAbsolutePath());
					if(!new File(fileChooser.getSelectedFile().getAbsolutePath() + "\\" + "Snapshot Repo").exists()) {
						new File(fileChooser.getSelectedFile().getAbsolutePath() + "\\" + "Snapshot Repo").mkdir();
					}
					if(!new File(fileChooser.getSelectedFile().getAbsolutePath() + "\\" + "Open Project.bat").exists()) {
						Writer writer = null;	
						try {
							String appPath = new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\TEA.exe";
							String projectPath = fileChooser.getSelectedFile().getAbsolutePath();
						    writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(fileChooser.getSelectedFile().getAbsolutePath() + "\\" + "Open Project.bat"), "utf-8"));
						    writer.write("\"" +appPath + "\" \"" + projectPath + "\"");
						} catch (Exception ex) {
						  // report
						} finally {
						   try {writer.close();} catch (Exception ex) {/*ignore*/}
						}
					}
					tcTree.setVisible(false);
					tcTree.setModel(null);
					TreeModel model = new FileTreeModel(new File(fileChooser.getSelectedFile().getAbsolutePath()));
					tcTree.setVisible(true);
					tcTree.setRootVisible(false);
					tcTree.setShowsRootHandles( true );
			        tcTree.putClientProperty( "JTree.lineStyle", "Angled" );
			        tcTree.setModel(model);
			        tcTree.repaint();
			        mntmImportALM.setEnabled(true);
					mntmImportExcel.setEnabled(true);
			        projectPath = fileChooser.getSelectedFile().getAbsolutePath();
			        RepaintManager.currentManager(tcTree).markCompletelyClean(tcTree);
			        
				}
			}
		});
		
		JMenu importMenu = new JMenu("Import From..");
		mnFile.add(importMenu);
		mntmImportALM = new JMenuItem("HP ALM");
		mntmImportALM.setEnabled(false);
		mntmImportALM.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_M, InputEvent.CTRL_MASK));
		mntmImportALM.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					if(!props.getProperty("ALMURL").trim().equals("")) {
						new ImportALM(projectPath, tcTree);
						tcTree.setVisible(false);
						tcTree.setModel(null);
						TreeModel model = new FileTreeModel(new File(projectPath));
						tcTree.setRootVisible(false);
						tcTree.setShowsRootHandles( true );
				        tcTree.putClientProperty( "JTree.lineStyle", "Angled" );
				        tcTree.setModel(model);
				        tcTree.setVisible(true);
				        tcTree.repaint();
				        RepaintManager.currentManager(tcTree).markCompletelyClean(tcTree);
				        
					}
					else {
						JOptionPane.showMessageDialog(null, "ALM Server URL is not available. Please add your ALM URL in Preferences window and try again.");
					}
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
			}
		});
		importMenu.add(mntmImportALM);
		
		mntmImportExcel = new JMenuItem("Excel");
		mntmImportExcel.setEnabled(false);
		KeyStroke importKS = KeyStroke.getKeyStroke(KeyEvent.VK_I, KeyEvent.CTRL_DOWN_MASK);
		mntmImportExcel.setAccelerator(importKS);
		mntmImportExcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new ImportExcel(projectPath, tcTree);
				tcTree.setVisible(false);
				tcTree.setModel(null);
				TreeModel model = new FileTreeModel(new File(projectPath));
				tcTree.setRootVisible(false);
				tcTree.setShowsRootHandles( true );
		        tcTree.putClientProperty( "JTree.lineStyle", "Angled" );
		        tcTree.setModel(model);
		        tcTree.setVisible(true);
		        tcTree.repaint();
		        RepaintManager.currentManager(tcTree).markCompletelyClean(tcTree);
			}
		});
		importMenu.add(mntmImportExcel);
		
		/*JMenu mnEdit = new JMenu("Edit");
		menuBar.add(mnEdit);*/
		
		
		mnTestSuite.setVisible(false);
		menuBar.add(mnTestSuite);
		
		mnTestCase.setVisible(false);
		menuBar.add(mnTestCase);
		
		/*JMenu mnTools = new JMenu("Tools");
		menuBar.add(mnTools);*/
		
		mntmQuickRun = new JMenuItem("Run");
		mntmQuickRun.setEnabled(false);
		//mntmQuickRun.setAction(action);
		mntmQuickRun.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_R, InputEvent.CTRL_MASK));
		mnTestCase.add(mntmQuickRun);
		mntmQuickRun.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				if(!tcpath.equals("")) {
					//setState(Frame.ICONIFIED);
					setVisible(false);
					new Run(tcpath, (JFrame)getRootPane().getTopLevelAncestor());
					//((JFrame)getRootPane().getTopLevelAncestor()).setVisible(true);
				}
				else{
					JOptionPane.showMessageDialog(null, "Please open a testcase to start Run");
				}
				
			}
		});
		
		JMenu mnOtherRun = new JMenu("Other Runs..");
		mnTestCase.add(mnOtherRun);
		
		mntmStepRun = new JMenuItem("Step Run");
		mntmStepRun.setEnabled(false);
		//mntmQuickRun.setAction(action);
		mntmStepRun.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_T, InputEvent.CTRL_MASK));
		mnOtherRun.add(mntmStepRun);
		mntmStepRun.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				if(!tcpath.equals("")) {
					setState(Frame.ICONIFIED);
					setVisible(false);
					new RunDialog(tcpath);
					setVisible(true);
				}
				else{
					JOptionPane.showMessageDialog(null, "Please open a testcase to start Step Run");
				}
				
			}
		});
		
		mntmFreeRun = new JMenuItem("Free Run");
		mntmFreeRun.setEnabled(false);
		//mntmQuickRun.setAction(action);
		mntmFreeRun.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_Q, InputEvent.CTRL_MASK));
		mnOtherRun.add(mntmFreeRun);
		
		
		mntmFreeRun.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				if(!projectPath.equals("")) {
					setState(Frame.ICONIFIED);
					setVisible(false);
					new QuickRun(projectPath, (JFrame)getRootPane().getTopLevelAncestor());
					setVisible(true);
				}
				else{
					JOptionPane.showMessageDialog(null, "Please open a project to start Free Run");
				}
				
			}
		});
		
		JMenu mnOptions = new JMenu("Options");
		menuBar.add(mnOptions);
		
		JMenuItem mntmPreferences = new JMenuItem("Preferences");
		mnOptions.add(mntmPreferences);
		mntmPreferences.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					new Prefs();
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
			}
		});
		
		JMenu mnHelp = new JMenu("Help");
		menuBar.add(mnHelp);
		JTabbedPane explorer = new JTabbedPane(JTabbedPane.TOP);
		explorer.setMinimumSize(new Dimension(240, 5));
		explorer.setBounds(0,27, 250,515);
		//contentPane.add(explorer);
		
		/* Main Panel*/
		JPanel tcExplorerPanel = new JPanel();
		tcExplorerPanel.setMinimumSize(new Dimension(270, 10));
		
		tcExplorerPanel.setBorder(null);
		explorer.addTab("Test Case Explorer", null, tcExplorerPanel, null);
		tcTree.getSelectionModel().setSelectionMode(TreeSelectionModel.DISCONTIGUOUS_TREE_SELECTION);
		tcTree.setCellRenderer(new CustomIconRenderer());
		tcTree.setBorder(null);
		tcTree.setVisible(false);
		
		
		
		TreeUI ui = tcTree.getUI();
		tcTree.addTreeWillExpandListener(new TreeWillExpandListener() {
			@Override
			public void treeWillExpand(TreeExpansionEvent event)
					throws ExpandVetoException {
					tcTree.setUI(null);
					TreePath path = event.getPath();
			        expandedModule =  path.getLastPathComponent().toString();
			}
			
			@Override
			public void treeWillCollapse(TreeExpansionEvent event)
					throws ExpandVetoException {
			}
		});
		
		tcTree.addTreeExpansionListener(new TreeExpansionListener() {
			
			@Override
			public void treeExpanded(TreeExpansionEvent event) {
				tcTree.setUI(ui);
				/*tcTree.revalidate();
				tcTree.repaint();*/
			}
			
			@Override
			public void treeCollapsed(TreeExpansionEvent event) {
				
			}
		});
		tcExplorerPanel.setLayout(null);
		
		JScrollPane scrollPane = new JScrollPane(tcTree);
		scrollPane.setViewportBorder(null);
		//scrollPane.setLayout(null);
		scrollPane.setBounds(0, 5, 930, 485);
		//scrollPane.getVerticalScrollBar().setUnitIncrement(16);
		tcExplorerPanel.add(scrollPane);
        
		ListSelectionListener cellChangeList = new ListSelectionListener() {	
				@Override
				public void valueChanged(ListSelectionEvent e) {
					if (tcTable.getSelectedRow() > -1) {
			            viewtextArea.setText(tcTable.getValueAt(tcTable.getSelectedRow(), tcTable.getSelectedColumn()).toString());
			        }
				}
		};
		Action runAction = new AbstractAction("") {
			public void actionPerformed(ActionEvent e) {
				//setState(Frame.ICONIFIED);
				setVisible(false);
				//new RunDialog(tcpath);
				new Run(tcpath, (JFrame)getRootPane().getTopLevelAncestor());
				//((JFrame)getRootPane().getTopLevelAncestor()).setVisible(true);
			}

		};
		
		JPopupMenu moduleMenu = new JPopupMenu();
		
		moduleMenu.add(refreshALM);
		refreshALM.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				int sync = JOptionPane.YES_OPTION;
				sync = JOptionPane.showConfirmDialog(null, "Synchronizing with ALM will add the updated steps from test plan and testcases from test lab. Are you sure want to continue?");
				if(sync == JOptionPane.YES_OPTION) {
					if(!prefs.get("UserID", "").equals("") && !prefs.get("pass", "").equals("")) {
						String modulePath = tcTree.getSelectionPath().getLastPathComponent().toString();
						if(checkALMconnection(modulePath)) {
							//String moduleName = modulePath.substring(modulePath.lastIndexOf("\\") + 1);
							new createProject().refreshProjectALM(modulePath, prefs.get("UserID", ""), prefs.get("pass", ""));
							
							tcTree.setVisible(false);
							tcTree.setModel(null);
							TreeModel model = new FileTreeModel(new File(projectPath));
							tcTree.setVisible(true);
							tcTree.setRootVisible(false);
							tcTree.setShowsRootHandles( true );
					        tcTree.putClientProperty( "JTree.lineStyle", "Angled" );
					        tcTree.setModel(model);
					        tcTree.repaint();
					        mntmImportALM.setEnabled(true);
							mntmImportExcel.setEnabled(true);
					        RepaintManager.currentManager(tcTree).markCompletelyClean(tcTree);
						}
						else {
							JOptionPane.showMessageDialog(null, "Failed to Authenticate. Please check your ALM credentials in Preferences window");
						}
					}
					else {
						JOptionPane.showMessageDialog(null, "Please set ALM credentials in Preferences window");
					}
				}
			}
		});
		
		
		moduleMenu.add(mntestData);
		
		JMenuItem mntmimportTestData = new JMenuItem("Import Test Data");
		mntestData.add(mntmimportTestData);
		mntmimportTestData.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				String modulePath = tcTree.getSelectionPath().getLastPathComponent().toString();
				if(new File(modulePath + "\\Test Data.xlsx").exists()) {
					importTestData(modulePath);
				}
				else {
					JOptionPane.showMessageDialog(null, "Test data setup not done in Testdata sheet");
				}
			}
		});
		
		
		JMenuItem mntmCreateImportSheet = new JMenuItem("Open Data Import Sheet");
		mntestData.add(mntmCreateImportSheet);
		mntmCreateImportSheet.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				String modulePath = tcTree.getSelectionPath().getLastPathComponent().toString();
				openTestDataSheet(modulePath);
			}
		});
		
		mnTestSuite.add(refreshALM);
		mnTestSuite.add(mntestData);
		
		JPopupMenu treeMenu = new JPopupMenu();
		uploadALMItem.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					File inputFile = new File(tcTree.getSelectionPath().getParentPath().getLastPathComponent().toString() + "\\moduleDetails.ts");
			        SAXReader reader = new SAXReader();
			        Document document = reader.read(inputFile);
			        Node node = (Node) document.selectNodes("/module").get(0);
		        	Element element = (Element) node;
		        	String domainName = element.selectSingleNode("ALMDomain").getText();
		        	String projectName = element.selectSingleNode("ALMProject").getText();
		        	String testlabPath = element.selectSingleNode("TestLab").getText();
					if(!props.getProperty("ALMURL").trim().equals("") && !prefs.get("UserID","").trim().equals("") && !prefs.get("pass", "").trim().equals("") && !domainName.equals("") && !projectName.equals("") && !testlabPath.equals("")) {
						String tcPath = tcTree.getSelectionPath().getLastPathComponent().toString();
						String modulePath = tcPath.substring(0, tcPath.lastIndexOf("\\"));
						if(checkALMconnection(modulePath)) {
							for(TreePath eachPath : tcTree.getSelectionPaths()) {
								if(eachPath.getPathCount() > 2) {
									uploadTC(eachPath.getLastPathComponent().toString(), domainName, projectName, prefs.get("UserID","").trim(), prefs.get("pass", "").trim(), testlabPath, props.getProperty("ALMURL").trim());
								}
							}
						}
						else {
							JOptionPane.showMessageDialog(null, "Failed to Authenticate! Please check the credentials in Preferences window.");
						}
					}
					else {
						JOptionPane.showMessageDialog(null, "Few ALM details are missing. Please enter the ALM details to upload TCs");
					}
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
				
			}
		});
		treeMenu.add(uploadALMItem);
		
		genReportItem.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				for(TreePath eachPath : tcTree.getSelectionPaths()) {
					if(eachPath.getPathCount() > 2) {
						try {
							generateReportWord(eachPath.getLastPathComponent().toString());
						} catch (Exception e1) {
							e1.printStackTrace();
						}
					}
				}
			}
		});
		treeMenu.add(genReportItem);
		
		treeMenu.add(changeStatusItem);
		
		for(String eachStatus: props.getProperty("Status").split(";")) {
			JMenuItem passItem = new JMenuItem(eachStatus);
			changeStatusItem.add(passItem);
			passItem.addActionListener(new ActionListener() {
				@Override
				public void actionPerformed(ActionEvent e) {
					for(TreePath eachPath : tcTree.getSelectionPaths()) {
						if(eachStatus.equalsIgnoreCase("passed") || eachStatus.equalsIgnoreCase("failed")) {
							for(int eachRow = 0; eachRow<tcTable.getRowCount(); eachRow++) {
								String actuals = "";
								if(eachStatus.equalsIgnoreCase("passed") ) {
									actuals =  props.getProperty("PassedActual");
								}
								else if(eachStatus.equalsIgnoreCase("failed")) {
									actuals = props.getProperty("FailedActual");
								}
								setStepDetails(eachPath.getLastPathComponent().toString(), eachRow, eachStatus, "", actuals);
							}
						}
						if(getTestDetails(eachPath.getLastPathComponent().toString(), "executedon").equals("")) {
							setTestDetails(eachPath.getLastPathComponent().toString(), "executedon", new SimpleDateFormat("MM/dd/yyyy HH:mm").format(new Date()));
						}
						else{
							setTestDetails(eachPath.getLastPathComponent().toString(), "executedon", getTestDetails("executedon") + ";" + new SimpleDateFormat("MM/dd/yyyy HH:mm").format(new Date()));
						}
						if(getTestDetails(eachPath.getLastPathComponent().toString(), "executedby").equals("")) {
							setTestDetails(eachPath.getLastPathComponent().toString(), "executedby", System.getProperty("user.name"));
						}
						else{
							setTestDetails(eachPath.getLastPathComponent().toString(), "executedby", getTestDetails("executedby") + ";" + System.getProperty("user.name"));
						}
						setTestDetails(eachPath.getLastPathComponent().toString(), "overallstatus", eachStatus);
					}
				}
			});
		}
		
		JMenuItem copyMenu = new JMenuItem("Copy");
		treeMenu.add(copyMenu);
		
		JMenu pasteMenu = new JMenu("Paste");
		treeMenu.add(pasteMenu);
		JMenuItem scrnmenuItem = new JMenuItem("Screenshot");
		pasteMenu.add(scrnmenuItem);
		
		JMenuItem commmenuItem = new JMenuItem("Screenshot & Comments");
		pasteMenu.add(commmenuItem);
		pasteMenu.setEnabled(false);
		scrnmenuItem.setVisible(false);
		commmenuItem.setVisible(false);
		
		copyMenu.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				TreePath eachPath = tcTree.getSelectionPath();
				copiedPath = eachPath.getLastPathComponent().toString();
				pasteMenu.setEnabled(true);
				scrnmenuItem.setVisible(true);
				commmenuItem.setVisible(true);
			}
		});
		
		scrnmenuItem.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				for(TreePath eachPath : tcTree.getSelectionPaths()) {
					pasteGeneral(copiedPath, eachPath.getLastPathComponent().toString(), false);
				}
				pasteMenu.setEnabled(false);
				scrnmenuItem.setVisible(false);
				commmenuItem.setVisible(false);
			}
		});
		
		commmenuItem.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				for(TreePath eachPath : tcTree.getSelectionPaths()) {
					pasteGeneral(copiedPath, eachPath.getLastPathComponent().toString(), true);
				}
				pasteMenu.setEnabled(false);
				scrnmenuItem.setVisible(false);
				commmenuItem.setVisible(false);
			}
		});
		
		mnTestCase.add(uploadALMItem);
		mnTestCase.add(genReportItem);
		mnTestCase.add(changeStatusItem);
		
		testCasePanel.setBounds(250, 27, 685, 515);
		//contentPane.add(testCasePanel);
		testCasePanel.setLayout(null);
		
		JPanel tcStepspanel = new JPanel();
		tcStepspanel.setBounds(0, 26, 685, 490);
		testCasePanel.add(tcStepspanel);
		tcStepspanel.setLayout(null);
		
		JButton btnExpnd = new JButton("");
		btnExpnd.setIcon(new ImageIcon(tea.class.getResource("/res/up-down15x15.png")));
		btnExpnd.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(btnExpnd.getBounds().getY() == 296)  {
					tabbedPane.setLocation(0, 467);
					tcScrollPane.setSize(685, 465);
					btnExpnd.setLocation(652, 467);
					screenType.setLocation(550, 467);
					btnArrange.setLocation(525, 467);
					tcScrollPane.revalidate();
					tcScrollPane.repaint();
					tcTable.repaint();
				}
				else{
					tabbedPane.setLocation(0, 296);
					tcScrollPane.setSize(685, 294);
					btnExpnd.setLocation(652, 296);
					screenType.setLocation(550, 296);
					btnArrange.setLocation(525, 296);
					tcScrollPane.revalidate();
					tcScrollPane.repaint();
					tcTable.repaint();
				}
			}
		});
		screenType.addActionListener(new ActionListener() {
			
			@SuppressWarnings("unchecked")
			@Override
			public void actionPerformed(ActionEvent e) {
				if(((JComboBox<String>) e.getSource()).getSelectedItem().toString().equalsIgnoreCase("general")) {
					generategeneralCarousel("general");
				}
				else{
					generategeneralCarousel("step");
				}
			}
		});
		
		
		screenType.setVisible(false);
		screenType.setModel(new DefaultComboBoxModel(new String[] {"General", "Step-wise"}));
		screenType.setBounds(550, 296, 100, 25);
		tcStepspanel.add(screenType);
		btnExpnd.setBounds(652, 296, 25, 25);
		tcStepspanel.add(btnExpnd);
		btnArrange.setToolTipText("Re-Arrange Screenshots");
		btnArrange.setIcon(new ImageIcon(tea.class.getResource("/res/reorder16x16.png")));
		
		
		btnArrange.setBounds(525, 296, 25, 25);
		tcStepspanel.add(btnArrange);
		btnArrange.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					new ImageArranger(tcpath);
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
			}
		});
		
        
		tcScrollPane = new JScrollPane(tcTable);
		tcScrollPane.setBounds(0, 0, 685, 296);
		tcStepspanel.add(tcScrollPane);
		tcTable.setModel(testStepsModel);
		
		TextAreaRenderer textAreaRenderer = new TextAreaRenderer();
        TableColumnModel cmodel = tcTable.getColumnModel();
        //cmodel.getColumn(0).setCellRenderer(textAreaRenderer); 
        cmodel.getColumn(1).setCellRenderer(textAreaRenderer); 
        cmodel.getColumn(2).setCellRenderer(textAreaRenderer);
        cmodel.getColumn(3).setCellRenderer(textAreaRenderer);
        cmodel.getColumn(4).setCellRenderer(textAreaRenderer);
        
        JPopupMenu tctableMenu = new JPopupMenu();
		JMenuItem passStep = new JMenuItem("Passed");
		tctableMenu.add(passStep);
		passStep.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				tcTable.getSelectedRows();
				for(int eachRow : tcTable.getSelectedRows()) {
					tcTable.setValueAt("Passed", eachRow, 4);
					tcTable.setValueAt(props.getProperty("PassedActual"), eachRow, 3);
					setStepDetails(tcpath, eachRow, "Passed", "", props.getProperty("PassedActual"));
				}
			}
		});
		JMenuItem failStep = new JMenuItem("Failed");
		tctableMenu.add(failStep);
		failStep.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				tcTable.getSelectedRows();
				for(int eachRow : tcTable.getSelectedRows()) {
					tcTable.setValueAt("Failed", eachRow, 4);
					tcTable.setValueAt(props.getProperty("FailedActual"), eachRow, 3);
					setStepDetails(tcpath, eachRow, "Failed", "", props.getProperty("FailedActual"));
				}
			}
		});
		JMenuItem noRunStep = new JMenuItem("No Run");
		tctableMenu.add(noRunStep);
		noRunStep.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				tcTable.getSelectedRows();
				for(int eachRow : tcTable.getSelectedRows()) {
					tcTable.setValueAt("No Run", eachRow, 4);
					tcTable.setValueAt("", eachRow, 3);
					setStepDetails(tcpath, eachRow, "", "", "");
				}
			}
		});
		
		 
		tcTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		/*tcTable.setFocusable(false);
		tcTable.setRowSelectionAllowed(true);*/
		tcTable.setCellSelectionEnabled(true);
		tcTable.addMouseListener(new MouseAdapter() {
		    public void mousePressed(MouseEvent me) {
		    	if (me.getClickCount() == 1 && SwingUtilities.isRightMouseButton(me)) {
			        tctableMenu.show(me.getComponent(), me.getX(), me.getY());
			    }
		    	else if (me.getClickCount() == 2) {
		        	tabbedPane.setSelectedIndex(4);
			        viewPanel.setFocusable(true);
			        /*int row = tcTable.getSelectedRow();
			        int col = tcTable.getSelectedColumn();*/
			        //tcTable.getSelectionModel().clearSelection();
			        tcTable.editCellAt(-1, -1);
			        viewtextArea.requestFocusInWindow();
		        }
		    }
		});
		tcTable.getSelectionModel().addListSelectionListener(cellChangeList);
		tcTable.getColumnModel().getSelectionModel().addListSelectionListener(cellChangeList);
		
		tabbedPane = new JTabbedPane(JTabbedPane.TOP);
		tabbedPane.addChangeListener(new ChangeListener() {
			public void stateChanged(ChangeEvent e) {
				JTabbedPane pane = (JTabbedPane) e.getSource();
                if(pane.getSelectedIndex() == 2) {
                	screenType.setVisible(true);
                	btnArrange.setVisible(true);
                }
                else{
                	screenType.setVisible(false);
                	btnArrange.setVisible(false);
                }
			}
		});
		
		tabbedPane.setBounds(0, 298, 685, 192);
		tcStepspanel.add(tabbedPane);
		
		JPanel tabDetails = new JPanel();
		tabbedPane.addTab("Details", null, tabDetails, null);
		tabDetails.setLayout(null);
		
		lblTestCaseName = new JLabel("Test Case Name :");
		lblTestCaseName.setHorizontalAlignment(SwingConstants.LEFT);
		lblTestCaseName.setBounds(5, 5, 107, 16);
		tabDetails.add(lblTestCaseName);
		
		lblDescription = new JLabel("Description :");
		lblDescription.setBounds(7, 27, 99, 16);
		tabDetails.add(lblDescription);
		
		descPane = new JTextPane();
		descPane.setBackground(SystemColor.menu);
		JScrollPane descscrollPane = new JScrollPane(descPane);
		descscrollPane.setBounds(7, 48, 373, 92);
		tabDetails.add(descscrollPane);
		
		JLabel lblStatus = new JLabel("Status :");
		lblStatus.setHorizontalAlignment(SwingConstants.RIGHT);
		lblStatus.setBounds(385, 13, 99, 16);
		tabDetails.add(lblStatus);
		
		JLabel lblExecutedBy = new JLabel("Executed by :");
		lblExecutedBy.setHorizontalAlignment(SwingConstants.RIGHT);
		lblExecutedBy.setBounds(385, 52, 99, 16);
		tabDetails.add(lblExecutedBy);
		
		JLabel lblExecutedon = new JLabel("Executed on :");
		lblExecutedon.setHorizontalAlignment(SwingConstants.RIGHT);
		lblExecutedon.setBounds(385, 105, 99, 16);
		tabDetails.add(lblExecutedon);
		
		lblTCName = new JTextArea("");
		lblTCName.setEditable(false);
		lblTCName.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblTCName.setLineWrap(true);
		lblTCName.setWrapStyleWord(true);
		JScrollPane tcNameScroll = new JScrollPane(lblTCName);
		tcNameScroll.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
		tcNameScroll.setBounds(116, 5, 261, 36);
		tabDetails.add(tcNameScroll);
		
		lblOverallStatus = new JComboBox<String>();
		lblOverallStatus.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if(lblOverallStatus.getSelectedItem()!=null && lblOverallStatus.isEnabled()) {
					if(lblOverallStatus.getSelectedItem().toString().equals("Passed")) {
						for(int eachRow = 0; eachRow<tcTable.getRowCount(); eachRow++) {
							tcTable.getModel().setValueAt(props.getProperty("PassedActual"), eachRow, 3);
							tcTable.getModel().setValueAt(lblOverallStatus.getSelectedItem().toString(), eachRow, 4);
							setStepDetails(eachRow, lblOverallStatus.getSelectedItem().toString(), "", props.getProperty("PassedActual"));
						}
					}
					else if(lblOverallStatus.getSelectedItem().toString().equals("Failed")) {
						/*for(int eachRow = 0; eachRow<tcTable.getRowCount(); eachRow++) {
							tcTable.getModel().setValueAt("", eachRow, 3);
							tcTable.getModel().setValueAt("", eachRow, 4);
							setStepDetails(eachRow, "", "", props.getProperty("FailedActual"));
						}*/
					}
					else{
						for(int eachRow = 0; eachRow<tcTable.getRowCount(); eachRow++) {
							tcTable.getModel().setValueAt("", eachRow, 3);
							tcTable.getModel().setValueAt("", eachRow, 4);
							setStepDetails(eachRow, "", "", "");
						}
					}
					if(!lblOverallStatus.getSelectedItem().toString().equals("No Run")) {
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
						/*try {
							if(props.getProperty("AutoReport").equalsIgnoreCase("Yes")) {
								generateReportWord();
							}
						}
						catch(Exception exp) {
							JOptionPane.showMessageDialog(null, "Unable to generate word report.");
						}*/
					}
					setTestDetails("overallstatus", lblOverallStatus.getSelectedItem().toString());
				}
				lblOverallStatus.setEnabled(false);
				editStatus.setSelected(false);
			}
		});
		/*lblOverallStatus.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent arg0) {
				lblOverallStatus.setEnabled(true);
			}
		});*/
		
		lblOverallStatus.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblOverallStatus.setBounds(496, 9, 135, 25);
		lblOverallStatus.setEnabled(false);
		tabDetails.add(lblOverallStatus);
		
		
		lblexecutedres = new JLabel("");
		lblexecutedres.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblexecutedres.setVerticalAlignment(SwingConstants.TOP);
		JScrollPane exeresPane = new JScrollPane(lblexecutedres);
		exeresPane.setBounds(494, 49, 172, 44);
		tabDetails.add(exeresPane);
		
		lblexecutedDt = new JLabel("");
		lblexecutedDt.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblexecutedDt.setVerticalAlignment(SwingConstants.TOP);
		JScrollPane exeonPane = new JScrollPane(lblexecutedDt);
		exeonPane.setBounds(494, 104, 172, 44);
		tabDetails.add(exeonPane);
		
		editStatus = new JToggleButton("");
		editStatus.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if(arg0.getStateChange() == ItemEvent.SELECTED) {
					lblOverallStatus.setEnabled(true);
				}
				else if(arg0.getStateChange() == ItemEvent.DESELECTED){
					lblOverallStatus.setEnabled(false);
				}
			}
		});
		editStatus.setIcon(new ImageIcon(tea.class.getResource("/res/highlight12x12.png")));
		editStatus.setBounds(638, 9, 25, 25);
		tabDetails.add(editStatus);
		
		JLabel lblalmRun = new JLabel("ALM Run:");
		lblalmRun.setBounds(5, 143, 52, 16);
		tabDetails.add(lblalmRun);
		
		lblRunID = new JLabel("");
		lblRunID.setForeground(new Color(0, 0, 0));
		lblRunID.setFont(new Font("SansSerif", Font.PLAIN, 12));
		lblRunID.setBounds(63, 143, 215, 16);
		tabDetails.add(lblRunID);
		
		JPanel datatTab = new JPanel();
		tabbedPane.addTab("Test Data", null, datatTab, null);
		datatTab.setLayout(null);
		
		dataTextArea = new JTextArea();
		dataTextArea.setWrapStyleWord(true);
		dataTextArea.setLineWrap(true);
		dataTextArea.setBorder(new LineBorder(new Color(0, 0, 0)));
		dataTextArea.setBackground(new Color(255, 255, 224));
		
		JScrollPane dataScroll = new JScrollPane(dataTextArea);
		dataScroll.setBounds(0, 20, 680, 142);
		datatTab.add(dataScroll);
		
		JButton btntestDataSave = new JButton("");
		btntestDataSave.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				for(int eachRow = 0; eachRow<tcTable.getRowCount(); eachRow++) {
					setTestDetails("testdata", dataTextArea.getText());
				}
			}
		});
		btntestDataSave.setIcon(new ImageIcon(tea.class.getResource("/res/Save15x15.png")));
		btntestDataSave.setToolTipText("Save");
		btntestDataSave.setBounds(3, -2, 26, 25);
		datatTab.add(btntestDataSave);
		
		
		JPanel screenshotsPanel = new JPanel();
		tabbedPane.addTab("Screenshots", null, screenshotsPanel, null);
		screenshotsPanel.setLayout(null);
		
		screenShtPanel = new JPanel();
		scrnShtScrlPane = new JScrollPane(screenShtPanel);
		screenShtPanel.setLayout(new FlowLayout(FlowLayout.CENTER));
		scrnShtScrlPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_ALWAYS);
		scrnShtScrlPane.setBounds(0, 0, 680, 162);
		screenshotsPanel.add(scrnShtScrlPane);
		
		JPanel reportsPanel = new JPanel();
		tabbedPane.addTab("Reports", null, reportsPanel, null);
		reportsPanel.setLayout(null);
		
		
		wordLbl.setBackground(Color.WHITE);
		wordLbl.setHorizontalAlignment(SwingConstants.CENTER);
		wordLbl.setIcon(new ImageIcon(tea.class.getResource("/res/word120x120.png")));
		wordLbl.setBounds(10, 19, 120, 120);
		wordLbl.setVisible(false);
		reportsPanel.add(wordLbl);
		wordLbl.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				if(!wordReport.equals("") && new File(wordReport).exists()) {
					try {
						Desktop.getDesktop().open(new File(wordReport));
					} catch (IOException e1) {
						JOptionPane.showMessageDialog(null, "Unable to open the report " + wordReport);
					}
				}
			}
		});
		
		pdfLbl.setBackground(Color.WHITE);
		pdfLbl.setHorizontalAlignment(SwingConstants.CENTER);
		pdfLbl.setIcon(new ImageIcon(tea.class.getResource("/res/pdf120x120.png")));
		pdfLbl.setBounds(150, 19, 120, 120);
		pdfLbl.setVisible(false);
		reportsPanel.add(pdfLbl);
		
		
		filesTable = new JTable();
		filesTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		filesTable.setCellSelectionEnabled(true);
		filesTable.setBorder(new LineBorder(new Color(0, 0, 0)));
		filesTable.setModel(filesModel);
		filesTable.setBounds(282, 26, 397, 133);
		filesTable.addMouseListener(new MouseAdapter() {
			  public void mouseClicked(MouseEvent e) {
				  try {
				    if (e.getClickCount() == 2) {
				      JTable target = (JTable)e.getSource();
				      int row = target.getSelectedRow();
				      int column = target.getSelectedColumn();
				      Desktop.getDesktop().open(new File(tcpath + "\\Reports\\" + filesModel.getValueAt(row, column)));
				    }
				  }
				  catch(Exception ex) {
					  ex.printStackTrace();
				  }
			  }
			});
		filesTable.setDropTarget(new DropTarget() {
		    public synchronized void drop(DropTargetDropEvent evt) {
		        try {
		            evt.acceptDrop(DnDConstants.ACTION_COPY);
		            if(!tcpath.equalsIgnoreCase("")) {
			            File reportPath = new File(tcpath + "\\Reports");
			            List<File> droppedFiles = (List<File>)evt.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);
			            int dialogResult = JOptionPane.showConfirmDialog (null, "Would you like to move the files into the reports folder of the selected testcase?");
			            if(dialogResult == JOptionPane.YES_OPTION){
			            	for (File file : droppedFiles) {
				                FileUtils.copyFileToDirectory(file, reportPath);
				            }
			            }
		            }
		        } catch (Exception ex) {
		            ex.printStackTrace();
		        }
		    }
		});
		filesTable.getColumnModel().getColumn(0).setPreferredWidth(390);
		JScrollPane filesScroll = new JScrollPane(filesTable);
		filesScroll.setBackground(Color.WHITE);
		filesScroll.setBounds(282, 26, 397, 133);
		reportsPanel.add(filesScroll);
		
		JButton uploadReportBtn = new JButton("");
		uploadReportBtn.setToolTipText("Attach Files");
		uploadReportBtn.setIcon(new ImageIcon(tea.class.getResource("/res/attach15x15.png")));
		uploadReportBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
			    try {
			    	JFileChooser fileChooser = new JFileChooser();
			        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
			        fileChooser.setMultiSelectionEnabled(true);
					int rVal = fileChooser.showOpenDialog(null);
					if (rVal == JFileChooser.APPROVE_OPTION) {
						if(!tcpath.equalsIgnoreCase("")) {
				            File reportPath = new File(tcpath + "\\Reports");
				            
				            File[] droppedFiles = fileChooser.getSelectedFiles();
				            if(droppedFiles.length>0) {
					            for (File file : droppedFiles) {
						            FileUtils.copyFileToDirectory(file, reportPath);
						        }
				            }
				            else {
				            	FileUtils.copyFileToDirectory(fileChooser.getSelectedFile(), reportPath);
				            }
						}
					}
			    } catch (Exception ex) { ex.printStackTrace();}
			}
		});
		uploadReportBtn.setBounds(282, 0, 27, 27);
		reportsPanel.add(uploadReportBtn);
		
		JButton btnDeleteFiles = new JButton("");
		btnDeleteFiles.setToolTipText("Delete Files");
		btnDeleteFiles.setIcon(new ImageIcon(tea.class.getResource("/res/delete-15x15.png")));
		btnDeleteFiles.setBounds(312, 0, 27, 27);
		reportsPanel.add(btnDeleteFiles);
		btnDeleteFiles.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					int total = filesTable.getSelectedRowCount();
					if(total > 0) {
						if(JOptionPane.showConfirmDialog(null, "Are you sure want to delete the Selected " + total + " files?") == JOptionPane.YES_OPTION) {
							for(int eachRow : filesTable.getSelectedRows()) {
								String fileName = filesTable.getValueAt(eachRow, 0).toString();
								if(new File(tcpath + "\\Reports\\" + fileName).exists()) {
									try {
										new File(tcpath + "\\Reports\\" + fileName).delete();
									}
									catch(Exception ex) {
										
									}
								}	
							}
						}
					}
				}
				catch(Exception exx) {
					exx.printStackTrace();
				}
				
			}
		});
		
		viewPanel = new JPanel();
		tabbedPane.addTab("View/Edit", null, viewPanel, null);
		viewPanel.setLayout(null);
		
		viewtextArea = new JTextArea();
		viewtextArea.setBackground(new Color(255, 255, 224));
		viewtextArea.setBorder(new LineBorder(new Color(0, 0, 0)));
		viewtextArea.setWrapStyleWord(true);
		viewtextArea.setLineWrap(true);
		JScrollPane viewScroll = new JScrollPane(viewtextArea);
		viewScroll.setBounds(0, 20, 680, 142);
		viewtextArea.getDocument().addDocumentListener(new DocumentListener() {
			
			@Override
			public void removeUpdate(DocumentEvent e) {
				int column= tcTable.getSelectedColumn();
				int row = tcTable.getSelectedRow();
				tcTable.getModel().setValueAt(viewtextArea.getText(), row, column);
			}
			
			@Override
			public void insertUpdate(DocumentEvent e) {
				int column= tcTable.getSelectedColumn();
				int row = tcTable.getSelectedRow();
				tcTable.getModel().setValueAt(viewtextArea.getText(), row, column);
			}
			
			@Override
			public void changedUpdate(DocumentEvent e) {
				
			}
		});
		viewPanel.add(viewScroll);
		
		JButton btnActualSave = new JButton("");
		btnActualSave.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				for(int eachRow = 0; eachRow<tcTable.getRowCount(); eachRow++) {
					setStepDetails(tcpath, eachRow, tcTable.getModel().getValueAt(eachRow, 4).toString(), "", tcTable.getModel().getValueAt(eachRow, 3).toString());
				}
			}
		});
		btnActualSave.setIcon(new ImageIcon(tea.class.getResource("/res/Save15x15.png")));
		btnActualSave.setToolTipText("Save");
		btnActualSave.setBounds(3, -2, 26, 25);
		viewPanel.add(btnActualSave);
		
		runbutton = new JButton("");
		runbutton.setEnabled(false);
		runbutton.setToolTipText("Start Execution");
		runbutton.setBounds(5, -2, 28, 28);
		testCasePanel.add(runbutton);
		runbutton.addActionListener(runAction);
		//runAction.putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke(KeyEvent.VK_R, KeyEvent.CTRL_DOWN_MASK));
		runbutton.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_R, KeyEvent.CTRL_DOWN_MASK), "runAction");
		runbutton.getActionMap().put("runAction", runAction);
		runbutton.setIcon(new ImageIcon(tea.class.getResource("/res/playBtn18x18.png")));
		
		refreshbtn = new JButton("");
		refreshbtn.setEnabled(false);
		refreshbtn.setToolTipText("Refresh");
		refreshbtn.setBounds(33, -2, 28, 28);
		testCasePanel.add(refreshbtn);
		refreshbtn.setIcon(new ImageIcon(tea.class.getResource("/res/Refresh18x18.png")));
		
		uploadALMBtn = new JButton("");
		uploadALMBtn.setEnabled(false);
		uploadALMBtn.setToolTipText("Upload to ALM");
		uploadALMBtn.setIcon(new ImageIcon(tea.class.getResource("/res/almUpload30x30.png")));
		uploadALMBtn.setBounds(61, -2, 28, 28);
		testCasePanel.add(uploadALMBtn);
		
		addSnapsRepo = new JButton("");
		addSnapsRepo.setEnabled(false);
		addSnapsRepo.setIcon(new ImageIcon(tea.class.getResource("/res/black10x10.png")));
		addSnapsRepo.setToolTipText("Open Snaps Repo");
		addSnapsRepo.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					new ImageRepo(tcpath, projectPath);
				} catch (Throwable e) {
					e.printStackTrace();
				}
			}
		});
		addSnapsRepo.setBounds(89, -2, 28, 28);
		testCasePanel.add(addSnapsRepo);
		
		reportButton = new JButton("");
		reportButton.setEnabled(false);
		reportButton.setToolTipText("Generate Report");
		reportButton.setIcon(new ImageIcon(tea.class.getResource("/res/word30x30.png")));
		reportButton.setBounds(117, -2, 28, 28);
		testCasePanel.add(reportButton);
		reportButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					reportButton.setEnabled(false);
					generateReportWord(tcpath);
					reportButton.setEnabled(true);
				} catch (Exception e1) {
					reportButton.setEnabled(true);
					e1.printStackTrace();
				}
			}
		});
		
		JLabel lblStatusBar = new JLabel("");
		lblStatusBar.setBounds(0, 0, 940, 20);
		contentPane.add(lblStatusBar);
		lblStatusBar.setBorder(new LineBorder(SystemColor.activeCaptionBorder));
		
		JSplitPane mainPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, explorer, testCasePanel);
		//mainPane.setMinimumSize(arg0);
		mainPane.setResizeWeight(0.15);
		//mainPane.setPreferredSize(new Dimension(1000, 517));
		mainPane.setBounds(0, 27, 935, 522);
		contentPane.add(mainPane);
		
		/*JPanel loader = new JPanel();
		loader.setBounds(250, 27, 685, 515);
		mainPane.add(loader);*/
		uploadALMBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					uploadALMBtn.setEnabled(false);
					File inputFile = new File(tcTree.getSelectionPath().getParentPath().getLastPathComponent().toString() + "\\moduleDetails.ts");
			        SAXReader reader = new SAXReader();
			        Document document = reader.read(inputFile);
			        Node node = (Node) document.selectNodes("/module").get(0);
		        	Element element = (Element) node;
		        	String domainName = element.selectSingleNode("ALMDomain").getText();
		        	String projectName = element.selectSingleNode("ALMProject").getText();
		        	String testlabPath = element.selectSingleNode("TestLab").getText();
					if(!props.getProperty("ALMURL").trim().equals("") && !prefs.get("UserID","").trim().equals("") && !prefs.get("pass", "").trim().equals("") && !domainName.equals("") && !projectName.equals("") && !testlabPath.equals("")) {
						String modulePath = tcpath.substring(0, tcpath.lastIndexOf("\\"));
						if(checkALMconnection(modulePath)) {
							uploadTC(tcpath, domainName, projectName, prefs.get("UserID","").trim(), prefs.get("pass", "").trim(), testlabPath, props.getProperty("ALMURL").trim());
						}
						else{
							JOptionPane.showMessageDialog(null, "Failed to Authenticate! Please check the credentials in Preferences window.");
						}
					}
					else {
						JOptionPane.showMessageDialog(null, "Few ALM details are missing. Please configure ALM details in Preferences Window");
					}
					uploadALMBtn.setEnabled(true);
				}
				catch(Exception ex) {
					ex.printStackTrace();
					uploadALMBtn.setEnabled(true);
				}
			}
		});
		refreshbtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					loadTC();
				} catch (Throwable e1) {
					e1.printStackTrace();
				}
			}
		});
		
		tcTable.getColumnModel().getColumn(0).setPreferredWidth(10);
		tcTable.getColumnModel().getColumn(1).setPreferredWidth(200);
		tcTable.getColumnModel().getColumn(2).setPreferredWidth(200);
		tcTable.getColumnModel().getColumn(3).setPreferredWidth(200);
		tcTable.getColumnModel().getColumn(4).setPreferredWidth(65);
		
		MouseListener ml = new MouseAdapter() {
			public void mousePressed(MouseEvent e) {
				try {
					if (e.getClickCount() == 1 && SwingUtilities.isRightMouseButton(e)) {
						TreePath selPath = tcTree.getPathForLocation(e.getX(), e.getY());
						if(new File(selPath.getLastPathComponent() + "\\moduleDetails.ts").exists()) {
							moduleMenu.show(e.getComponent(), e.getX(), e.getY());
						}
						else {
							treeMenu.show(e.getComponent(), e.getX(), e.getY());
						}
				    }
					else if (e.getClickCount() == 1 && SwingUtilities.isLeftMouseButton(e)) {
						if(tcTree.getPathForLocation(e.getX(), e.getY())!=null) {
						TreePath selPath = tcTree.getPathForLocation(e.getX(), e.getY());
						if(new File(selPath.getLastPathComponent() + "\\moduleDetails.ts").exists()) {
							mnTestSuite.setVisible(true);
							mnTestCase.setVisible(false);
						}
						else {
							mnTestSuite.setVisible(false);
							mnTestCase.setVisible(true);
						}
						}
				    }
					else if (e.getClickCount() == 2) {
						props.load(new FileInputStream(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\config.cfg"));
						TreePath selPath = tcTree.getPathForLocation(e.getX(), e.getY());
						if(new File(selPath.getLastPathComponent() + "\\moduleDetails.ts").exists()) {
							tcStepspanel.setVisible(false);
							runbutton.setVisible(false);
							uploadALMBtn.setVisible(false);
							addSnapsRepo.setVisible(false);
							refreshbtn.setVisible(false);
							reportButton.setVisible(false);
							try {
								if(reportSumm != null) {reportSumm.getParent().remove(reportSumm);reportSumm = null;}
								reportSumm = new ReportPanel(projectPath, tcTree);
								testCasePanel.add(reportSumm);
								reportSumm.setVisible(true);
							} catch (Exception e1) {
								e1.printStackTrace();
							}
						}
						else {
							if(tcpath.equals("")) {
								mnTestCase.setVisible(true);
								mntmFreeRun.setEnabled(true);
								mntmQuickRun.setEnabled(true);
								mntmStepRun.setEnabled(true);
								runbutton.setEnabled(true);
								uploadALMBtn.setEnabled(true);
								addSnapsRepo.setEnabled(true);
								refreshbtn.setEnabled(true);
								reportButton.setEnabled(true);
							}
							tcStepspanel.setVisible(true);
							runbutton.setVisible(true);
							uploadALMBtn.setVisible(true);
							addSnapsRepo.setVisible(true);
							refreshbtn.setVisible(true);
							reportButton.setVisible(true);
							if(reportSumm != null) {reportSumm.setVisible(false);reportSumm.getParent().remove(reportSumm);reportSumm = null;}
							editStatus.setSelected(false);
							File nodeObj = (File)selPath.getLastPathComponent();
							String jTreeVarSelectedPath = nodeObj.getAbsolutePath();
							while (testStepsModel.getRowCount() > 0) {
								testStepsModel.removeRow(0);
							}
							File inputFile = new File(jTreeVarSelectedPath + "\\testDetails.tc");
							if(inputFile.exists()) {
								tcpath = jTreeVarSelectedPath;
								lblOverallStatus.removeAllItems();
								for(String eachStatus : props.getProperty("Status").split(";")) {
									lblOverallStatus.addItem(eachStatus);
								}
								loadTC();
							}
						}
						testCasePanel.revalidate();
						testCasePanel.repaint();
					}
				}
				catch(Exception ex) {
					ex.printStackTrace();
				}
		};
	};
	tcTree.addMouseListener(ml);
	
	
	if(!projectPath.equals("")) {
		if(!new File(projectPath + "\\" + "Snapshot Repo").exists()) {
			new File(projectPath + "\\" + "Snapshot Repo").mkdir();
		}
		tcTree.setVisible(false);
		tcTree.setModel(null);
		TreeModel model = new FileTreeModel(new File(projectPath));
		tcTree.setVisible(true);
		tcTree.setRootVisible(false);
		tcTree.setShowsRootHandles( true );
        tcTree.putClientProperty( "JTree.lineStyle", "Angled" );
        tcTree.setModel(model);
        tcTree.repaint();
        RepaintManager.currentManager(tcTree).markCompletelyClean(tcTree);
	}
	
}
	
	
	
	//function to generate carousel 
    @SuppressWarnings("unchecked")
	public void generategeneralCarousel(String type){
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
        int x = 10;
    	int width = 175;
    	int height = 135;
    	int y = 5;
    	screenShtPanel.removeAll();
    	pic = new JLabel[nodes.size()];
        for(int eachNode =0;eachNode<nodes.size();eachNode++) {
        	Element element = (Element) nodes.get(eachNode);
        	List<Node> scrnNodes = element.selectNodes("screenshot");
        	for(Node scrnNode : scrnNodes) {
	        	if(new File(tcpath + "\\Screenshots\\" + scrnNode.getText()).exists()) {
			        	ImageIcon icon = new ImageIcon(tcpath + "\\Screenshots\\" + scrnNode.getText());
			        	pic[eachNode] = new JLabel("");
			        	pic[eachNode].setBounds(x, y, width, height);
			            Image img = icon.getImage();
			            Image newImg = img.getScaledInstance(pic[eachNode].getWidth(), pic[eachNode].getHeight(), Image.SCALE_SMOOTH);
			            ImageIcon newImc = new ImageIcon(newImg);
			            pic[eachNode].setIcon(newImc);
			            pic[eachNode].setVisible(true);
			            pic[eachNode].addMouseListener(imageEditor);
			            pic[eachNode].setToolTipText(scrnNode.getText());
			            screenShtPanel.add(pic[eachNode]);
			            x = x + width + 10;
	        	}
        	}
        }
        screenShtPanel.revalidate();
        screenShtPanel.repaint();
        scrnShtScrlPane.revalidate();
        scrnShtScrlPane.repaint();
    	}
    	catch(Exception e) {
    		e.printStackTrace();
    	}
        
    }
    
    @SuppressWarnings("unchecked")
	private static String getScreenshots(String type) {
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
		        	value = value + ";" + element.getText();
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
    
    MouseListener imageEditor = new MouseListener() {
		
		@Override
		public void mouseClicked(MouseEvent e) {
			//openImageEditor(((JLabel) e.getComponent()).getToolTipText(), screenType.getSelectedItem().toString());
			try {
				String scrnName = ((JLabel) e.getComponent()).getToolTipText();
					if(scrnName.startsWith("Screenshot")) {
						scrns  = getScreenshots("general").split(";");
						File inputFile = new File(tcpath + "\\testDetails.tc");
				        SAXReader reader = new SAXReader();
				        Document document = reader.read(inputFile);
				        List<Node> nodes = document.selectNodes("/testcase/general");
				        if(nodes.size()!=0) {
							curIndex = ArrayUtils.indexOf(scrns, scrnName);
							ImageEditor editor = new ImageEditor(scrns[curIndex], tcpath, "general");
				        }
				        else{
				        	JOptionPane.showMessageDialog(null, "No screenshots captured for this testcase yet!");
				        }
					}
			        else {
						scrns  = getScreenshots("step").split(";");
						File inputFile = new File(tcpath + "\\testDetails.tc");
				        SAXReader reader = new SAXReader();
				        Document document = reader.read(inputFile);
				        List<Node> nodes = document.selectNodes("/testcase/teststep");
				        if(nodes.size()!=0) {
							curIndex = ArrayUtils.indexOf(scrns, scrnName);
					        /*Element lastEle = (Element) nodes.get(nodes.size() - 1);
					        String lastScrnName = lastEle.selectSingleNode("screenshot").getStringValue();*/
							ImageEditor editor = new ImageEditor(scrns[curIndex], tcpath, "step");
				        }
				        else{
				        	JOptionPane.showMessageDialog(null, "No screenshots captured for this testcase yet!");
				        }
			        }
			}
			catch(Exception ex) {
				ex.printStackTrace();
			}
		}

		@Override
		public void mousePressed(MouseEvent e) {
			
		}

		@Override
		public void mouseReleased(MouseEvent e) {
			
		}

		@Override
		public void mouseEntered(MouseEvent e) {
			
		}

		@Override
		public void mouseExited(MouseEvent e) {
			
		}
	};
    private JScrollPane tcScrollPane;
    private JTabbedPane tabbedPane;
    private JTextArea viewtextArea;
    private JPanel viewPanel;
    private JToggleButton editStatus;
    private JTextArea dataTextArea;
    private JMenuItem mntmImportALM;
    private JMenuItem mntmImportExcel;
    private JButton runbutton;
    private JButton refreshbtn;
    private JButton uploadALMBtn;
    private JButton addSnapsRepo;
    private JButton reportButton;
    private JMenuItem mntmQuickRun;
    private JMenuItem mntmStepRun;
    private JMenuItem mntmFreeRun;
    private JTable filesTable;
    DefaultTableModel filesModel = new DefaultTableModel(new String[] {"File Name"}, 0) { 
    	@Override
        public boolean isCellEditable(int row, int column) {
           return false;
        }
    };
    private JLabel lblRunID;
    //private final Action action = new SwingAction();
    
    
    
    @SuppressWarnings({ "unchecked", "unused" })
	public void loadTC() {
    	try {
    		int rowCount = testStepsModel.getRowCount();
    		for(int eachRow = 0; eachRow<rowCount; eachRow++) {
	        	 testStepsModel.removeRow(0);
	         }
    		File inputFile = new File(tcpath + "\\testDetails.tc");
	         SAXReader reader = new SAXReader();
	         Document document = reader.read( inputFile );
	         List<Node> nodes = document.selectNodes("/testcase");
	         String tcName = ((Element)nodes.get(0)).selectSingleNode("testcasename").getText();
	         lblTCName.setText(tcName);
	         descPane.setText(((Element)nodes.get(0)).selectSingleNode("testdescription").getText());
	         if(((Element)nodes.get(0)).selectSingleNode("overallstatus").getText().trim().equals("") || !((Element)nodes.get(0)).selectSingleNode("overallstatus").getText().trim().equals("No Run")) {
	        	 lblOverallStatus.setSelectedItem(((Element)nodes.get(0)).selectSingleNode("overallstatus").getText());
	         }
	         lblexecutedres.setText("<html>" + ((Element)nodes.get(0)).selectSingleNode("executedby").getText().replace(";", "<br/>") + "</html>");
	         lblexecutedDt.setText("<html>" + ((Element)nodes.get(0)).selectSingleNode("executedon").getText().replace(";", "<br/>") + "</html>");
	         dataTextArea.setText(((Element)nodes.get(0)).selectSingleNode("testdata").getText().trim());
	         if(!((Element)nodes.get(0)).selectSingleNode("report").getText().trim().equals("")) {
	        	 if(new File(tcpath +"\\Reports\\" + ((Element)nodes.get(0)).selectSingleNode("report").getText().trim()).exists()) {
	        		 wordLbl.setVisible(true);
	        		 wordReport = tcpath +"\\Reports\\" + ((Element)nodes.get(0)).selectSingleNode("report").getText().trim();
	        	 } 
	        	 else {
	        		 wordLbl.setVisible(false);
	        	 }
	         }
	         else{
	        	 wordLbl.setVisible(false);
	         }
	         int tcrow = 0;
	         List<Node> tsnodes = document.selectNodes("/testcase/teststep");
	         
	         for (Node node : tsnodes) {
	        	Element element = (Element) node;
	            String tcDesc = element.selectSingleNode("stepdescription").getText();
	            String tcExp = element.selectSingleNode("expectedresult").getText();
	            String tcActual = element.selectSingleNode("actualresult").getText();
	            String tcStatus = element.selectSingleNode("stepstatus").getText();
	            /*JTextArea tcDescLbl = new JTextArea(tcDesc);
	            tcDescLbl.setWrapStyleWord(true);
	            tcDescLbl.setLineWrap(true);
	            tcDescLbl.setMaximumSize(new Dimension(200, 200));*/
	            testStepsModel.addRow(new Object[] { "", tcDesc, tcExp, tcActual, tcStatus});
	            tcrow++;
	         }
	         generategeneralCarousel("general");
	         
	         int rowfilesCount = filesModel.getRowCount();
    		 for(int eachRow = 0; eachRow<rowfilesCount; eachRow++) {
    		 	filesModel.removeRow(0);
	         }
	         File reportsFolder = new File(tcpath + "\\Reports");
	         for(File eachFile : reportsFolder.listFiles()) {
	        	 if(!eachFile.isHidden() && eachFile.isFile()){
	        		 filesModel.addRow(new Object[] {eachFile.getName()});
	        	 }
	         }
	         lblRunID.setText(((Element)nodes.get(0)).selectSingleNode("runid").getText());
			
    	}
    	catch(Exception e) {
    		e.printStackTrace();
    	}
    }
	
	@SuppressWarnings({ "unused", "unchecked" })
	public void generateReportWord() throws Exception{
		  CustomXWPFDocument document= new CustomXWPFDocument();
		  XWPFRun header = document.createParagraph().createRun();
		  File inputFile = new File(tcpath + "\\testDetails.tc");
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
			  XWPFRun expRun = exp.createRun();
			  expRun.setText("Expected Result: ");
			  expRun.setBold(true);
			  XWPFRun expCont = exp.createRun();
			  expCont.setText(element.selectSingleNode("expectedresult").getText());
			  expCont.addBreak();
			  
			  XWPFRun actRun = exp.createRun();
			  actRun.setText("Actual Result: ");
			  actRun.setBold(true);
			  XWPFRun actCont = exp.createRun();
			  actCont.setText(element.selectSingleNode("actualresult").getText());
			  actCont.addBreak();
			  
			  XWPFRun statRun = exp.createRun();
			  statRun.setText("Step Status: ");
			  statRun.setBold(true);
			  XWPFRun statCont = exp.createRun();
			  statCont.setText(element.selectSingleNode("stepstatus").getText());
			  statCont.addBreak();
			  
			  XWPFRun scrnRun = exp.createRun();
			  scrnRun.setText("Screenshots: ");
			  scrnRun.setBold(true);
			  List<Node> scrnNodes = element.selectNodes("screenshot");
			  for(Node EachScrn : scrnNodes) {
				  /*XWPFRun scrnCont = exp.createRun();*/
				  File scrnPath = new File(tcpath + "\\Screenshots\\" + EachScrn.getText());
				  if(scrnPath.exists()) {
					  String blipId = exp.getDocument().addPictureData(new FileInputStream(scrnPath), XWPFDocument.PICTURE_TYPE_PNG);
					  document.createPicture(blipId, document.getNextPicNameNumber(XWPFDocument.PICTURE_TYPE_PNG), 640, 360);
				  }
				  //scrnRun.addBreak();
			  }
			  StepCount++;
		  }
	      //Write the Document in file system
		  File wordDoc= new File(tcpath + "\\Reports\\"  + tcName + "_" + new SimpleDateFormat("hh_mm_ss").format(new Date()) + ".docx");
	      FileOutputStream out = new FileOutputStream(wordDoc);
	      document.write(out);
	      out.close();
	      wordReport = wordDoc.getName();  
	      setTestDetails("report", wordDoc.getName());
	}
	
	@SuppressWarnings({ "unused", "unchecked" })
	public void generateReportWord(String testcasePath) throws Exception{
		  File inputFile = new File(testcasePath + "\\testDetails.tc");
		  SAXReader reader = new SAXReader();
	      Document xmldocument = reader.read(inputFile);
	      Node tcnode = (Node) xmldocument.selectNodes("/testcase").get(0);
	      Element tcelement = (Element) tcnode;
	      String tcName = tcelement.selectSingleNode("testcasename").getText();
		  int confirm = JOptionPane.YES_OPTION;
		  if(!getTestDetails(testcasePath, "report").equals("")){
			  confirm = JOptionPane.showConfirmDialog(null, "Do you want over write the existing report for the testcase " + tcName + " ?");
		  }
		  if(confirm == JOptionPane.YES_OPTION) {
		      String reportPath = testcasePath + "\\Reports\\"  + tcName + "_" + new SimpleDateFormat("hh_mm_ss").format(new Date()) + ".docx";
			  FileUtils.copyFile(new File(new File(tea.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParentFile().getPath() + "\\Report\\cover_template.docx"), new File(reportPath));
			  CustomXWPFDocument document = new CustomXWPFDocument(new FileInputStream(new File(reportPath)));
			  String descData = tcelement.selectSingleNode("testdescription").getText();
			  String statusOverAll = tcelement.selectSingleNode("overallstatus").getText();
			  String testdata = tcelement.selectSingleNode("testdata").getText();
			  replaceText(document, tcName, "", descData, statusOverAll, testdata);
			  List<Node> stepnodes = xmldocument.selectNodes("/testcase/teststep");
			  int StepCount = 1;
			  for(Node eachnode : stepnodes){
				  Element element = (Element) eachnode;
				  XWPFParagraph exp = document.createParagraph();
				  XWPFRun stepNum = null;
				  if(props.getProperty("StepName").equalsIgnoreCase("yes")) {
					  stepNum = exp.createRun();
					  if(StepCount!=1)stepNum.addBreak();
					  stepNum.setText("Step: " + StepCount);
					  stepNum.setBold(true);
					  stepNum.addBreak();
				  }
				  if(props.getProperty("ExpectedResult").equalsIgnoreCase("yes")) {
					  stepNum = exp.createRun();
					  XWPFRun expRun = exp.createRun();
					  expRun.setText("Expected Result: ");
					  expRun.setBold(true);
					  XWPFRun expCont = exp.createRun();
					  expCont.setText(element.selectSingleNode("expectedresult").getText());
					  expCont.addBreak();
				  }
				  if(props.getProperty("ActualResult").equalsIgnoreCase("yes")) {
					  stepNum = exp.createRun();
					  XWPFRun actRun = exp.createRun();
					  actRun.setText("Actual Result: ");
					  actRun.setBold(true);
					  XWPFRun actCont = exp.createRun();
					  actCont.setText(element.selectSingleNode("actualresult").getText());
					  actCont.addBreak();
					  XWPFRun statRun = exp.createRun();
					  statRun.setText("Step Status: ");
					  statRun.setBold(true);
					  XWPFRun statCont = exp.createRun();
					  statCont.setText(element.selectSingleNode("stepstatus").getText());
					  statCont.addBreak();
				  }
				  if(props.getProperty("Screenshots").equalsIgnoreCase("all")) {
					  XWPFRun scrnRun = exp.createRun();
					  scrnRun.setText("Screenshots: ");
					  scrnRun.setBold(true);
					  List<Node> scrnNodes = element.selectNodes("screenshot");
					  for(Node EachScrn : scrnNodes) {
						  XWPFRun scrnCont = exp.createRun();
						  File scrnPath = new File(testcasePath + "\\Screenshots\\" + EachScrn.getText());
						  if(scrnPath.exists() && scrnPath.getName().endsWith("png")) {
							  String blipId = exp.getDocument().addPictureData(new FileInputStream(scrnPath), XWPFDocument.PICTURE_TYPE_PNG);
							  document.createPicture(blipId, document.getNextPicNameNumber(XWPFDocument.PICTURE_TYPE_PNG), 640, 360);
						  }
						  //scrnRun.addBreak();
					  }
				  }
				  StepCount++;
			  }
			  
			  List<Node> screennodes = xmldocument.selectNodes("/testcase/general");
			  for(Node eachnode : screennodes){
				  Element element = (Element) eachnode;
				  XWPFParagraph exp = document.createParagraph();
				  XWPFRun screenRun = exp.createRun();
				  screenRun.setText("Comments : ");
				  screenRun.setBold(true);
				  XWPFRun expCont = exp.createRun();
				  expCont.setText(element.selectSingleNode("comment").getText());
				  expCont.addBreak();
				  Node eachScreen = element.selectSingleNode("screenshot");
				  File scrnPath = new File(testcasePath + "\\Screenshots\\" + eachScreen.getText());
				  if(scrnPath.exists() && scrnPath.getName().endsWith("png")) {
					  String blipId = exp.getDocument().addPictureData(new FileInputStream(scrnPath), XWPFDocument.PICTURE_TYPE_PNG);
					  document.createPicture(blipId, document.getNextPicNameNumber(XWPFDocument.PICTURE_TYPE_PNG), 640, 360);
				  }
			  }
			  
		      //Write the Document in file system
			  //File wordDoc= new File(testcasePath + "\\Reports\\"  + tcName + "_" + new SimpleDateFormat("hh_mm_ss").format(new Date()) + ".docx");
			  File wordDoc= new File(reportPath);
		      FileOutputStream out = new FileOutputStream(wordDoc);
		      document.write(out);
		      out.close();
		      wordReport = wordDoc.getName();  
		      setTestDetails(testcasePath, "report", wordDoc.getName());
		  }
	}
	
	private String getTestDetails(String tagName) {
		String value = "";
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
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
	
	private void setTestDetails(String tagName, String value) {
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
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
	
	
	private String getTestDetails(String testName, String tagName) {
		String value = "";
		try {
			File inputFile = new File(testName + "\\testDetails.tc");
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
	
	private void setTestDetails(String testName, String tagName, String value) {
		try {
			File inputFile = new File(testName + "\\testDetails.tc");
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
	
	public String getStepDetails(int stepNum, String tagName) {
		String tcStepDetails =  "";
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectNodes("/testcase/teststep").get(stepNum);
        	Element element = (Element) node;
        	tcStepDetails = element.selectSingleNode(tagName).getText();
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return tcStepDetails;
	}
	
	public String getgeneralDetails(int stepNum, String tagName) {
		String tcStepDetails =  "";
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectNodes("/testcase/general").get(stepNum);
        	Element element = (Element) node;
        	tcStepDetails = element.selectSingleNode(tagName).getText();
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return tcStepDetails;
	}
	
	
	
	public void uploadTC(String testpath, String domainName, String projectName, String userName, String pass, String testlabPath, String ALMUrl) {
		try{
			String tcName = "\"" + testpath.substring(testpath.lastIndexOf("\\") + 1) + "\"";
			int confirm = JOptionPane.YES_OPTION;
		    if(!getTestDetails(tcpath, "runid").equals("")){
			    confirm = JOptionPane.showConfirmDialog(null, "A Run instance is already available for testcase " + tcName + " Do you want to create a new run?");
		    }
		    if(confirm == JOptionPane.YES_OPTION) {
		    	ITDConnection itdc= ClassFactory.createTDConnection();
		    	try{
			        itdc.initConnectionEx(ALMUrl);
			        itdc.connectProjectEx(domainName, projectName, userName, pass);
		    	}
		    	catch(Exception e) {
		    		JOptionPane.showMessageDialog(null, "Unable to connect to ALM with the credentials provided in the Preferences. Please check!");
		    		return;
		    	}
		        itdc.ignoreHtmlFormat(true);
		        ITestSetTreeManager sTestSetTree = itdc.testSetTreeManager().queryInterface(ITestSetTreeManager.class);
		        String testLabPath = testlabPath;
		        String testLabName = testLabPath.substring(testLabPath.lastIndexOf("\\") + 1);
		        String testLabFolder = testLabPath.substring(0, testLabPath.lastIndexOf("\\"));
		        ITestSetFolder sTestFolder = sTestSetTree.nodeByPath(testLabFolder).queryInterface(ITestSetFolder.class);
		        IList foundList = sTestFolder.findTestSets(null, true, null).queryInterface(IList.class);
		        File inputFile = new File(testpath + "\\testDetails.tc");
		        //String tcName = "\"" + testpath.substring(testpath.lastIndexOf("\\") + 1) + "\"";
		        SAXReader reader = new SAXReader();
		        Document document = null;
				try {
					document = reader.read(inputFile);
				} catch (DocumentException e1) {
					e1.printStackTrace();
				}
		        for(int eachTLcnt = 1; eachTLcnt<=foundList.count(); eachTLcnt++) {
		        	ITestSet tsfac = ((Com4jObject) foundList.item(eachTLcnt)).queryInterface(ITestSet.class);
		        	if(tsfac.name().equals(testLabName)) {
		        		IBaseFactory tsfc = tsfac.tsTestFactory().queryInterface(IBaseFactory.class);
		        		ITDFilter tcFilter = tsfc.filter().queryInterface(ITDFilter.class);
		        		tcFilter.filter("TS_NAME",tcName);
			            IList testList = tcFilter.newList();
			            ITSTest tsTest = ((Com4jObject)testList.item(1)).queryInterface(ITSTest.class);
			            
			            IRunFactory runFac = tsTest.runFactory().queryInterface(IRunFactory.class);
			            ITest designTest = tsTest.test().queryInterface(ITest.class);
			            IRun newRun = runFac.addItem(runFac.uniqueRunName()).queryInterface(IRun.class);
			            document.selectSingleNode("/testcase/runid").setText(newRun.name());
			            IBaseFactory stepFac = newRun.stepFactory().queryInterface(IBaseFactory.class);
			            if(!document.selectSingleNode("/testcase/overallstatus").getText().equals("")) {
				            if(!document.selectSingleNode("/testcase/overallstatus").getText().equalsIgnoreCase("")){
				            	newRun.status(document.selectSingleNode("/testcase/overallstatus").getText());
			        		}
				            IBaseFactory desginstepFac = designTest.designStepFactory().queryInterface(IBaseFactory.class);
				            IList desginstepsList = desginstepFac.newList("").queryInterface(IList.class);
				            for(int eachIStep = 1; eachIStep<=desginstepsList.count(); eachIStep++) {
				            	Node node = (Node) document.selectNodes("/testcase/teststep").get(eachIStep - 1);
				            	Element element = (Element) node;
				            	String tcDesc = element.selectSingleNode("stepdescription").getText();
				            	String tcStatus = element.selectSingleNode("stepstatus").getText();
				            	String tcExp = element.selectSingleNode("expectedresult").getText();
				            	String tcActual = element.selectSingleNode("actualresult").getText();
			    	        	//IDesignStep eachIStepN = ((Com4jObject)desginstepsList.item(eachIStep)).queryInterface(IDesignStep.class);
			    	        	try {
			    	        		stepFac.addItem(Integer.toString(eachIStep)).queryInterface(IStep.class);
			    		            IList newStepList = stepFac.newList("").queryInterface(IList.class);
			    		            IStep newStep = ((Com4jObject) newStepList.item(eachIStep)).queryInterface(IStep.class);
			    		            newStep.field("ST_STATUS",tcStatus);
			    		            newStep.field("ST_STEP_NAME","");
			    		            newStep.field("ST_DESCRIPTION",tcDesc);
			    		            newStep.field("ST_EXPECTED",tcExp);
			    		            newStep.field("ST_ACTUAL",tcActual);
			    		            newStep.post();
			    	        	}
			    	        	catch(Exception e) {
			    	        		e.printStackTrace();
			    	        	}
			    	        
				            }
				            
				            if(props.getProperty("AttachTestRun").equalsIgnoreCase("no")) {
					            IAttachmentFactory reportattachs = tsTest.attachments().queryInterface(IAttachmentFactory.class);
					            IList attachList = reportattachs.newList("").queryInterface(IList.class);
					            for(int eachAttach = 1; eachAttach<=attachList.count(); eachAttach++) {
					            	IAttachment eachAttachment = ((Com4jObject)attachList.item(eachAttach)).queryInterface(IAttachment.class);
					            	reportattachs.removeItem(eachAttachment.id());
					            }
					            if(!document.selectSingleNode("/testcase/report").getText().equals("")) {
					            	if(new File(testpath + "\\Reports\\" + document.selectSingleNode("/testcase/report").getText()).exists()) {
					            		File reportFile = new File(testpath + "\\Reports\\" + document.selectSingleNode("/testcase/report").getText());
					            		IAttachment attachreport = reportattachs.addItem(reportFile.getName()).queryInterface(IAttachment.class);
					            		//attachreport.fileName(tcpath + "\\Reports\\" + document.selectSingleNode("/testcase/report").getText());
					            		IExtendedStorage extAttach = attachreport.attachmentStorage().queryInterface(IExtendedStorage.class);
					            		extAttach.clientPath(reportFile.getParent());
					            		extAttach.save(reportFile.getName(), true);
					            		attachreport.post();
					            	}
					            }
					            for(File eachFile : new File(testpath + "\\Reports").listFiles()) {
					            	if(eachFile.isFile() && !eachFile.isHidden() && !eachFile.getName().contains(".docx")) {
					            		File reportFile = eachFile;
					            		IAttachment attachreport = reportattachs.addItem(reportFile.getName()).queryInterface(IAttachment.class);
					            		//attachreport.fileName(tcpath + "\\Reports\\" + document.selectSingleNode("/testcase/report").getText());
					            		IExtendedStorage extAttach = attachreport.attachmentStorage().queryInterface(IExtendedStorage.class);
					            		extAttach.clientPath(reportFile.getParent());
					            		extAttach.save(reportFile.getName(), true);
					            		attachreport.post();
					            	}
					            }
				            }
				            else if(props.getProperty("AttachTestRun").equalsIgnoreCase("yes")) {
				            	IAttachmentFactory reportattachs = newRun.attachments().queryInterface(IAttachmentFactory.class);
					            if(!document.selectSingleNode("/testcase/report").getText().equals("")) {
					            	if(new File(testpath + "\\Reports\\" + document.selectSingleNode("/testcase/report").getText()).exists()) {
					            		File reportFile = new File(testpath + "\\Reports\\" + document.selectSingleNode("/testcase/report").getText());
					            		IAttachment attachreport = reportattachs.addItem(reportFile.getName()).queryInterface(IAttachment.class);
					            		//attachreport.fileName(tcpath + "\\Reports\\" + document.selectSingleNode("/testcase/report").getText());
					            		IExtendedStorage extAttach = attachreport.attachmentStorage().queryInterface(IExtendedStorage.class);
					            		extAttach.clientPath(reportFile.getParent());
					            		extAttach.save(reportFile.getName(), true);
					            		attachreport.post();
					            	}
					            }
					            for(File eachFile : new File(testpath + "\\Reports").listFiles()) {
					            	if(eachFile.isFile() && !eachFile.isHidden() && !eachFile.getName().contains(".docx")) {
					            		File reportFile = eachFile;
					            		IAttachment attachreport = reportattachs.addItem(reportFile.getName()).queryInterface(IAttachment.class);
					            		//attachreport.fileName(tcpath + "\\Reports\\" + document.selectSingleNode("/testcase/report").getText());
					            		IExtendedStorage extAttach = attachreport.attachmentStorage().queryInterface(IExtendedStorage.class);
					            		extAttach.clientPath(reportFile.getParent());
					            		extAttach.save(reportFile.getName(), true);
					            		attachreport.post();
					            	}
					            }
					            
				            }
				            /*List<Node> nodes =  document.selectNodes("/testcase/teststep/screenshot");
				            for(Node eachScrn : nodes) {
				            	Element element = (Element) eachScrn;
				            	if(!element.getText().equals("")) {
				            		if(new File(testpath + "\\Screenshots\\" + element.getText()).exists()) {
				            			File eachScreen = new File(testpath + "\\Screenshots\\" + element.getText());
				            			IAttachment attachscreen = reportattachs.addItem(eachScreen.getName()).queryInterface(IAttachment.class);
				            			//attachscreen.fileName(tcpath + "\\Screenshots\\" + element.getText());
				            			IExtendedStorage extAttach = attachscreen.attachmentStorage().queryInterface(IExtendedStorage.class);
					            		extAttach.clientPath(eachScreen.getParent());
					            		extAttach.save(eachScreen.getName(), true);
				            			attachscreen.post();
				            		}
				            	}
				            }*/
				            newRun.post();
			            }
		            }
		        }
		        FileOutputStream fos = new FileOutputStream(inputFile);
		        OutputFormat format = OutputFormat.createPrettyPrint();
		        XMLWriter writer = new XMLWriter(fos, format);
		        writer.write(document);
		        writer.close();
		    }
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/*private class SwingAction extends AbstractAction {
		public SwingAction() {
			putValue(NAME, "SwingAction");
			putValue(SHORT_DESCRIPTION, "Some short description");
		}
		public void actionPerformed(ActionEvent e) {
		}
	}*/
	
	public void setStepDetails(int stepNum, String Status, String Screenshotname, String actuals) {
		try {
			File inputFile = new File(tcpath + "\\testDetails.tc");
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
	
	
	public void setStepDetails(String testName, int stepNum, String Status, String Screenshotname, String actuals) {
		try {
			File inputFile = new File(testName + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectNodes("/testcase/teststep").get(stepNum);
        	Element element = (Element) node;
        	element.selectSingleNode("actualresult").setText(actuals);
        	if(!Status.equals("")) element.selectSingleNode("stepstatus").setText(Status);
        	if(!Screenshotname.equals("")) {
	        	if(element.selectSingleNode("screenshot").getText().equals(""))  {
	        		element.selectSingleNode("screenshot").setText(Screenshotname);
	        	}
	        	else {
	        		element.addElement("screenshot").setText(Screenshotname);
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
	
	/*@SuppressWarnings("unchecked")
	private String getOverAllStatus() {
		String status = "";
		try {
			String value = "";
			File inputFile = new File(tcpath + "\\testDetails.tc");
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
	
	
	@SuppressWarnings("unchecked")
	private String getOverAllStatus(String testCasePath) {
		String status = "";
		try {
			String value = "";
			File inputFile = new File(testCasePath + "\\testDetails.tc");
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
	}*/
	
	public class CustomIconRenderer extends DefaultTreeCellRenderer {
		   ImageIcon passedIcon;
		   ImageIcon failedIcon;
		   ImageIcon ncIcon;
		   ImageIcon blockedIcon;
		
		    public CustomIconRenderer() {
		        passedIcon = new ImageIcon(tea.class.getResource("/res/passed12x12.png"));
		        failedIcon = new ImageIcon(tea.class.getResource("/res/failed12x12.png"));
		        ncIcon = new ImageIcon(tea.class.getResource("/res/notcompleted12x12.png"));
		        blockedIcon = new ImageIcon(tea.class.getResource("/res/blocked12x12.png"));
		    }
		
		    public Component getTreeCellRendererComponent(JTree tree,
		      Object value,boolean sel,boolean expanded,boolean leaf,
		      int row,boolean hasFocus) {
		
		        super.getTreeCellRendererComponent(tree, value, sel, 
		          expanded, leaf, row, hasFocus);
			        if(value instanceof File) {
				        File valueName= (File)value;
				        setText(valueName.getName());
				        if (leaf) {
				        	if(new File(((File) value).getAbsolutePath() + "\\testDetails.tc").exists()) {
				        		String status = getTestDetails(((File) value).getAbsolutePath(), "overallstatus");
						        if(status.equals("Passed")) {
						            setIcon(passedIcon);
						        } else if(status.equals("Failed")) {
						            setIcon(failedIcon);
						        } else if(status.equals("Not Completed")) {
						            setIcon(ncIcon);
						        } else if(status.equals("Blocked")) {
						            setIcon(blockedIcon);
						        } 
				        	}
				        }
			        }
			        return this;
			    }
		    
		}
	
	private XWPFDocument replaceText(XWPFDocument doc, String testcaseName, String moduleName, String description, String Status, String testData) {
		   for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains("<TESTCASENAME>")) {
                        text = text.replace("<TESTCASENAME>", testcaseName);
                        r.setText(text, 0);
                    }
                    if (text != null && text.contains("<MODULENAME>")) {
                        text = text.replace("<MODULENAME>", moduleName);
                        r.setText(text, 0);
                    }
                    if (text != null && text.contains("<TESTDESCRIPTION>")) {
                        text = text.replace("<TESTDESCRIPTION>", description);
                        r.setText(text, 0);
                    }
                    if (text != null && text.contains("<OVERALLSTATUS>")) {
                        text = text.replace("<OVERALLSTATUS>", Status);
                        r.setText(text, 0);
                    }
                    if (text != null && text.contains("<TESTDATA>")) {
                        text = text.replace("<TESTDATA>", testData);
                        r.setText(text, 0);
                    }
                }
            }
        }

        for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null && text.contains("<TESTCASENAME>")) {
                                text = text.replace("<TESTCASENAME>", testcaseName);
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("<MODULENAME>")) {
                                text = text.replace("<MODULENAME>", moduleName);
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("<TESTDESCRIPTION>")) {
                                text = text.replace("<TESTDESCRIPTION>", description);
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("<OVERALLSTATUS>")) {
                                text = text.replace("<OVERALLSTATUS>", Status);
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("<TESTDATA>")) {
                                text = text.replace("<TESTDATA>", testData);
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
	    }
		return doc;
	   }
	
	public boolean checkALMconnection(String modulePath) {
		boolean connected = false;
		try {
			File inputFile = new File(modulePath + "\\moduleDetails.ts");
	        SAXReader reader = new SAXReader();
	        Document document = reader.read(inputFile);
	        Node node = (Node) document.selectNodes("/module").get(0);
        	Element element = (Element) node;
        	String domainName = element.selectSingleNode("ALMDomain").getText();
        	String projectname = element.selectSingleNode("ALMProject").getText();
			ITDConnection itdc= ClassFactory.createTDConnection();
	        itdc.initConnectionEx(props.getProperty("ALMURL").trim());
	        itdc.connectProjectEx(domainName, projectname, prefs.get("UserID", ""), prefs.get("pass", ""));
	        if(itdc.connected()) {
	        	connected = true;
	        }
	        else {
	        	connected = false;
	        	prefs.put("pass", "");
	        }
		}
		catch(Exception e) {
			connected = false;
			prefs.put("pass", "");
		}
		return connected;
	}
	
	public boolean checkALMconnection(String domainName, String projectname, String userID, String pass) {
		boolean connected = false;
		try {
			ITDConnection itdc= ClassFactory.createTDConnection();
	        itdc.initConnectionEx(props.getProperty("ALMURL").trim());
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
	
	public void openTestDataSheet(String modulePath) {
		try {
			if(!new File(modulePath + "\\Test Data.xlsx").exists()) {
				XSSFWorkbook tcBook = new XSSFWorkbook();
				XSSFSheet tcSheet = tcBook.createSheet();
				XSSFRow headerRow = tcSheet.createRow(0);
				headerRow.createCell(0).setCellValue("Test Case Name");
				headerRow.createCell(1).setCellValue("Test Data");
				
				int rowNum = 1;
				File moduleFolder = new File(modulePath);
				if(moduleFolder.isDirectory()) {
					for(File eachFolder : moduleFolder.listFiles()) {
						if(new File(eachFolder.getAbsolutePath() + "\\testDetails.tc").exists()) {
							XSSFRow curRow = tcSheet.createRow(rowNum);
							curRow.createCell(0).setCellValue(eachFolder.getName());
							rowNum++;
						}
					}
				}
				tcSheet.autoSizeColumn(0);
				FileOutputStream fileOutputStream = null;
				try	{
					fileOutputStream = new FileOutputStream(modulePath + "\\Test Data.xlsx");
					tcBook.write(fileOutputStream);
					fileOutputStream.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			
			Desktop.getDesktop().open(new File(modulePath + "\\Test Data.xlsx"));
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public void importTestData(String modulePath) {
		try {
			if(new File(modulePath + "\\Test Data.xlsx").exists()) {
				FileInputStream tcStream = new FileInputStream(new File(modulePath + "\\Test Data.xlsx"));
				XSSFWorkbook tcBook = new XSSFWorkbook(tcStream);
				XSSFSheet tcSheet = tcBook.getSheetAt(0);
				FormulaEvaluator formulaEvaluator = tcBook.getCreationHelper().createFormulaEvaluator();
				for(int rowIndex = 1; rowIndex<=tcSheet.getLastRowNum(); rowIndex++) {
					if(tcSheet.getRow(rowIndex).getCell(1)!=null && tcSheet.getRow(rowIndex).getCell(1)!=null){
						if(new File(modulePath + "\\" +  getCellValueAsString(tcSheet.getRow(rowIndex).getCell(0), formulaEvaluator) + "\\testDetails.tc" ).exists()) {
							setTestDetails(modulePath + "\\" + getCellValueAsString(tcSheet.getRow(rowIndex).getCell(0), formulaEvaluator), "testdata", getCellValueAsString(tcSheet.getRow(rowIndex).getCell(1), formulaEvaluator));
						}
					}
				}
			}
			else {
				JOptionPane.showMessageDialog(null, "Test Data Import Sheet not available for this module");
			}
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	private static String getCellValueAsString(XSSFCell xssfCell, FormulaEvaluator formulaEvaluator)
	{
		if (xssfCell == null || xssfCell.getCellType() == XSSFCell.CELL_TYPE_BLANK) {
			return "";
		} else {
			DataFormatter dataFormatter = new DataFormatter();
			return dataFormatter.formatCellValue(formulaEvaluator.evaluateInCell(xssfCell));
		}
	}
	
	
	
	public void pasteGeneral(String sourceTC, String targetTC, boolean comments) {
		try {
			File inputFile = new File(sourceTC + "\\testDetails.tc");
			File destFile = new File(targetTC + "\\testDetails.tc");
	        SAXReader reader = new SAXReader();
	        Document sourceDoc = reader.read(inputFile);
	        List<Object> allNodes = sourceDoc.selectNodes("/testcase/general");
	        Document destDoc = reader.read(destFile);
	        Node destNode = destDoc.selectSingleNode("/testcase");
	        Element destElement = (Element)destNode;
			for(Object eachNode : allNodes) {
				Element destGeneral = destElement.addElement("general");
				Node node = (Node)eachNode;
				Element sourceElement = (Element)node;
				FileUtils.copyFileToDirectory(new File(sourceTC + "\\Screenshots\\" + sourceElement.selectSingleNode("screenshot").getText()), new File(targetTC + "\\Screenshots"));
				destGeneral.addElement("screenshot").setText(sourceElement.selectSingleNode("screenshot").getText());
				if(comments) {
					destGeneral.addElement("comment").setText(sourceElement.selectSingleNode("comment").getText());
				}
				else {
					destGeneral.addElement("comment").setText("");
				}
			}
			FileOutputStream fos = new FileOutputStream(destFile);
			OutputFormat format = OutputFormat.createPrettyPrint();
	        XMLWriter writer = new XMLWriter(fos, format);
	        writer.write(destDoc);
	        writer.close();
		}
		catch(Exception ex) {
			ex.printStackTrace();
		}
	}
	

}