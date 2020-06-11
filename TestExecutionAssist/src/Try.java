package src;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.Image;
import java.awt.Paint;
import java.awt.Point;
import java.awt.Toolkit;
import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.prefs.Preferences;

import javax.swing.BorderFactory;
import javax.swing.ButtonGroup;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JPopupMenu;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.JViewport;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.Border;
import javax.swing.border.EmptyBorder;
import javax.swing.plaf.ColorUIResource;
import javax.swing.table.DefaultTableModel;
import javax.swing.tree.TreePath;

import net.coderazzi.filters.gui.AutoChoices;
import net.coderazzi.filters.gui.TableFilterHeader;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.DialShape;
import org.jfree.chart.plot.MeterInterval;
import org.jfree.chart.plot.MeterPlot;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.xy.XYLineAndShapeRenderer;
import org.jfree.chart.title.TextTitle;
import org.jfree.data.Range;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.general.DefaultValueDataset;
import org.jfree.data.time.Day;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;

import javax.swing.border.BevelBorder;

import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
/*import com.jgoodies.forms.layout.FormLayout;
import com.jgoodies.forms.layout.ColumnSpec;
import com.jgoodies.forms.layout.RowSpec;*/
/*import net.miginfocom.swing.MigLayout;*/
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.event.MouseMotionAdapter;
import java.awt.event.MouseMotionListener;
import java.awt.font.TextAttribute;

import javax.swing.JProgressBar;
import javax.swing.SwingConstants;

import java.awt.FlowLayout;

import javax.swing.border.LineBorder;
import javax.swing.JTextPane;
import javax.swing.ScrollPaneConstants;

import java.awt.ComponentOrientation;

public class Try {
	String tcPath = "";
	int previousY = 0;
	ArrayList<Component> selectedComps = new ArrayList<Component>();
	
	/**
	 * @wbp.parser.entryPoint
	 */
	public static void main(String[] args) throws Exception {
		//Try tryObj =  new Try("T:\\Operations\\IT\\IT Shared\\CL\\Test\\Carefirst Automation_WF_Sel\\Sangeeth\\TEA_Demo\\TEA_Demo\\TesCase_05");
		System.out.println(System.getProperty("sun.arch.data.model"));
		
	}
	
	/**
	 * @wbp.parser.entryPoint
	 */
	public Try(String tcpath) throws Exception{
		/*for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
            if ("Nimbus".equals(info.getName())) {
                UIManager.setLookAndFeel(info.getClassName());
                if (uilayout.equals("Nimbus")) {
                    tweakNimbusUI();
                }
                break;
            }
        }
		
		JDialog domainDialog = new JDialog();
		domainDialog.setIconImage(Toolkit.getDefaultToolkit().getImage(Try.class.getResource("/res/logo48x48.png")));
		domainDialog.setTitle("Domain and Project");
		domainDialog.setSize(300, 164);
		FlowLayout flowLayout = new FlowLayout();
		flowLayout.setVgap(15);
		flowLayout.setHgap(10);
		domainDialog.getContentPane().setLayout(flowLayout);
		
		JLabel lblDomain = new JLabel("Domain :");
		domainDialog.getContentPane().add(lblDomain);
		
		JComboBox comboDomain = new JComboBox();
		comboDomain.setPreferredSize(new Dimension(200, 26));
		domainDialog.getContentPane().add(comboDomain);
		
		JLabel lblProject = new JLabel("Project   :");
		domainDialog.getContentPane().add(lblProject);
		
		JComboBox comboProject = new JComboBox();
		comboProject.setPreferredSize(new Dimension(200, 26));
		domainDialog.getContentPane().add(comboProject);
		
		JButton btnTest = new JButton("Test");
		btnTest.setPreferredSize(new Dimension(100, 28));
		domainDialog.getContentPane().add(btnTest);*/
		
		//JFrame mainFrame = new JFrame();
		tcPath = tcpath;
		/*mainFrame.getContentPane().setLayout(null);
		mainFrame.setBounds(0 ,0 ,650, 600);
		JPanel panel = new JPanel();
		JScrollPane panelScroll = new JScrollPane(panel);
		panelScroll.getVerticalScrollBar().setUnitIncrement(16);
		panelScroll.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
		panelScroll.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
		panelScroll.setBounds(0, 30, 640, 550);
		//panelScroll.setBounds(0, 0, 690, 480);
		mainFrame.getContentPane().add(panelScroll);
		panel.setLayout(new WrapFlowLayout(WrapFlowLayout.LEFT));
		
		
		
		JPopupMenu arrangePopup = new JPopupMenu();
		JMenuItem cutMenu = new JMenuItem("Cut");
		arrangePopup.add(cutMenu);
		JMenuItem insertAboveMenu = new JMenuItem("Insert Above");
		arrangePopup.add(insertAboveMenu);
		JMenuItem insertBelowMenu = new JMenuItem("Insert Below");
		arrangePopup.add(insertBelowMenu);
		insertAboveMenu.setEnabled(false);
		insertBelowMenu.setEnabled(false);
		
		
		cutMenu.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				selectedComps.clear();
				for(Component eachCOmp : panel.getComponents()) {
					if(eachCOmp instanceof JPanel) {
						if(((JCheckBox)((JPanel)eachCOmp).getComponent(1)).isSelected()) {
							selectedComps.add(eachCOmp);
							insertAboveMenu.setEnabled(true);
							insertBelowMenu.setEnabled(true);
						}
					}
				}
			}
		});
		insertAboveMenu.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				JPanel invokerPanel = (JPanel)((JPopupMenu)((JMenuItem)e.getSource()).getParent()).getInvoker();
				int index = 0;
				ArrayList<Component> prevList = new ArrayList<Component>();
				for(Component eachCOmp : panel.getComponents()) {
					if(eachCOmp instanceof JPanel) {
						prevList.add(eachCOmp);
					}
				}
				int indexNum = 0;
				index = getComponentIndex(invokerPanel);
				panel.removeAll();
				for(int intCOmp = 0; intCOmp < prevList.size(); intCOmp++) {
					Component eachCOmp = prevList.get(intCOmp);
					if(eachCOmp instanceof JPanel) {
						if(!selectedComps.contains(eachCOmp)) {
							if(index == 0 && intCOmp == 0) {
								for(Component seleComp : selectedComps) {
									((JCheckBox)((JPanel)seleComp).getComponent(1)).setSelected(false);
									panel.add(seleComp, indexNum);
									indexNum++;
								}
								panel.add(eachCOmp, indexNum);
								indexNum++;
							}
							else if(intCOmp == index-1) {
								panel.add(eachCOmp, indexNum);
								indexNum++;
								for(Component seleComp : selectedComps) {
									((JCheckBox)((JPanel)seleComp).getComponent(1)).setSelected(false);
									panel.add(seleComp, indexNum);
									indexNum++;
								}
							}
							else {
								panel.add(eachCOmp, indexNum);
								indexNum++;
							}
						}
					}
				}
				panel.revalidate();
				panel.repaint();
				insertAboveMenu.setEnabled(false);
				insertBelowMenu.setEnabled(false);
			}
		});
		
		insertBelowMenu.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				JPanel invokerPanel = (JPanel)((JPopupMenu)((JMenuItem)e.getSource()).getParent()).getInvoker();
				int index = 0;
				ArrayList<Component> prevList = new ArrayList<Component>();
				for(Component eachCOmp : panel.getComponents()) {
					if(eachCOmp instanceof JPanel) {
						prevList.add(eachCOmp);
					}
				}
				int indexNum = 0;
				index = getComponentIndex(invokerPanel);
				panel.removeAll();
				for(int intCOmp = 0; intCOmp < prevList.size(); intCOmp++) {
					Component eachCOmp = prevList.get(intCOmp);
					if(eachCOmp instanceof JPanel) {
						if(!selectedComps.contains(eachCOmp)) {
							if(index == prevList.size() - 1 && intCOmp == prevList.size() - 1) {
								panel.add(eachCOmp, indexNum);
								indexNum++;
								for(Component seleComp : selectedComps) {
									((JCheckBox)((JPanel)seleComp).getComponent(1)).setSelected(false);
									panel.add(seleComp, indexNum);
									indexNum++;
								}
							}
							else if(intCOmp == index+1) {
								for(Component seleComp : selectedComps) {
									((JCheckBox)((JPanel)seleComp).getComponent(1)).setSelected(false);
									panel.add(seleComp, indexNum);
									indexNum++;
								}
								panel.add(eachCOmp, indexNum);
								indexNum++;
							}
							else {
								panel.add(eachCOmp, indexNum);
								indexNum++;
							}
						}
					}
				}
				
				panel.revalidate();
				panel.repaint();
				insertAboveMenu.setEnabled(false);
				insertBelowMenu.setEnabled(false);
			}
		});
		
		 MouseListener imageListener = new MouseAdapter() {
		    	public void mousePressed(MouseEvent e) {
		    		try {
			            if (e.getClickCount() == 1) {
			            	if(e.getSource() instanceof JLabel) {
			            		new ImageEditor(((JLabel)e.getSource()).getToolTipText() ,tcPath, "general");
			            	}
					    }
		    		}
		    		catch(Exception ex) {
		    			ex.printStackTrace();
		    		}
		        }
		 };
		
		
		MouseMotionListener motionListener = new MouseMotionAdapter() {
	        public void mouseDragged(MouseEvent me) {
        		JPanel sourceComp = (JPanel)me.getSource();
        		panel.setComponentZOrder(sourceComp, 0);
	            me.translatePoint(me.getComponent().getLocation().x, me.getComponent().getLocation().y);
	            sourceComp.setLocation(sourceComp.getX(), me.getY());
	        }
	    };
	    
	    MouseListener dropListener = new MouseAdapter() {
	    	public void mousePressed(MouseEvent e) {
	            previousY = e.getComponent().getLocation().y;
	            if (e.getClickCount() == 1 && SwingUtilities.isRightMouseButton(e)) {
	            	arrangePopup.show(e.getComponent(), e.getX(), e.getY());
			    }
	        }

	        public void mouseReleased(MouseEvent e) {
	        	int y = e.getComponent().getLocation().y;
	        	if(y != previousY) {
		        	JPanel sourceComp = (JPanel)e.getSource();
		        	Component dropComp = null;
		            if (y < previousY) {
		            	dropComp = getComponentAt(sourceComp, panel, new Point(e.getX(), e.getComponent().getLocation().y));
		            	panel.add(sourceComp, getComponentIndex(dropComp) - 1);
		            } else if (y > previousY) {
		            	dropComp = getComponentAt(sourceComp, panel, new Point(e.getX(), e.getComponent().getLocation().y + e.getComponent().getHeight()));
		            	panel.add(sourceComp, getComponentIndex(dropComp));
		            }
		        	panel.revalidate();
		        	panel.repaint();
		        }
	        }
	    };
	    
	    File inputFile = new File(tcPath + "\\testDetails.tc");
        SAXReader reader = new SAXReader();
        Document document = reader.read(inputFile);
        List<Object> allNodes = document.selectNodes("/testcase/general");
		for(Object eachNode : allNodes) {
			Node node = (Node)eachNode;
			Element element = (Element) node;
			JPanel panel1 = new JPanel();
			panel1.setBorder(new BevelBorder(BevelBorder.RAISED, null, null, null, null));
			FlowLayout flowLayout = (FlowLayout) panel1.getLayout();
			flowLayout.setVgap(10);
			flowLayout.setHgap(20);
			panel1.addMouseMotionListener(motionListener);
			panel1.addMouseListener(dropListener);
			panel.add(panel1);
			JLabel lblNewLabel = new JLabel("+");
			panel1.add(lblNewLabel);
			JCheckBox selectCheck = new JCheckBox("");
			panel1.add(selectCheck);
			JLabel imgLbl = new JLabel("");
			imgLbl.setHorizontalAlignment(SwingConstants.TRAILING);
			imgLbl.setPreferredSize(new Dimension(180, 150));
			imgLbl.setToolTipText(element.selectSingleNode("screenshot").getText());
			imgLbl.addMouseListener(imageListener);
			ImageIcon icon = new ImageIcon(new File(tcPath + "\\Screenshots\\" + element.selectSingleNode("screenshot").getText()).getAbsolutePath());
			Image img = icon.getImage();
            Image newImg = img.getScaledInstance(imgLbl.getPreferredSize().width, imgLbl.getPreferredSize().height, Image.SCALE_SMOOTH);
            ImageIcon newImc = new ImageIcon(newImg);
            imgLbl.setIcon(newImc);
			panel1.add(imgLbl);
			
			JTextArea snapComments = new JTextArea();
			JScrollPane commentScroll = new JScrollPane(snapComments);
			commentScroll.setPreferredSize(new Dimension(300, 80));
			snapComments.setText(element.selectSingleNode("comment").getText());
			panel1.add(commentScroll);
		}
			
		JButton btnSave = new JButton("Save");
		btnSave.setBounds(6, 1, 69, 28);
		mainFrame.getContentPane().add(btnSave);
		btnSave.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					File inputFile = new File(tcPath + "\\testDetails.tc");
			        SAXReader reader = new SAXReader();
			        Document document = reader.read(inputFile);
			        List<Node> allNodes = document.selectNodes("/testcase/general");
			        if(allNodes.size() == panel.getComponents().length) {
			        	int eachPanel = 0;
			        	for(Node eachNode : allNodes) {
			        		Element element = (Element) eachNode;
			        		JPanel curPanel = (JPanel)panel.getComponent(eachPanel);
			        		element.selectSingleNode("screenshot").setText(((JLabel)curPanel.getComponent(2)).getToolTipText());
			        		element.selectSingleNode("comment").setText(((JTextArea)((JViewport)((JScrollPane)curPanel.getComponent(3)).getComponent(0)).getView()).getText());
			        		eachPanel++;
			        	}
			        	FileOutputStream fos = new FileOutputStream(inputFile);
						OutputFormat format = OutputFormat.createPrettyPrint();
				        XMLWriter writer = new XMLWriter(fos, format);
				        writer.write(document);
				        writer.close();
			        }
			        else {
			        	JOptionPane.showMessageDialog(null, "Something went wrong. Please re-open the arragement window and try again!");
			        }
				}
				catch(Exception ex) {
					
				}
			}
		});
		mainFrame.setVisible(true);*/
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
	   
   public int getComponentIndex(Component component) {
	    if (component != null && component.getParent() != null) {
	      Container c = component.getParent();
	      for (int i = 0; i < c.getComponentCount(); i++) {
	        if (c.getComponent(i) == component)
	          return i;
	      }
	    }

	    return -1;
	  }
   public Component getComponentAt(Component source, Container parent, Point p) {
       Component comp = null;
       for (Component child : parent.getComponents()) {
    	   if(child!=source) {
	           if (child.getBounds().contains(p)) {
	               comp = child;
	           }
    	   }
       }
       return comp;
   }
}
