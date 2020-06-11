package src;


import java.awt.Component;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Image;
import java.awt.Point;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.event.MouseMotionAdapter;
import java.awt.event.MouseMotionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JViewport;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.border.BevelBorder;

import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;
/*import com.jgoodies.forms.layout.FormLayout;
import com.jgoodies.forms.layout.ColumnSpec;
import com.jgoodies.forms.layout.RowSpec;*/
/*import net.miginfocom.swing.MigLayout;*/

public class ImageArranger {
	String tcPath = "";
	int previousY = 0;
	ArrayList<Component> selectedComps = new ArrayList<Component>();
	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();

	public ImageArranger(String tcpath) throws Exception{

		JFrame mainFrame = new JFrame();
		tcPath = tcpath;
		mainFrame.setTitle("Snapshot Arrange");
		mainFrame.setIconImage(Toolkit.getDefaultToolkit().getImage(ImageArranger.class.getResource("/res/logo48x48.png")));
		mainFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		mainFrame.setResizable(false);
		mainFrame.getContentPane().setLayout(null);
		mainFrame.setSize(650, 600);
		mainFrame.setLocation(dim.width/2- mainFrame.getSize().width/2, dim.height/2-mainFrame.getSize().height/2);
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
			JLabel lblNewLabel = new JLabel();
			lblNewLabel.setIcon(new ImageIcon(ImageArranger.class.getResource("/res/reorder16x16.png")));
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
		mainFrame.setVisible(true);
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

