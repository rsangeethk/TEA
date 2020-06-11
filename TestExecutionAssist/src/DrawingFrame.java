package src;

import java.awt.*;
import java.awt.event.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.border.Border;

import  java.util.List;

import org.apache.commons.lang.ArrayUtils;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

@SuppressWarnings("serial")
public class DrawingFrame extends JPanel {
	public Color DRAWING_COLOR = new Color(255, 100, 200);
	

	public static BufferedImage backgroundImg;
	public int prefW;
	public int prefH;
	private Point startPt = null;
	private Point endPt = null;
	private Point currentPt = null;
	static Border emptyBorder = BorderFactory.createEmptyBorder();
	

	public DrawingFrame() throws IOException {
		try {
			 
			JToggleButton mark1 = new JToggleButton();
			add(mark1);
			MyMouseAdapter myMouseAdapter = new MyMouseAdapter();
			mark1.addActionListener(new ActionListener() {
				@Override
				public void actionPerformed(ActionEvent e) {
					AbstractButton abstractButton = (AbstractButton) e.getSource();
					if (abstractButton.getModel().isSelected()) {
						addMouseMotionListener(myMouseAdapter);
						addMouseListener(myMouseAdapter);
					} else {
						removeMouseMotionListener(myMouseAdapter);
						removeMouseListener(myMouseAdapter);
					}

				}
			});
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	

	@Override
	   protected void paintComponent(Graphics g) {
		  Graphics2D g2D = (Graphics2D) g;
	      super.paintComponent(g2D);
	      if (backgroundImg != null) {
	    	  g2D.drawImage(backgroundImg, 0, 0, this);
	      }
	      if (startPt != null && currentPt != null) {
	    	 float thickness = 1;
	    	 if(((JPanel)this).getParent() !=null) {
				  if(((JPanel)this).getParent().getParent().getComponent(16) instanceof JComboBox) {
					  JComboBox eachComp = (JComboBox) ((JPanel)this).getParent().getParent().getComponent(16);
					  JComboBox<Object> tglbtn = (JComboBox<Object>) eachComp;
					  if(tglbtn.getSelectedIndex() == 0) {
						  thickness = 1;
					  }
					  else {
						  thickness = 2;
					  }
				  }
	    	 }
	    	 g2D.setStroke(new BasicStroke(thickness));
	    	 g2D.setColor(DRAWING_COLOR);
	         int x = Math.min(startPt.x, currentPt.x);
	         int y = Math.min(startPt.y, currentPt.y);
	         int width = Math.abs(startPt.x - currentPt.x);
	         int height = Math.abs(startPt.y - currentPt.y);
	         g2D.drawRect(x, y, width, height);
	         
	      }
	   }

	   @Override
	   public Dimension getPreferredSize() {
	      return new Dimension(prefW, prefH);
	   }

	   public void drawToBackground() {
		   try {
			  Color FINAL_DRAWING_COLOR = null;
			  float thickness = 1;
		      Graphics2D g = (Graphics2D) backgroundImg.getGraphics();
		      if(((JPanel)this).getParent() !=null) {
	    		  for(Component eachComp : ((JPanel)this).getParent().getParent().getComponents()) {
	    			  if(eachComp instanceof JToggleButton) {
	    				  JToggleButton tglbtn = (JToggleButton) eachComp;
	    				  if(tglbtn.getToolTipText()!=null) {
	    				  if(tglbtn.getToolTipText().equals("Selected")) {
	    					  String selectedColor = tglbtn.getName();
	    					  switch(selectedColor)	{
	    					  case "red" :
	    						  FINAL_DRAWING_COLOR = Color.red;
	    						  break;
	    					  case "blue" :
	    						  FINAL_DRAWING_COLOR = Color.blue;
	    						  break;
	    					  case "green" :
	    						  FINAL_DRAWING_COLOR = Color.green;
	    						  break;
	    					  case "black" :
	    						  FINAL_DRAWING_COLOR = Color.black;
	    						  break;
	    					  default :
	    						  FINAL_DRAWING_COLOR = Color.red;
	    						  break;
	    					  }
	    				  }
	    				  }
	    			  }
	    			  else if(eachComp instanceof JComboBox) {
	    				  JComboBox<Object> tglbtn = (JComboBox<Object>) eachComp;
	    				  if(tglbtn.getSelectedIndex() == 0) {
	    					  thickness = 1;
	    				  }
	    				  else {
	    					  thickness = 2;
	    				  }
	    			  }
	    		  }
	    	  }
		      if(FINAL_DRAWING_COLOR == null) {
		    	  FINAL_DRAWING_COLOR = Color.red;
		      }
		      
		      g.setStroke(new BasicStroke(thickness));
		      g.setColor(FINAL_DRAWING_COLOR);
		      int x = Math.min(startPt.x, endPt.x);
		      int y = Math.min(startPt.y, endPt.y);
		      int width = Math.abs(startPt.x - endPt.x);
		      int height = Math.abs(startPt.y - endPt.y);
		      g.drawRect(x, y, width, height);
		      g.dispose();
		      //TODO improve image scaling following this answer http://stackoverflow.com/a/34266703
		      startPt = null;
		      repaint();
		   }
		   catch(Throwable e) {
			   
		   }
	   }

	   private class MyMouseAdapter extends MouseAdapter {
	      @Override
	      public void mouseDragged(MouseEvent mEvt) {
	         currentPt = mEvt.getPoint();
	         DrawingFrame.this.repaint();
	      }

	      @Override
	      public void mouseReleased(MouseEvent mEvt) {
	         endPt = mEvt.getPoint();
	         currentPt = null;
	         drawToBackground();
	      }

	      @Override
	      public void mousePressed(MouseEvent mEvt) {
	         startPt = mEvt.getPoint();
	      }
	}
}
