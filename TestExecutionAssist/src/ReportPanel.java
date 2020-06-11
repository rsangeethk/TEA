package src;

import java.awt.Color;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Rectangle;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.font.TextAttribute;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Map;

import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTree;
import javax.swing.JViewport;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.BevelBorder;
import javax.swing.event.RowSorterEvent;
import javax.swing.event.RowSorterListener;
import javax.swing.plaf.ColorUIResource;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import javax.swing.table.TableModel;
import javax.swing.tree.TreeNode;
import javax.swing.tree.TreePath;

import net.coderazzi.filters.IFilter;
import net.coderazzi.filters.IFilterObserver;
import net.coderazzi.filters.gui.AutoChoices;
import net.coderazzi.filters.gui.IFilterEditor;
import net.coderazzi.filters.gui.IFilterHeaderObserver;
import net.coderazzi.filters.gui.TableFilterHeader;
import net.coderazzi.filters.gui.editor.FilterEditor;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.SAXReader;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.xy.XYLineAndShapeRenderer;
import org.jfree.chart.title.TextTitle;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.time.Day;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;

import javax.swing.ScrollPaneConstants;
import javax.swing.JButton;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

public class ReportPanel extends JPanel{
	static String tcPath = "";
	private JTable summaryTable;
	JComboBox<String> moduleCombo = new JComboBox<String>();
	private JLabel lblTotalCount;
	JTree curtree = null;
	
	/**
	 * @wbp.parser.entryPoint
	 */
	public ReportPanel(String modulePath, JTree tree) throws Exception{
		this.curtree = tree;
		//frame.getContentPane().setLayout(null);
		ColorUIResource colorResource = new ColorUIResource(Color.green.darker());
		UIManager.put("nimbusOrange", colorResource);
		DefaultPieDataset dataset = new DefaultPieDataset();
		HashMap<String, Integer> statusMap = new HashMap<>();
		HashMap<String, Integer> dateMap = new HashMap<>();
		statusMap.put("Passed", 0);
		statusMap.put("Failed", 0);
		statusMap.put("No Run", 0);
		statusMap.put("Blocked", 0);
		statusMap.put("Not Completed", 0);
		moduleCombo.addItem("All");
		setBounds(0, 0, 685, 515);
		setLayout(null);
		
		JTabbedPane reportTabs = new JTabbedPane(JTabbedPane.TOP);
		reportTabs.setBounds(10, 6, 667, 515);
		add(reportTabs);
		
		JPanel summaryTab = new JPanel();
		reportTabs.addTab("Summary", null, summaryTab, null);
		summaryTab.setLayout(null);
		
		JPanel summaryPanel = new JPanel();
        summaryPanel.setPreferredSize(new Dimension(660, 800));
        summaryPanel.setLayout(null);
		
		DefaultTableModel moduleTableModel = new DefaultTableModel(new Object[] {"Module Name", "Passed", "Failed", "No Run", "Blocked", "Not Completed", "N/A", "Total"}, 0);
        JTable moduleTable = new JTable(moduleTableModel);
        moduleTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        moduleTable.getColumnModel().getColumn(0).setPreferredWidth(160);
        moduleTable.getColumnModel().getColumn(1).setPreferredWidth(60);
        moduleTable.getColumnModel().getColumn(2).setPreferredWidth(60);
        moduleTable.getColumnModel().getColumn(3).setPreferredWidth(60);
        moduleTable.getColumnModel().getColumn(4).setPreferredWidth(70);
        moduleTable.getColumnModel().getColumn(5).setPreferredWidth(70);
        moduleTable.getColumnModel().getColumn(6).setPreferredWidth(60);
        moduleTable.getColumnModel().getColumn(7).setPreferredWidth(60);
        JScrollPane moduleStatusScroll = new JScrollPane(moduleTable);
        moduleStatusScroll.setBounds(29, 470, 600, 324);
        summaryPanel.add(moduleStatusScroll);
		
		JPanel detailReport = new JPanel();
		reportTabs.addTab("Detailed Report", null, detailReport, null);
		detailReport.setLayout(null);
		
		
		
		DefaultTableModel tableModel = new DefaultTableModel(new Object[] {"Module Name", "Testcase Name", "Status", "Uploaded", "Last Run ID", "Report Name", "Last Executed By", "Last Executed On"}, 0);
		File projectFolder = new File(modulePath);
		HashMap<String, Integer> moduleStatus = new HashMap<>();
		int TotalCount = 0;
		for(File eachModule : projectFolder.listFiles()) {
			if(eachModule.isDirectory() && new File(eachModule.getAbsolutePath() + "\\moduleDetails.ts").exists()) {
				moduleCombo.addItem(eachModule.getName());
				moduleStatus.clear();
				moduleStatus.put("Passed", 0);
				moduleStatus.put("Failed", 0);
				moduleStatus.put("No Run", 0);
				moduleStatus.put("Blocked", 0);
				moduleStatus.put("Not Completed", 0);
				moduleStatus.put("Total", 0);
				for(File eachTestcase : eachModule.listFiles()) {
					if(eachTestcase.isDirectory() && new File(eachTestcase.getAbsolutePath() + "\\testDetails.tc").exists()) {
						tcPath = eachTestcase.getAbsolutePath();
						String tcStatus = getTestDetails("overallstatus");
						moduleStatus.put("Total", moduleStatus.get("Total") + 1);
						if(!tcStatus.equals("")) {
							statusMap.put(tcStatus, statusMap.get(tcStatus) + 1); 
							moduleStatus.put(tcStatus, moduleStatus.get(tcStatus) + 1);
						}
						else {
							statusMap.put("No Run", statusMap.get("No Run") + 1);
							moduleStatus.put("No Run", moduleStatus.get("No Run") + 1);
						}
						String tcRunID = getTestDetails("runid");
						String reportName = getTestDetails("report");
						String executedDates = getTestDetails("executedon");
						String executedBy = getTestDetails("executedby");
						String lastExecutedOn = executedDates.equals("") ? "" : executedDates.split(";")[executedDates.split(";").length - 1];
						if(!lastExecutedOn.equals("")) {
							if(!dateMap.containsKey(lastExecutedOn.split(" ")[0])) {
								dateMap.put(lastExecutedOn.split(" ")[0], 0);
							}
							dateMap.put(lastExecutedOn.split(" ")[0], dateMap.get(lastExecutedOn.split(" ")[0]) + 1); 
						}
						tableModel.addRow(new Object[] {eachModule.getName(), eachTestcase.getName(), tcStatus, "", tcRunID, reportName, executedBy, lastExecutedOn});
						TotalCount++;
					}
				}
				moduleTableModel.addRow(new Object[] {eachModule.getName(), moduleStatus.get("Passed"), moduleStatus.get("Failed"), moduleStatus.get("No Run"), moduleStatus.get("Blocked"), moduleStatus.get("Not Completed"), "0", moduleStatus.get("Total")});
			}
		}
		
		JPopupMenu reportMenu = new JPopupMenu();
		
		moduleTableModel.addRow(new Object[] {"Total", statusMap.get("Passed"), statusMap.get("Failed"), statusMap.get("No Run"), statusMap.get("Blocked"), statusMap.get("Not Completed"), "0", statusMap.get("Total")});
		JTable table = new JTable();
		table.setModel(tableModel);
		
		//TableRowFilterSupport.forTable(table).searchable(true).apply();
		TableFilterHeader filter = new TableFilterHeader(table, AutoChoices.ENABLED);
		
        moduleTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        filter.addFilter();
        
		table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		table.getColumnModel().getColumn(0).setPreferredWidth(150);
		table.getColumnModel().getColumn(1).setPreferredWidth(200);
		table.getColumnModel().getColumn(2).setPreferredWidth(100);
		table.getColumnModel().getColumn(3).setPreferredWidth(80);
		table.getColumnModel().getColumn(4).setPreferredWidth(150);
		table.getColumnModel().getColumn(5).setPreferredWidth(200);
		table.getColumnModel().getColumn(6).setPreferredWidth(100);
		table.getColumnModel().getColumn(7).setPreferredWidth(100);
		JScrollPane reportTableScroll = new JScrollPane(table);
		reportTableScroll.setBounds(0, 0, 667, 465);
		detailReport.add(reportTableScroll);
		table.getRowSorter().addRowSorterListener(new RowSorterListener() {
			@Override
			public void sorterChanged(RowSorterEvent e) {
				lblTotalCount.setText(filter.getTable().getRowCount() + " testcases");
				detailReport.revalidate();
				detailReport.repaint();
				lblTotalCount.revalidate();
				lblTotalCount.repaint();
			}
		});
		
		
		lblTotalCount = new JLabel("");
		lblTotalCount.setBounds(6, 464, 151, 20);
		detailReport.add(lblTotalCount);
		lblTotalCount.setText(table.getRowCount() + " testcases");
		
		JMenuItem openTest = new JMenuItem("Open Test Case");
		openTest.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				String moduleName = table.getValueAt(table.getSelectedRow(), 0).toString();
				String tcName = table.getValueAt(table.getSelectedRow(), 1).toString();
				File curNode = new File(new File(modulePath).getAbsolutePath() + "\\" + moduleName);
				File curChildNode = new File(curNode.getAbsolutePath() + "\\" + tcName);
				TreePath treePath = new TreePath(new Object[]{curtree.getModel().getRoot(),curNode, curChildNode});
				tree.expandPath(new TreePath(treePath));
				tree.setSelectionPath(treePath);
				Rectangle nodeLoc = curtree.getPathBounds(treePath);
				if(nodeLoc!=null) {
					((JViewport)curtree.getParent()).scrollRectToVisible(nodeLoc);
					MouseEvent me = new MouseEvent(curtree, 0, 0, 0, nodeLoc.x + 2, nodeLoc.y + 2, 2, false);
					for(MouseListener ml: curtree.getMouseListeners()){
					    ml.mousePressed(me);
					}
				}
				
			}
		});
		reportMenu.add(openTest);
		JMenuItem openReport = new JMenuItem("Open Test Report");
		openReport.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
			try {
				String moduleName = table.getValueAt(table.getSelectedRow(), 0).toString();
				String tcName = table.getValueAt(table.getSelectedRow(), 1).toString();
				tcPath = modulePath + "\\" + moduleName + "\\" + tcName;
				String reportName = getTestDetails("report");
				if(!reportName.equals("")) {
					File reportFile = new File(tcPath + "\\Reports\\" + reportName);
					if(reportFile.exists()){
						Desktop.getDesktop().open(reportFile);
					}
					else {
						JOptionPane.showMessageDialog(null, "Generated report is not available in the reports folder!");
					}
				}
				else {
					JOptionPane.showMessageDialog(null, "Report not available for the testcase");
				}
			}
			catch(Exception ex) {
				ex.printStackTrace();
			}
			}
		});
		reportMenu.add(openReport);
		
		table.addMouseListener(new MouseAdapter() {
		    public void mousePressed(MouseEvent me) {
		    	if (me.getClickCount() == 1 && SwingUtilities.isRightMouseButton(me)) {
			        reportMenu.show(me.getComponent(), me.getX(), me.getY());
			    }
		    }
		});
		
		
		//Pie Chart
		dataset.setValue( "Passed" , new Double( statusMap.get("Passed") ) );             
	    dataset.setValue( "Failed" , new Double( statusMap.get("Failed") ) );             
	    dataset.setValue( "No Run" , new Double( statusMap.get("No Run") ) );             
	    dataset.setValue( "Blocked" , new Double( statusMap.get("Blocked") ) );
	    dataset.setValue( "Not Completed" , new Double( statusMap.get("Not Completed") ) );
        JFreeChart chart = ChartFactory.createPieChart(      
                "Status Split-up", 
                dataset,    
                true, 
                true, 
                false);
        chart.removeLegend();
        Map<TextAttribute, Object> map = new Hashtable<TextAttribute, Object>();
        map.put(TextAttribute.WEIGHT, TextAttribute.WEIGHT_MEDIUM);
        TextTitle my_Chart_title=new TextTitle("Status Split-Up", new Font (map));
        chart.setTitle(my_Chart_title);
        PiePlot plot = (PiePlot) chart.getPlot();
        plot.setSectionPaint(0, Color.GREEN);
        plot.setSectionPaint(1, Color.RED);
        plot.setSectionPaint(2, Color.GRAY);
        plot.setSectionPaint(3, Color.ORANGE);
        plot.setSectionPaint(4, Color.YELLOW);
        plot.setLabelGenerator(new
        	    StandardPieSectionLabelGenerator("{0}: {1}, {2}"));

        TimeSeries exeSeries = new TimeSeries("Execution count");
        for(String eachValue : dateMap.keySet()) {
        	exeSeries.add(new Day(new SimpleDateFormat("MM/dd/yyyy").parse(eachValue)),new Integer(dateMap.get(eachValue)));
        }
        TimeSeriesCollection dateset = new TimeSeriesCollection();
        dateset.addSeries(exeSeries);
        JFreeChart timechart = ChartFactory.createTimeSeriesChart(
            "", // Chart
            "Date", // X-Axis Label
            "Count", // Y-Axis Label
            dateset);
        TextTitle exeTrendtitle=new TextTitle("Execution Trend", new Font (map));
        timechart.setTitle(exeTrendtitle);
        XYPlot timeplot = (XYPlot)timechart.getPlot(); 
        NumberAxis yAxis = (NumberAxis)timeplot.getRangeAxis();
        yAxis.setStandardTickUnits(NumberAxis.createIntegerTickUnits());
        timechart.removeLegend();
        
        XYLineAndShapeRenderer renderer =
                (XYLineAndShapeRenderer) timeplot.getRenderer();
        renderer.setSeriesShapesVisible(0, true);
        
        JProgressBar plannedProgress = new JProgressBar();
        plannedProgress.setBounds(53, 390, 589, 26);
        plannedProgress.setMinimum(0);
        plannedProgress.setMinimum(100);
        
        
        JScrollPane summaryScroll = new JScrollPane(summaryPanel);
        summaryScroll.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        summaryScroll.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
        summaryScroll.setPreferredSize(new Dimension(680, 1000));
        //summaryPanel.setBounds(0, 0, 680, 1000);
        
        moduleCombo.setBounds(6, 6, 154, 26);
        summaryPanel.add(moduleCombo);
        
        summaryTable = new JTable();
        summaryTable.setBounds(202, 62, 275, 147);
        summaryPanel.add(summaryTable);
        summaryTable.setGridColor(new Color(169, 169, 169));
        summaryTable.setFont(new Font("Segoe UI", Font.PLAIN, 13));
        summaryTable.setBorder(new BevelBorder(BevelBorder.RAISED, null, null, null, null));
        summaryTable.setShowHorizontalLines(true);
        summaryTable.setShowVerticalLines(true);
        summaryTable.setModel(new DefaultTableModel(
        	new Object[][] {	
        		{"Passed", null},
        		{"Failed", ""},
        		{"No Run", null},
        		{"Blocked", null},
        		{"Not Completed", null},
        		{"Total Executed", null},
        		{"Execution %", null},
        		{"Passed %", null},
        		{"Total Count", null}
        	},
        	new String[] {
        		"Type", "Count"
        	}
        ));
        //plot.setLegendLabelGenerator(new
        //	    StandardPieSectionLabelGenerator("{0}: {1}, {2}"));
        
        ChartPanel chartPanel = new ChartPanel(chart);
        chartPanel.setBounds(17, 225, 300, 195);
        summaryPanel.add(chartPanel);
        
        ChartPanel timechartPanel = new ChartPanel(timechart);
        timechartPanel.setBounds(334, 225, 300, 195);
        summaryPanel.add(timechartPanel);
        //summaryTab.add(plannedProgress);
        
        JProgressBar actualProgress = new JProgressBar();
        actualProgress.setBounds(66, 438, 568, 26);
        summaryPanel.add(actualProgress);
        actualProgress.setStringPainted(true);
        actualProgress.setMaximum(0);
        actualProgress.setMaximum(100);
        
        JLabel lblProgress = new JLabel("Progress");
        lblProgress.setBounds(6, 441, 55, 16);
        summaryPanel.add(lblProgress);
        lblProgress.setFont(new Font("SansSerif", Font.BOLD, 12));
        
        JLabel lblNewLabel = new JLabel("Execution Summary");
        lblNewLabel.setBounds(201, 31, 275, 29);
        summaryPanel.add(lblNewLabel);
        lblNewLabel.setBorder(new BevelBorder(BevelBorder.RAISED, null, null, null, null));
        lblNewLabel.setFont(new Font("Segoe UI", Font.BOLD, 23));
        lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
        
        summaryTable.getColumnModel().getColumn(0).setPreferredWidth(157);
        summaryTable.getColumnModel().getColumn(1).setPreferredWidth(114);
        
        summaryTable.setValueAt(statusMap.get("Passed"), 0, 1);
        summaryTable.setValueAt(statusMap.get("Failed"), 1, 1);
        summaryTable.setValueAt(statusMap.get("No Run"), 2, 1);
        summaryTable.setValueAt(statusMap.get("Blocked"), 3, 1);
        summaryTable.setValueAt(statusMap.get("Not Completed"), 4, 1);
        summaryTable.setValueAt(statusMap.get("Passed") + statusMap.get("Failed"), 5, 1);
        summaryTable.setValueAt(Double.toString(Math.round((new Double(statusMap.get("Passed") + statusMap.get("Failed")))/new Double(table.getRowCount()) * 100)) + "%", 6, 1);
        summaryTable.setValueAt(Double.toString(Math.round((new Double(statusMap.get("Passed"))/new Double(table.getRowCount()) * 100))) + "%", 7, 1);
        summaryTable.setValueAt(TotalCount, 8, 1);
        
        moduleCombo.addItemListener(new ItemListener() {
			@Override
			public void itemStateChanged(ItemEvent e) {
				for(int eachFilter = 0; eachFilter < tableModel.getColumnCount(); eachFilter++) {
					filter.getFilterEditor(eachFilter).resetFilter();
				}
				statusMap.put("Passed", 0);
				statusMap.put("Failed", 0);
				statusMap.put("No Run", 0);
				statusMap.put("Blocked", 0);
				statusMap.put("Not Completed", 0);
				dateMap.clear();
				int TotalCount = 0;
				for(int eachRow = 0; eachRow<table.getRowCount(); eachRow++) {
					try {
						if(e.getItem().toString().equals(table.getValueAt(eachRow, 0)) || e.getItem().toString().equals("All")) {
							String tcStatus = table.getValueAt(eachRow, 2).toString();
							if(!tcStatus.equals("")) {
								statusMap.put(tcStatus, statusMap.get(tcStatus) + 1); 
							}
							else {
								statusMap.put("No Run", statusMap.get("No Run") + 1);
							}
							
							String lastExecutedOn = table.getValueAt(eachRow, 7).toString();
							if(!lastExecutedOn.equals("")) {
								if(!dateMap.containsKey(lastExecutedOn.split(" ")[0])) {
									dateMap.put(lastExecutedOn.split(" ")[0], 0);
								}
								dateMap.put(lastExecutedOn.split(" ")[0], dateMap.get(lastExecutedOn.split(" ")[0]) + 1); 
							}
							exeSeries.clear();
							for(String eachValue : dateMap.keySet()) {
					        	exeSeries.add(new Day(new SimpleDateFormat("MM/dd/yyyy").parse(eachValue)),new Integer(dateMap.get(eachValue)));
					        }
							TotalCount++;
						}
					}
					catch(Exception ex) {
						ex.printStackTrace();
					}
				}
				summaryTable.setValueAt(statusMap.get("Passed"), 0, 1);
				summaryTable.setValueAt(statusMap.get("Failed"), 1, 1);
				summaryTable.setValueAt(statusMap.get("No Run"), 2, 1);
				summaryTable.setValueAt(statusMap.get("Blocked"), 3, 1);
				summaryTable.setValueAt(statusMap.get("Not Completed"), 4, 1);
				summaryTable.setValueAt(statusMap.get("Passed") + statusMap.get("Failed"), 5, 1);
				summaryTable.setValueAt(Double.toString(Math.round((new Double(statusMap.get("Passed") + statusMap.get("Failed")))/new Double(TotalCount) * 100)) + "%", 6, 1);
				summaryTable.setValueAt(Double.toString(Math.round((new Double(statusMap.get("Passed"))/new Double(TotalCount) * 100))) + "%", 7, 1);
				summaryTable.setValueAt(TotalCount, 8, 1);
				
				dataset.setValue( "Passed" , new Double( statusMap.get("Passed") ) );             
			    dataset.setValue( "Failed" , new Double( statusMap.get("Failed") ) );             
			    dataset.setValue( "No Run" , new Double( statusMap.get("No Run") ) );             
			    dataset.setValue( "Blocked" , new Double( statusMap.get("Blocked") ) );
			    dataset.setValue( "Not Completed" , new Double( statusMap.get("Not Completed") ) );
			    actualProgress.setValue((int)Double.parseDouble(summaryTable.getValueAt(6, 1).toString().replace("%", "")));
				chart.fireChartChanged();
				timechart.fireChartChanged();
				chartPanel.revalidate();
				chartPanel.repaint();
				timechartPanel.revalidate();
				timechartPanel.repaint();
			}
		});
        actualProgress.setValue((int)Double.parseDouble(summaryTable.getValueAt(6, 1).toString().replace("%", "")));
        summaryScroll.setBounds(0, 0, 670, 485);
        summaryTab.add(summaryScroll);
        
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
	
	private static void writeToExcel(JTable table, File path) throws FileNotFoundException, IOException {
		Workbook wb = new HSSFWorkbook(); // Excell workbook
		Sheet sheet = wb.createSheet(); // WorkSheet
		Row row = sheet.createRow(1); // Row created at line 3
		TableModel model = table.getModel(); // Table model

		Row headerRow = sheet.createRow(0); // Create row at line 0
		for (int headings = 0; headings < model.getColumnCount(); headings++) { // For
																				// each
																				// column
			headerRow.createCell(headings).setCellValue(model.getColumnName(headings));// Write
																						// column
																						// name
		}

		for (int rows = 0; rows < model.getRowCount(); rows++) { // For each
																	// table row
			for (int cols = 0; cols < table.getColumnCount(); cols++) { // For
																		// each
																		// table
																		// column
				row.createCell(cols).setCellValue(model.getValueAt(rows, cols).toString()); // Write
																							// value
			}
			// Set the row to the next one in the sequence
			row = sheet.createRow((rows + 2));
		}
		FileOutputStream fileout = new FileOutputStream(path);
		wb.write(fileout);// Save the file
		fileout.flush();//Flush the output stream
		fileout.close();// Close the output Stream
	}
}
