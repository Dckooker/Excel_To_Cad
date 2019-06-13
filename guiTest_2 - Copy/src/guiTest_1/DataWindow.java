package guiTest_1;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JButton;
import java.awt.SystemColor;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedWriter;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
//import java.io.File;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.io.Writer;
import java.util.ArrayList;

import javax.swing.JTextField;
import javax.swing.SwingConstants;
import java.awt.Font;
import javax.swing.JToggleButton;
import java.awt.Color;
import javax.swing.JCheckBoxMenuItem;
import javax.swing.JCheckBox;
import javax.swing.JList;
import javax.swing.JRadioButton;


 
public class DataWindow {

	private JFrame frame;
	private JTextField textField;
	static String pathway= " ";
	private static ArrayList<String> CADscript = new ArrayList<String>();
	private static boolean spacing = true;
	// Launch the application.

	//   X:\DKooker\Programs\NPtest.xlsx
	public static void main(String[] args) throws IOException  {
	
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					DataWindow window = new DataWindow();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public DataWindow() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 450, 300);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		JLabel userInstructionLabel = new JLabel("Paste excel pathway below");
		userInstructionLabel.setFont(new Font("Tahoma", Font.PLAIN, 12));
		userInstructionLabel.setHorizontalAlignment(SwingConstants.CENTER);
		userInstructionLabel.setOpaque(true);
		userInstructionLabel.setBackground(SystemColor.activeCaptionBorder);
		userInstructionLabel.setBounds(130, 22, 187, 23);
		frame.getContentPane().add(userInstructionLabel);
		
		textField = new JTextField();
		textField.setBackground(SystemColor.activeCaption);
		textField.setBounds(76, 56, 291, 44);
		frame.getContentPane().add(textField);
		textField.setColumns(10);
		
		JButton goButton = new JButton("press for script");
		goButton.setFont(new Font("Tahoma", Font.PLAIN, 12));
		goButton.setBounds(161, 111, 125, 37);
		frame.getContentPane().add(goButton);
		
		JButton fineTuneUp = new JButton("^");
		fineTuneUp.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
			}
		});
		fineTuneUp.setBounds(351, 192, 57, 23);
		frame.getContentPane().add(fineTuneUp);
		
		JButton fineTuneDown = new JButton("V");
		fineTuneDown.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			}
		});
		fineTuneDown.setBounds(351, 213, 57, 23);
		frame.getContentPane().add(fineTuneDown);
		
		JLabel lblFineTuning = new JLabel("Fine Tune Spacing");
		lblFineTuning.setOpaque(true);
		lblFineTuning.setBackground(Color.LIGHT_GRAY);
		lblFineTuning.setHorizontalAlignment(SwingConstants.CENTER);
		lblFineTuning.setBounds(316, 167, 108, 14);
		frame.getContentPane().add(lblFineTuning);
		
		JLabel lblCadScript = new JLabel("CAD_script_creator");
		lblCadScript.setFont(new Font("Trebuchet MS", Font.ITALIC, 9));
		lblCadScript.setBounds(0, 0, 102, 14);
		frame.getContentPane().add(lblCadScript);
		
		JCheckBox AutoSpacingCheck = new JCheckBox("Auto Spacing");
		AutoSpacingCheck.setBackground(Color.LIGHT_GRAY);
		AutoSpacingCheck.setHorizontalAlignment(SwingConstants.CENTER);
		AutoSpacingCheck.setBounds(316, 137, 108, 23);
		frame.getContentPane().add(AutoSpacingCheck);
		
		JLabel tuningCounter = new JLabel("");
		tuningCounter.setOpaque(true);
		tuningCounter.setBackground(Color.LIGHT_GRAY);
		tuningCounter.setHorizontalAlignment(SwingConstants.CENTER);
		tuningCounter.setBounds(299, 213, 42, 14);
		frame.getContentPane().add(tuningCounter);
		
		JLabel lblSpacing = new JLabel("Spacing");
		lblSpacing.setHorizontalAlignment(SwingConstants.CENTER);
		lblSpacing.setBounds(295, 196, 46, 14);
		frame.getContentPane().add(lblSpacing);
		
		JLabel lblNewLabel = new JLabel("mounting type");
		lblNewLabel.setBackground(Color.LIGHT_GRAY);
		lblNewLabel.setOpaque(true);
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel.setBounds(20, 178, 82, 14);
		frame.getContentPane().add(lblNewLabel);
		
		JRadioButton AdhesiveRadio = new JRadioButton("Adhesive");
		AdhesiveRadio.setHorizontalAlignment(SwingConstants.CENTER);
		AdhesiveRadio.setBounds(20, 192, 82, 23);
		frame.getContentPane().add(AdhesiveRadio);
		
		
		JRadioButton ScrewOnRadio = new JRadioButton("Screw-on");
		ScrewOnRadio.setHorizontalAlignment(SwingConstants.CENTER);
		ScrewOnRadio.setBounds(6, 213, 114, 23);
		frame.getContentPane().add(ScrewOnRadio);
		 AdhesiveRadio.setSelected(true);
		
		//Radio events
		AdhesiveRadio.addActionListener(new ActionListener()
		{
		    public void actionPerformed(ActionEvent e)
		    {
		       if(AdhesiveRadio.isEnabled()== true) {
		    	   ScrewOnRadio.setSelected(false);
		       }	
		    }
		});
		
		ScrewOnRadio.addActionListener(new ActionListener()
		{
		    public void actionPerformed(ActionEvent e)
		    {
		       if(ScrewOnRadio.isEnabled()== true)  {
		    	   AdhesiveRadio.setSelected(false);
		       }    	
		    }
		});
		
		//Button event//////////////////////////////////////////
		goButton.addActionListener(new ActionListener()
		{
		    public void actionPerformed(ActionEvent e)
		    {
		       
		    	pathway = textField.getText();
	        	ReadExcel exl = new ReadExcel(pathway);
	        	int row = exl.getRowCount("Nameplates");
	        	System.out.println("Total rows: " + (row -1) );
	        	System.out.println("Line count: " + exl.getLineCount(exl) );
		        	
		        	//before each nameplate set
		        	CADscript.add("zoom all");
	        		CADscript.add("-layer s ETCH");
	        		CADscript.add(" ");
	        		////////////////////////////////////////////////////////////////////
		        	for(int i =2; i < row+1;i++) {
		        		
		        		try {
		        		String Panel=  exl.getData("Nameplates", "Panel"   ,i); 
		        		String NPNum=  exl.getData("Nameplates", "NP #"    ,i);
	    				String Line1=  exl.getData("Nameplates", "Line 1"  ,i);
	    				String Line2=  exl.getData("Nameplates", "Line 2"  ,i);
	    				String Line3=  exl.getData("Nameplates", "Line 3"  ,i);
	    				String Line4=  exl.getData("Nameplates", "Line 4"  ,i);
	    				String Size1=  exl.getData("Nameplates", "Size L1" ,i);
	    				String Size2=  exl.getData("Nameplates", "Size L2" ,i);
	    				String Size3=  exl.getData("Nameplates", "Size L3" ,i); 
	    				String Size4=  exl.getData("Nameplates", "Size L4" ,i);
	    				String Height=  exl.getData("Nameplates", "Height" ,i);
	    				String Width=  exl.getData("Nameplates", "Width"   ,i);
		        		
	    				double NPwidth = Double.parseDouble(Width);
	    				double NPheight = Double.parseDouble(Height);
	    				double NPpnaelNum = Double.parseDouble(Panel);
	    				int panelLengthX= (int)(11/NPheight);
	    				int panelLengthY= (int)(23/NPwidth);
	    				double totalPerSheet= panelLengthX*panelLengthY;
	    				int hor = 0;
	    				int vert = 0;
		        	 
	        			CADscript.add("-layer s ETCH");
	        			CADscript.add(" ");
	        			
	        			//get NamePlate
	        			hor =  (int) ( (NPwidth * 10000 / 2) + (NPwidth * 10000 * ((i - 2) % panelLengthY)) ); 
	        			vert = (int) (NPheight * 10000 * ((i-2)/ panelLengthY));
	        		    CADscript.addAll(exl.getCadStrings(hor, vert,NPheight , spacing, exl));
	        		    CADscript.add("-layer s TEXT");
	        		    CADscript.add(" ");
	        		     
	        		    //get NamePlate Number
	        		    hor = (int) (500 + (NPwidth * 10000 * ((i - 2) % panelLengthY)) ); 
	        		    vert = (int) (NPheight * 10000 * ((i-2)/ panelLengthY)) + 500;
	        		    CADscript.add("text j TL " + hor + ",-" + vert + " 1200 0 " + (int)NPpnaelNum + "/" + (i-1) );
	        		    
	        		    //get Screw-on Circles
	        		    
		        		}
		        		catch(NullPointerException c) { } 
		        	}
		        	
		        	
		        //draw border lines for whole program
		        	
		        	
		        	CADscript.addAll(exl.getBorderLines( exl ));
		       
		        	
		        	//write to text file first
		        	PrintWriter writer;
					try {
						writer = new PrintWriter ("NPtext.txt");
						 
						 for(int i = 0; i <= CADscript.size()-1; i ++) {
							 writer.println(CADscript.get(i));
				         }
						 
						 writer.close ();       
					} catch (FileNotFoundException d) { }
		        		
		        	//creates CAD script from text file
		        	exl.makeScriptFile();
		    }
		});
	     //   X:\DKooker\Programs\NPtest.xlsx 
		//////////////////////////////////////////////////////////////
	}
}
