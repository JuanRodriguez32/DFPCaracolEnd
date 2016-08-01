package gui;

import java.awt.Event;
import java.awt.EventQueue;
import java.awt.TextArea;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Properties;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JTextArea;
import javax.swing.KeyStroke;

import world.DFPCaracolWorld;
import world.DFPCaracolWorld2;

import com.google.api.ads.common.lib.auth.OfflineCredentials;
import com.google.api.ads.common.lib.auth.OfflineCredentials.Api;
import com.google.api.ads.dfp.axis.factory.DfpServices;
import com.google.api.ads.dfp.lib.client.DfpSession;
import com.google.api.client.auth.oauth2.Credential;

public class GUI implements ActionListener{

	private JFrame frmDfpReporterFor;
	public final static String END = "End";
	public final static String PATH = "Path";
	public final static String START = "Start";
	private static final String SEARCH = "SEARCH";
	private TextArea textArea;
	File path2 = null;
	File path3 = null;
	File file2 = null;
	File p = null;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					GUI window = new GUI();
					window.frmDfpReporterFor.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public GUI() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmDfpReporterFor = new JFrame();
		frmDfpReporterFor.setResizable(false);
		frmDfpReporterFor.setTitle("DFP Reporter For Caracol Next");
		frmDfpReporterFor.setIconImage(Toolkit.getDefaultToolkit().getImage(GUI.class.getResource("/gui/iconnext.jpg")));
		frmDfpReporterFor.setBounds(100, 100, 498, 410);
		frmDfpReporterFor.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmDfpReporterFor.getContentPane().setLayout(null);
		
		JLabel label = new JLabel("");
		label.setIcon(new ImageIcon(GUI.class.getResource("/gui/nban.jpg")));
		label.setBounds(0, 0, 490, 77);
		frmDfpReporterFor.getContentPane().add(label);
		
		JButton btnStart = new JButton("Start");
		btnStart.setActionCommand(START);
		btnStart.addActionListener(this);
		btnStart.setBounds(383, 317, 89, 23);
		frmDfpReporterFor.getContentPane().add(btnStart);
		
		textArea = new TextArea();
		textArea.setBounds(10, 88, 351, 239);
		frmDfpReporterFor.getContentPane().add(textArea);
		
		JMenuBar menuBar = new JMenuBar();
		frmDfpReporterFor.setJMenuBar(menuBar);
		
		JMenu mnArchivo = new JMenu("Archivo");
		menuBar.add(mnArchivo);
		
		JMenuItem mntmOutputPath = new JMenuItem("Output Path");
		mntmOutputPath.setActionCommand(PATH);
		mntmOutputPath.addActionListener(this);
		mntmOutputPath.setMnemonic('O');	
		mntmOutputPath.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_P, Event.CTRL_MASK));
		mnArchivo.add(mntmOutputPath);
		
		JMenuItem mntmSalir = new JMenuItem("Salir");
		mntmSalir.setMnemonic('S');	
		mntmSalir.setActionCommand(END);
		mntmSalir.addActionListener(this);
		mntmSalir.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_F4, Event.ALT_MASK));
		
		mnArchivo.add(mntmSalir);
		
		JMenu mnAbout = new JMenu("Search");
		menuBar.add(mnAbout);
		
		JMenuItem mntmAbouthelp = new JMenuItem("By id");
		
		mnAbout.add(mntmAbouthelp);
		mntmAbouthelp.setActionCommand(SEARCH);
		mntmAbouthelp.addActionListener(this);
	}
	
	public void pintarTextArea(String text)
	{
		textArea.append(text + "\n");
		textArea.repaint();
		textArea.revalidate();
		
	}

	public void actionPerformed(ActionEvent arg0) {
		// TODO Auto-generated method stub
		String commd = arg0.getActionCommand();
		if(commd.equals(END))
		{
			System.exit(0);
		}
		else if(commd.equals(PATH))
		{
			try
			{	
				
				Properties prop = new Properties();
				Properties prop2 = new Properties();
				File path = null;
				
				String title = "";
//			 Change to your file location.
				for(int i = 0; i<4; i++)
				{
					JFileChooser chooser = new JFileChooser();
				    chooser.setCurrentDirectory(new java.io.File("."));
				    if( i == 0)
				    {
				    	   title = "Select destiny Parent path";
				    	   chooser.setDialogTitle(title);
						    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
						    if(chooser.showOpenDialog(chooser)== JFileChooser.APPROVE_OPTION)
						    {
						    	path = chooser.getSelectedFile();
						    	
						    }
						    prop.setProperty("ParentPath", path.getPath()+"\\");
				    }
				    
				    if( i == 1)
				    {
				    	  title = "Select destiny report path";
				    	  chooser.setDialogTitle(title);
						    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
						    if(chooser.showOpenDialog(chooser)== JFileChooser.APPROVE_OPTION)
						    {
						    	path = chooser.getSelectedFile();
						    }
						    prop.setProperty("ReportPath", path.getPath()+"\\");
				    }
				  
				    if( i == 2)
				    {
				    	   title = "Select destiny PreDesignFormat path";
				    	   chooser.setDialogTitle(title);
						    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
						    if(chooser.showOpenDialog(chooser)== JFileChooser.APPROVE_OPTION)
						    {
						    	path = chooser.getSelectedFile();
						    	path3 = chooser.getSelectedFile();
						    	
						    }
						    prop.setProperty("PlantPath", path.getPath()+"\\");
						    
				    }
				  
				    if( i == 3)
				    {
				    	   title = "Select JSON KEY";
				    	   chooser.setDialogTitle(title);
						    chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
						    if(chooser.showOpenDialog(chooser)== JFileChooser.APPROVE_OPTION)
						    {
						    	path = chooser.getSelectedFile();
						    	path2 = chooser.getSelectedFile();
						    }
						    prop2.setProperty("api.dfp.jsonKeyFilePath", path.getPath() );
						    prop2.setProperty("api.dfp.applicationName", "JUAN" );
						    String code = JOptionPane.showInputDialog("dfp nteworkCode");
						    prop2.setProperty("api.dfp.networkCode", code );
						   
						   
						 
						    
				    }
				    
				    
						
				}
		    
				File file = new File(path3.getPath()+"\\config.properties");
				file2 = new File(path3.getPath()+"\\ads.properties");
				FileOutputStream fileOut = new FileOutputStream(file);
				FileOutputStream fileout = new FileOutputStream(file2);
				prop.store(fileOut, "Paths");
				prop2.store(fileout, "Ads Properties");
				p = file;
				fileOut.close();
				fileout.close();
				
			}catch (Exception e)
			{
				e.printStackTrace();
			}
			
		}
		else if ( commd.equals(START))
		{
			try
			{
				
				
			  Credential oAuth2Credential = new OfflineCredentials.Builder()
		        .forApi(Api.DFP).fromFile(file2)
		        .build()
		        .generateCredential();

		
		    
		    
		    // Construct a DfpSession.
		    DfpSession session = new DfpSession.Builder()
		        .fromFile(file2)
		        .withOAuth2Credential(oAuth2Credential)
		        .build();

		    DfpServices dfpServices = new DfpServices();
			
			DFPCaracolWorld world = new DFPCaracolWorld(dfpServices, session, textArea, this , p);
			}
			catch(Exception e)
			{
				e.getMessage();
				e.printStackTrace();
			}
			
		}
		
		else if(commd.equals(SEARCH))
		{
			System.out.println("entra");
			try
			{
				
				int id = Integer.parseInt(JOptionPane.showInputDialog("Input ID"));
			  Credential oAuth2Credential = new OfflineCredentials.Builder()
		        .forApi(Api.DFP).fromFile(file2)
		        .build()
		        .generateCredential();

		
		    
		    
		    // Construct a DfpSession.
		    DfpSession session = new DfpSession.Builder()
		        .fromFile(file2)
		        .withOAuth2Credential(oAuth2Credential)
		        .build();

		    DfpServices dfpServices = new DfpServices();
			
			DFPCaracolWorld2 world = new DFPCaracolWorld2(dfpServices, session, textArea, this , id , p);
			}
			catch(Exception e)
			{
				e.getMessage();
				e.printStackTrace();
			}
		}
		
	}
}
