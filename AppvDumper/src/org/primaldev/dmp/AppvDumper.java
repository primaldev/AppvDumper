package org.primaldev.dmp;

import java.awt.EventQueue;







import javax.swing.JFrame;
import javax.swing.JMenuBar;
import javax.swing.JMenu;
import javax.swing.JMenuItem;
import javax.swing.AbstractAction;

import java.awt.event.ActionEvent;

import javax.swing.Action;

import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;

import org.eclipse.swt.widgets.Shell;

import java.awt.TextArea;
import java.awt.BorderLayout;
import java.io.IOException;
import java.awt.Font;

public class AppvDumper {

	private JFrame frmAppvDumper;
	private final Action action = new SwingAction();
	Display d;
	Shell s;
	TextArea infoDump;
	private final Action createDoc = new CreateDocAction();
	private final Action showHelp = new ShowHelp();
	ParseAppv parseAppv;
	JMenuItem mntmCreateDoc;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					AppvDumper window = new AppvDumper();
					window.frmAppvDumper.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public AppvDumper() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmAppvDumper = new JFrame();
		frmAppvDumper.setTitle("PrimalDev Appv Dumper");
		frmAppvDumper.setBounds(100, 100, 541, 406);
		frmAppvDumper.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		d = new Display();
	    s = new Shell(d);
	    s.setSize(400, 400);
		
		JMenuBar menuBar = new JMenuBar();
		frmAppvDumper.setJMenuBar(menuBar);
		
		JMenu mnFile = new JMenu("File");
		menuBar.add(mnFile);
		
		JMenuItem mntmOpen = new JMenuItem("Open");
		mntmOpen.setAction(action);
		mnFile.add(mntmOpen);
		
		 mntmCreateDoc = new JMenuItem("Create Doc");
		mntmCreateDoc.setAction(createDoc);
		mnFile.add(mntmCreateDoc);
		mntmCreateDoc.setEnabled(false);
		
		JMenu mnHelp = new JMenu("Help");
		menuBar.add(mnHelp);
		
		JMenuItem mntmAbout = new JMenuItem("About");
		mntmAbout.setAction(showHelp);
		mnHelp.add(mntmAbout);
		
		infoDump = new TextArea();
		infoDump.setText("Choose File > Open and select an Appv File.");
		infoDump.setFont(new Font("Monospaced", Font.PLAIN, 12));
		
		frmAppvDumper.getContentPane().add(infoDump, BorderLayout.CENTER);
		
	}

	private class SwingAction extends AbstractAction {
		public SwingAction() {
			putValue(NAME, "Open");
			putValue(SHORT_DESCRIPTION, "Open Appv File");
		}
		public void actionPerformed(ActionEvent e) {
			FileDialog fd = new FileDialog(s, SWT.OPEN);
	        fd.setText("Open");
	        fd.setFilterPath("C:/");
	        String[] filterExt = { "*.appv"};
	        fd.setFilterExtensions(filterExt);
	        String selected = fd.open();
	       
	         parseAppv = new ParseAppv(selected);
	        infoDump.setText(parseAppv.getSummaryText() + "\nChoose File -> Create Doc to generate a Word Document \n");
	        mntmCreateDoc.setEnabled(true);
	        
		}
	}
	
	private class CreateDocAction extends AbstractAction {
		public CreateDocAction() {
			putValue(NAME, "Create Doc");
			putValue(SHORT_DESCRIPTION, "Some short description");
		}
		public void actionPerformed(ActionEvent e) {
			FileDialog fd = new FileDialog(s, SWT.SAVE);
	        fd.setText("Save");
	        fd.setFilterPath("C:/");
	        String[] filterExt = { "*.doc"};
	        fd.setFilterExtensions(filterExt);
	        fd.setFileName(parseAppv.getAppvDocName());
	        String path=fd.open();
	        if (path!=null){
	        	try {
					parseAppv.saveDoc(path);
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
	        }
		}
	}
	private class ShowHelp extends AbstractAction {
		public ShowHelp() {
			putValue(NAME, "About");
			putValue(SHORT_DESCRIPTION, "Generate a Word document.");
		}
		public void actionPerformed(ActionEvent e) {
			AboutDialog aboutDialog = new AboutDialog();
			aboutDialog.setVisible(true);
			
			
			
		}
	}
}
