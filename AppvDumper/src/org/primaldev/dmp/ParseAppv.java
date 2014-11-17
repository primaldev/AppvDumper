package org.primaldev.dmp;

import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.Document;


import org.primaldev.appv.AppvAppPath;
import org.primaldev.appv.AppvFileTypeAssociation;
import org.primaldev.appv.AppvShellCommand;
import org.primaldev.appv.AppvShortCut;
import org.primaldev.appv.ParseAppvManifest;
import org.w3c.dom.NodeList;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;

import word.api.interfaces.IDocument;
import word.w2004.Document2004;
import word.w2004.elements.BreakLine;
import word.w2004.elements.Heading1;
import word.w2004.elements.Heading2;
import word.w2004.elements.PageBreak;
import word.w2004.elements.Paragraph;
import word.w2004.elements.Table;
import word.w2004.elements.tableElements.TableEle;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;


public class ParseAppv {
	//This is transformed from a groovy script

	
	String appvDocName;

	IDocument myDoc;	
	
	String summaryText = "";
	
	
	static String appvMountRoot;

	ParseAppvManifest appvManifest;
	
	public ParseAppv(String appvFileName) {


		extractAppvInfo(appvFileName);

	/////////////////////////////////////////////////////
	// Documents must be saved as word 2003 xml format //
	/////////////////////////////////////////////////////


	//read template

/*	

	try {
		BufferedReader br = new BufferedReader(new FileReader(templateAppvDocName));
		try {
			StringBuilder sb = new StringBuilder();
			String line = br.readLine();

			while (line != null) {
				sb.append(line);
				sb.append(System.lineSeparator());
				line = br.readLine();
			}
			xmlTemplate = sb.toString();
		} finally {
			br.close();
		}
	} catch (FileNotFoundException e1) {
		// TODO Auto-generated catch block
		e1.printStackTrace();
	} catch (IOException e1) {
		// TODO Auto-generated catch block
		e1.printStackTrace();
	}

*/
	// Replace template values
	myDoc = new Document2004();
	myDoc.addEle(Heading1.with(appvManifest.getDisplayName()).create());
	appvDocName = appvManifest.getDisplayName() + ".doc";

	
	
	//xmlTemplate = replacePh(xmlTemplate, "ppmAppName", projectName);
	replacePh("Mount Root: ", appvMountRoot);
	replacePh("COM Mode: ", appvManifest.getAppvCOMMode());
	replacePh("COM Mode Inprocess:", String.valueOf(appvManifest.isComModeInProcessEnabled()));
	replacePh("COM Mode OutOfProcess: ", String.valueOf(appvManifest.isComModeOutOfProcessEnabled()));
	replacePh("Package Description: ", appvManifest.getAppVPackageDescription());
	replacePh("PackageId: ", appvManifest.getAppvPackageId());
	replacePh("VersionId: ", appvManifest.getAppvVersionId());
	replacePh("Publisher: ",appvManifest.getAppvPublisher());	
	replacePh("Description: ", appvManifest.getDescription());
	replacePh("DisplayName: ", appvManifest.getDisplayName());
	replacePh("Language: ", appvManifest.getLanguage());
	replacePh("Name: ", appvManifest.getName());
	replacePh("OS Max Version Tested: ", appvManifest.getoSMaxVersionTested());
	replacePh("OS Min Version: ", appvManifest.getoSMinVersion());
	replacePh("Package Version: ", appvManifest.getPackageVersion());
	replacePh("Publisher Display Name: ", appvManifest.getPublisherDisplayName());
	replacePh("Sequencing Station Processor: ", appvManifest.getSequencingStationProcessorArchitecture());
	replacePh("Full VFS WriteMode: ", String.valueOf(appvManifest.getFullVFSWriteMode()));

	Table appPathTable = new Table();
	appPathTable.addTableEle(TableEle.TH, "App ID", "Application Path", "Prefix");
	for (AppvAppPath appvPath : appvManifest.getAppvAppPaths()){
		appPathTable.addTableEle(TableEle.TD,appvPath.getApplicationId(),appvPath.getApplicationPath(),appvPath.getPATHEnvironmentVariablePrefix());
	}
	
	myDoc.addEle(Heading2.with("Applications").create());
	replacePh( appPathTable);
	
	Table appvSc = new Table();
	appvSc.addTableEle(TableEle.TH,"Link File", "Target", "Working Dir", "Arguments", "Icon" );

	List<AppvShortCut> bufest = appvManifest.getAppvShortcuts();
	
	for (AppvShortCut appvShortcut : bufest){
		appvSc.addTableEle(TableEle.TD,appvShortcut.getFile(),appvShortcut.getTarget(),appvShortcut.getWorkingDirectory(),appvShortcut.getArguments(),appvShortcut.getIcon());
		
	}
	 myDoc.addEle(PageBreak.create());
	 myDoc.addEle(Heading2.with("Shortcuts").create());
	replacePh( appvSc);
	
	

	Table appFtaTable = new Table();
	appFtaTable.addTableEle(TableEle.TH, "Name","ProgramId", "Discription","Mime");

	for (AppvFileTypeAssociation appvFta :appvManifest.getAppvFileTypeAssociations()) {
		
		
		appFtaTable.addTableEle(TableEle.TF,appvFta.getName(),appvFta.getProgIdName(), appvFta.getProgIdDescription(), String.valueOf(appvFta.isMimeAssociation()));
		if(appvFta.isShellCommds()){
			for (AppvShellCommand appvCmda : appvFta.getAppvShellCommands()) {
				
				if(appvCmda.getFriendlyName() == null || appvCmda.getFriendlyName().length() < 1){
					//print "Shell name " + appvCmda.getName() + " : " + appvCmda.getCommandLine() + " : " + appvCmda.getApplicationId() + "\n"
					appFtaTable.addTableEle(TableEle.TD, "", appvCmda.getName(), appvCmda.getCommandLine(), appvCmda.getApplicationId());
				}else if (appvCmda.getFriendlyName() != null){
					//print "Shell name " + appvCmda.getName() + " : " + appvCmda.getFriendlyName() + "\n"
				    
					appFtaTable.addTableEle(TableEle.TD, "Shell",appvCmda.getName(), appvCmda.getFriendlyName().replaceAll("[^\\w\\s\\-_]", ""), "" );
				}
				
				//appFtaTable.addTableEle(TableEle.TF, "Shell", appvCmda.getName(), "dsaf", "asdf");
			}
		
		}

	}
	 myDoc.addEle(PageBreak.create());
	 myDoc.addEle(Heading2.with("File Type Associations").create());
	 replacePh(appFtaTable);


	}
	
	
	

	public String getAppvDocName() {
		return appvDocName;
	}




	public void saveDoc(String docPath) throws IOException {

		//File tempFile = File.createTempFile("ppm", ".doc");
		File tempFile = new File(docPath);
		FileWriter fileWriter = new FileWriter( tempFile, true);
		
		BufferedWriter bw = new BufferedWriter(fileWriter);
		bw.write(myDoc.getContent());
		bw.close();
		
		//Files.copy(tempFile.toPath(), new File(docpath).toPath(), StandardCopyOption.REPLACE_EXISTING );
		//tempFile.delete();

	}
	
	

	

	private void extractAppvInfo(String appvPackage) {
		ZipFile zipFile = null;
		try {
			zipFile = new ZipFile(appvPackage);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		Enumeration<? extends ZipEntry> entries = zipFile.entries();
		while(entries.hasMoreElements()){
			ZipEntry entry = entries.nextElement();
			//Get xml data:
			if (entry.getName().equalsIgnoreCase("FilesystemMetadata.xml")){
				try {
					parseFileMetaData(zipFile.getInputStream(entry));
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			if (entry.getName().equalsIgnoreCase("AppxManifest.xml")){
				//parseAppxManifest(zipFile.getInputStream(entry));
				ParseAppvManifest parseAppvManifest = null;
				try {
					parseAppvManifest = new ParseAppvManifest(zipFile.getInputStream(entry));
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				this.appvManifest = parseAppvManifest;
			}

			//InputStream stream = zipFile.getInputStream(entry);
		}

	}



	private void parseFileMetaData(InputStream stream) {
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		Document doc = null;
		try {
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			doc = dBuilder.parse(stream);
			
			stream.close();
		} catch (ParserConfigurationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		NodeList nList = doc.getElementsByTagName("Filesystem");

		
			Element eElement = (Element) nList.item(0);
			appvMountRoot=(eElement.getAttribute("Root"));
		

	}

	
	
	
	public String getSummaryText() {
		return summaryText;
	}

	public void setSummaryText(String summaryText) {
		this.summaryText = summaryText;
	}

	private void replacePh( String placeHolder, String value) {
		myDoc.addEle(Paragraph.with(placeHolder + value));
		myDoc.addEle(BreakLine.times(1).create());
		
		summaryText +=String.format("%-30s %s", placeHolder , value)  + "\n";
	}

	private void replacePh(Table tbl){
		
		myDoc.addEle(tbl);
		myDoc.addEle(BreakLine.times(1).create());
	}
	

}
