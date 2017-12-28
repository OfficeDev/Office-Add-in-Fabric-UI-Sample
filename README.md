# Office Add-in Fabric UI Sample
This is a Word task pane add-in sample that uses [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric). It inserts citations into the body of the document using the Office.js APIs.

For general information about using Fabric, see [Using Office UI Fabric in Office Add-Ins](https://msdn.microsoft.com/EN-US/library/office/mt450443.aspx). 

![Screenshot of running sample](/Images/citationssample.png)

## Try it out
### Get web application running
1. Open command prompt and navigate to project directory
2. Run 'npm install' in root directory
3. Run 'gulp serve-static' to run the node web server

### Deploy add-in manifest
The simplest way to deploy and test your add-in is to copy the files to a network share.

1.  Create a folder on a network share (for example, \\\MyShare\addins) and copy all the files in the Code Editor folder. 
2.  Edit the <SourceLocation> element of the manifest file so that it points to the share location from step 1. 
3.  Copy the manifest (manifest.xml) to a network share (for example, \\\MyShare\addins).
4.  Add the share location that contains the manifest as a trusted app catalog in Excel.

    a.  Launch Excel and open a blank spreadsheet.  
    
    b.  Choose the **File** tab, and then choose **Options**.
    
    c.  Choose **Trust Center**, and then choose the **Trust Center Settings** button.
    
    d.  Choose **Trusted Add-in Catalogs**.
    
    e.  In the **Catalog Url** box, enter the path to the network share you created in step 3, and then choose **Add Catalog**.
    
   f.  Select the **Show in Menu** check box, and then choose **OK**. A message appears to inform you that your settings will be applied the next time you start Office. 
        

    
### Test add-in in Word

    a.  In the **Insert tab** in Excel 2016, choose **My Add-ins**. 
    
    b.  In the **Office Add-ins** dialog box, choose **Shared Folder**.
    
    c.  Choose **Citations Sample**>**Insert**. The add-in opens in a task pane and shows the add-in. 
      


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
