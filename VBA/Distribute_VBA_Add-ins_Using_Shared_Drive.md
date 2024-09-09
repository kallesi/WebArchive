
> "Work smarter, not harder."

Improving efficiency & decreasing errors are vital requirements for any company's ongoing success in the business world.

Providing & maintaining an Add-On in a corporate environment can be a painstaking & headache-enducing experience for many reasons.

1. First, you have to figure out some way to distribute the Add-On, be it by email, flash drive, shared network folder, etc.
2. Then, you have to worry about the users correctly installing it (we'll cover what I mean be "correctly" later on) or you have to run around to each person's computer and install it yourself.
3. Finally, what happens when you have to update, add to, or fix any of the code? Then, you have to repeat the entire process all over again.

Well, instead of worrying about these hassles I'm going to share a streamlined way to distribute & maintain an Add-On in a corporate environment between multiple computers/users with ease. Here is an example of one of my Add-Ons I've created for my company ([link](http://i.imgur.com/ifYMoce.png))

_"I don't know how to actually code an Add-On (in C#)."_ That was my first thought too when even beginning to consider trying to build an Add-On for my co-workers. I had no idea at the time how much you could still accomplish with an Add-On solely coded using VBA. Yes, some of the more verbose options may not be available as they are when coding an add-on using C#, but to provide macros and everyday functions to improve efficiency, save time & reduce errors, coding in VBA still more than gets the job done.

**In short, the method I'm going to explain goes like this:**

- There is an easy one-click install method that does everything for the end-user, so you don't have to worry about them installing it incorrectly.
- There is a public version of the Add-On (This is the version your end-users will be using)
- There is a private/development version of the Add-On (This is the version you will maintain, make updates to, and deploy. This should be kept locally on your computer, so no one else has access to it)

**Prerequisites:**

- VBA knowledge (obviously)
- A Public/Shared Network drive location that all of your intended users (and yourself) have access to. This is where we will keep the public version of the Add-On.
- Some knowledge of XML
- Custom UI Editor Tool (This is the tool we will use to make the ribbon and the elements that appear on the ribbon)
    - You can download the Custom UI Editor Tool that we'll use to create our ribbon and its contents from this site ([link](http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2006/05/26/customuieditor.aspx))
    - However, seeing how that website is shutting down and not knowing when/if that download page will be removed I have also hosted the file on my personal dropbox account ([link](https://www.dropbox.com/s/ouje1bpmoj1p440/OfficeCustomUIEditorSetup.msi?dl=0))

**Once all the prerequisites are met, here's what you should do**

1. Open up Excel (it's best to only have one instance/window open)
2. Go into the code editor by right-clicking on a worksheet tab and selecting _View Code_
3. Insert a New Module & place/create you sub-routines in there. You can create as many Modules as you like.
4. For each subroutine that you are going to connect to a button on the ribbon you need to add a parameter. For regular buttons you would add `control As IRibbonControl` between the sub's parenthesis, so the sub would look like this `Public Sub MissingImageReport(control As IRibbonControl)`
    - Certain buttons, such as toggle buttons, have multiple parameters, but I can go into more detail on that in another post upon request.
5. Once you're done adding all your Modules & code add an additional module and call it something like `Deployment` and place the code below inside it. Modify the paths & filenames to match your files and paths. This is the sub that you will run whenever you are deploying an update. I'd suggest making it private & locking your add-on.

```vba
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 '''''''''''''''''''''''''''''Add-In Deployment''''''''''''''''''''''''''''''''''''

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 Public Const strAddinPublicPath = "Q:\Supplier's Material\Imports-Exports\0 Export-Import Info\Documentation\ESP Assistant Resources\"

 Private Sub DeployAddIn()

 'Macro Purpose: To deploy finished/updated add-in to a network

 '               location as a read only file

 Dim strAddinDevelopmentPath As String

 'strAddinPublicPath declared as Public variable above

  

 'Set development and public paths

 strAddinDevelopmentPath = ThisWorkbook.Path & Application.PathSeparator

  

 'Turn off alert regarding overwriting existing files

 Application.DisplayAlerts = False

  

 'Save the add-in

 With ThisWorkbook

  'Save to ensure work is okay in case of a crash

  .Save

  

  'Save read only copy to the network (remove read only property

  'save the file and reapply the read only status)

  On Error Resume Next

  SetAttr strAddinPublicPath & .Name, vbNormal

  On Error GoTo 0

  .SaveCopyAs Filename:=strAddinPublicPath & .Name

  SetAttr strAddinPublicPath & .Name, vbReadOnly

 End With

 'Copy the updated documentation to the public folder

 Dim updateDoc As Object: Set updateDoc = VBA.CreateObject("Scripting.FileSystemObject")

 On Error Resume Next

 SetAttr strAddinPublicPath & "ESP Assistant Documentation.docx", vbNormal

 On Error GoTo 0

 updateDoc.CopyFile strAddinDevelopmentPath & "ESP Assistant Documentation.docx", strAddinPublicPath & "ESP Assistant Documentation.docx"

 SetAttr strAddinPublicPath & "ESP Assistant Documentation.docx", vbReadOnly

  

 'Resume alerts

 Application.DisplayAlerts = True

 MsgBox "Update successfully deployed.", vbOKOnly, "Deployment Complete"

 End Sub
```

6. Once you've done all of this, Save As and select Excel Add-On (xlam). Save it to your local path because this will become the developer version.
7. Next thing is creating the ribbon to go along with our Add-On, so download & install the Custom Ribbon UI Tool using one of the links above if you haven't already done so.
8. Once you have it installed go to File>Open and navigate to your Add-On file.
9. When you've opened your Add-On file go to Insert>Office 2010 Custom UI Part, then paste the XML code below into the window & Save
   
```xml
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">

  <ribbon startFromScratch="false">

    <tabs>

      <tab id="CustomTab" label="My Tab">

        <group id="SimpleControls" label="My Group">

          <button id="test1" label="Btn 1" imageMso="HappyFace" screentip="Happy!" size="large" onAction="YourAddOnName.xlam!ModuleName.TheSubRoutineToRun"/>

          <button id="test2" label="Btn 2" imageMso="HappyFace" screentip="Look at me!" size="large" onAction="YourAddOnName.xlam!ModuleName.AnotherSubRoutineToRun"/>

          <button id="test3" label="Btn 3" imageMso="HappyFace" screentip="Hi there!" size="large" onAction="YourAddOnName.xlam!ModuleName.YetAnotherSubRoutineToRun"/>

        </group>

      </tab>

    </tabs>

  </ribbon>

</customUI>
```

- **Important Note:** XML is very picky. One wrong character or forgotten quote will cause your ribbon to not show up at all! If this happens I recommend copying the xml code into an online validator. I recommend W3School's online validator ([link](http://www.w3schools.com/xml/xml_validator.asp)).
- You can find a list of all the elements that can be added to the ribbon on Microsoft's Custom UI page [here](https://msdn.microsoft.com/en-us/library/dd926139.aspx).
- You can find all the stock microsoft office icons and their corresponding `imageMso`s to use for your button icons on this handy site ([link](http://soltechs.net/CustomUI/)).

Once you've saved the XML, if you open up Excel and open your Add-On you should see the new tab called "My Tab", which will have a group called "My Group" and inside that group will have the 3 buttons we created. Now you can move onto deploying the Add-On, so nagivate to the Deployment Module, click into the subroutine & run it. This will create the public version at the public path you previously specified. Now, when you run this subroutine in the future it will simply overwrite the existing public version.

**Lastly, we need to create the file that you tell your co-workers/employees to run that will install the add-on for them.**

_How to create the One-Click install file._

1. Open up a text file
2. Paste the following code
3. Change the path to point to wherever you have the public add-on. You can change the wording of the msgboxes to suit your needs.
4. Essentially, what this code does is
	- Tells the user to close all excel files (all excel instances will be terminated after they click ok on the first prompt)
	- Opens Excel & points to the Add-On to install
	- DOES NOT COPY the file to the user's personal add-on folder, simply creates a connection to the public filepath (this is where most users mess up). This is vital to being able to effortlessly update the add-on in the future.
	- Then, closes & restarts Excel, so the installation can complete. Once it's done it closes out Excel and tells the user the installation is complete.

- One thing to note, there are some instances where certain Excel installations are not successful the first time around due to some registry issues. To resolve this I've created a second small vbs file to refresh the registry values. If this occurs, that file will run, then the user will be told to re-run this installation file. If after the Registry Refresh file is run and the error is still occurring (this may happen in 2010s sometimes depending on settings), then you'll have to manually do the install. I never said anything was fool-proof.

1. Save as a ".vbs" file

**One-Click Installation.vbs File Code**
```vba
'Ask user to save all Excel documents
y=msgbox("Please save all of your work before continuing. All instances of Excel will be terminated before the installation begins." ,0, "Preparation")

'Kill all instances of Excel
Dim objWMIService, objProcess, colProcess
Dim strComputer, strProcessKill
strComputer = "."
strProcessKill = "'EXCEL.exe'"
 
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 
Set colProcess = objWMIService.ExecQuery _
("Select * from Win32_Process Where Name = " & strProcessKill )
For Each objProcess in colProcess
objProcess.Terminate()
Next

'Launch Excel
set objExcel = createobject("Excel.Application")
strAddIn = "ESP Assistant.xlam"
' ~~> Path where the XLAM resides
SourcePath = "Q:\Supplier's Material\Imports-Exports\0 Export-Import Info\Documentation\ESP Assistant Resources\" & strAddIn

'Add the AddIn
On Error Resume Next
With objExcel
	'Add Workbook
	.Workbooks.Add
	'Show Excel
	objExcel.Visible = True
	.AddIns.Add(SourcePath, False).Installed = True
End With

If Err.Number <> 0 Then
	Dim shell
	Set shell = CreateObject("WScript.Shell")
	shell.Run "Q:\Supplier's Material\Imports-Exports\0 Export-Import Info\Documentation\ESP Assistant Resources\Excel Registry Refresh.vbs"
	z=msgbox("Now that Excel's Registry Values have been refreshed please try to rerun this file. If you are still having issue email {your name & email here}" ,0, "Refresh Complete - Please Rerun")
	Err.Clear
	objExcel.Quit
	Set objExcel = Nothing
	wscript.quit
End If

objExcel.Quit
Set objExcel = Nothing

x=msgbox("The ESP Assistant Add-In has successfully been installed." ,0, "Add-In Installation")
```

**Excel Registry Refresh.vbs File Code**

```vba
'File to use just in case Add-In installation fails
'Refreshes Excel Registry Entries to allow for clean install of Add-In
Dim objFSO, objShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject ("WScript.shell")
objShell.Run "cmd /c ""C:\Program Files (x86)\Microsoft Office\Office14\excel.exe"" /unregserver && timeout /t 3 && tskill excel && ""C:\Program Files (x86)\Microsoft Office\Office14\excel.exe"" /regserver",1,True
Set objFSO = Nothing
Set objShell = Nothing
x=msgbox("Excel registry refreshed." ,0, "Registry Update")
wscript.quit
```


Once they run the install file they should be good to go and will always have the updated version of the add-on (unless you push out an update while they have Excel already open, but I'll explain that in a later post upon request). Also, if anyone would like another post covering some more of this process and also how to add a button on the toolbar to indicate to the user when there is an update (basically letting them know to restart excel to get the most up-to-date version) please let me know in the comments.

Possible future topics (upon request, let me know) include:

- Adding other elements to the ribbon
    
- Using ribbon callbacks
    
- Adding section to the ribbon to let users know they don't have the most up-to-date version of the Add-On
    

**If anyone has any questions or if things seem to be a bit unclear, let me know and I'll be happy to help!**


