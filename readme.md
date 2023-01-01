<script language="JSCRIPT">function checkExpand( ) { if ("" != event.srcElement.id) { var ch = event.srcElement.id + "Child"; var el = document.all[ch]; if (null != el) { el.style.display = "none" == el.style.display ? "" : "none"; if (el.style.display != "none") event.returnValue=false; } } }</script>

## Visual Basic Version 6.0 Readme

© 1998 Microsoft Corporation. All rights reserved.

Other product and company names herein may be the trademarks of their respective owners.

_Visual Basic® Readme_ includes updated information for the documentation provided with Microsoft<font face="Symbol"><span style="font-family:Symbol">®</span></font> Visual Studio<font face="Symbol"><span style="font-family:Symbol">™</span></font> -- Development System for Windows® and the Internet. The information in this document is more up-to-date than the information in the Help system.

### 

Contents <font size="2" face="Verdana,Arial,Helvetica" color="#000000">- Click any of the items below</font>

**NOTE:** Be sure all headings in the table of contents are expanded when you search this ReadMe for a topic. In this way, you'll know when the search finds the topic among the TOC headings.

##### [Important Issues -- Please Read First!](# "Click to expand or collapse.")

<div id="CriticalIssuesChild">

> [Sample File Locations](#Samples1)

> [Passing User-Defined Types to Procedures](#Critical1)

> [Incompatibilities with Data-bound Controls](#DataBind2)

> [Searching Online by Topic Title](#Critical2)

> [Cross References to Internet Client SDK Refer to the Internet/Intranet/Extranet SDK](#Critical3)

> [Context-Sensitive Help](#Critical4)

> [Sample Code Sometimes Does Not Cut and Paste Properly](#Critical5)

> [Locate Button Disabled for Reference Topics](#Critical6)

> [Finding Help For ADO Objects](#DataBind7)

> [SQL Server OLE DB Provider Requires New instcat.sql](#DataBind8)

</div>

##### [Data Access Issues and DataBinding Tips](# "Click to expand or collapse.")

<div id="DataBindIssuesChild">

> [Error in Data Environment Designer Code Example](#DataBind1)

> [Incompatibilities with Data-bound Controls](#DataBind2)

> [Binding to Properties of Objects May Yield Unexpected Results](#DataBind3)

> [Complex Binding to an ADO Recordset Requires CursorType](#DataBind4)

> [Creating Visual Basic Data Sources: Type field as adVarChar for SQL Server and Access Databases Instead of adBSTR](#DataBind5)

> [Incorrect References for Creating OLE DB Providers](#DataBind6)

> [Finding Help For ADO Objects](#DataBind7)

> [SQL Server OLE DB Provider Requires New instcat.sql](#DataBind8)

> [Setup for Data Access Applications May Fail on Windows 95/98](#DataBind9)

</div>

##### [Controls Issues](# "Click to expand or collapse.")

<div id="ControlsIssuesChild">

> [Lightweight Controls Must Be Borderless](#Controls1)

> [Run Time Error 711: Compiled .Exe Doesn't Contain Information About Unreferenced Control Causing Controls.Add to Fail](#Controls2)

> [Hierarchical FlexGrid Control: ColWordWrapOption, ColWordWrapOptionBand, ColWordWrapOptionFixed, ColWordWrapOptionHeader Properties](#Controls3)

> [Hierarchical FlexGrid Control: ColIsVisible and RowIsVisible Properties are Read Only](#Controls4)

> [DataRepeater Control: Setting Public Properties Affect only the Current Control](#Controls5)

> [Data Report Designer: Error in Event Handling Code](#Controls6)

> [RichTextBox Control: SelPrint Method Has New Argument](#Controls7)

> [Visual Basic 5 Version of MSChart Control Available in Tools Directory](#Controls8)

> [Toolbar Control: Style Property Settings Changed](#Controls9)

> [Visual Basic Run-Time Error 720: Attempting to Add Anything Except a Control to Controls Collection Causes Run-Time Error](#Controls10)

> [Hierarchical FlexGrid Control: Correcting Errors Binding a Recordset to the HFlexGrid](#Controls11)

> [Hierarchical FlexGrid Control: How to Change the Font of Individual Bands](#Controls12)

> [Hierarchical FlexGrid Control: Avoiding the display of duplicate headers](#Controls13)

> [ADO Data Control: FetchProgress and FetchComplete Events Not Implemented](#Controls14)

> [DataGrid: SizeMode and Size Properties Do Not Accept Value of 2 (dbgNumberOfColumns)](#Controls15)

> [Controls: ImageList Control on Page Designer](#Controls16)

> [Page Designer: Control Issues](#PageDsr12)

> [MSComm Control: EOFEnable Property Doesn't Stop Data Input](#Controls17)

> [Treeview Control: Node Object's Visible Property is Read-Only](#Controls18)

> [SysInfo Control: Constants Not Supported](#Controls19)

> [User Control: Binary Persistence of PropertyBag Data Causes Page Designer to Fail](#Controls20)

</div>

##### [Language Issues](# "Click to expand or collapse.")

<div id="LanguageIssuesChild">

> [InStr Function and Locale-Specific Comparisons](#Language1)

> [SendKeys Statement Gives Invalid Procedure Call Error](#Language2)

> [Type Statement Clarification](#Language3)

> [Decimal Data Type Stored As Signed Integer](#Language4)

> [DateSerial Function and Windows 98/Windows NT 5](#Language5)

> [Code Window "Find Next" Keyboard Shortcut](#Language6)

> [Add Method (Folders) Syntax](#Language7)

</div>

##### [Samples Issues](# "Click to expand or collapse.")

<div id="SamplesIssuesChild">

> [Obtaining Updated Versions of Sample Applications](#Samples9)

> [Sample File Locations](#Samples1)

> [Visual Basic Sample: Biblio and Mouse Samples Omitted](#Samples2)

> [Visual Basic Samples: ChrtSamp Description](#Samples3)

> [Visual Basic Samples: CtlsAdd Sample: Controls.mdb Must Be Read/Write Enabled on Hard Disk](#Samples4)

> [DHSHOWME.VBP Sample: You May Need to Reset the SourceFile Property For This Sample to Work Correctly in Design Mode](#Samples5)

> [PROPBAG.VBP Sample: Possible Error on Loading the Module for this Sample](#Samples6)

> [Running the IObjSafe Sample Application](#Samples7)

</div>

##### [Wizard Issues](# "Click to expand or collapse.")

<div id="WizardIssuesChild">

> [Setup for Data Access Applications May Fail on Windows 95/98](#DataBind9)

> [Package and Deployment Wizard: Automatically Pick Up Files From Redist directory](#Wizard1)

> [Package and Deployment Wizard Has [Do Not Redist] Section](#Wizard2)

> [Package and Deployment Wizard: In Silent Mode the Notification Window May Not Be First in the Window Z-order](#Wizard3)

> [Package and Deployment Wizard: Command Line Mode Argument Added for Specifying Executable Path](#Wizard4)

> [Package and Deployment Wizard: Manually Add User Control License Files](#Wizard5)

> [Package and Deployment Wizard: Steps in the Web Deployment Process](#Wizard6)

> [Package and Deployment Wizard: Web Deployment Tips for HTTP Protocol](#Wizard7)

> [Package and Deployment Wizard: Start Menu Items: Run Option Not Supported](#Wizard8)

> [System Configurations for WebPost's Posting Acceptor](#Wizard9)

> [Package and Deployment Wizard: Edit Setup.lst file if you rebuild cabs from batch file](#Wizard10)

> [Package And Deployment Wizard: Error 80042114](#Wizard11)

> [Package and Deployment Wizard: Use Mdac_typ.Cab to Distribute Data Access Components](#Wizard12)

> [.ASP Files Not Included in Standard Packages](#WebClass21)

> [Package and Deployment Wizard: Manually Include .ASP and .HTM Files For IIS Applications When Using Standard Setup](#Wizard13)

> [Package and Deployment Wizard: Bad Date and Time Formats](#Wizard14)

> [Package and Deployment Wizard: Unable to Run Setup.exe on First Windows 95 Version](#Wizard15)

> [Package and Deployment Wizard: Packaging ActiveX Documents](#Wizard16)

</div>

##### [Error Message Issues](# "Click to expand or collapse.")

<div id="ErrorIssuesChild">

> [No Help Topic for the Following Error Messages](#Error1)

> [DHTML Page Designer Error Messages](#PageDsr1)

> [Run Time Error 711: Compiled .Exe Doesn't Contain Information About Unreferenced Control Causing Controls.Add to Fail](#Controls2)

</div>

##### [WebClass Designer Issues](# "Click to expand or collapse.")

<div id="WebClassIssuesChild">

> [Webclasses: "Me." Not Supported](#WebClass1)

> [Webclasses: Invalid HTML Syntax Can Cause Unspecified Error](#WebClass2)

> [Webclasses: Avoid Using Global or Static Variables in a Webclass](#WebClass3)

> [Webclasses: Some External HTML Changes are not Detected Automatically](#WebClass4)

> [Webclasses: IIS Administration Console File Settings are not Acknowledged for Templates](#WebClass5)

> [Webclasses: Unattended Execution](#WebClass6)

> [Webclasses: Retain in Memory](#WebClass7)

> [Webclasses: Accounting for Differences Between the Debug and Compiled Versions](#WebClass8)

> [Webclasses: Performance Tips](#WebClass9)

> [Webclasses: Miscellaneous Issues](#WebClass10)

> [Webclasses: Articles of Interest](#WebClass11)

> [Webclasses: Formatting in Source HTM File](#WebClass12)

> [Webclasses: Cannot Support HTML's LINK Element](#WebClass13)

> [Webclasses: When Using Visual SourceSafe with Webclass Projects, You Must Manually Check in the Project's .HTM Files](#WebClass14)

> [Webclasses: TagPrefix Should Be WC:](#WebClass15)

> [Webclasses: Variant Parameter in URLFor Method](#WebClass16)

> [Webclasses: Sequencing Data is Passed Using the &WCU Parameter](#WebClass17)

> [Webclasses: StateManagement Property Constants Contain Incorrect Property Reference](#WebClass18)

> [Webclasses: State and the Session Object](#WebClass19)

> [Webclasses: Code Corrections in Help Topic "Defining Webclass Events at Run Time"](#WebClass20)

> [Webclasses: HTM and ASP Files Not Included in Standard Packages](#WebClass21)

> [Webclasses: Unspecified Error](#WebClass22)

</div>

##### [DHTML Page Designer Issues](# "Click to expand or collapse.")

<div id="PageDsrIssuesChild">

> [Page Designer Error Messages](#PageDsr1)

> [Page Designer: "Me." Not Supported](#PageDsr2)

> [Page Designer: Cannot Access HTML Elements from Forms or External Objects](#PageDsr3)

> [Page Designer: Image SourceFiles are Resolved Incorrectly](#PageDsr4)

> [Page Designer: Binary Persistence Issue](#PageDsr5)

> [Page Designer: Type Library Problems are Preventing Some Help Topics from Appearing](#PageDsr6)

> [Page Designer: Modal Prompts Appear Behind Browser in Run Mode](#PageDsr7)

> [Page Designer: Do Not Watch Objects of Type HTMLDocument](#PageDsr8)

> [Page Designer: Cannot Design Frames Within A Page Designer](#PageDsr9)

> [Page Designer: Cannot use Visual Basic Code with the SetInterval Method](#PageDsr11)

> [Page Designer: Control Issues](#PageDsr12)

> [Page Designer: Miscellaneous HTML Issues](#PageDsr13)

> [Page Designer: Miscellaneous Debugging Issues](#PageDsr14)

> [Page Designer: Syntax for Navigating Programmatically](#PageDsr15)

> [Page Designer: Location of Project Files in .CAB Deployment](#PageDsr24)

> [Page Designer: OBJECT Tag Insertion](#PageDsr17)

> [Page Designer: Cannot CTRL+TAB Through Windows When Page Designer Has Focus](#PageDsr18)

> [Page Designer: Cannot Delete a Table if it Contains No Columns](#PageDsr19)

> [Page Designer: When Using Visual SourceSafe with Page Designer Projects, You Must Manually Check in the Project's .HTM Files](#PageDsr20)

> [Page Designer: Unqualified BuildFile Property Results in DLL](#PageDsr21)

> [Page Designer: Use of Load Event With Asynchronous Property](#PageDsr22)

> [Page Designer: Property Page Issues](#PageDsr23)

> [Page Designer: Name of Save Option Changed](#PageDsr26)

> [Page Designer: Problems Seeing Code Changes When Switching from Compiled to Debug Mode](#PageDsr27)

> [Page Designer: Help for Most Language Elements is Available in the Platform SDK](#PageDsr28)

</div>

##### [Extensibility Issues](# "Click to expand or collapse.")

<div id="ExtensibilityIssuesChild">

> ["Command-Line Safe" Add-In Behavior](#Ext1)

> [Manually Setting Add-In Registry Values](#Ext2)

> [Using the Add-In Designer](#Ext3)

> [Add-In Designer: More Information About Specifying Satellite DLL](#Ext4)

</div>

##### [Miscellaneous Issues](# "Click to expand or collapse.")

<div id="MiscIssuesChild">

> [Returning an Error Value from a DLL](#Misc1)

> [Data Access Guide: DataFormats tutorial has wrong file extension](#Misc2)

> [External Editor Field Added to Options Dialog Box](#Misc3)

> [CodeBase Fixup Utility in Internet Component Download](#Misc4)

> [Text in Project Properties/Open Dialogs Truncated in Japanese, Chinese, and Korean versions of Windows](#Misc6)

> [Finding Help For ADO Objects](#DataBind7)

> [Avoid Using Repository Add-In with ActiveX Designers](#Misc7)

</div>

##### [Microsoft Transaction Server (MTS) Issues](# "Click to expand or collapse.")

<div id="MTSIssuesChild">

> [Building and Debugging MTS Components in Visual Basic 6.0](#MTS1)

</div>

##### [Dictionary Object](# "Click to expand or collapse.")

<div id="DictIssuesChild">

> [Introducing the Dictionary Object](#Dict1)

</div>

##### [Visual Component Manager](# "Click to expand or collapse.")

> <div id="VCMChild">
> 
> [Known Problems](# "Click to expand or collapse.")
> 
> > <div id="VCMKnownProblemsChild">
> > 
> > ["Related Files Tab (Component Properties Dialog Box)" Topic Incorrect](#RelatedFilesTab)
> > 
> > [Removing Repository 1.0 Registry Keys](#RemovingRepository10RegistryKeys)
> > 
> > [Adding Repository Tables to an Existing .mdb File](#AddingrepositorytablestoanexistingMDBfile)
> > 
> > </div>
> 
> </div>

##### [Application Performance Explorer](# "Click to expand or collapse.")

> <div id="APEChild">
> 
> [Known Problems](# "Click to expand or collapse.")
> 
> > <div id="APEKnownProblemsChild">
> > 
> > [Configuring Remote Automation Security When Using Remote APE Components](#ConfiguringRemoteAutomationSecurityWhenUsingRemoteAPEComponents)
> > 
> > [Compatibility Issues Between the Application Performance Explorer that Ships with Visual Studio 6.0 and the Version that Shipped with Visual Basic 5.0](#compatibilityissuesbetweenvs6apeandvb5ape)
> > 
> > [Adjusting Default Settings To Use APE and MTS](#AdjustingdefaultsettingstouseAPEandMTS)
> > 
> > [Application Performance Explorer Server-Side Setup May Generate Error](#ApplicationPerformanceExplorerServerSideSetup)
> > 
> > </div>
> 
> </div>

For installation issues pertaining to Microsoft® Visual SourceSafe™, see [Visual SourceSafe Notes](http:\\msdn.microsoft.com/ssafe/ "Jumps to the Visual SourceSafe installation readme.").

For _general installation issues_ on the Visual Studio 6.0 suite of products, including side by side product installation, see [Installation Notes](install.htm "Jumps to the installation readme (install.htm).") readme (install.htm).

##### For other issues on the Help system of the Visual Studio suite of products, go to:

[MSDN<font face="Symbol"><span style="font-family:Symbol">ä</span></font>, the Microsoft Developer Network Readme.](readmeDN.htm "Jumps to readmeDN.htm")

* * *

### Important Issues

#### <a name="Critical1"></a>Passing User-Defined Types to Procedures

With Visual Basic 6.0 it is possible to pass a user defined type (UDT) as an argument to a procedure or function, however there is a restriction. Passing a UDT to a procedure in an out-of-process component or across threads in a multi-threaded component requires an updated version of DCOM for Windows 95 and Windows 98, or Service Pack 4 for Windows NT 4.0\. This update is required on your development computer as well as on any computer that will run your application. A run-time error will occur if the required files are not installed.

The above does not apply to passing UDT's within a single-threaded application; this will work without updating.

The Package and Deployment Wizard will not determine the dependencies for the necessary components - it is up to you to make sure that the files are on the end user's computer. You can test for the existence of the components by trapping for run-time error 458 - "Variable uses an Automation type not supported in Visual Basic". If this error occurs, the DCOM or Service Pack components must be updated; the update procedure differs depending on the operating system:

##### Windows 95 / Windows 98

DCOM98.EXE is a self-extracting executable that installs the updated DCOM components for Windows 95 or Windows 98\. It can be found in the DCOM98 directory of the Visual Basic 6.0 CD. This file may be freely distributed with your Visual Basic application.

##### Windows NT 4.0

The updated DCOM components are automatically installed with Service Pack 4 (SP4). You can download the Service Pack from Microsoft's web site.

#### <a name="Critical2"></a>Searching Online by Topic Title

**To search for a topic when you have the title**

1.  In the navigation pane of the MSDN window, click the **Search** tab and then type or paste the title of the topic you want to find. Enclose the search string in quotation marks.
2.  Click **Search Titles Only**.
3.  Click **List Topics**. (If your search returns more than one hit, you can sort the topic list by clicking the **Title** or **Location** column heading).
4.  Select the title of the topic you want and then click **Display**.

**To find where a topic is located in the table of contents**

*   Click the **Locate** button on the toolbar. The table of contents will synchronize with the topic you are viewing.

**Note**   The **Locate** button is unavailable for the topics in the Reference node of the Visual Basic documentation.

#### <a name="Critical3"></a>Cross References to Internet Client SDK Refer to the Internet/Intranet/Extranet SDK

In the Building Internet Applications book within the Component Tools Guide, multiple cross references are made to a part of MSDN referred to as the "Internet Client SDK." The correct name for this SDK is the "Internet/Intranet/Extranet SDK." When searching for an Internet Client SDK reference in MSDN, please look in this section.

#### <a name="Critical4"></a>Context-Sensitive Help

To use Help buttons and the F1 key to access Help without having the MSDN CD in your CD drive, you **must** choose the Custom install option during setup of the MSDN Library. Check the boxes labeled "VB Documentation," "VB Product Samples," and "VS Shared Documentation." You may also want to check "VSS Documentation" if you are using Visual SourceSafe.

#### <a name="Critical5"></a>Sample Code Sometimes Does Not Cut and Paste Properly

Line breaks and formatting information may not copy correctly when you copy and paste sample code from the MSDN Library Visual Studio documentation to your code editor. To work around this issue, do one of the following:

*   Manually edit the line breaks after you copy the code.
*   View the sample code source, copy the entire code sample, including the **<pre>** and **</pre>** tags, paste it to your code editor, and then delete the unwanted sections from the pasted version.

#### <a name="Critical6"></a>Locate Button Disabled for Reference Topics

When you find a language reference topic in MSDN through the Search tab, you cannot use the Locate button to find where the topic is located in the MSDN Table of Contents tree.

* * *

### Data Access Issues and DataBinding Tips

#### <a name="DataBind1"></a>Error in Data Environment Designer Code Example

In the topic, "Programmatically Accessing Objects in Your Data Environment Designer," the example under "Executing a Command Object with Multiple Parameters" erroneously uses the **Open** method:

<pre> <font face="Courier">MyDE.Commands("InsertCustomer").Parameters("ID").value = "34"
	MyDE.Commands("InsertCustomer").Parameters("Name").value = "Fred"
	MyDE.Commands("InsertCustomer").**Open**</font> </pre>

There is no **Open** method for the **Commands** object. You must use the **Execute** method instead.  

#### <a name="DataBind2"></a>Incompatibilities with Data-bound Controls

Due to changes in Visual Basic 6.0, not all data-bound controls are compatible with all data sources. This incompatibility is due to a difference in the internal binding mechanisms of ADO versus DAO/RDO. Controls that were created specifically to work with DAO/RDO can't be bound to an ADO Data control; controls created for use with ADO can't be bound to the standard Data control or the Remote Data Control.

This incompatibility primarily applies to complex-bound controls such as grids or lists that bind to multiple fields in a data source; simple-bound controls such as text boxes or labels that bind to a single field will work with either type of data source. Some examples are as follows:

*   The Microsoft Data Bound Grid control (Dbgrid32.ocx) can be bound to the DAO or RDO Data controls; it can't be bound to the ADO Data control.
*   The Microsoft DataGrid control (Msdatgrd.ocx) can be bound to the ADO Data control; it can't be bound to the DAO or RDO Data controls.
*   The Microsoft Masked Edit Control (Msmask32.ocx) can be bound to any of the Data controls.
*   The intrinsic controls (TextBox, PictureBox, Label, and so on) can be bound to any of the Data controls.
*   Third-party controls and Visual Basic-authored User controls should be tested on a case-by-case basis.

When attempting to bind a control to a data source at design time, you may encounter a "No compatible data source" error message. In this case, you will need to substitute another control that is compatible with your data source.

#### <a name="DataBind3"></a>Binding to Properties of Objects May Yield Unexpected Results

While it's possible to bind any object to any other object, the results may not always be what you expect. Some properties are read-only bindable and will not update their bound source.

For example, if you were to bind the Caption property of a Frame control to a field named Foo in an ADO Recordset object, the Caption would change to reflect the value of Foo as you scrolled through the Recordset. If, however, you changed the Caption property programmatically (Frame1.Caption = "Bar"), the value of Foo would not be updated. Because the Caption property of the Frame is read-only bindable, it doesn't provide notification that its data has changed.

This isn't a problem for Visual Basic-authored objects, since you can call the PropertyChanged method in your object's code. For other objects, you can determine if a property is update bindable by checking the DataBindings collection. If a property is enumerated in the DataBindings collection, it is update bindable and the data source will receive updates to data; if it isn't enumerated, the property is read-only bindable.

#### <a name="DataBind4"></a>Complex Binding to an ADO Recordset Requires CursorType

When binding an ADO Recordset object to a complex-bound control (such as a Grid control), it is necessary to explicitly set the CursorType property to either adOpenStatic or adOpenKeyset. If you don't set this property, no data will be displayed. The following code shows the use of the CursorType property.

<pre><font size="2">

> Private Sub DataClass_Initialize()
>    Set cn = New ADODB.Connection
>    Set rs = New ADODB.Recordset
>    rs.CursorType = adOpenStatic
>    cn.Open "northwind"
>    rs.Open "customers", cn
> End Sub

</font></pre>

Binding to a simple-bound control (such as a TextBox) doesn't require a specific CursorType.

#### <a name="DataBind5"></a>Creating Visual Basic Data Sources: Type the fields as adVarChar for SQL Server and Access Databases Instead of adBSTR

When appending fields to an ADO Recordset object for use with a SQL Server or Access database, type the fields as adVarChar instead of adBSTR (as shown in some sample code). When reading data out of either SQL Server or Access databases, ADO will use the adVarChar type.

#### <a name="DataBind6"></a>Incorrect References for Creating OLE DB Providers

The documentation erroneously states that it is possible to set a class module's **DataSourceBehavior** property to _2 - vbOLEDBProvider_ to create an OLE DB data provider. The correct values for **DataSourceBehavior** are _0 - vbNone_ and _1 - vbDataSource_.

The documentation also erroneously refers to a non-existent event in class modules called OnDataConnection.

Finally, in the topic "Creating the MyDataSource Class," the step-by-step example incorrectly states that you should set **DataSourceBehavior** to _2 - vbOLEDBProvider_. Instead, you should set **DataSourceBehavior** to _1 - vbDataSource_.

To create OLE DB data providers using Visual Basic, use the Provider Writer Toolkit included with the OLE DB SDK. For more information, see the OLE DB Simple Provider Toolkit in the Platform SDK Documentation on MSDN.

#### <a name="DataBind7"></a>Finding Help For ADO Objects

When using the ADO objects, (for example, Recordset, Connection, Command, Parameter, ADOR, RDS, and RDS Server object), you cannot get context-sensitive help on the object or its properties, events, or methods. That is, if you have a reference to the object and you use one of its features, selecting the code and pressing F1 does not result in a help topic. Instead, you will get either a wrong topic or the "Keyword Not Found" topic.

However, you can get help on any of the object's properties, events, or methods by using the online documentation Index:

1.  If the MSDN documentation viewer is not open, on the **Help** menu, click **Contents**.
2.  Click the **Index** tab.
3.  Type the name of the property, event, or method including the word "collection", "property", "event", or "method" as appropriate.
4.  From the list of available topics, select the topic that includes "ADO" in its title.

> **Note**   You can also find additional help on other ADO topics, such as the ADO object model, by looking in the MSDN Library Table of Contents: open **Platform SDK**, and under **Database and Messaging Services**, go to **Microsoft Data Access SDK**.

#### <a name="DataBind8"></a>SQL Server OLE DB Provider Requires New instcat.sql

Before using the SQL Server OLE DB data provider, you must run the version of instcat.sql distributed with Microsoft Visual Basic 6.0 on SQL Server (version 6.5 and later). Instcat.sql is distributed with Visual Basic 6.0 and can be found in the \winnt\system32 directory upon installation.

If Instcat.sql is not run on your SQL Server, the provider is unable to retrieve metadata from the SQL Server, and thus will not be able to connect to that server.

#### <a name="DataBind9"></a>Setup for Data Access Applications May Fail on Windows 95/98

When redistributing a VB 6.0 application that includes data access components, setup will fail if DCOM for Windows 95 and Windows 98 isn't present on Windows 9x client machines.

The file Mdac_typ.exe is added to your setup package by the Package & Deployment Wizard if your project includes references to ADO, OLEDB, or ODBC (you can check for this on the Included Files page of the wizard). This file installs MDAC 2.0 files on the client computer. MDAC 2.0 requires DCOM for Windows 95 and Windows 98 in order to function properly, however it does not perform a check for this during setup. The setup will fail if DCOM for Windows 95 and Windows 98 isn't present on the client machine. Some of the older data access components will be overwritten prior to the failure, possibly causing older data access applications on the client to fail.

When distributing data access applications for Windows 9x, you need to make sure that DCOM for Windows 95 and Windows 98 is installed on the client. DCOM98.EXE is a self-extracting executable file that installs the updated DCOM components for Windows 95 or Windows 98\. It can be found in the DCOM98 directory of the Visual Basic 6.0 CD. This file may be freely distributed with your Visual Basic application.

* * *

### Controls Issues

#### <a name="Controls1"></a>Lightweight Controls Must Be Borderless

When creating a lightweight User control by setting the Windowless property to True, the BorderStyle property is invalidated. By definition a lightweight control has no border.

If you first set the BorderStyle property to anything other than 0 - None and subsequently change the Windowless property to True, you will receive an error message "Windowless UserControls only support BorderStyle = None".

#### <a name="Controls2"></a>Run Time Error 711: Compiled .Exe Doesn't Contain Information About Unreferenced Control Causing Controls.Add to Fail

**

Problem:

**

1.  Create a new Standard Exe.
2.  Add a user control to the project.
3.  Add the following code:

<pre>

> <font size="2" face="Courier">Dim WithEvents x as VBControlExtender
> 
> Private Sub Form_Load ()
>    Set x = Controls.Add ("Project1.Usercontrol1", "XX")
>    x.Visible = True
> End Sub</font>

</pre>

5.  On the **File** menu, click **Make Project1.exe** (Don't run the project.)
6.  Run the exe.

**Result:** You get an error (711) stating that Project1.Usercontrol1 is an invalid ProgID since no info about it can be found in the exe.

**Solution:** Before compiling the project, under the **Project** menu, click **Project1 Properties**. On the **Make** tab, clear the "Remove information about unused ActiveX controls" check box.

**

Remarks

**

By default ActiveX controls that are referenced but not placed on any type of form at design time are not available for Controls.Add at runtime or in an executable.

#### <a name="Controls3"></a>Hierarchical FlexGrid Control: ColWordWrapOption, ColWordWrapOptionBand, ColWordWrapOptionFixed, ColWordWrapOptionHeader Properties

The following properties are part of the Hierarchical FlexGrid control's feature set but are not documented in the control's help: ColWordWrapOption, ColWordWrapOptionBand, ColWordWrapOptionFixed, ColWordWrapOptionHeader. Descriptions and syntaxes for these properties are found below. Settings for all properties are the same, and can be found at the bottom of the topic.

**

ColWordWrapOption Property

**

Returns or sets a value that specifies how text is wrapped in a specified column.

**

Syntax

**

_object_.**ColWordWrapOption** (_Index_) = _integer_

The **ColWordWrapOption** property syntax has these parts:

<table width="559" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="20%" valign="TOP">

<font size="2">Part</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Description</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_object_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">An object expression that evaluates to a Hierarchical FlexGrid control.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_Index_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Long. The number of the column to get or set word wrap on. The value must be in the range of -1 to Cols - 1\. Setting this value to –1 selects all columns.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_integer_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">A numeric expression that determines how words will wrap, as shown in settings.</font>

</td>

</tr>

</tbody>

</table>

**

ColWordWrapOptionBand Property

**

Returns or sets a value that specifies how text is wrapped in a specified band.

**

Syntax

**

_object_.**ColWordWrapOptionBand (**_BandNumber_, _BandColIndex_**)** = _integer_

The **ColWordWrapOption** property syntax has these parts:

<table width="559" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="20%" valign="TOP">

<font size="2">Part</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Description</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_object_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">An object expression that evaluates to a Hierarchical FlexGrid control.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_BandNumber_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Long. The number of the band to get or set word wrap on. The value must be in the range of 0 to Bands - 1\.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_BandColIndex_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Long. The number of the column to get or set word wrap on. This optional parameter defaults to –1, indicating all columns in the band. Valid values are –1 to Cols –1.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_integer_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">A numeric expression that determines how words will wrap, as shown in settings.</font>

</td>

</tr>

</tbody>

</table>

**

ColWordWrapOptionFixed Property

**

Returns or sets a value that specifies how text is wrapped in a specified fixed column.

**

Syntax

**

_object_.**ColWordWrapOptionFixed(**_index_**)** = _integer_

The **ColWordWrapOptionFixed** property syntax has these parts:

<table width="559" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="20%" valign="TOP">

<font size="2">Part</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Description</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_object_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">An object expression that evaluates to a Hierarchical FlexGrid control.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_index_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Long. The number of the column to get/set word wrap on. This optional parameter defaults to –1\. Valid values are –1 to Cols –1.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_integer_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">A numeric expression that determines how words will wrap, as shown in settings.</font>

</td>

</tr>

</tbody>

</table>

**

ColWordWrapOptionHeader Property

**

Returns or sets a value that specifies how text is wrapped in column headers.

**

Syntax

**

_object_.**ColWordWrapOptionHeader(**_BandNumber, BandColIndex_**)** = _integer_

The **ColWordWrapOptionHeader** property syntax has these parts:

<table width="559" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="20%" valign="TOP">

<font size="2">Part</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Description</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_object_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">An object expression that evaluates to a Hierarchical FlexGrid control.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_BandNumber_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Long. The number of the band to get/set word wrap on. The value must be in the range of 0 to Bands - 1\.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_BandColIndex_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Long. The number of the column to get/set word wrap on. This optional parameter defaults to –1 indicating all column headers in the band. Valid values are –1 to Cols –1.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">_integer_</font>

</td>

<td width="80%" valign="TOP">

<font size="2">A numeric expression that determines how words will wrap, as shown in settings.</font>

</td>

</tr>

</tbody>

</table>

**

Settings

**

The settings for _integer_ are:

<table width="559" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="35%" valign="TOP">

**<font size="2">Constant</font>**

</td>

<td width="10%" valign="TOP">

**<font size="2">Value</font>**

</td>

<td width="55%" valign="TOP">

**<font size="2">Description</font>**

</td>

</tr>

<tr>

<td width="35%" valign="TOP">

**<font size="2">flexSingleLine</font>**

</td>

<td width="10%" valign="TOP">

<font size="2">0</font>

</td>

<td width="55%" valign="TOP">

<font size="2">(Default) Displays text on a single line only.</font>

</td>

</tr>

<tr>

<td width="35%" valign="TOP">

**<font size="2">flexWordBreak</font>**

</td>

<td width="10%" valign="TOP">

<font size="2">1</font>

</td>

<td width="55%" valign="TOP">

<font size="2">The lines are automatically broken between words.</font>

</td>

</tr>

<tr>

<td width="35%" valign="TOP">

**<font size="2">flexWordEllipsis</font>**

</td>

<td width="10%" valign="TOP">

<font size="2">2</font>

</td>

<td width="55%" valign="TOP">

<font size="2">Truncates text that does not fit in the rectangle and adds ellipsis.</font>

</td>

</tr>

<tr>

<td width="35%" valign="TOP">

**<font size="2">flexWordBreakEllipsis</font>**

</td>

<td width="10%" valign="TOP">

<font size="2">3</font>

</td>

<td width="55%" valign="TOP">

<font size="2">Breaks words between lines and adds ellipsis if text doesn't fit in the rectangle.</font>

</td>

</tr>

</tbody>

</table>

#### <a name="Controls4"></a>Hierarchical FlexGrid Control: ColIsVisible and RowIsVisible Properties are Read Only

The **ColIsVisible** and **RowIsVisible** properties are read-only properties and cannot be set programmatically. You can use the properties to test whether a column or row is visible, and hide the column or row, if appropriate, as show below:

<pre>

> <font size="2" face="Courier">With MSHFlexGrid1
>    If .ColIsVisible(1) Then .ColWidth(1) = 0
>    If .RowIsVisible(1) Then .RowHeight(1) = 0
> End With</font>

</pre>

**

Hierarchical FlexGrid Control: Additional Settings for GridLines, GridLinesBand, GridLinesFixed, GridLinesHeader, GridLinesIndent, and GridLinesUnpopulated Properties

**

Two additional settings are possible for the following properties: GridLines, GridLinesBand, GridLinesFixed, GridLinesHeader, GridLinesIndent, GridLinesUnpopulated Properties. The possible settings are show in the table below:

<table width="559" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="35%" valign="TOP">

**<font size="2">Constant</font>**

</td>

<td width="10%" valign="TOP">

**<font size="2">Value</font>**

</td>

<td width="55%" valign="TOP">

**<font size="2">Description</font>**

</td>

</tr>

<tr>

<td width="35%" valign="TOP">

**<font size="2">flexGridDashes</font>**

</td>

<td width="10%" valign="TOP">

<font size="2">4</font>

</td>

<td width="55%" valign="TOP">

<font size="2">Dashed Lines. Sets line style between cells to dashed lines.</font>

</td>

</tr>

<tr>

<td width="35%" valign="TOP">

**<font size="2">flexGridDots</font>**

</td>

<td width="10%" valign="TOP">

<font size="2">5</font>

</td>

<td width="55%" valign="TOP">

<font size="2">Dotted Lines. Sets line style between cells to dotted lines.</font>

</td>

</tr>

</tbody>

</table>

**

Remarks

**

These settings can be used in addition to flexGridNone, flexGridFlat, flexGridInset and flexGridRaised.

#### <a name="Controls5"></a>DataRepeater Control: Setting Public Properties Affect only the Current Control

When creating a user control to be used in a **DataRepeater** control, be aware that public properties of the control will be set only on the current control (the "live" control with the focus). For example, if you expose the Font property of a user control, at run time, resetting that property (as shown in the example code below) will only affect the current control in the **DataRepeater** control. The font of repeated controls will not be affected.

<pre>

> <font size="2" face="Courier">Private Sub Command1_Click()
>    ' Only the current control's Font will be affected.
>    DataRepeater1.RepeatedControl.FontName = "Courier"
> End Sub</font>

</pre>

The corresponding code in the user control would resemble the following:

<pre>

> <font size="2" face="Courier">Public Property Get FontName() As String
>    FontName = txtProductName.Font.Name
> End Property
> 
> Public Property Let FontName(ByVal newFontName As String)
>    txtProductName.Font.Name = newFontName
> End Property</font>

</pre>

**

TabStrip Control: Separators show only when the Style property = TabFlatButtons

**

Separators will only appear on a **TabStrip** control when the **Style** property is set to **TabFlatButtons**. An example is shown below:

<pre>

> <font size="2" face="Courier">Private Sub Form_Load()
>    With TabStrip1
>       .Style = tabFlatButtons
>       .Separators = True
>    End With
> End Sub</font>

</pre>

#### <a name="Controls6"></a>Data Report Designer: Error in Event Handling Code

In the topic titled _Data Report Events_, there is an error in code that shows how to handle asynchronous errors. For more information, [search online](#Critical2), with **Search titles only** selected, for "Data Report Events" in the MSDN Library Visual Studio 6.0 documentation.

The code is found under the heading "Error Events—for Asynchronous Events."

The code omits a "Select Case ErrObj.ErrorNumber" statement. The corrected code is:

<pre>

> <font size="2" face="Courier">Private Sub DataReport_Error(ByVal JobType As _ MSDataReportLib.AsyncTypeConstants, ByVal Cookie As Long, _
>  ByVal ErrObj As MSDataReportLib.RptError, ShowError As Boolean)
>    Select Case ErrObj.ErrorNumber
>       Case rptErrPrinterInfo ' 8555
>          MsgBox "A printing error has occurred. " & _
>           "You may not have a Printer installed."
>          ShowError = False
>          Exit Sub
>       Case Else ' handle other cases here.
>          ShowError = True
>    End Select
> End Sub</font>

</pre>

#### <a name="Controls7"></a>RichTextBox Control: SelPrint Method Has New Argument

The **SelPrint** method now features a second, optional argument. The syntax and part descriptions are shown below:

**

Syntax

**

_object_.**SelPrint(**_lHDC_ **As Long**, [_vStartDoc_]**)**

The **SelPrint** method syntax has these parts:

<table width="559" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="20%" valign="TOP">

**<font size="2">Part</font>**

</td>

<td width="80%" valign="TOP">

**<font size="2">Description</font>**

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

_<font size="2">object</font>_

</td>

<td width="80%" valign="TOP">

<font size="2">An object expression that evaluates to a **RichTextbox** control.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

_<font size="2">lHDC</font>_

</td>

<td width="80%" valign="TOP">

<font size="2">Long. The device context of the device you plan to use to print the contents of the control.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

_<font size="2">vStartDoc</font>_

</td>

<td width="80%" valign="TOP">

<font size="2">Boolean. Specifies the behavior of the control regarding _startdoc_ and _enddoc_ printer control operations, as shown in settings.</font>

</td>

</tr>

</tbody>

</table>

**

Settings

**

The settings for _vStartDoc_ are:

<table width="559" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="21%" valign="TOP">

**<font size="2">Constant</font>**

</td>

<td width="9%" valign="TOP">

**<font size="2">Value</font>**

</td>

<td width="71%" valign="TOP">

**<font size="2">Description</font>**

</td>

</tr>

<tr>

<td width="21%" valign="TOP">

**<font size="2">True</font>**

</td>

<td width="9%" valign="TOP">

<font size="2">-1</font>

</td>

<td width="71%" valign="TOP">

<font size="2">(Default) The control retains its original behavior and sends _startdoc_ and _enddoc_ commands to the printer.</font>

</td>

</tr>

<tr>

<td width="21%" valign="TOP">

**<font size="2">False</font>**

</td>

<td width="9%" valign="TOP">

<font size="2">0</font>

</td>

<td width="71%" valign="TOP">

<font size="2">The control doesn't send _startdoc_ and _enddoc_ commands, but sends only _startpage_ and _endpage_ commands to the printer.</font>

</td>

</tr>

</tbody>

</table>

**

Remarks

**

The argument was added to remedy situations when printers do not print with the default behavior. When the **SelPrint** method is invoked, both Visual Basic and the RichTextBox control send _startdoc_ and _enddoc_ commands to the printer resulting in a nested pair of _startdoc_/_enddoc_ commands. Some printers respond only to the first pair of commands and thereby become disabled when the RichTextbox control sends the second pair. In that case, setting the _vStartDoc_ argument to **False** prevents the second pair of commands from being sent.

#### <a name="Controls8"></a>Visual Basic 5 Version of MSChart Control Available in Tools Directory

For pre-release users of Visual Basic only:

A Visual Basic 5.0x version of the MSChart control is now included with Visual Basic. If you need a Visual Basic 5 version of the Chart control, and you have installed the pre-release version of the MSChart control, please overwrite the pre-release version with the version contained in the Tools directory of the Visual Basic CD.

#### <a name="Controls9"></a>Toolbar Control: Style Property Settings Changed

The Style property settings for the **Toolbar** control have been changed. The help topic for the property lists **tbrTransparent** and **tbrRight** as possible settings, however these are not implemented in the current version. The actual possible settings and descriptions are shown below:

<table width="559" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="21%" valign="TOP">

**<font size="2">Constant</font>**

</td>

<td width="9%" valign="TOP">

**<font size="2">Value</font>**

</td>

<td width="71%" valign="TOP">

**<font size="2">Description</font>**

</td>

</tr>

<tr>

<td width="21%" valign="TOP">

**<font size="2">tbrStandard</font>**

</td>

<td width="9%" valign="TOP">

<font size="2">0</font>

</td>

<td width="71%" valign="TOP">

<font size="2">(Default) Standard toolbar.</font>

</td>

</tr>

<tr>

<td width="21%" valign="TOP">

**<font size="2">tbrFlat</font>**

</td>

<td width="9%" valign="TOP">

<font size="2">1</font>

</td>

<td width="71%" valign="TOP">

<font size="2">Flat. The borders of a button dynamically appear when the cursor hovers over the button.</font>

</td>

</tr>

</tbody>

</table>

#### <a name="Controls10"></a>Visual Basic Run-Time Error 720: Attempting to Add Anything Except a Control to Controls Collection Causes Run-Time Error

Attempting to add an object that is not a control to the Controls collection causes run-time error 720\. You can only add Visual Basic intrinsic controls or ActiveX controls to the collection.

**

To reproduce:

**

1.  Create a new Standard Exe.
2.  Add the following code:

<pre>

> <font size="2" face="Courier">Private Sub Form_Load ()
>    Controls.Add "Excel.Application", "MyExcelApp")
> End Sub</font>

</pre>

4.  Run the exe.

**Result:** You get an error (720): Invalid class string.

#### <a name="Controls11"></a>Hierarchical FlexGrid Control: Correcting Errors Binding a Recordset to the HFlexGrid

If you receive the following error when trying to bind the **Hierarchical FlexGrid** to an ADO Recordset object, "DataSource settings may be incorrect", try changing some of the behavioral properties associated with the ADO Recordset Object or Command. For example, change the **CursorLocation** property to **adUseNone** or **adUseClient**.

#### <a name="Controls12"></a>Hierarchical FlexGrid Control: How to Change the Font of Individual Bands

Since the same font object is used for the entire grid object, you must create a new font object to change the fonts of individual bands, rather than changing the font directly.

For example, this way will not change the font for the individual band:

<pre><font size="2" face="Courier">

> MSHFlexGrid1.FontBand(1).Name = "Arial"

</font></pre>

Since you are directly modifying the font object, this will change the fonts in all of the bands to Arial.

To change an individual band, first create a new Font object, then assign that Font object to the **FontBand** property:

<pre>

> <font size="2" face="Courier">Dim ft As New StdFont
> ft.Name = "Arial"
> Set MSHFlexGrid1.FontBand(1) = ft</font>

</pre>

This will change just the band's font to Arial.

#### <a name="Controls13"></a>Hierarchical FlexGrid Control: Avoiding the display of duplicate headers

By default, the Hierarchical FlexGrid control uses the first FixedRow in the Hierarchical FlexGrid as a set of headers (which means it displays the names of the fields bound to each column in this row). Since by default the HFlexGrid control displays one FixedRow, if you enable the display of headers on Band 0, it will appear as though the header is being duplicated twice. To avoid this, set the FixedRow property to 0, or clear out the text values in the first FixedRow using code.

#### <a name="Controls14"></a>ADO Data Control: FetchProgress and FetchComplete Events Not Implemented

Although the reference topic for the **ADO Data Control** includes links to the FetchProgress and FetchComplete events, the events are not implemented for the control.

#### <a name="Controls15"></a>DataGrid: SizeMode and Size Properties Do Not Accept Value of 2 (dbgNumberOfColumns)

The reference topics for the **Split** object's **SizeMode** and **Size** properties refer to a non-existent property value of 2 (**dbgNumberOfColumns**). Please ignore this value.

#### <a name="Controls16"></a>Controls: ImageList Control on Page Designer

When using the ImageList control on a DHTML Page designer, images cannot be added at design time. If you try to use the following code in an uncompiled .dll project, you will get the run-time error: -2147418113 (8000ffff), "Method 'Add' of object images failed".

<pre>

> <font size="2" face="Courier">Private Sub DHTMLPage_Load()
>    ImageList1.ListImages.Add , , LoadPicture("C:\Winnt\winnt.bmp")
> End Sub</font>

</pre>

However, the code will work when the .dll project is compiled.

#### <a name="Controls17"></a>MSComm Control: EOFEnable Property Doesn't Stop Data Input

The **EOFEnable** property determines if the OnComm event occurs when an EOF character is detected. Contrary to the documentation for the property, however, input does not stop.

#### <a name="Controls18"></a>Treeview Control: Node Object's Visible Property is Read-Only

The **Visible** property of the **Treeview** control's **Node** object is a read-only property. If the node is not visible, you can use the **EnsureVisible** method to make it visible, as shown in the example:

<pre>

> <font size="2" face="Courier">Private Sub Command1_Click()
>    If Not TreeView1.Nodes(10).Visible Then
>       TreeView1.Nodes(10).EnsureVisible
>    End If
> End Sub</font>

</pre>

#### <a name="Controls19"></a>SysInfo Control: Constants Not Supported

The reference topics for the following events:

*   DeviceArrival Event
*   DeviceOtherEvent Event
*   DeviceQueryRemove Event
*   DeviceQueryRemoveFailed Event
*   DeviceRemoveComplete Event
*   DeviceRemovePending Event

have lists of constants that identify devices and device data. Contrary to the documentation, however, these constants are not supported by the events or the SysInfo control. The values associated with the constants listed in the help topics are valid, but the constant names are not.

#### <a name="Controls20"></a>User Control: Binary Persistence of PropertyBag Data Causes Page Designer to Fail

The PropertyBag saves data in binary format. Due to a known problem with binary persistence and the DHTML page designer, however, such data causes the page designer and Visual Basic to fail. See [Page Designer: Binary Persistence Issue](#PageDsr5) for more information.

* * *

### Language Issues

#### <a name="Language1"></a>InStr Function and Locale-Specific Comparisons

To use locale-specific rules in a comparison, enter a valid LCID (LocaleID).

#### <a name="Language2"></a>SendKeys Statement Gives Invalid Procedure Call Error

The short form of the code for sending an Insert, {INS}, results in an "Invalid procedure call" error under Windows NT 4.0 Service Pack 3\. To work around this problem, use the long code for Insert, {Insert}.

#### <a name="Language3"></a>Type Statement Clarification

The last sentence of the "Type statement" Help topic states: "The setting of the **Option Base** statement determines the lower bound for arrays." This sentence is incorrect and should be ignored. The **Option Base** setting has no effect on arrays in user-defined types.

#### <a name="Language4"></a>Decimal Data Type Stored As Signed Integer

The "Decimal Data Type" Help topic states that **Decimal** variables are stored as unsigned integers, which is incorrect. **Decimal** variables are stored as signed integers.

#### <a name="Language5"></a>DateSerial Function and Windows 98/Windows NT 5

For the **_year_** argument, two-digit years are interpreted based on user-defined machine settings (the default range is 1930-2029). The range settings are defined in the Regional settings of the Microsoft Windows Control Panel.

#### <a name="Language6"></a>Code Window "Find Next" Keyboard Shortcut

<dl>The "Code Window Keyboard Shortcuts" Help topic incorrectly states the Find Next keyboard shortcut is SHIFT+F4\. The correct keyboard shortcut for Find Next is F3.</dl>

#### <a name="Language7"></a>Add Method (Folders) Syntax

In the "Add Method (Folders)" Help topic, the syntax shown is incorrect. The correct syntax is:

<pre><font size="2">_object_.**Add** _foldername_</font> </pre>

* * *

### Samples Issues

#### <a name="Samples1"></a>Sample File Locations

If you choose to include Visual Basic samples in your MSDN setup, they are installed to the directory:

> C:\Program Files\Microsoft Visual Studio\MSDN98\98VS\1033\Samples\VB98\.

If you choose not to include the Visual Basic samples in your MSDN setup, you can find the Visual Basic samples on the MSDN CD at:

> D:\Samples\Vb98

**Note:** The drive letters mentioned above may vary on your system.

#### <a name="Samples2"></a>Visual Basic Sample: Biblio and Mouse Samples Omitted

The Biblio sample program, found in the Visual Basic documentation table of contents, is no longer included with the Visual Basic product. The Mouse sample, metnioned in "Responding to Mouse and Keyboard Events" is also no longer included with the product.

#### <a name="Samples3"></a>Visual Basic Samples: ChrtSamp Description

ChrtSamp is a new sample program included with Visual Basic that demonstrates the major features of the MSChart control. If you have installed the Visual Basic samples, the sample can be found in the following location on your hard disk:

\\Program Files\Microsoft Visual Studio\Msdn98\98vs\1033\Samples\Vb98\ChrtSamp

If you have not installed the Visual Basic samples on your hard disk, the sample can be found on the MSDN CD at the following location:

Samples\Vb98\ChrtSamp

The sample uses an Excel spreadsheet to supply data for a chart. The sample also allows you to display multi-series charts by clicking various buttons. Finally, the sample demonstrates 3D features of the control by setting the ChartType property to an appropriate value.

<table width="559" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="20%" valign="TOP">

**<font size="2">File</font>**

</td>

<td width="80%" valign="TOP">

**<font size="2">Description</font>**

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">Chrtsamp.vbp</font>

</td>

<td width="80%" valign="TOP">

<font size="2">The project file for the sample.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">Frmchart.frm</font>

</td>

<td width="80%" valign="TOP">

<font size="2">The main form for the sample.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">Frmchart.frx</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Binary data for the form.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">Gas.xls</font>

</td>

<td width="80%" valign="TOP">

<font size="2">The Excel worksheet containing the data.</font>

</td>

</tr>

<tr>

<td width="20%" valign="TOP">

<font size="2">Modchart.bas</font>

</td>

<td width="80%" valign="TOP">

<font size="2">Code module containing functions and procedures for the sample.</font>

</td>

</tr>

</tbody>

</table>

**

To Run

**

Press F5 to run the sample. After all data has loaded, click on the "Chart Type" to change the type. To see the three-dimensional features of the sample, click the Chart Type combo box and select a 3-D chart type, such as **3dArea**. While holding down the CTRL key, right-click the chart, and use the cursor to change the aspect of the chart.

#### <a name="Samples4"></a>Visual Basic Samples: CtlsAdd Sample: Controls.mdb Must Be Read/Write Enabled on Hard Disk

If you attempt to run the CtlsAdd samples from the MSDN CD, an error will occur if you attempt to use the controls.mdb database found on the CD. Because CtlCfg.vbp sample uses an Access database (controls.mdb) to store license key information about controls, the database must be installed on a hard disk. Copy the controls.mdb file to the hard disk and make it writable.

#### <a name="Samples5"></a>DHSHOWME.VBP Sample: You May Need to Reset the SourceFile Property For This Sample to Work Correctly in Design Mode

If your page designer samples appear blank when you open them in design mode, please reset the SourceFile property to reflect the location on your computer to which the project's HTML files were installed. You must reset this property for each designer in the project. Your sample should then work correctly.

To reset the SourceFile property for a designer, either type a path directly into the SourceFile property in the Properties window, or select the **Project Properties** icon from the toolbox, click **Save as an External File**, then click **Open** and navigate to the correct .htm file.

#### <a name="Samples6"></a>PROPBAG.VBP Sample: Possible Error on Loading the Module for this Sample

Propbag.vbp references a module (Module1.bas) that is located in the default installation directory for MSDN samples. If you move this sample to another directory, you will receive an error when you open the project that the module path is incorrect. To fix this, load the project without the module, then re-add the module from the directory to which you installed your samples.

#### <a name="Samples7"></a>Running the IObjSafe Sample Application

</font><font size="2" face="Verdana,Arial">

Due to some late-breaking changes, the IObjSafe sample application (IObjSafe.vbp) will not run properly unless you first make the following modifications:

1.  Load the IObjSafe.vbp project file into the development environment.
2.  Choose **ucObjSafety Properties** from the **Project** menu. On the **Debugging** tab, delete the path and file name from the **Start program** text box.
3.  Type the actual path and file name for IObjSafe.htm in the **Start browser with URL** text box. The actual path may vary depending on your installation.
4.  Choose **Options** from the **Tools** menu. On the **General** tab, choose the **Break on Unhandled Errors** option button.

The application should now run properly.

An updated version of the IObjSafe sample application is available online at the [Microsoft Visual Basic samples page](http://msdn.microsoft.com/vbasic/downloads/).

#### <a name="Samples9"></a>Obtaining Updated Versions of Sample Applications

Updated versions of many of the Visual Basic sample applications as well as additional samples not included on the CD are available online at the [Microsoft Visual Basic samples page](http://msdn.microsoft.com/vbasic/downloads/).

* * *

### Wizard Issues

#### <a name="Wizard1"></a>Package and Deployment Wizard: Automatically Pick Up Files From Redist directory

A new feature of the Package and Deployment wizard is its ability pick up files from the Redist folder. An example of how the feature is used would follow the steps below:

1.  You have created an application that's ready for packaging.
2.  The application depends on a certain system file "MySysFile.dll" which, until recently, has been a stand-alone file. But now a recent system update has made that file dependent on other files to function properly.
3.  However, a stand-alone version of the system file is also supplied.
4.  To simplify setups, you place that stand-alone version into the Redist folder.
5.  When creating your deployment package, you navigate to the system file to instruct the wizard that the system file should be included.
6.  Instead of picking up the system file (which is dependent on other files), the Package and Deployment wizard will pick up the stand-alone version that is in the Redist directory.

By default, no files are placed in the redist folder.

#### <a name="Wizard2"></a>Package and Deployment Wizard Has [Do Not Redist] Section

A new section has been added to the Package and Deployment VB6Dep.ini file: [Do Not Redist]

Two kinds of files are listed under the new section:

1.  Files that are not needed at run time. For example, ActiveX designers often require two file references, one for design time, and one for run time. Since the design time reference is not needed, it is listed in the section and is not included in the package.
2.  Files that cannot be redistributed by the Package and Deployment wizard. Some dependent files are not installed by Visual Basic, but must be installed by another component. For example, the following files are installed by Internet Explorer 4.x: Shdocvw.dll and Mshtml.dll.

#### <a name="Wizard3"></a>Package and Deployment Wizard: In Silent Mode the Notification Window May Not Be First in the Window Z-order

With Visual Basic, it's possible to run the Package and Deployment Wizard in _silent mode_ from a command prompt. When doing so, you can also set a path for logging the wizard's output (/**l** ). With the path set, the wizard will create a log of events. If you do not use the argument, the wizard will instead display a dialog box to notify you of the wizard's completion. However the window is not always obvious because it may be underneath other windows. In order to see it, you may need to close or minimize all other windows.

For more information, search online for "running as a stand-alone" in _Visual Basic Concepts_ in the MSDN Library Visual Studio 6.0 documentation.

#### <a name="Wizard4"></a>Package and Deployment Wizard: Command Line Mode Argument Added for Specifying Executable Path

An argument has been added to the command line mode of the Package and Deployment wizard. The argument is shown in the table below:

<table width="590" cellspacing="1" cellpadding="7" border="">

<tbody>

<tr>

<td width="17%" valign="TOP">

**<font size="2">Argument</font>**

</td>

<td width="83%" valign="TOP">

**<font size="2">Description</font>**

</td>

</tr>

<tr>

<td width="17%" valign="TOP">

**<font size="2">/e _path_</font>**

</td>

<td width="83%" valign="TOP">

<font size="2">Use the argument to set the path for the project's executable file if it is different from the project's path.</font>

</td>

</tr>

</tbody>

</table>

<font size="2">**

Remarks

**

The argument allows the command line mode to be used in a multi-developer environment where a computer is dedicated to compiling the project and creating the package.

For more information, search online, with **Package and Deployment Wizard** selected, for _running as a stand-alone_ in the MSDN Library Visual Studio 6.0 documentation.

#### <a name="Wizard5"></a>Package and Deployment Wizard: Manually Add User Control License Files

When creating a package for a user control that requires a license key, the license file (.vbl) is not automatically included. Instead, you must add the file manually.

When you are at the dialog titled "Package and Deployment Wizard – Included Files," click the **Add** button and search for the appropriate .vbl file.

</font>

#### <a name="Wizard6"></a>Package and Deployment Wizard: Steps in the Web Deployment Process

When you post an Internet package to a Web server using the Package and Deployment Wizard, the wizard uses a technology known as WebPost to copy your package to the server and process it appropriately. These are the steps that the WebPost process goes through when processing your .cab:

1.  It extracts the .cab file into a temporary directory.
2.  It locates the .inf file for the .cab file.
3.  Based on the contents of the .inf file, the WebPost process installs application files (based on the RInstallApplicationFiles section of the .inf file), installs system files (based on the RIinstallSystemFiles section), and installs shared files (based on the RInstallSharedFiles section). In the process, it registers any necessary files.

> **Note** The DefaultInstall section of the .inf file is not run, because the instructions it contains often require user input. The WebPost process also does not create a virtual directory for your application, if one is required; directories must be set up in advance.

#### <a name="Wizard7"></a>Package and Deployment Wizard: Web Deployment Tips for HTTP or FTP Protocol

*   If the .cab file you are attempting to deploy to a Web server is copied to the server but is not unpacked, make sure you included the .cab file on the Items to Deploy screen, and that you used HTTP Post as the protocol on the Web Publishing Site screen. In addition, you must have checked the Unpack and Install Server-Side Cab check box on the Web Publishing Site screen. If you did not, try re-deploying your package again with this option selected.

> **Note** Cab unpacking is supported by Posting Acceptor 2.0 running on IIS 4.0.

*   If you receive an error saying that the Web server you selected does not support the service provider you selected, there are several things you can do to try to fix this problem:

1.  If you are using the HTTP Post protocol, make sure that Posting Acceptor is installed on your Web server. If you are running IIS 4.0, install Posting Acceptor 2.0\. If you are running IIS 3.0, install Posting Acceptor 1.0\. Cab unpacking is not supported by Posting Acceptor 1.0.
2.  Ensure that your URL is correct. If you are using HTTP Post, ensure that your URL begins with http://. If you are using FTP, ensure that your URL begins with ftp://.
3.  If you are uploading to a server location that has Posting Acceptor 1.0 installed, you cannot select the option to unpack and install your cabinet files when you deploy your package using the wizard. If you receive an error C0042115 that the query string INSTALL is invalid, check your server configuration to determine what version of posting acceptor is installed. If it is version 1.x, deploy your cab again, making sure to deselect the **Unpack and Install Server-Side Cab** option.
4.  If you need to use the FTP protocol to post to a URL beginning with http://, you may be able to resolve this error by adding the following entry to the end of your Web server's postinfo.asp file, usually located in the scripts directory on the Web server:  

    <!--This entry is for the FTP protocol -->  
    [{02B5E1D1-8B7C-11D0-AD45-00AA00A219AA}]  
    ServerName="<%= Request.ServerVariables("SERVER_NAME") %>"

*   If you receive an error that some of the files in the INF are busy and require a reboot, change the **ComponentPermitReboot** entry in the reg key HKEY_LOCAL_MACHINE\Software\Microsoft\Publishing\RemoteInstaller from No to **Yes**. This will allow cab unpacking and installation to proceed when files are busy, but a manual reboot of the server may be necessary in order for the installation to finish successfully.
*   If you use the HTTP Post protocol and receive an error that you do not have write access for the Web server, open the Internet Service Manager on the server computer, access the node for your Default Web Site (Console Root \ Internet Information Server \ machinename \ Default Web Site), choose Properties, select the Home Directory tab, and check the Write check box.
*   If your files are read-only when you try to post, you will receive an error C0042116 announcing that processing has stopped. Change your file attributes to proceed.
*   Be aware that if you post a file to a directory where the same file already exists, the file on the server will be overwritten and no warning message will be displayed.
*   If you use the FTP protocol and receive an error that access is denied, open the Internet Service Manager on the server computer, access the node for your Default FTP Site (Console Root \ Internet Information Server \ machinename \ Default FTP Site), choose Properties, select the Home Directory tab, and check the Write check box.
*   If you use the FTP protocol and receive additional errors besides the one described in the previous bullet, make sure that you have properly configured your FTP service on the Web server. To do so, start the Microsoft Management Console (MMC), then follow these steps:

1.  Right-click on **Default FTP Site** and choose **New Virtual Directory**.
2.  Enter an alias to be used to access the virtual directory, then click **Next**.
3.  Enter the physical path of the directory to which to map the virtual directory -- for example, c:\inetpub\ftproot -- then click **Next**.
4.  Select the appropriate access permissions, making sure to allow write access so that your deployments can proceed without errors, then click **Finish**.
5.  Right-click on **Default FTP Site** and choose **Stop**.
6.  Right-click on **Default FTP Site** and choose **Start**.  

    When you deploy to the FTP server using the Package and Deployment Wizard, use the site FTP://servername/_alias_ where _alias_ is the alias you assigned in step 2\. Use the username "anonymous" and the password "me@somewhere" for anonymous login.

#### <a name="Wizard8"></a>Package and Deployment Wizard: Start Menu Items: Run Option Not Supported

When creating a package for deployment, you can also create a Start menu item. Although it is mentioned in the help topic for the dialog box, the Run option is not supported.

#### <a name="Wizard9"></a>System Configurations for WebPost's Posting Acceptor

When you deploy your packages to a Web server, the Package and Deployment Wizard uses a technology known as WebPost to move your files to the desired location. WebPost consists of two main components:

*   The Package and Deployment wizard, which sends your content to a qualified site.
*   A Posting Acceptor, located on the Web server, that enables the posting of content to an IIS server.

There are multiple versions of Posting Acceptor available. You must make sure you install the correct version on your Web server based on your machine configuration. The following table lists the appropriate configurations:

<table width="561" cellspacing="1" cellpadding="7" border="1">

<tbody>

<tr>

<td width="25%" valign="TOP">

<font size="2">**Use this**</font>

</td>

<td width="30%" valign="TOP">

<font size="2">**If you are running**</font>

</td>

<td width="45%" valign="TOP">

<font size="2">**Comments**</font>

</td>

</tr>

<tr>

<td valign="TOP">

<font size="2">Posting Acceptor 2.0</font>

</td>

<td valign="TOP">

<font size="2">Windows NT 4.0 with SP3  
IIS 4.0</font>

</td>

<td valign="TOP">

<font size="2">Posting Acceptor 2.0 supports both content posting and the unpacking of cabinet (.cab) files on the server.</font>

</td>

</tr>

<tr>

<td valign="TOP">

<font size="2">Posting Acceptor 1.0</font>

</td>

<td valign="TOP">

<font size="2">Windows NT 4.0 with SP3  
IIS 3.0</font>

</td>

<td valign="TOP">

<font size="2">You cannot unpack .cab files with this version of posting acceptor. Use this for content posting only. You can move your files to the server and then manually register any necessary files that would have been registered by the .cab process.</font>

</td>

</tr>

</tbody>

</table>

> **Note:** Posting Acceptor does not work on any platforms that are using Personal Web Server or Peer Web Server. You must use IIS.

You can install Posting Acceptor 2.0 from the Deploy folder of Visual Studio installation CD number 2\. Posting Acceptor 1.0 can be installed from the [Microsoft Posting Acceptor Information Website](http://www.microsoft.com/windows/software/webpost/post_accept.htm). If you want to install Posting Acceptor 2.0 on a computer that already has the Windows NT Option Pack, you should first check to see if version 1.0 of the Posting Acceptor is already installed. If so, remove it before installing version 2.0\. To determine if 1.0 is installed, select the **NT Option Pack** on the **Add/Remove Programs** mechanism of the **Control Panel**. Click **Add/Remove**, then look for Posting Acceptor 1.0 among the components listed. You may find it under Microsoft Site Server. If it is installed, remove it. You can then install version 2.0 by running PASetup.exe.

#### <a name="Wizard10"></a>Package and Deployment Wizard: Edit Setup.lst file if you rebuild cabs from batch file

After creating a standard setup package using the Package and Deployment wizard, you can manually recreate the setup files (Setup.exe, Setup.lst, and all .cab files) by running the batch file found in the in the Support folder. Doing this allows you to manually customize the package or to recreate a package without running the wizard again. Running the batch file will copy setup.exe and setup.lst from the Support folder to the Package folder and will generate the cab file(s) in the Package folder. However, once the batch file is finished, setup.lst will not know how many .cab files were generated. Unless this problem is remedied, the setup program will fail.

To remedy this situation, do the following:

1.  Before running the batch file, delete all .cab files in the Package folder.
2.  After running the batch file, count the number of cabs produced by the batch operation.
3.  Open the Setup.lst file in the Package folder with a text editor. Note: there are two Setup.lst files. One is found in the Support folder. The second is found outside the Support folder, in the Package folder where the .cab files are created. Be sure to open the Setup.lst file in the Package folder.
4.  In the text editor, look for the following lines (which are at the top of the file):

<pre> <font size="2" face="Courier">> [Bootstrap]
> SetupTitle=Install
> SetupText=Copying Files, please stand by.
> CabFile=Projec1.CAB
> Spawn=Setup1.exe
> Uninstal=st6unst.exe
> TmpDir=msftqws.pdw</font></pre>

6.  After the last line (<font size="2" face="Courier">TmpDir=msftqws.pdw</font>), insert the following line:

</font>

> <font size="2" face="Courier">Cabs=</font><font size="2" face="Verdana,Arial">_N_</font>

<font size="2" face="Verdana,Arial">

> where N equals the number of cabs generated.

The Setup.lst should now be up to date, and the setup won't fail.

#### <a name="Wizard11"></a>Package And Deployment Wizard: Error 80042114

If you are creating a package for Internet deployment and you get the following error:

> Unexpected error 80042114 has occurred: The Web server you selected does not indicate support for the service provider you selected. Do you want to proceed anyway?

This error occurs because you have specified that your package should be posted to an FTP site using an HTTP URL. If you are certain that you have access privileges to the web server, you can click "Yes" and the deployment will proceed.

To avoid this error in the future, when creating a package for deployment on the same server, specify the site and its protocol. When you do so, the following dialog box will appear:

> The specified URL and publishing method can be saved in the registry as a Web publishing site. This ensures that the URL and publishing method are valid and saves you time in future deployments to this site. Do you want to store this information as a Web publishing site?

If you select "yes", this saves the site information and you will no longer receive the 80042114 error.

#### <a name="Wizard12"></a>Package and Deployment Wizard: Use Mdac_typ.Cab and Mdac20.Cab to Distribute Data Access Components

When your Internet Package includes any of the following four files, the Wizard will by default set these files to be downloaded from [http://activex.microsoft.com/controls/vb6/mdac_typ.cab](http://activex.microsoft.com/controls/vb6/mdac_typ.cab).

<pre><font size="2" face="Courier">MSDAOSP.dll 
MSADO15.dll 
MSADCF.dll 
ODBC32.dll</font> </pre>

Similarly, the Wizard will, by default, set the following files to be downloaded from [http://activex.microsoft.com/controls/vb6/mdac20.cab](http://activex.microsoft.com/controls/vb6/mdac20.cab):

<pre><font size="2" face="Courier">MSADOR15.dll 
MSADCO.dll</font> </pre>

In both cases, these defaults are indicated as the "Download from Microsoft Web site" option on the File Source screen. These cab files (mdac_typ.cab and mdac20.cab) perform special handling that should not be attempted manually. In order to ensure that the proper handling takes place, your Internet cab should not include these files but rather should reference the cabs. Therefore, you should never choose the "Include in this cab" option for any of these files. In addition, if you choose the "Download from alternate Web site" option, you should be careful to specify cabs that are copies of these cabs to ensure that the proper handling takes place.

Do not change the default settings for these files.

#### <a name="Wizard13"></a>Package and Deployment Wizard: Manually Include .ASP and .HTM Files For IIS Applications When Using Standard Setup

If you use the Package and Deployment wizard's standard setup to deploy an IIS application, you must manually include any .asp or .htm files with the package. Add the files using the Include Files dialog box.

#### <a name="Wizard14"></a>Package and Deployment Wizard: Bad Date and Time Formats

In certain situations, the Package and Deployment Wizard will incorrectly write time and date information to the Setup.lst file. When this occurs, the setup will fail because the dates will be written in a form that the setup.exe can't read. The problem occurs when you create a deployment package using the **US** version of Visual Basic on:

*   A computer using the German version of Windows (Note: the **German** version of Visual Basic will work correctly.)
*   Any computer where the date separator isn't either a forward slash ("/") or a dash ("-").

**

To fix the bad formats:

**

1.  From the Start menu, open **Control Panel**.
2.  Click the **Regional Settings** icon.
3.  On the Date tab, change the Date separator to "/".
4.  On the Time tab, change the Time separator to ":".
5.  Run the Package and Deployment Wizard.
6.  Restore the date and time settings.

#### <a name="Wizard15"></a>Package and Deployment Wizard: Unable to Run Setup.exe on First Windows 95 Version

Any setup package built with the wizard will not launch on certain installations of Windows 95 due to lack of support for an API in the original version of the Oleaut32.dll. This failure will **not** occur on OS release 2 of Windows 95 or any versions of Windows NT 4.0 and later, and will not occur if Microsoft Office 97 or Internet Explorer 3.0 or 4.0 has been installed. Any installation of a Visual Basic 5.0 application will also remedy the situation. You can also work around this failure by first overwriting the older version of Oleaut32.dll with the latest version. Be sure to shut down all applications before attempting to manually update Oleaut32.dll.

#### <a name="Wizard16"></a>Package and Deployment Wizard: Packaging ActiveX Documents

The Visual Basic 6.0 Package and Deployment Wizard can insert CODEBASE= and VERSION= information directly embedded into the .vbd file for your ActiveX Document project. This eliminates the need for having the extra .htm file previously used to launch ActiveX Documents. The embedded information allows Internet Explorer to read the .cab file name for your ActiveX Document code and version information from the .vbd file and carry out the installation. You can now navigate directly to the .VBD file and your User document code will download if necessary.

The same functionality was available for Visual Basic 5.0 setup using the SetCodeBase utility found on the Visual Basic Owners Area.

The following are some issues with Visual Basic 6.0 Package and Deployment Wizard generated setups for User documents:

*   Internet Explorer 3.x. 4.0, and 4.01 cannot read the embedded information in the VBD file. The wizard also generates an old (VB5) style .htm file. This .htm file can be used to launch ActiveX Documents. You must first modify the .htm file, however, since most of the code is commented out. After removing the comments and an extra <A> tag(<a href=xxxx.VBD>xxxxx.VBD</a>) from the top, the file will only be an OBJECT tag with the CLSID of your User document and some script code for Window_OnLoad event.

*   Internet Explorer 4.01 with Service Pack 1 and later will read this information correctly from the .vbd file.

* * *

### Error Message Issues

#### <a name="Error1"></a>No Help Topic for the Following Error Messages

There are currently no Help topics for the following error messages:

*   "Object module needs to implement '<name>' for interface '<classname>'."

An interface is a collection of unimplemented procedure prototypes. This error occurs when you specified an interface in an Implements statement, but you failed to add code for all the procedures in the interface.

You must write code for all procedures specified in the interface. An empty procedure containing only a comment is sufficient.

For additional information, select the item in question and press F1.

*   "Private Enum types and Enum types defined in standard modules or private classes cannot be used in public object modules as parameters or return types for public procedures, as public data members, or as fields of public user defined types."

This error occurs when you attempt to use an Enum type (or private Enum type) as:

*   A parameter for a public object module
*   A return type for a public procedure
*   A public data member
*   Fields of public user-defined types

Avoid using Enum or private Enum types in these circumstances.

*   "Can't ReDim, Erase, or assign to Variant that contains array whose element is With object."

This error occurs when you attempt to ReDim, Erase, or assign to a Variant a variable whose element is a With object. For example, the following code will produce this error:

<pre> <font size="2">Type Test
      Name as Integer
   End Type

   Sub Main()
      Dim c(0) As Test
      Dim e
      e = c
      With e(0)
         ReDim e(1)
      End With
   End Sub</font> </pre>

* * *

### WebClass Designer Issues

#### <a name="WebClass1"></a>Webclasses: "Me." Not Supported

You cannot use the "Me" reference in your webclass code to reference the webclass object. For example, the documentation frequently shows that you can write code such as "Me.URLData = _value_". This is not supported. Instead of using Me, you must use the "Webclass" statement. For example, instead of Me.URLData, you would use Webclass.URLData.

#### <a name="WebClass2"></a>Webclasses: Invalid HTML Syntax Can Cause Unspecified Error

If one of the templates you add to your webclass contains badly formed HTML, you will sometimes receive an error message on loading the template. The message states only that an unspecified error has occurred. For example, in older pages there may be two BODY tags, one that specifies a background GIF and one that specifies a color. You may also have errors in unmatched opening and closing tags, invalid nesting, or other syntax issues. If you receive this message on loading a template, check your HTML carefully or run the file through an HTML syntax checker, then reload the template.

#### <a name="WebClass3"></a>Webclasses: Avoid Using Global or Static Variables in a Webclass

One allocation of global variables occurs per thread in a multi-threaded environment. For more information, [search online](#Critical2), with **Search titles only** selected, for "Scalability and Multithreading" in the MSDN Library Visual Studio 6.0 documentation.

#### <a name="WebClass4"></a>Webclasses: Some External HTML Changes are not Detected Automatically

When working on an HTML template in the webclass designer, any changes made to the HTML file outside of Visual Basic (for example, in an external HTML editor) are usually detected by Visual Basic when you return to the designer. In these cases you are prompted to reload the changed file.

In some cases external changes are not detected. The most common occurrence of this problem is when you set focus to a Visual Basic window other than the webclass designer before switching to an external editor. Upon return to Visual Basic, the refresh prompt does not appear. This could result in the external changes being overwritten when you save the project, unless you refresh the file on your own.

> **Note** You may also see this situation if you edit the template while your project is running.

In cases where you make changes to the HTML and are not prompted to refresh, you can refresh manually by selecting Refresh HTML Template from the template's shortcut menu.

**Note:** When you navigate to your external HTML editor, it is best ot use the Edit HTML toolbar button or shortcut menu command. If you use the taskbar or th ALT+TAB to navigate to an editor, make sure to save your project before leaving Visual Basic, or you could lose changes you made in the designer.

#### <a name="WebClass5"></a>Webclasses: IIS Administration Console File Settings are not Acknowledged for Templates

The IIS Administration Console allows the server administrator to specify properties for files that are available on the IIS server. These properties include HTTP headers, file security, and custom errors. These properties will not be set on a webclass template file if that file is sent to the client by the webclass runtime.

#### <a name="WebClass6"></a>Webclasses: Unattended Execution

A project containing a webclass must have the Unattended Execution option selected in the Project Properties dialog box. This property has the following benefits:

*   Setting this property allows the webclass to be run as an apartment model object. This allows the webclass to service an HTTP request on the thread on which the request was received instead of processing all requests on a single thread.

> **Note** You must set the Threading Model property in the Project Properties dialog box to Apartment Threaded to run as an apartment-model object.

*   Setting this property causes the Visual Basic run-time DLL to log all run-time errors to the event log instead of displaying the error in a prompt. Displaying the message in a prompt would hang the IIS thread.

*   Setting this property causes any call to the Visual Basic MsgBox function to log its message to the event log instead of displaying a prompt. Displaying the message in a prompt would hang the IIS thread.

#### <a name="WebClass7"></a>Webclasses: Retain in Memory

In standard Visual Basic projects, projects are unloaded from threads or processes as soon as they are no longer being used. In a webclass project, this model can cause performance issues because the server must create an object, invoke a method on it, and destroy it. You can optimize your webclasses by setting a project property called Retain In Memory. The Retain In Memory property prevents the project from being unloaded until the thread or process in which it is running terminates.

#### <a name="WebClass8"></a>Webclasses: Accounting for Differences Between the Debug and Compiled Versions

Visual Basic provides the ability to debug components running under a Windows NT service. One of the most common uses of this feature is to debug an IIS Application. Visual Basic achieves this by running the component in the Visual Basic IDE. When the component runs, IIS creates a proxy object supplied by Visual Basic, which in turn creates the real object running in the Visual Basic IDE. IIS then communicates with the object through DCOM.

This debugging behavior is very different from how the project runs as a compiled DLL. Certain behavior that is present in debug mode works differently when you run the compiled version of the project. Because of this, you must keep the compiled behavior of the project in mind when you build your webclass.

The following are key areas in which you must tailor your application to the behavior the webclass displays as a compiled application:

*   Use only system DSNs because other DSNs will not work beyond debug mode.
*   Do not use an Access database on a remote computer in your project. While this will work in debug mode, you will not be able to use the database in the compiled application.
*   Do not allow the webclass to add itself or other Visual Basic components to the Active Server Page's Application object. Attempting to do so will generate an error when you run the compiled application.
*   Understand the security context of the compiled application. Please refer to the "Webclasses: Articles of Interest" section, below, for information on an article about security.
*   Keep in mind that your compiled webclass will be accessed from multiple threads, rather than through the same thread, as is the case in debug mode. Static and global variables will not be kept across threads. For more information, [search online](#Critical2), with **Search titles only** selected, for "Scalability and Multithreading" in the MSDN Library Visual Studio 6.0 documentation.
*   Understand that although you will see message prompts in debug mode, the compiled webclass writes all errors as entries in the NT event log or in a log file created in the Windows directory. No prompt appears for errors in compiled mode.
*   While Unattended Execution must be set for webclasses, you will not see the side effects of failing to set this property in debug mode. See the Unattended Execution section, above, for details.

#### <a name="WebClass9"></a>Webclasses: Performance Tips

The following are miscellaneous tips you can incorporate to improve the performance of your IIS applications:

*   Make sure that the Unattended Execution and Retain In Memory options are selected in the Project Properties dialog box for your application.
*   If your application does not include any text replacements, set the TagPrefix property to an empty string. This prevents the webclass from performing unnecessary scans.
*   Do not store Visual Basic objects (or any other apartment-model COM object) to the Active Server Pages' Session object. This may affect scalability. You can store strings in the Session object without adverse effects. Refer to the IIS documentation for more details.
*   Limit the use of variants in your application.
*   When the webclass's StateManagement property is set to wcRetainInstance, performance will decrease when the number of clients significantly increases.
*   If your application is performing a client-side transaction to a webclass template that does not contain any replacements or use the URLData property, you should access the template directly through a URL.
*   When using the URLFor method, specify the webitem by the string name rather than by an object reference.
*   Use specific types when creating and invoking other components.

#### <a name="WebClass10"></a>Webclasses: Miscellaneous Issues

*   When the debugger for your IIS application hits a breakpoint in any event, pressing F5 to continue does not return the focus to Internet Explorer. You must switch to Internet Explorer manually after continuing.
*   Webclass names and tag names are case insensitive. You cannot rename a webclass to the same name it had previously, changing only the case. For example, if you change a webclass named Orderentry to OrderEntry, the original name remains unchanged.
*   Avoid running multiple browser instances during debugging. If more than one instance of Internet Explorer is open, Visual Basic does not keep track of which browser is running the webclass project. If you have two browsers open, one pointing to your project and one pointing to another page, both browsers will be affected when you end your debugging session.
*   You may receive an error if you attempt to compile your IIS application project from the command line. One way to work around this is to open the project in Visual Basic, dirty the designer in some insignificant way, and then resave the project. You can then restart the compile from the command line and it should work correctly.
*   If you want to program buttons on the HTML templates in your webclass, you must be aware of two items. First, your buttons must be of type SUBMIT. You can set this by adding a parameter to the HTML for your button element that says type=SUBMIT. Second, you cannot code for the button directly; instead you must connect their form element. You can either place each button in a separate form or you can use the Request object's form collection to determine the button from which the event originated.

#### <a name="WebClass11"></a>Webclasses: Articles of Interest

Webclasses tie together several distinct technologies, including Visual Basic, Active Server Pages, Internet Information Server, and Windows NT. There are several articles available on Microsoft's Web site that may be useful to you as you learn about the technologies behind webclasses. Some of the articles that may be particularly helpful are listed below:

*   SiteBuilder Network main site.  
    [(http://www.microsoft.com/workshop/server/toc.asp.)](http://www.microsoft.com/workshop/server/toc.asp)
*   "Implementing a Secure Site with ASP"  
    [(http://www.microsoft.com/isn/techcenter/security.asp.)](http://www.microsoft.com/isn/techcenter/security.asp)
*   "Security Issues with Objects in ASP and ISAPI Extensions"  
    [(KB article Q172925)](http://support.microsoft.com/support/kb/articles/q172/9/25.asp)
*   "COM Security Frequently Asked Questions"  
    [(KB article Q158508)](http://support.microsoft.com/support/kb/articles/q158/5/08.asp)
*   "Descriptions of Workings of OLE Threading Models"  
    [(KB article Q150777)](http://support.microsoft.com/support/kb/articles/q150/7/77.asp)
*   "Automate Printing in ASP from COM Servers"  
    (KB article Q184291)
*   "COM Servers Activation and NT Windows Stations"  
    [(KB article Q169321)](http://support.microsoft.com/support/kb/articles/q169/3/21.asp)
*   "Launching ActiveX Servers from ISAPI Extensions"  
    [(KB article Q156223)](http://support.microsoft.com/support/kb/articles/q156/2/23.asp)
*   "Security Ramifications for IIS Applications"  
    [(KB article Q158229)](http://support.microsoft.com/support/kb/articles/q158/2/29.asp)

#### <a name="WebClass12"></a>Webclasses: Formatting in Source HTM File

You may see a loss of some formatting in your HTML source code after you add a template file to the webclass designer. For example, the webclass may remove some extraneous white spaces from your original file. This will not affect the functioning of your HTML page in any way.

#### <a name="WebClass13"></a>Webclasses: Cannot Support HTML's LINK Element

LINK tags are used in an HTML page to reference style sheets. While your HTML pages in a webclass project can contain this tag, you cannot use the designer to access the LINK element and process Visual Basic code for it. If you need to manipulate a LINK tag in your code, you can manually add event notation to the tag as shown in the online documentation. To see the notation, [search online](#Critical2), with **Search titles only** selected, for "Manually Adding Event Notation to an .HTM File" in the MSDN Library Visual Studio 6.0 documentation.

#### <a name="WebClass14"></a>Webclasses: When Using Visual SourceSafe with Webclass Projects, You Must Manually Check in the Project's .HTM Files

When you check an IIS application project into Visual Source Safe, the HTML pages associated with the project are not automatically checked into the Source Safe tree with the rest of the project files. You must manually add them to the tree as related files.

#### <a name="WebClass15"></a>Webclasses: TagPrefix Should Be WC:

Although the default value for the TagPrefix property for your webclass templates is **WC@**, it is preferable for you to use **WC:** whenever possible to indicate text replacements in your template files.  

#### <a name="WebClass16"></a>Webclasses: Variant Parameter in URLFor Method

The WebItem parameter of the URLFor method can accept a WebItem object or the name of a WebItem as a string. For performance reasons, you should use the string form when referencing multiple webitems within one request.

#### <a name="WebClass17"></a>Webclasses: Sequencing Data is Passed Using the &WCU Parameter

In "Handling Sequencing in Webclasses" section of the Building Internet Applications book in MSDN's Component Tools Guide, the documentation incorrectly states that you can move data between the client and the server using a **?Data** parameter appended on to your URL request. In fact, you must use a **&WCU** parameter instead of **?Data**. The correct syntax for the request is:

<tt><font size="2">

> WCI=webitem1?WCE=event1&WCU=01

</font></tt>

#### <a name="WebClass18"></a>Webclasses: StateManagement Property Constants Contain Incorrect Property Reference

The "StateManagement Property Constants" topic incorrectly states that the **RetainInstance** constant causes the webclass to retain state data until the webclass object calls the **SetComplete** method. This should say that data is maintained until the webclass object calls the **ReleaseInstance** method. To see the erroneous Help topic, [search online](#Critical2), with **Search titles only** selected, for "StateManagement Property Constants" in the MSDN Library Visual Studio 6.0 documentation.

#### <a name="WebClass19"></a>Webclasses: State and the Session Object

If the WebClass's StateManagementType is **wcRetainInstance**, a separate instance of the WebClass will be maintained in the ASP Session object per user session. In some cases, it may appear to you that state is not being maintained when you actually have two instances of a webclass in your Session object. One situation in which this might occur is when you have two virtual directories that both point to the same location. If you create one virtual directory when you begin your debugging session and reference the second in your code, you will actually start a second instance of the webclass when the code is activated.

Please refer to the Active Server Pages documentation in MSDN for details on how the Active Server Pages Session object is implemented.

#### <a name="WebClass20"></a>Webclasses: Code Corrections in Help Topic "Defining Webclass Events at Run Time"

In the topic "Defining Webclass Events at Run Time," the sample code shows a statement that reads:

<pre><font size="2">

> rs = New ADO.Recordset

</font></pre>

The correct syntax for this line should be:

<pre><font size="2">

> Set rs = New ADODB.Recordset

</font></pre>

#### <a name="WebClass21"></a>Webclasses: HTM and ASP Files Not Included in Standard Packages

When you package an IIS application into a standard package using the Package and Deployment Wizard, the wizard does not automatically include the .htm and .asp files for the project in the .cab file it creates. You must include these files manually while you are packaging the application.

#### <a name="WebClass22"></a>Webclasses: Unspecified Error

An "unspecified error" occurs if you add an existing webclass to a new project and then click on the template icon before the project has been saved. If you receive a prompt saying that "An unspecified error has occurred" in this context, save your project.

* * *

### DHTML Page Designer Issues

#### <a name="PageDSR1"></a>Page Designer Error Messages

There are currently no Help topics for the following error messages:

*   "Save internal HTML to File?"

This occurs if you specify a new, existing HTML file in the Property Page.

If you have HTML code in an instance of the Page Designer and you decide to use a different, existing HTML file with the designer, you must open the property page and specify the name of the existing file. The Page Designer must then know what to do with the internal HTML that is currently in the designer. If you answer Yes to this question, Visual Basic will save the HTML to a file after you specify a filename.

*   "DHTML Page Designers cannot be private"

This occurs if you attempt to change a DHTML Page Designer's **Public** property to **False**. While ActiveX designers can be either public or private, they must be Public to be accessible in a DHMTL application. Private designers can be useful in certain cases, such as in the Data Environment when you want to encapsulate your data access logic in the Data Environment, but not make it public.

*   "Reload changed HTML file?"

This occurs if the Page Designer notices that the HTML file referenced by the designer has changed on disk. This message confirms whether you want to load the modified HTML back into the designer.

#### <a name="PageDSR2"></a>Page Designer: "Me." Not Supported

You cannot use the "Me" reference in your page designer code to reference the DHTMLPage object. For example, the documentation frequently shows that you can write code such as "Me.Document._item_". This is not supported. Instead of using Me, you must use the "DHTMLPage" statement. For example, instead of Me.Document, you would use DHTMLPage.Document. If the designer is private, "Me" can be used, but private is not allowed in DHTML pages.

#### <a name="PageDSR3"></a>Page Designer: Cannot Access HTML Elements from Forms or External Objects

You cannot write code in a form, message box, or other object that references HTML elements on the page. Page elements are only accessible within the page itself.

#### <a name="PageDSR4"></a>Page Designer: Image SourceFiles are Resolved Incorrectly

Often, as you work in the page designer, you will want to reference an image that is located in the same directory as your current HTML page. Normally, when you reference such an image, you do not include a path in the image element's SRC attribute. The lack of a path tells the browser to pull the image from the page's current directory. For example, you might enter "myimage.gif" in the SRC property for an image of this type, rather than specifying "c:\mydhtml\myimage.gif." This is called a relative path, because it does not specify the full location of the image file.

At design time, the page designer does not correctly display these images with relative paths. When you enter such a property value in the SRC property or display a page that contains such a reference, two things will happen:

*   The image element will appear empty at design time.
*   The page designer appends the word "About:" to the value of your SRC property. For example, you might see "About:myimage.gif."

Despite these problems, the image appears correctly when you run the project. So you can ignore the "About:" keyword that appears in the SRC property and the fact that the image doesn't appear in the designer. When you save the HTML file, the image tag will be written correctly.

#### <a name="PageDSR5"></a>Page Designer: Binary Persistence Issue

Some controls attempt to save some or all of their properties in a binary format that cannot be directly represented in HTML. Because of this situation, some control properties may not be saved after you run your project. In order to correct this problem, set the property at run-time. The following list shows the known properties for which this problem occurs:

<table width="623" cellspacing="1" cellpadding="7" border="1">

<tbody>

<tr>

<td width="42%" valign="MIDDLE">

**<font size="2">Control</font>**

</td>

<td width="58%" valign="MIDDLE">

**<font size="2">Item</font>**

</td>

</tr>

<tr>

<td width="42%" valign="MIDDLE">

<font size="2">Tabbed Dialog Control</font>

</td>

<td width="58%" valign="MIDDLE">

<font size="2">TabsPerRow property does not persist</font>

</td>

</tr>

<tr>

<td width="42%" valign="MIDDLE">

<font size="2">Windowless Controls Listbox</font>

</td>

<td width="58%" valign="MIDDLE">

<font size="2">List items do not persist</font>

</td>

</tr>

<tr>

<td width="42%" valign="MIDDLE">

<font size="2">Common Controls</font>

</td>

<td width="58%" valign="MIDDLE">

<font size="2">Tabstrip settings do not persist</font>

</td>

</tr>

<tr>

<td width="42%" valign="MIDDLE">

<font size="2">Common Controls</font>

</td>

<td width="58%" valign="MIDDLE">

<font size="2">Toolbar settings do not persist</font>

</td>

</tr>

<tr>

<td width="42%" valign="MIDDLE">

<font size="2">Common Controls</font>

</td>

<td width="58%" valign="MIDDLE">

<font size="2">StatusBar settings do not persist</font>

</td>

</tr>

<tr>

<td width="42%" valign="MIDDLE">

<font size="2">ADODC Control</font>

</td>

<td width="58%" valign="MIDDLE">

<font size="2">ConnectionString property does not persist</font>

</td>

</tr>

</tbody>

</table>

In addition, for many controls, font properties do not persist. This occurrence is the most common manifestation of the binary persistence problem.

#### <a name="PageDSR6"></a>Page Designer: Type Library Problems are Preventing Some Help Topics from Appearing

When you begin a page designer project, you will see two DHTML type libraries in the Object Browser -- the **DHTMLPAGELIB** library, and a **DHTMLProject** library. Help is enabled for the DHTMLPAGELIB library, but not for the DHTMLProject library. Most elements in DHTMLProject should be available within the other library.

In other cases, type library problems are preventing certain pieces of the help to appear or function correctly. This happens in the following cases:

*   When you try to access help from the Properties window for an ActiveX control's properties.
*   When you try to access help from the Code Editor window, for the Document property.

In most cases, you should be able to locate the topic you want through the index in the MSDN Help viewer.

#### <a name="PageDSR7"></a>Page Designer: Modal Prompts Appear Behind Browser in Run Mode

If you add a message box to your DHTML page, the message box appears behind Internet Explorer when activated in run mode. In addition, you cannot move Internet Explorer out of the way in order to view the message, because the message is modal. When you hear the beep that indicates a message is on the screen, select the Visual Basic application from your taskbar to view and clear the message.

#### <a name="PageDSR8"></a>Page Designer: Do Not Watch Objects of Type HTMLDocument

Expanding a watch on objects of type HTMLDocument may cause problems within the IDE. Avoid watching objects of this type.

#### <a name="PageDSR9"></a>Page Designer: Cannot Design Frames Within A Page Designer

When you create web pages in the DHTML Page Designer, you cannot insert framesets within the page and fill their contents. You can, however, display the pages you create for your DHTML application within a frameset created outside of Visual Basic. If you want to design and debug frames, the process is as follows:

1.  In Visual Basic, design the contents of each frame with an individual DHTML Page Designer.
2.  In an external editing program, design the frameset document as a separate .htm file and save it in the temp directory. Set the SRC attributes for each frame to point to your content pages, using the names you defined in their BuildFile properties.
3.  In Visual Basic, choose "Start Browser with URL" on the Debugging tab of the Project Properties dialog box and enter the path of the frameset file.
4.  Enter run mode. Visual Basic launches Internet Explorer and loads the frameset page you designed in an external program. The frameset document should then load your page designer pages in the appropriate frames.
5.  Debug your application as usual.

#### <a name="PageDSR11"></a>Page Designer: Cannot use Visual Basic Code with the SetInterval Method

If you use the SetInterval method of the Internet Explorer BaseWindow object on a timer within your DHTML applications, you must set the first parameter of the method to point to a Javascript or VBScript time routine contained within the HTML page. This method cannot reference a Visual Basic routine within your page designer code.

#### <a name="PageDSR12"></a>Page Designer: Control Issues

The following are items to be aware of as you work with ActiveX Controls in your DHTML applications:

*   You must use the compiled version of any ActiveX controls you add to your DHTML pages.
*   You cannot use a DHTML page within an ActiveX Control project in this release. The reverse is also true -- that is, the ActiveX control project cannot be part of the project group in which you are working. The safest and most reliable way to build and debug a user control is to compile the user control project, then close that project, add the control to the page designer, and proceed from there. Interactive debugging and design between the two projects is not possible at this point.
*   When you embed an ActiveX control on a page in your DHTML application, not all of the appropriate type library information is copied with it. This can cause errors when you try to access a link in the object browser to the information for that control, and it can prevent you from accessing extended properties and methods for the control in the statement completion window that appears when you write code.
*   Some ActiveX controls will not work correctly with the page designer. You cannot use the following controls in a DHTML page: **MS Chart**, **Script Debugger**, **Hierarchical FlexGrid**, **SrcEdit OC**, **LayoutDTC**, **Tabular Data control**, **PageNavbarDTC**, or **IE Popup Window**. In addition, you cannot use private or uncompiled user controls on pages in your DHTML applications. In addition, it is possible that some controls you have obtained from third parties may not work correctly with the page designer.  
    Most controls that work with Internet Explorer 4 should work with the page designer. If you buy a third-party control that does not work with the page designer, test it within Internet Explorer and then contact the control vendor. For information on how to build ActiveX controls that work with the page designer and Internet Explorer, please see the article "Building ActiveX Controls for Internet Explorer 4.0," available in your MSDN library or on Microsoft's Web site in the MSDN Online section.
*   If you have an ActiveX control for which you have set the background transparent, it will appear opaque at design time when added to your DHTML application.
*   If your DHTML page contains a **Listview** control and you try to access the ColHeader.Width property, you will encounter an error saying that the object does not support this property. The property does appear on the Auto List Members drop-down list, but you should not use it in your code.
*   Some Visual Basic controls, such as the common dialog or the sysinfo control, are invisible at run time. If you add one of these objects to your HTML page, you cannot select it and move it around within the page after you initially draw it. You can, however, select the control in the treeview and either delete it or access its properties.
*   If you add a **File Upload** control (FlUpl) to your DHTML designer and you click it during design time, it will activate and display the standard file chooser dialog box. You can cancel this dialog box and move on.
*   There is a nonfunctional InnerText property available in the Auto List Members drop-down list for the horizontal rule element. Do not set a value for this property as it will produce no result.
*   When you add an ActiveX control to your page, the Height and Width properties do not update automatically as the control is resized in the designer. To refresh the numbers in the height and width properties, click in a blank area of the treeview after resizing your control.
*   When you click a control in the toolbox and then move the cursor to the Design pane of the designer, the cursor does not change from an I-beam to a pointer to indicate that you can click and draw your control. You can still click and drag your control to the desired shape, even though the cursor is still an I-beam.
*   When you copy controls between two DHTML projects in the same project group, the Components dialog box is not updated. You will need to set the necessary references to that control manually.
*   If you use the ImageList control, its picture will not be loaded if the project is Run in debug mode. However, the image will work correctly when you run the compiled DLL.

#### <a name="PageDSR13"></a>Page Designer: Miscellaneous HTML Issues

 The following are items to be aware of as you create and work with HTML in the DHTML page designer:

*   If your page contains a CHARSET metatag, this tag will be stripped out when the page is read into the designer. This should not cause a problem in most pages.
*   The column widths you see in tables within the designer may not match the column widths you see when you run your page in the browser. You may be able to avoid this problem either by setting your column widths in percentages rather than pixels, or by making sure the left-most column does not have a width measurement set for it.
*   Do not assign ID property values with greater than 117 characters. If you do, you may receive an error when running your application.
*   If the first paragraph element on your page does not contain text and is followed by another HTML element, you may see unexpected results if you delete the first <P> tag. In this case, other elements on the page may also be deleted.
*   You may receive inconsistent results when trying to delete a DIV tag using the designer's treeview panel. If you have trouble deleting a DIV tag, use the Launch Editor function and remove the <DIV> and </DIV> tags manually.
*   When you run your project, table borders will appear only for cells that contain text. You can force a table border to appear for empty cells in your tables by placing a <BR> tag in each unused table cell.
*   Property values set for the TITLE element on your HTML page may not persist in the property grid after you run the project. However, if you examine the HTML source code in the browser you will see that your title properties are still present in the HTML stream.
*   Elements on a page are always anchored by the upper left corner. Therefore, if you resize the element by stretching the top or left borders, you will see an element of the size you indicated when you release the cursor but it will still be positioned from the original top, left corner. This may produce unexpected results. You can avoid this by first positioning your control with the upper left corner in the appropriate place, then sizing it correctly.
*   You may have difficulty removing font formatting from some elements. For example, if you add an H1 heading to your page, press ENTER, and type a few sentences below it, the paragraph will receive the same formatting as the heading -- that is, it will appear large and in bold. You cannot remove this formatting using the bold icons on the formatting toolbar. You can remove this kind of incorrect formatting either by using the Launch Editor feature and manually correcting the formatting tags, or by changing the style of the selected text in the designer to Normal.
*   Some properties that appear in the Properties window for the Document object are not valid. These properties begin with the letters "on". If you set a value for any of these properties, you will see no results based on the value you set when you run the project or examine the HTML source.
*   For a check box HTML element, a property called "Indeterminate" is used to set the check box to a grayed-out state. However, if you set this property to True and then click the check box to move it to the blank state, the Indeterminate property does not reset to False. You can reset it manually or click a third time to reset the value.
*   For absolutely positioned controls with Height and Width properties that appear in the Properties window, you cannot resize the control by changing the values of those properties directly. Instead, resize the control by dragging one of its borders in the designer.
*   If you insert an input image element onto your page from the HTML toolbox, then use the Shape property to change it from another shape to a polygon, you will receive an error.

*   The treeview pane of the designer may not reflect the exact order of items as they appear in the design pane. The treeview shows the structural relationship of elements on the page in the order in which they appear in the HTML stream. The position of an element in the HTML stream does not always correspond to the position of an element on the displayed page, because you can change an element's position with attributes and inline styles. Your page will appear in Internet Explorer as it does in the design pane of the designer.

#### <a name="PageDSR14"></a>Page Designer: Miscellaneous Debugging Issues

The following are items to be aware of as you debug and test your DHTML applications:

*   In the Locals window, functions used in your DHTML application are listed twice. This may cause difficulty because it allows you to seemingly set two different values for Booleans.
*   Several properties listed in the Locals window will produce an error if you try to edit their value. These include the following properties:

> > DHTMLPage.BaseWindow.document.activeElement.all.length

> > DHTMLPage.BaseWindow.document.activeElement.offsetParent.all.length

> > DHTMLPage.BaseWindow.document.activeElement.children.length

*   If you refresh Internet Explorer while your DHTML project is in break mode, then you return to run mode in Visual Basic, your breakpoints will no longer be hit and stop instructions will no longer be followed. Stop debug mode, then press F5 again to begin your debug process a second time.

#### <a name="PageDSR15"></a>Page Designer: Syntax for Navigating Programmatically

In the "Navigating in DHTML Applications" topic in the Developing DHTML Applications section of the Building Internet Applications book, the documentation states that you can programmatically move between pages by using this syntax in your code:

    Private Function Button1_onclick() As Boolean
       BaseWindow.navigate "Project1.DHTMLPage2.html"
    End Function

This syntax is incorrect, because an underscore, rather than a period, should separate the project name and the page name. The correct syntax to use is shown below:

    Private Function Button1_onclick() As Boolean
       BaseWindow.navigate "Project1_DHTMLPage2.html"
    End Function

#### <a name="PageDSR17"></a>Page Designer: OBJECT Tag Insertion

In the topic "Testing your DHTML Projects," the documentation states that an OBJECT tag is inserted into your page during debugging but does not state that this tag is enclosed in a METADATA tag that is also inserted at this time. In addition, the following clarifications may be helpful:

*   The contents of the OBJECT tag's CODEBASE attribute are filled in during packaging, when you use the Package and Deployment Wizard to prepare your project for distribution.
*   During debugging, the OBJECT tag is actually inserted into a temporary copy of the HTML file. The final tag is set up during the packaging process.

#### <a name="PageDSR18"></a>Page Designer: Cannot CTRL+TAB Through Windows When Page Designer Has Focus

When the Page Designer window has focus, you cannot use the CTRL+TAB keystroke to move through all open windows in the Visual Basic IDE as you can with other projects.

#### <a name="PageDSR19"></a>Page Designer: Cannot Delete a Table if it Contains No Columns

If you delete all of the columns from a table on your page and then try to delete the table, it will still appear in the treeview pane of the designer. To delete the table from the treeview, you must delete the paragraph element under which the table element appears.

#### <a name="PageDSR20"></a>Page Designer: When Using Visual SourceSafe with Page Designer Projects, You Must Manually Check in the Project's .HTM Files

When you check a DHTML project into Visual Source Safe, the HTML pages associated with the project are not automatically checked into the Source Safe tree with the rest of the project files. You must manually add them to the tree as related files.

> **Note** This applies only when you have saved your HTML pages to external files. If you have saved them within the project there are no .htm files to check in for your design-time project.

#### <a name="PageDSR21"></a>Page Designer: Unqualified BuildFile Property Results in DLL Building To Desktop

If you edit the BuildFile property so that it does not specify the drive and directory in which the file should be placed, the system will build the resulting files directly to your desktop. Always enter the fully qualified path when entering values for this property.

#### <a name="PageDSR22"></a>Page Designer: Use of Load Event With Asynchronous Property

In the "Load Event" topic within the Reference section of the Visual Basic documentation in MSDN, the documentation incorrectly states the following:

> "Programmers can use this event when running asynchronously (when the **AsyncLoad** property is set to **True**) as a notification that all elements have been loaded onto the page."

In fact, this should say that programmers can use the Load event when running synchronously (when AsyncLoad is set to False) as a notification that all elements have been loaded.

#### <a name="PageDSR23"></a>Page Designer: Property Page Issues

 The following are items to be aware of as you work in the property pages for various HTML elements in the DHTML Page Designer:

*   Some elements have additional tabs in their property pages that do not appear when you view the property page in a DHTML project. In many cases, you can get to these tabs, even if they are not visible, by pressing the TAB key.
*   If you make changes to the custom property page for an ActiveX control you have added to your DHTML application, you may not always see those changes reflected in your DHTML project. To correct this problem, close your DHTML project and recompile the ActiveX control project. Your changes should then appear in the DHTML project. Repeat this process each time you make changes to the property page for the control.
*   The word "Test" will appear in the body of some property pages. This will not cause any problems in your application.

#### <a name="PageDSR24"></a>Page Designer: Location of Project Files in .CAB Deployment

In the topic "Deploying your DHTML Projects," the documentation states that the project DLL file, the Visual Basic run-time, and the .dsr and .dsx files for the project are all placed into the same .cab file during the deployment process. This is not correct. In fact, the .dsr and .dsx files for the project are not placed into the .cab.

#### <a name="PageDSR26"></a>Page Designer: Name of Save Option Changed

In the documentation for the page designer, the procedures for choosing save options tell you to select a Page Properties dialog box option called "Save HTML Within the Designer" if you do not want to save to an external HTML file. This option is actually named "Save HTML as part of the VB Project."

#### <a name="PageDSR27"></a>Page Designer: Problems Seeing Code Changes When Switching from Compiled to Debug Mode

You may, on occasion, encounter a problem where you do not see changes you have made to the code in your DHTML application when you have run the compiled DLL, made a change to the code, then entered debug mode. In this situation, you can try either of the following fixes:

*   Clear the Use Existing Browser check box on the Debugging tab of the Project Properties dialog before you start the debug process, after making changes to your code.
*   Exit the browser after running the compiled DLL, before opening the IDE to make code changes.

#### <a name="PageDSR28"></a>Page Designer: Help on most Language Elements is Available in the Platform SDK

Most language elements in the page designer are inherited from the Internet Explorer's document object model. F1 for these elements is not available. To get help on these topics, open the "Platform SDK" node in the MSDN table of contents, then open the "Internet/Intranet/Extranet Services" node and look for the Dynamic HTML section. You can also use the Index to do a search for the name of a particular language element.

* * *

### Extensibility Issues

#### <a name="Ext1"></a>"Command-Line Safe" Add-In Behavior

You can use the Load Behavior box in the Add-In Manager to control how and when an add-in loads in Visual Basic.

*   **Loaded/Unloaded** -- either loads or unloads a selected add-in when the box is checked or unchecked.
*   **Load On Startup** -- indicates whether the selected add-in should load when the Visual Basic IDE is started.
*   **Command Line** -- indicates whether an add-in should load when Visual Basic is started from a command line, either through a DOS prompt or a script.

When you select Command Line load behavior for an add-in, you may get the following warning message:

<dl>

<dd>"The selected add-in has not been confirmed to be 'command-line safe', and may require some user intervention (possible UI). Do you wish to proceed?"</dd>

</dl>

This occurs when you select an add-in for Command Line load behavior that was not declared by the author of the add-in to be "command line safe" when it was created. (This can be indicated with the Add-In Designer through a checkbox.)

"Command-line safe" means that the add-in is registered in a way to indicate that it contains no user interfaces that require user input when Visual Basic is invoked through a command-line. A user interface can interfere with the operation of unattended processes (such as build scripts).

If you don't indicate that an add-in is command-line safe (even if it _is_ command-line safe), when a user selects your add-in and then Command Line in the Load Behavior box, they'll receive the warning message. This isn't a serious problem, but merely a warning to the user that the selected add-in might possibly contain UI elements that can pop up unexpectedly and halt their automated scripts by pausing for user input.

#### <a name="Ext2"></a>Manually Setting Add-In Registry Values

You can also manually set the command-line safe flag (as well as the other values) for an add-in through the Windows registry.

**Note:** You should not attempt to directly manipulate any Windows registry entries unless you are familiar with doing so. Setting an invalid registry entry can cause problems with Windows, even preventing you from being able to load Windows.

In Visual Basic 6.0, the key that holds add-in information is located in HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Visual Basic\6.0\Addins\<add-in.name>. For Visual Basic 6.0, the LoadBehavior DWord values are:

*   None = 0
*   Startup = 1
*   Command Line = 4
*   Command Line / Startup = 5

There is also an additional DWord value that indicates whether the add-in is command-line safe: CommandLineSafe. A value of 1 indicates the add-in is command-line safe, while a value of 0 (the default) indicates that it is not command-line safe. A value of 0 is implied if you forget to check the command-line safe box in the Add-In Designer since the default value of 0 is assumed, and the add-in isn't considered command-line safe.

So, to demonstrate how to use these values to indicate that a ficticious add-in (My.Addin) is command-line safe and to have it load when Visual Basic is started by command-line, you would set the following registry values, using a tool such as RegEdit:

    HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Visual Basic\6.0\Addins\My.Addin
       "FriendlyName"="A friendly name for your add-in"
       "Description"="This value describes the add-in"
       "LoadBehavior"=dword:4
       "CommandLineSafe"=dword:1

#### <a name="Ext3"></a>Using the Add-In Designer

Visual Basic 6.0 includes a new tool, called the Add-In designer, to aid you in creating add-ins. To open it:

*   Create a new add-in project.
*   In the Project Explorer, under Designers, is a designer called Connect. Double-click it to activate the Add-In designer.

Unfortunately, context-sensitive help currently isn't available for the Add-In designer. Help topics are available, however. You can find the appropriate topics by searching for **Add-In Designer** in the MSDN index. You should see a list of three associated topics:

*   "Using the Add-In Designer"
*   "General Tab (Add-In Designer)"
*   "Advanced Tab (Add-In Designer)"

For more information, [search online](#Critical2), with **Search titles only** selected, for "Registering Add-Ins" in the MSDN Library Visual Studio 6.0 documentation.

#### <a name="Ext4"></a>Add-In Designer: More Information About Specifying Satellite DLL

When creating an add-in with the Add-In designer, you can specify a DLL on the Advanced tab. Be sure, however, to type only the name of the DLL file, and not its fully-qualified path. For example:

> MyAddinName.DLL

not:

> Addins\MyAddinName\MyAddinName.DLL

**

Localized Satellite DLLs

**

If you create a localized satellite DLL, you should also create a resources directory and a locale ID directory for the satellite DLL and install the DLL in the appropriate directory. The schematic for such a path is:

> <AddIn Directory>\Resources\<Locale ID>\<MySatellite.DLL>

For example, a satellite DLL for the German version (Locale ID = 1031) would go into the directory:

> C:\Program Files\MyAddin\Resources\1031\MyAddinName.DLL

* * *

### Miscellaneous Issues

#### <a name="Misc1"></a>Returning an Error Value from a DLL

To return an error value from a dynamic link library (DLL) procedure, the C language prototype must be coded so that the return value is an HRESULT. Refer to the Microsoft Press OLE 2 Programmer's Reference, Volume 2 for more information on how to do this.

#### <a name="Misc2"></a>Data Access Guide: DataFormats tutorial has wrong file extension

The topic named "Format Objects Tutorial" contains a wrong reference to a file with the extension .mdl. The actual file extension is .udl. For more information, [search online](#Critical2), with **Search titles only** selected, for "Format Objects Tutorial" in the MSDN Library Visual Studio 6.0 documentation.

The file in question is listed as "Northwind.mdl," but should be "Northwind.udl."

#### <a name="Misc3"></a>External Editor Field Added to Options Dialog Box

The Advanced tab of the Options dialog box has a new text box called **External HTML Editor**. This option allows you to select the HTML editing program that appears when you select **Launch Editor** from either the DHTML Page Designer or the Webclass Designer. You must enter the drive, path, and executable name of the program you want to use. You can choose an HTML editing program, a word processing program, or the text editor you prefer to use.

#### <a name="Misc4"></a>CodeBase Fixup Utility in Internet Component Download

The "Downloading ActiveX Components" section of the Building Internet Applications book makes reference to a utility called the CodeBase Fixup Utility that can be used to manually set codebase information in an ActiveX document. This information is incorrect. The utility is not shipped in the \Tools directory with Visual Basic, and you do not need to perform this process manually for Internet Explorer 4.0 because the Package and Deployment wizard automatically inserts the appropriate codebase information for these and other applicable projects.

#### <a name="Misc6"></a>Text in Project Properties/Open Dialogs Truncated in Japanese, Chinese, and Korean versions of Windows

When you run Visual Basic in the Japanese, Chinese, or Korean version of Windows, you may notice that text in the Project Properties or Open dialogs is truncated. If this occurs, shut down Windows, restart it, then restart Visual Basic and the problem will be fixed.

#### <a name="Misc7"></a>Avoid Using Repository Add-In with ActiveX Designers

You should avoid using the Repository add-in with projects that contain ActiveX designers.

For a complete list of available designers, on the Project menu in Visual Basic, click Components, then click the Designers tab in the Components dialog box.

* * *

### Microsoft Transaction Server (MTS) Issues

#### <a name="MTS1"></a>Building and Debugging MTS Components in Visual Basic 6.0

</font><font size="2" face="Verdana,Arial">

Visual Basic 6.0 supports the debugging of Microsoft Transaction Server (MTS) components, but there are several issues to keep in mind. The following issues apply only to MTS components running in the debugger.

</font><font face="Verdana,Arial">

##### Windows NT 4.0 SP4 Required

</font><font size="2" face="Verdana,Arial">

MTS debugging support requires Windows NT 4.0 Service Pack 4 (SP4) or later. MTS debugging is not supported under Windows 95 or Windows 98.

</font><font face="Verdana,Arial">

##### MTSTransactionMode Property

</font><font size="2" face="Verdana,Arial">

Visual Basic 6.0 introduces a new MTSTransactionMode property on classes that allows you to set the Microsoft Transaction Server (MTS) transaction support required for the class. The values for this property are equivalent to the property in the MTS explorer. However, the names of these properties in the Visual Basic IDE are not exactly the same as the names used in the MTS explorer. The mapping of names is as follows:

</font>

<table width="638" cellspacing="0" cellpadding="7" border="0">

<tbody>

<tr>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">**VB Property Value**</font>

</td>

<td width="50%" valign="TOP">

**<font size="2" face="Verdana,Arial">Option in MTS Explorer</font>**

</td>

</tr>

<tr>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">0 - NotAnMTSObject</font>

</td>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">N/A</font>

</td>

</tr>

<tr>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">1 - NoTransactions</font>

</td>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">Does not support transactions</font>

</td>

</tr>

<tr>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">2 - RequiresTransaction</font>

</td>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">Requires a transaction</font>

</td>

</tr>

<tr>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">3 - UsesTransaction</font>

</td>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">Supports transactions</font>

</td>

</tr>

<tr>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">4 - RequiresNewTransaction</font>

</td>

<td width="50%" valign="TOP">

<font size="2" face="Verdana,Arial">Requires a new transaction</font>

</td>

</tr>

</tbody>

</table>

<font size="2" face="Verdana,Arial">

The Transaction attributes of a class are imported into MTS only if the component is added to a Package with the Add File utility. If the component is brought into a package via the registered component list, the MTS attributes are not reflected in MTS Explorer.

</font><font size="2" face="Verdana,Arial">

##### Enabling MTS Debugging

To debug MTS components with Visual Basic 6.0, set the MTSTransactionMode property to a value other than 0-NotAnMTSObject. When you hit F5 to begin debugging, Visual Basic will now activate your component inside of the Microsoft Transaction Server runtime.

##### Single Client, Server, and Thread

</font><font size="2" face="Verdana,Arial">

Debugging is supported only for a single client and a single MTS server component at a time operating on a single thread. For situations requiring multiple clients or MTS servers or multiple threads, you should debug the Visual Basic component in the Visual C++ development environment. For details on debugging Visual Basic components in the Visual C++ environment, see the Visual C++ documentation.

</font><font face="Verdana,Arial">

##### Build Requirements for Debugging

</font>

<font size="2" face="Verdana,Arial">To build and debug an MTS component in Visual Basic, you must build your component into a DLL and</font><font size="2" face="Verdana,Arial">set binary compatibility on the project. If you do not set binary compatibility and you add interfaces to, or remove them from, the component, these changes may not be detected by MTS.</font>

<font face="Verdana,Arial">

##### Debugging Limitations on Class_Initialize and Class_Terminate Events

</font><font size="2" face="Verdana,Arial">

You should not put code in the Class_Initialize and Class_Terminate events of an MTS component that attempts to access the object or its corresponding context object. The Visual Basic run-time environment calls Class_Initialize before the object and its context are activated, so any operations that Class_Initialize attempts to perform on the object or its object context will fail. Similarly, the object and its context are deactivated before Class_Terminate is called, so operations that this method attempts on the object and its context will also fail.

You should not set a breakpoint in the Class_Terminate event of an MTS component. When the debugger reaches the breakpoint, it will attempt to activate the object, an attempt that will fail and cause Visual Basic to stop.

</font><font face="Verdana,Arial">

##### Watching MTS Objects

</font><font size="2" face="Verdana,Arial">

During debugging, do not watch object variables returned from the MTS runtime, including return values from SafeRef, GetObjectContext, CreateInstance, and other functions that return objects wrapped by MTS.

To simulate the runtime environment more effectively, the Microsoft Transaction Server runtime pauses operation each time that Visual Basic breaks in the debugger. Internally, Visual Basic makes method calls on objects that are being watched in the debugger. Since the MTS runtime is paused as you look at watch variables, the calls that Visual Basic makes to these objects may fail.

If you do add MTS-wrapped objects to the watch window or watch via other means, it may cause an inconsistent state to be detected by MTS, and the process will be terminated.

</font><font face="Verdana,Arial">

##### Registration and Debugging

</font><font size="2" face="Verdana,Arial">

The debugging facilities in Visual Basic allow an MTS component to be debugged without being installed in the MTS explorer. When you start debugging, Visual Basic will automatically call into MTS to run your component in the MTS runtime.

Depending on your debugging requirements, you may also want to install your component into the MTS explorer. There are a few issues to keep in mind when doing this. If you make changes in the Visual Basic IDE in your component’s interfaces, class names, project names, transactional support or other settings, there may be mismatches between the configuration data in the MTS explorer and the actual configuration running in the Visual Basic debugger. It is also possible that the component could be launched in MTS while you are debugging it. Further, if you export a package while you are debugging a component in the package, MTS will treat the Visual Basic development environment as the component.

These problems can be avoided by making sure the component to be debugged is not registered in the MTS explorer. As noted later in this section, if you change the configuration of an installed component in the debugger, you may have to remove and reinstall the component.

</font><font face="Verdana,Arial">

##### Component Changes Made During Debugging

</font><font size="2" face="Verdana,Arial">

In Visual Basic, you can modify transactional attributes on an MTS component during debugging. Visual Basic does not register these changes in the MTS explorer.

If during debugging you make a source code change that requires Visual Basic to generate a new CLSID or ProgID or that changes the transactional attribute of any MTS class, you must use MTS Explorer to delete and reinstall the package containing the class. If you have set binary compatibility for the component, you will be warned that changes have occurred.

</font><font face="Verdana,Arial">

##### Starting Debugger While a Component is Running in MTS

</font><font size="2" face="Verdana,Arial">

If you are running a component outside the debugger and then decide to begin debugging, an instance of the component may still be running in MTS when you start it in the debugger. MTS will detect this condition and attempt to silently shut down the instance it controls. To avoid this problem, remove the component from the MTS explorer before you begin debugging.

</font><font face="Verdana,Arial">

##### Debugging Unregistered MTS Components

</font><font size="2" face="Verdana,Arial">

An MTS component can run in the Visual Basic debugger without having been registered in the MTS catalog. In this case, the component will not be visible in the MTS Explorer. It is preferable to debug components that are not registered, as it avoids a number of problems discussed elsewhere in this section.

</font><font face="Verdana,Arial">

##### Deployment and Debugging

</font><font size="2" face="Verdana,Arial">

To properly deploy an MTS component, you need to build the component as a DLL, be sure the component is not running in any debug session, and then run the Package and Deployment wizard. The component can be open in Visual Basic, but it cannot be active in a debug session.

</font><font face="Verdana,Arial">

##### MTS Components in Debugger Run As If in Library Package

</font><font size="2" face="Verdana,Arial">

The MTS run-time environment treats Visual Basic components being debugged as if they belong to a library package, even if the components are registered with MTS as belonging to a server package. Library packages do not support component tracking, role checking, or process isolation.

Because MTS components being debugged behave as if they are in a library package, you cannot do security debugging in the Visual Basic development environment. Remote activation of the debugged component will use the security attributes of Visual Basic. Remote activation of a component running in the MTS run-time environment (mtx.exe), however, will use the security attributes set up for the particular package in the MTS explorer. To debug security issues, you should use the Visual C++ development environment.

</font><font face="Verdana,Arial">

##### Component Failure Causes Visual Basic to Stop Running

</font>

<font size="2" face="Verdana,Arial">An MTS component being debugged runs in the same process as the Visual Basic development environment, so a component failure will also cause Visual Basic to stop running. Also, the MTS runtime environment automatically shuts down the runtime process when it detects an inconsistent state internally. In these cases, MTS will display a dialog box explaining the situation, the Visual Basic window will disappear, and an event will be recorded in the Windows NT system log. Check the Windows NT Event Viewer as well as other topics in this document for possible explanations of the problem.</font>

<font face="Verdana,Arial">

##### No Support for Transacted Web Classes

</font><font size="2" face="Verdana,Arial">

Transacted Visual Basic Web Classes are not supported in Visual Basic 6.0.

</font><font face="Verdana,Arial">

##### RunWithoutContext Registry Key Is Ignored

</font><font size="2" face="Verdana,Arial">

Visual Basic 6.0 ignores the RunWithoutContext registry key. This key is no longer needed with Visual Basic 6.0’s integrated debugging of MTS objects, as the functionality provided by the context object is now available during debugging.

</font><font face="Verdana,Arial">

##### Using IObjectControl

</font><font size="2" face="Verdana,Arial">

If you need to execute code during startup and shutdown of your MTS object, you should implement the IObjectControl interface (from the Microsoft Transaction Server Type Library) and use the Activate and Deactivate functions. These functions are called by the MTS runtime during startup and shutdown of your object. Using the IObjectControl functions is preferable to using Class_Initialize and Class_Terminate due to the limitations described below.

You can place code that accesses the object context in the Activate and Deactivate functions. However, due to the way that the MTS runtime activates objects, you should not put breakpoints on IObjectControl::Deactivate or IObjectControl::CanBePooled.

</font><font face="Verdana,Arial">

##### Debugger May Reactivate Objects Released by MTS

</font><font size="2" face="Verdana,Arial">

Visual Basic 6.0 may reactivate MTS objects while you are debugging single-step through a client. Because of the way that Visual Basic 6.0 discovers information about objects, this is expected behavior. For example, consider the following code:

</font>

<pre><font size="2">

> Dim x as object
> Set x = CreateObject("MyApp.Class")
> x.Test
> Set x = Nothing

</font></pre>

<font size="2" face="Verdana,Arial">

If the x.Test method calls SetComplete, MTS immediately frees x from memory, but x has not yet been set to Nothing. When x.Test returns, the Visual Basic debugger calls QueryInterface on x for the IProvideClassInfo interface. The context wrapper associated with x creates a new instance of MyApp.Class to service the QueryInterface call. As a result, you will see this uninitialized object in the debugger after x.Test has returned. This object appears only in the debugger and is removed by the subsequent instruction to set x to Nothing.

* * *

### Dictionary Object

#### <a name="Dict1"></a>Introducing the Dictionary Object

Visual Basic 6.0 includes a new Dictionary object that can be used instead of a Collection object or an Array. A Dictionary object is the equivalent of an associative array as used in PERL and other languages. The Dictionary object provides additional properties and methods that aren't available for Collections or Arrays.

The Dictionary object is contained in the Microsoft Scripting Runtime (Scrrun.dll) that ships with Visual Basic 6.0\. You must add a reference to Scrrun.dll in order to use the Dictionary object in your project.

To learn more about the Dictionary object, [search online](#Critical2), with **Search titles only** selected, for "Dictionary Object" in the MSDN Library Visual Studio 6.0 documentation.

* * *

### <a name="rmmscVisualComponentManager"></a>Visual Component Manager

#### <a name="rmmscKnownProblemsinVisualComponentManager"></a>Known Problems in Visual Component Manager

##### <a name="RelatedFilesTab"></a>"Related Files Tab (Component Properties Dialog Box)" Topic Incorrect

Visual Component Manager User Interface Reference: The topic "Related Files Tab (Component Properties Dialog Box)" incorrectly states that the tab is used to display and enter files that are related to the selected component. In fact, none of the information displayed on this tab can be modified. You can add related files to a component only when publishing or re-publishing the component. For more information, search online, with **Search titles only** selected, for "Publishing Components" in the MSDN Library Visual Studio 6.0 documentation.

##### <a name="RemovingRepository10RegistryKeys"></a>Removing Repository 1.0 Registry Keys

If you installed VCM 5.0 (previously available for web download) you will have the following Windows Registry keys setup. They were necessary for VCM 5.0 and the 1.0 version of the Repository. If you find the following Registry entries then it safe to remove them and may, in fact, improve VCM 6.0 performance.

*   HKEY_LOCAL_MACHINE\Software\Microsoft\Repository\CacheMaxAnnProps
*   HKEY_LOCAL_MACHINE\Software\Microsoft\Repository\CacheMaxObjects
*   HKEY_LOCAL_MACHINE\Software\Microsoft\Repository\CacheRelshipMaxCollections
*   HKEY_LOCAL_MACHINE\Software\Microsoft\Repository\CacheRelshipMaxRows
*   HKEY_LOCAL_MACHINE\Software\Microsoft\Repository\MaxRowCacheAge

##### <a name="AddingrepositorytablestoanexistingMDBfile"></a>Adding Repository Tables to an Existing .mdb File

If you try to open an existing .mdb file from within VCM that is not a repository database (i.e., it does not contain the repository structure/tables), you will be asked if you want the repository tables added to the database. You should not do this for normal use; the repository should generally be in a separate database. This will work, but it can take as long as 10 minutes to create the repository structure in an existing .mdb file.

To create a brand new .mdb file containing the repository structure, right-click in the folder outline, click Repository, click New, and then enter the name of the file you want to create.

* * *

### <a name="rmmscApplicationPerformanceExplorer"></a>Application Performance Explorer

#### <a name="rmmscKnownProblemsinApplicationPerformanceExplorer"></a>Known Problems in Application Performance Explorer

##### <a name="ConfiguringRemoteAutomationSecurityWhenUsingRemoteAPEComponents"></a>Configuring Remote Automation Security When Using Remote APE Components

In order to use Remote Automation (RA) to communicate with remote APE components, you may have to configure RA security using the Remote Automation Connection Manager (Racmgr32.exe).

**To configure RA security**

1.  Start Racmgr32.exe and click the Client Access tab.
2.  Select either "Allow All Remote Creates" or "Allow Remote Creates by Key".
3.  If "Allow Remote Creates by Key" is selected, make sure the "Allow Remote Activation" check box is checked for each APE component.

RA supports the following levels of authentication:

<table>

<tbody>

<tr>

<td width="22%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">**Name**</font>

</td>

<td width="10%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">**Value**</font>

</td>

<td width="68%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">**Description**</font>

</td>

</tr>

<tr>

<td width="22%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Default</font>

</td>

<td width="10%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">0</font>

</td>

<td width="68%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Use Network default.</font>

</td>

</tr>

<tr>

<td width="22%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">None</font>

</td>

<td width="10%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">1</font>

</td>

<td width="68%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">No authentication.</font>

</td>

</tr>

<tr>

<td width="22%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Connect</font>

</td>

<td width="10%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">2</font>

</td>

<td width="68%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Connection to the server is authenticated.</font>

</td>

</tr>

<tr>

<td width="22%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Call</font>

</td>

<td width="10%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">3</font>

</td>

<td width="68%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Authenticates only at the beginning of each remote procedure call, when the server receives the request. Does not apply to connection-based protocol sequences (those that start with the prefix "ncacn").</font>

</td>

</tr>

<tr>

<td width="22%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Packet</font>

</td>

<td width="10%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">4</font>

</td>

<td width="68%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Verifies that all data received is from the expected client.</font>

</td>

</tr>

<tr>

<td width="22%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Packet Integrity</font>

</td>

<td width="10%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">5</font>

</td>

<td width="68%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Verifies that none of the data transferred between client and server has been modified.</font>

</td>

</tr>

<tr>

<td width="22%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Packet Privacy</font>

</td>

<td width="10%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">6</font>

</td>

<td width="68%" valign="top">

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Verifies all other levels and encrypts the argument values of each remote procedure call.</font>

</td>

</tr>

</tbody>

</table>

<font link="#0000FF" vlink="#660066" size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">

APE profiles are initially installed with an authentication level of 1 ("None") because Windows 95 supports only that level of authentication. However, if additional security is desired, the level of authentication of a profile can be changed by modifying the profile collection file (the Aemanagr.ini file) by using a text editor such as Notepad.

Each profile in the profile collection file begins with the name of the profile within square brackets, such as [Peak performance, synchronous (CPU, Pool)]. The attributes of the profile follow, using the format <name>=<value> (such as "Task Duration=1"). To change the authentication level, change the value of the "Authentication" attribute of the selected profile and save the file.

##### <a name="compatibilityissuesbetweenvs6apeandvb5ape">Compatibility Issues Between the Application Performance Explorer (APE) that Ships with Visual Studio 6.0 and the Version that Shipped with Visual Basic 5.0</a>

There are known compatibility issues between the Application Performance Explorer (APE) that ships with Visual Studio 6.0 and the APE that shipped with Visual Basic 5.0.

**To avoid the compatibility issues, do one of the following:**

*   Before installing Visual Studio 6.0 and APE on the computer that has the version of APE shipped with VB 5, first uninstall APE from VB, and then install Visual Studio and APE.
*   If you have installed Visual Studio 6.0 and APE on the same computer that has the VB5 APE, uninstall the VB APE and then reinstall the Visual Studio APE.

##### <a name="AdjustingdefaultsettingstouseAPEandMTS">Adjusting Default Settings To Use APE and MTS</a>

After installing the APETEST database onto your SQL Server, you must adjust some of the default settings in order to use APE and MTS.

> **Note**   If you haven't already installed the APETEST database on your SQL Server, you should do that first. To learn how to install the APETEST database, search for the topic "APE Database Setup Wizard" in _MSDN Library Visual Studio 6.0_.

**To configure the APETEST database installation to work with MTS**

1.  Start Microsoft SQL Enterprise Manager.
2.  In the Databases folder, right-click the APETEST database and click Edit.
3.  Click the Options tab.
4.  Select the Truncated Log on Checkpoint check box and click OK.
5.  In the Databases folder, right-click the tempdb database and click Edit.
6.  Click Expand.
7.  In the Data Device box, select <new>.
8.  In the New Database Device dialog box, in the Name box, type tempdbData.
9.  In the Size (MB) box, type 10.
10.  Click Create Now, and finally click OK.
11.  Click Expand Now.
12.  Click Expand.
13.  In the Log Device box, select <new>.
14.  In the New Database Device dialog box, in the Name box, type tempdbLog.
15.  In the Size (MB) box, type 10.
16.  Click Create Now, and finally click OK.

**To configure the allowable number of user connections**

1.  Start Microsoft SQL Enterprise Manager.
2.  Right-click the server and click Configure.

For example, if your server is named CORONA, in the Server Manager child window, right-click CORONA and then click Configure.

4.  Click the Configuration tab.
5.  In the Configuration box, increase the number of user connections by at least 15.

> **Note**   If you are running APETEST on an established production database server, you may not have access permission to adjust the current number of user connections. In this case, you should ask your database administrator to increase the number of current user connections by at least 15 connections to support APE testing.

##### <a name="ApplicationPerformanceExplorerServerSideSetup">Application Performance Explorer Server-Side Setup May Generate Error</a>

While installing the APE server-side components, you may see an error referring to an incorrect version of OLEAUT32.dll. You may dismiss this error and continue with the installation.

However, this error message may indicate that the Microsoft Transaction Server Package was not installed correctly. To confirm that it was installed correctly, run the Transaction Server Explorer and look for all installed MTS packages on your computer. Visual Studio APE Package should be listed.

To install the package, AEMTSSVC.pkg, run the MTS Transaction Server Explorer from the Start menu and install the package to the local computer using the MTS Explorer.
