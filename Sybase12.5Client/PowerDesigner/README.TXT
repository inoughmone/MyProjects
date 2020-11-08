Release Notes for PowerDesigner[TM] Version 9.5.2 (Wintel)

(c) 1991-2003 Sybase, Inc. and its subsidiaries. All rights reserved.
Updated 2003-04-11

Thank you for choosing PowerDesigner as your design tool!  
Please read this document to learn about last minute updates to the product and documentation.
 
This version is certified for the following operating systems:
Windows 98
Windows NT 4.00
Windows Me
Windows 2000
Windows XP


Installation
------------
Installation provides an install for Microsoft Data Access Component (MDAC) that may prompt to reboot your computer to be completed.  If you have launched PowerDesigner installation from a network drive, this drive must be reconnected automatically at logon for the MDAC install to complete properly.
Some problems may occur with MDAC installation on a computer having VirusScan version 4.5 from McAfee running. In that case, deactivate the anti-virus to allow MDAC installation.

Under Windows 2000, if you are using an NTFS drive, the user that will install PowerDesigner must have write access to the Windows temporary directory (\WinNt\temp). 
This user can belong either to the Administrator group or to the PowerUser group.

Under Windows NT4, 2000 and XP, non Administrator users may receive a warning message during installation. Setup requires write access on certain directory and on the registry that a non Administrator users may not have. Setup can go on despite that warning message.

Under Windows NT 4.0, at the end of the Acrobat Reader installation, Acrobat setup may ask to reboot your computer. 
You should answer No to the question "Do you want to reboot".

At the end of the Microsoft Windows script 5.5 installation, Microsoft setup may ask you to reboot your computer. 
You should answer No to the question "Do you want to reboot".


PowerDesigner 8 or previous installed
-------------------------------------
PowerDesigner 9.5.2 cannot be installed over version 8 and previous.

PowerDesigner 9.0 or 9.5 installed
----------------------------------
PowerDesigner 9.5.2 is a minor release and therefore it is supposed to be installed in the same directory as PowerDesigner 9.0. or 9.5.0. PowerDesigner 9.0, 9.5.x and 9.5.2 share the same registry and named path information.

Database and Language support
-----------------------------
* Non Certified Database Definitions
Some database definition files are delivered with "Not Certified" in their names.
Sybase will perform all possible validation checks, however Sybase does not maintain specific environments to fully certify these databases.  Sybase will support the definition by accepting bug reports and will provide fixes as per standard policy, with the exception that there will be no final environmental validation of the fix.  Users are invited to assist Sybase by testing fixes of the definition provided by Sybase and report any continuing inconsistencies.

* Non Certified Object Languages
Some object language are delivered with "Not certified" in their names.
Sybase will perform all possible validation checks, however Sybase does not maintain specific environments to fully certify these object languages.  Sybase will support the object language by accepting bug reports and will provide fixes as per standard policy, with the exception that there will be no final environmental validation of the fix.  Users are invited to assist Sybase by testing fixes of the object language provided by Sybase and report any continuing inconsistencies.

* List of modified Database Definitions between 9.5.1 and 9.5.2
all Sybase IQ family
all Sybase ASA family
all Sybase ASE family
all Oracle family
all Microsoft SQL Server family
all Microsoft Access family
all DB2 family
all Informix family
Allbase

* In order to have access to the most current OOM or PDM generation features, you must use the DBMS and Object languages delivered with version 9.5.2.

For DBMS and Object Language Definitions copied in models from previous PowerDesigner version make sure you change target language or DBMS when you open your old models including the DBMS/Object Language definitions copied within the models.

The Java language file delivered with PowerDesigner 8 is no longer supported in version 9.x. You must change your target language to use the new language file delivered with version 9.x

* Customized XDB
If you have customized your XDB, taking advantage of the mechanism of user-defined queries during reverse engineering, you must look at the corresponding section in Advanced User Documentation: chapter DBMS Reference Guide, section "Extension mechanism for ODBC reverse engineering queries" and modify your XDB as described.

* Optimization for Reverse engineering
PowerDesigner 9.5.2 has a new mechanism to retrieve data during reverse engineering.
It has been optimized for the major databases.
However for Microsoft SQLServer 2000, it is recommended you change a parameter in the  registry for best results.
[HKCU\Software\Sybase\PowerDesigner 9\FolderOptions\Physical Objects]
“RevOdbcMinPerct"=20

Changes in PowerDesigner
------------------------
* EAS support
PowerDesigner is delivered with an extended model definition that provides support for EAS 4.1.3 and above.
EAS 4.1.2 is no longer supported.

* Language report
Some new report items are not translated in the German report language.
All report titles have been modified to adopt a more generic formula in the English report language. The English report language delivered in 9.5.1 and earlier will be delivered under the name 'English (deprecated)". If you wish to revert to previous title naming, select the "English (deprecated)" language in the report property sheet. This report language will not be delivered in the next PowerDesigner version.

* Symbol behavior change
Priority is given to the symbols on top: when symbols overlap, it is no longer possible to select and resize the symbol in the background.  Although the grabber handles are visible, when you click to perform the resize, focus immediately goes to the object in the front.

* Start and Ends
Starts and ends can now have graphical synonyms in the Activity diagram and the Business process model.

* Decisions
Decisions can now have several input links and several output links in the Activity diagram and the Business process model.

* Auto-adjust to text
If Auto-adjust to text is selected, the symbol size is recalculated based on the symbol content.
If Auto-adjust to text is deselected, the symbol size is never recalculated.
In PowerDesigner 9.5.2, when you manually resize a symbol, the check box auto-adjust to text is deselected and symbol size will not be modified in case of object renaming or content modification.
If you wish to re-adjust the symbol to text, you have to use the auto-adjust to text function (Ctrl-J). The auto-adjust to text check box will then be reselected.

* Package symbol
Package symbols can be transparent. It can be useful when you wish to use InBox UML representation. You have to deselect the Fill Color check box in the Fill page of the Format dialog box.


* PDM preview
PDM models now have a Preview page that displays what will be generated for the entire model during database generation.

* User and Group
Password box is not encrypted anymore.

* Sybase ASA, Sybase ASIQ trigger and procedure
In previous versions of PowerDesigner, triggers and procedures generation scripts used to be generated with a '/' at the end.
This is no longer the case. PowerDesigner generate a ';' at the end of procedures and triggers which is closer to common usage and easier to execute for standard interpreters.

The scripts generated by previous versions of PowerDesigner will not be reversed properly unless you change the command terminator definition in the XDB file from ';' to '/' or define the "/" terminator in the option tab (Command zone) of the reverse engineering window.

ASE reverse-engineering
-----------------------
When reverse-engineering an Sybase ASE database you may have the following message.
<<The database has been partially reverse engineered: 1 error(s) occurred>>

with the following ODBC error on Sybase ASE 11.x
<<Unable to reverse engineer table attributes.
[SYBASE][ODBC Sybase driver][SQL Server]The column prefix 'p' does not match with a table name or alias name used 
in the query. Either the table is not specified in the FROM clause or it has a correlation name which must be used
instead.
SQLSTATE = S1000>>

with the following ODBC error on Sybase ASE 12.x
<<Unable to reverse engineer table attributes.
[SYBASE][ODBC Sybase driver][SQL Server]SELECT permission denied on object syspartitions, database sdp, owner dbo
SQLSTATE = S1000>>

To reverse engineer an Sybase ASE database without the above mentioned message, you need to grant select permission to public
group on some system views.
Using isql, connect as "sa" using appropriate password then run these statements - replace <working database>
with the database name you actually use.

use <working_database>
go
grant select on syspartitions to public
go
grant select on sysroles to public
go
use master
go
grant select on syssrvroles to public
go
grant select on sysloginroles to public
go


This problem has been identified and will be corrected in a coming release of Sybase ASE.

Known issues
------------
* Menu Execute commands
Previous PowerDesigner installations have created entries in the registry to insert lines in the "Execute commands" menu.
The 9.5.2 installation will change these commands.
However, only the user who installs PowerDesigner can see these changes correctly.
Others users may see some residual commands (under Tools>Execute Commands>ebXML>New model). They have to suppress them manually, in the Customize command list.


* Repository on Oracle 
DO NOT USE the driver named "PB DATADIRECT OEM 4.10 32-BITS Oracle WIRE PROTOCOL" delivered with PowerDesigner 9.5.2.
If you do, you will not be able to consolidate any models in the repository.
You must use the driver named "PB DATADIRECT OEM 4.10 32-BITS Oracle"

* Repository on DB2
DO NOT USE the driver named "PB DATADIRECT OEM 4.10 32-BITS DB2 WIRE PROTOCOL" delivered with PowerDesigner 9.5.2.
If you do, you will not be able to freeze models in the repository.
You must use the driver provided by IBM.

* Help menu - videos
The menu "Help > Where to start > Videos" requests that you have PowerDesigner CD in CD drive.
However it will not work properly if you are using Netscape 7 or Mozilla.

* HTML report
By default
- HTML generated from Chinese Windows is only viewable on English Windows using NetScape 6.2 or Mozilla 0.9.9
- HTML generated from Hebrew Windows is not viewable using NetScape 4.6 and the characters are inverted with Opera 6.01
- HTML generated from Russian Windows is viewable on all browsers (Mozilla 0.9.9, IE 5.5, IE 6, Opera, Netscape 6.2 and NetScape 4.6)

In order to generate a valid code page that will let you see all valid characters, you must create the value ReportHtmlCharset in [HKEY_CURRENT_USER\Software\Sybase\PowerDesigner 9\General], and give it the value of the code page that should be generated in HTML report.

Some samples of HTML charset can be found at http://www.w3.org/International/O-charset-lang.html

* PowerBuilder support
- When generating into, or reversing from PowerBuilder libraries, only targets/applications already opened by PowerBuilder are listed.
- Dynamic PowerBuilder libraries (PBD) are not supported: PBD objects cannot be reversed!

* Samples for OLE Automation
Executable delivered in previous version of PowerDesigner will not be updated by 9.5.2 setup and will need to be recompiled to function properly.

Windows XP support
------------------
PowerDesigner has been tested under Windows XP.
New XP look and feel can be obtained by adding a manifest file to PowerDesigner installation.
This file is located in the Support directory on the PowerDesigner CD. You must copy it in PowerDesigner main directory.
Nevertheless, you cannot use the manifest file on a network installation and you may have trouble with ODBC data source administrator.


Printable documents
-------------------
All documentation is delivered in PDF format. If you choose to install PDFs during setup, the Help menu will give you a direct access to these documents. They can be viewed or printed using the Acrobat Reader.
These documents are stored in the PowerDesigner\Printable Docs folder.

Tools
-----
The Tools folder in the PowerDesigner directory contains some utilities to help repository installation, database generation and reverse. A text file (tools.txt) describes its content.

VB Script
---------
The VBScript folder in the PowerDesigner directory contains some Visual Basic scripts that can be used as samples or tools. A text file (VBSreadme.txt) describes its content.

OLE Automation
--------------
The OLE Automation folder in the PowerDesigner directory contains some samples using OLE automation to access PowerDesigner objects.

ODBC
----
PowerDesigner is delivered with a set of ODBC drivers that can only be used with PowerDesigner, which appear under the "PB DATADIRECT OEM 4.10 32-BITS" name in the list of drivers, if you select to install them during PowerDesigner setup.


Repository
----------
PowerDesigner 9.5.2 is delivered with an authenticated Adaptive Server Anywhere server 8.01, which can be used to serve as PowerDesigner repository database.
Please refer to PowerDesigner version 9.5.2 Installation Guide or to online help (search for Authenticated keyword) for installation procedure.
The Adaptive Server Anywhere engine delivered with PowerDesigner cannot be used for any purpose other than a PowerDesigner repository database.

Set the time
------------
PowerDesigner uses modification time comparison in various functions. Be sure that all computers installed with PowerDesigner are set with accurate time.

Sybase Web Site
---------------

Find the latest information on PowerDesigner at Sybase's award winning Portal http://www.sybase.com - go into the 'Products' section and look for 'Enterprise Modeling'. 

You will access news, white papers, latest technical documents as well as useful downloads. You will also be able to download an evaluation version of all modules of PowerDesigner - Evaluation version fully usable for 15 days. 

In addition, you can register to Sybase Portal and select PowerDesigner in your profile - you will then have easier access and even automatically receive all the latest news and technical information on PowerDesigner.

Finally, access the 'Support' section and receive even more technical information on PowerDesigner!

The Help menu of PowerDesigner also gives you access to various interesting pages on the Sybase sites : PowerDesigner page, Support, online books...
