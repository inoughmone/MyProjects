  Read Me First:

  SQL-SOMBRERO/OCX (1.0.9) for DB-LIBRARY(4.2) 32 Bit


  The SQL-Sombrero/OCX (DB-Library) 32 Bit package consists of the following:

	1 - diskette  3.5

  NOTE THAT THIS PACKAGE CONTAINS AN ONLINE HELP FILE EQUIVALENT TO THE MANUAL

  SQL-SOMBRERO/OCX (DB-Library) INSTALLATION

  
  -  Microsoft* Windows NT* release 3.1 or later.

  -  80386 (20MHz) or higher processor.

  -  4 megabytes of RAM.

  -  2 megabytes of free disk space for software and environment.

  -  Microsoft or SYBASE* SQL Server* Open Client* software 
  (runtime DB-Library and Net-Library for Windows NT and compatible network DLL).

  The required Microsoft DB-Library DLL is NTWDBLIB.DLL.  SQL-Sombrero/OCX will
  work with either the 4.2 or 6.0 version of this DLL.

  The required Sybase DB-Library DLL is LIBSYBDB.DLL.  SQL-Sombrero/OCX will
  work with the NT version of this DLL.  Please note that Sybase recommends that
  you always work with the latest version of their Open Client Software.

  ******************************************************************************
  ******************************************************************************

       The SQL-Sombrero/OCX exposes the DB-Library 4.2 API 

  ******************************************************************************
  ******************************************************************************

  The SQL-Sombrero/OCX has one function which you must be used prior to executing
  any other SQL-Sombrero/OCX function.  This is the OCXInit function.  This function
  tells SQL-Sombrero/OCX whether you are using the Microsoft or Sybase 32 bit DB-Library
  DLL.  The syntax of the function is:

	object.OCXInit(DLLType%)

	If you set DLLType% to 1 then you will use the Microsoft 32 bit DB-Library
	DLL - NTWDBLIB.DLL.

	If you set DLLType% to 2 then you will use the Sybase 32 bit DB-Library
	DLL - LIBSYBDB.DLL.

  You must perform this function or your program will not execute correctly.

  The following files are required to run OLE2 applications. 
  These files should be installed by any application which is 
  an OLE2 Automation controller such as Microsoft* Excel 
  5.0.

  Files You Need

  To expose or access OLE Automation objects, you need the 
  following files, which are provided with OLE enabled applications.
 
   OLEAUT32.DLL
   Accesses type libraries.

   OLEAUT32.DLL
   Provides functions for creating OLE Automation 
   objects and retrieving active objects at run time. 
   Accesses OLE Automation objects by invoking 
   methods and properties.

   OLE32.DLL
   Provides OLE functions that may be used by OLE 
   objects or containers. Supports component object creation and access.
   Supports access to subfiles, such as type libraries, within compound documents.


   OLEPRX32.DLL
   Coordinates object access across processes.

  ------------------------------------------
  Run INSTALL.EXE on the installation diskette (disk #1) from within Windows.
  Everything is installed in the directories indicated below. This installation
  will runs as a Windows NT 3.1/3.5 application or as a Win95 application.

  The installation will allow you to install any of the following portions of
  the SQL-Sombrero/OCX (DB-Library) product.

  1.	SQL-Sombrero/OCX (32Bit DB-Library)
  2.	SQL-Sombrero/OCX Help
  3.	SQL-Sombrero/OCX Samples


  The installation procedure will allow you to view the "readme" file (this file) prior
  to performing the installation.

  The installation will then ask for your name and company name.

  You are then presented with a dialog which will allow you to select the components
  which you wish to install.  The default is a full installation of all components.  If 
  you wish to select the components to install click the "Custom Install" check box.  You
  may then see the list of components by clicking the "+" sign beside the "Custom Install".
  You may also click the "+" sign beside the "Samples" to see the available samples.
  Simply click the check box beside any component to select it or deselect it.  When you
  are finished click "OK" to proceed.

  You will then be asked to select a directory which to install the SQL-Sombrero/OCX product.
  into. The default installation directory is C:\SQL\SQLOCXDB.  You may
  change this directory by typing in a new directory.  If you have already have or plan to
  install the 16Bit version of SQL-Sombrero/OCX for DB-Library you should use the same
  installation directory.  The context sensitive help in SQL-Sombrero/OCX depends on
  the "Help" file being in a particular directory under the default directory.  To 
  share the "Help" file between 16 and 32 bit versions of SQL-Sombrero/OCX the suggested
  directory structure must be followed.

  Click "OK" to proceeed with the installation.  At this point all the selected components
  will be installed.  

  Next the installation will ask you if you wish to register each of the SQL-Sombrero/OCX's
  which were selected for installation into the registration database.  This step is
  necessary in order to use the SQL-Sombrero/OCX.  Note that the required DB-Library DLL's 
  for each version of the SQL-Sombrero/OCX must be present in your path in order for the 
  registration to succeed.

  The installation is now complete.


  If you select the SQL-Sombrero/OCX files for 32 bit applications the OCX files 
  will be installed along with other DLL's which are required for their use.  
  The appropriate development licence file is also installed.

  Prior to using the OCXs you must register the OCX into the system registration
  database.  This procedure will be done for you if you select the SQL-Sombrero/OCX 32 Bit
  component. 

  If you chose not to register the SQL-Sombrero/OCX files at installation time the
  instructions for installing the product are found in the Online help file.  It is 
  recommended that you install the Online help.  If you do you will be able to get context
  sensitive help for each SQL-Sombrero/OCX function from a product such as Visual Basic 4.0.

  *****************************************************************************
  *****  Changes for 1.0.9
  *****
  *****
  *****************************************************************************

  May 10,1996
  
  The values of the ErrorMsgCount and MsgCount properties of the Session object
  were not correct in version prior to 1.0.9.
  
  The SqlColLen method of the Connection object was returning improper values.
  
  The SqlServerEnum method of the Session object was not returning anything.
  
  *****************************************************************************
  *****  Changes for 1.0.8
  *****
  *****
  *****************************************************************************

  Feb 29,1996

  The function SqlRetData would always return an empty string no matter what the ret paramter
  was.  This was fixed and all properties on the OCX have been hidden from the VB property
  window.

  *****************************************************************************
  *****  Changes for 1.0.6
  *****
  *****
  *****************************************************************************

  Dec 27,1995

  The Text and Image functions which retrieve the Text/Image pointer and the Text/Image timestamp
  have been fixed so that when a NULL pointer is retrieved an empty string "" will be returned.
  The functions changed are:

	SqlTxTimeStamp
	SqlTxTsNewVal
	SqlTsNewVal
	


  *****************************************************************************
  *****  Changes for 1.0.6
  *****
  *****
  *****************************************************************************

  Oct 2,1995
  The following changes have been made for version 1.0.6                  *****

  The 32 bit OCX now functions for both 32 Microsoft and 32 bit Sybase NT. You must use the
  OCXInit method in order to load the correct 32 bit DB-Library DLL.

  The syntax of the OCXInit method is:

	object.OCXInit(DLLType%)

	If you set DLLType% to 2 then you will use the Sybase 32 bit DB-Library
	DLL - LIBSYBDB.DLL.

  A new property was added to the Session object and to the Connection object.  The DLLType
  property can be used to indicate whether to use the Microsoft or Sybase 32 bit DB-Library
  DLL.  You must have loaded the DLL by using the OCXInit function prior to setting this
  property.  When this property is set in the Session object all connections which are 
  opened from the Session object will inherit the DLLType property.  The DLLType property
  is read-only in the Connection object.

  The SqlWinExit function when used with the Sybase DB-Library DLL has no effect.
  The function was left in the OCX for compatablity with current source code.

  The SqlOpenConnection was not freeing the Login record that it created to make a connection.
  If you connect and disconnect in a program you would eventually run out of DB-Library
  resouces and receive a DB-Library error 10000 when you tried to connect.

  The SqlExit function will now unload the DB-Library DLL that was loaded with the OCXInit
  method.  If you use the SqlExit function and wish to execute more SQL-Sombrero/OCX functions
  you must reload the DB-Library DLL by using the OCXInit method.

  *****************************************************************************
  *****  Changes for 1.0.5
  *****
  *****
  *****************************************************************************

  Sept 13,1995
  The following changes have been made for version 1.0.5                  *****

  SqlLogin  now returns a long value rather than an integer if the function was
  successful.  A failure still returns 0.

  SqlOpen   now expects a long value for the login record.  This the long value
  returned from the SqlLogin method as described above.

  SqlFreeLogin now expects a long value for the login record.  This the long value
  returned from the SqlLogin method as described above.


  ********************************************************************************
  ***                                                                          ***
  ***   Note that the help file has been updated.  The help file will always   ***
  ***   contain the latest documentation on any of the SQL-Sombrero/OCX        ***
  ***   methods.  Please refer to the Help file before calling Tech support    ***
  ***   with a problem.                                                        ***
  ***                                                                          ***
  ********************************************************************************




  The following examples have been provided.

  Access examples:
  ----------------
     1.  FROMSVR.TXT is code which will copy a SQL Server table to a corresponding
                     Access 2.0+ table.

     2.	 TOSVR.TXT   is code which will copy an Access 2.0+ table to a
		     corresponding SQL Server table.

  Excel example:
  --------------
     1.  TESTOBJ.XLS is an Excel 5.0+ work sheet containing a code sheet with an
                     example of a SQL-Sombrero/OCX program which will populate
                     a spread sheet with data obtain from a SQL Server using an
                     ad hoc query.

  Visual Basic 4.0 examples:
  --------------------------

     1.  SAMPLEAPP   is an example of a program which maintains a database table. It
                     shows how to login to the database, handle error messages, send
                     SQL commands to the server and retrieve data from the server.

     2.  BCPAPP      is an example of a program which will use the SQL-Sombrero/OCX
                     routines to either bulk copy a file into a database table or to
                     bulk copy a table out to a disk file.

     3.  STORPROC    is an example of a program which will create a stored procedure
                     and then execute the stored procedure while providing the stored
                     procedure with required parameters.

     4.	 TEXTIMAG    is an example of a program using the Text and Image routines of
                     SQL-Sombrero/OCX.  This application will save large files in a 
                     SQL database and will allow the user to retrieve those files from
                     the database and recreate the original file.

     5.  SAMPLEAPP32 is an example using the Win95 ListView control showing how to 
                     issue SQL queries and then retrieve the results and populate
                     a ListView control.
  The installation will install the following files on your system depending on which
  components are installed. The term root is used to denote the destination directory
  which was chosen for the installation of the SQL-Sombrero/OCX product.

  Always installed:
  -----------------

	ReadMe.Txt	- this file			- root
	Install.Log	- list of files installed	- root

  SQL-Sombrero/OCX (32Bit)
  ------------------------

	OC30.DLL	- OCX runtime DLL		- Windows\System32
	REGSVR32.EXE	- OCX registration program	- root\ocxdb32
	SQLOCXDB.LIC	- OCX development licence	- root\ocxdb32
	SQLOCD32.OCX	- OCX 32 Bit 			- root\ocxdb32
	MFCANS32.DLL	- OCX Runtime DLL		- Windows\System32
	MSVCRT20.DLL	- OCX Runtime DLL		- Windows\System32


   SQL-Sombrero/OCX Help
   ---------------------

	SQLOCXDB.HLP	- OCX Help files		- root\help

   SQL-Sombrero/OCX Samples
   ------------------------

	Excel Example					- root\samples\excelapp
	   testobj.xls


	Access Example					- root\samples\access
	   FROMSVR.TXT
	   TOSVR.TXT

	Bulk Copy Example				- root\samples\bcpapp
	   bcpwin.mak
	   main.frm
	   main.frx
	   login.frm
	   about.frm
	   about.frx
	   bcpgloba.bas
	   helpform.frm
	   helpform.frx

	Table Maintenance Example			- root\samples\samplapp
	   gbass.bas
	   global.bas
	   logon.frm
	   mainform.frm
	   sample10.mak
	   helpform.frm
	   helpform.frx

	Execute Stored Procedure Example		- root\samples\storproc
	   storedpr.frm
	   storproc.mak
	   global.bas
	   helpform.frm
	   helpform.frx

	Image and Text Example				- root\samples\textimag
	   gbass.bas
	   logon.frm
	   mainform.frm
	   textimag.mak
	   helpform.frm
	   helpform.frx
	   compress.frm
	   compress.frx

	Win95 Query Form Example			- root\samples\sample95
	   gbass.bas
	   global.bas
	   logon.frm
	   mainform.frm
	   sample10.mak
	   helpform.frm
	   helpform.frx

  If you are using a product with an Object Browser you can paste function definitions
  and Constants from the SQL-Sombrero/OCX for DB-Library.  Please refer to the
  documentation for the container product for instructions.  The SQL-Sombrero/OCX
  product is also help context sensitive.  If the container product you are using
  supports context sensitive help then you may jump to the online help directly from
  the container.  This will only work if the Help files are left in the default
  directory configuration.

  ----------------------------------------------------------------------------
  Your input on any subject is welcomed and appreciated!
  ----------------------------------------------------------------------------

  You can get support:
     Internet:   support@sfi-software.com
     Compuserve: 71162,1050
     Phone:      819 778-5045
     Fax:        819 778-7943

  For updates:
  FTP: http://ftp.sfi-software.com
  WWW: http://www.sfi-software.com

  For information on other SFI Products you can reach us by fax at 819 778-7943
  and by voice at 819 778-5045 or via internet at info@sfi-software.com
  You can use our special Compuserve account for SFI products at 71162,1050. 
  You can access our FAX ON DEMAND system by calling 819 778-5045 and selecting
  the option "1" when entering the phone system.  Request document index 50 for
  the latest list of available documents.
  ----------------------------------------------------------------------------
  
  Happy SQL-Programming with SQL-Sombrero!

  Thank you!
  SFI

  ----------------------------------------------------------------------------

                   Copyright 1994-1995 Sylvain Faust Inc.
SQL-Programmer, SQL-Sombrero and CompressIT are Trademarks of Sylvain Faust Inc.
      Sylvain Faust Inc. claims copyright in this program and documentation.
   Claim of copyright does not imply waiver of Sylvain Faust Inc. other rights.

