Registration Utilities
----------------------

This directory contains three tools for registering in process ole servers.
In process servers are OLE DLLs or OLE controls.

REGSVR32.EXE
------------

RegServer is a windows program that allows you to register and un-register in
process servers.  REGSRVR32.EXE will display dialogs indicating if it was
successful unless you use the silent option /s.  To register a server use the
following format: REGSVR32.EXE MyServer.dll  To un-register a server use the /u 
option in the following format: REGSVR32.EXE /u MyServer.dll.

REGOCX32.EXE
------------

RegOCX is a windows program specifically designed for use by setup programs
when installing OCXes.  REGOCX32.EXE does not display dialogs.  To register an
OCX you use the following format: REGOCX32.EXE MyCtrl.ocx

REGIT.EXE
---------

RegIt is a command line utility that you can use to register one or more in
process servers.  Regit accepts wildcards.  For instance you can use
REGIT.EXE *.OCX to register all of the OCX files in a directory.  

Installation
------------

To install these utilities copy the files to a directory on you hard drive.
You may want to put these utilities to a directory in your PATH if you use
them often enough.  It is also useful to associate the .DLL and .OCX file
extensions with REGSVR32.EXE so that you can double click on DLLs and OCXes
to register them.
