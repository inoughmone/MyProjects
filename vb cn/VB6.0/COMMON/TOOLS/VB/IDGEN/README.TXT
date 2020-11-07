IDGEN
-----

Both GUIDGEN and UUIDGEN are tools for generating globally unique identifiers
known as GUIDs.  GUIDS are commonly used in OLE to identify classes (CLSID) or
interfaces (IID.)  These utilities are included with VB for developers who want
to generate IDL (Interface Description Language) or ODL (Object Description
Language.)  IDL and ODL are used with the MIDL.EXE and MKTYPLIB.EXE tools to
generate type libraries that can be used with VB.  

GUIDGEN.EXE
-----------

GUIDGEN is a windows program that generates GUIDs in several different formats.
GUIDGEN places GUIDs in the clipboard so that you can paste them where you need
to use them.

UUIDGEN.EXE
-----------

UUIDGEN is a command line utility that also generates GUIDs in different formats.
You use UUIDGEN as follows:

UUIDGEN [-isonvh?]

 i - Output UUID in an IDL interface template
 s - Output UUID as an initialized C struct
 o<filename> - redirect output to a file, specified immediately after o
 n<number> - Number of UUIDs to generate, specified immediately after n
 v - display version information about uuidgen
 h,? - Display command option summary