Print 'Before'
Go
Select * From TSYSPROF 
Where PROGRAM_NAME = "EMMA2" and BUILD_VERSION = 'E2.CHI.G4.131.000.00' 
Go
Delete From TSYSPROF  Where PROGRAM_NAME = "EMMA2" and BUILD_VERSION =  'E2.CHI.G4.131.000.00' 
Go

Insert Into TSYSPROF 
Select
	PROGRAM_NAME		= "EMMA2",
	BUILD_VERSION  		= "E2.CHI.G4.131.000.00",
	REF_NO			= "Build",
	REQ_FORM		= "PCR",
	DESCRIPTION		= "VB Upgrade To Windows Server 2016 & Userall group add authority",
	REMARKS			= "ITSR013622 , ITSR013528",
	USERID_CD		= SUSER_NAME(),
	TIMESTAMP		= getdate()
Go

Print 'After'
Go
Select * From TSYSPROF Where PROGRAM_NAME = "EMMA2" and BUILD_VERSION = 'E2.CHI.G4.131.000.00' 
Go