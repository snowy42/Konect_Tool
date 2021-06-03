/* ===========================================================================================================================================================================================================================================
==============================================================================================================================================================================================================================================
==============================================================================================================================================================================================================================================
==============================================================================================================================================================================================================================================
==============================================================================================================================================================================================================================================
==============================================================================================================================================================================================================================================
==============================================================================================================================================================================================================================================
==============================================================================================================================================================================================================================================
*/ ;==========================================================================================================================================================================================================================================


#SingleInstance force
SetBatchLines -1
konectDir:="C:\Program Files\Global GBM\Konect\Konect Manager\Konect.Manager.Shell.exe"
SetWorkingDir %A_ScriptDir%
thisUser:= A_Username
plsWait:=false
uploading:= 0
IniRead, importerDir, cfg\kon-ie.cfg, IMPORT, TOOLS
IniRead, exporterDir, cfg\kon-ie.cfg, EXPORT, TOOLS
FileRead, infotxt, cfg\info.cfg
if GetKeyState("Shift") ;manually override the current username (directory name, such as msnow or lgriggs)
{
	inputbox, overrideUser, , ,hide, 100,100 ;hidden input requesting override username
	if (overrideUser = "") ; checks if input is nothing as sets to current user
		overrideUser := A_Username
	IniRead, currentUser, cfg\users.cfg, % overrideUser, username ;reads the overridden username
	IniRead, activeStatus, cfg\users.cfg, % overrideUser, active ;gets active status
	IniRead, userRoles, cfg\users.cfg, % overrideUser, roles ;gets active status
	thisUser:=overrideUser
}
else
{
	IniRead, currentUser, cfg\users.cfg, % A_Username, username ;reads the data for current user
	IniRead, activeStatus, cfg\users.cfg, % A_Username, active
	IniRead, userRoles, cfg\users.cfg, % A_Username, roles	;gets active status
}

if (InStr(userRoles, "Urban") && InStr(userRoles, "Reserves") && InStr(userRoles, "Rural") && InStr(userRoles, "Buildings"))
	userRoles := 0 ;master
if (InStr(userRoles, "Urban") || InStr(userRoles, "Rural"))
	userRoles := 1 ;Roads
if (InStr(userRoles, "Reserves") || InStr(userRoles, "Buildings"))
	userRoles := 2 ;Tania
	
if (currentUser = "ERROR") ;if the user doesn't exist in config file, then decline access and exit the application
{
	msgbox You are not registered to use this system.  Please contact administration.
	ExitApp
}

if (activeStatus = "No") ;if user is inactive
{
	msgbox Your user accounts has been deactivated.  Please contact administration.
	ExitApp
}

Menu, FileMenu, Add, Start Konect, runKonect ;add menu items to main ui
Menu, FileMenu, Add, Update List, getData
Menu, FileMenu, Add, Reload Application, RL
Menu, FileMenu, Add, Close, closeApp
Menu, EditMenu, Add, Add/Modify User, addUser
Menu, EditMenu, Add, Change Supervisor Email, changeEmail
Menu, EditMenu, Add, Set Import Tool .exe Location, impLoc
Menu, EditMenu, Add, Set Export Tool .exe Location, expLoc
Menu, EditMenu, Add, Set My Email, myEmail
Menu, HelpMenu, Add, Send Test Email, sendTest
Menu, HelpMenu, Add, About, aboutSys
Menu, MyMenuBar, Add, &File, :FileMenu
Menu, MyMenuBar, Add, &Edit, :EditMenu
Menu, MyMenuBar, Add, &Help, :HelpMenu

guiWidth:=1115
listWidth:=guiWidth-115
buttonX:=listWidth + 10

Gui, -Resize ;allow maximising
Gui, Menu, MyMenuBar ; add the above menu items to gui
Gui, Add, Text, , %infotxt%
Gui, Add, ListView, w%listWidth% x5 y40 r20 gListWKO -LV0x10 -Multi, Priority|Data Template|Description|Asset No|Asset Description|Issued By|Assigned To|ID
Gui, Add, Button, x%buttonX% y40 w100 gCopyDesc, Copy Description
Gui, Add, Button, x%buttonX% y65 w100 gAssignWK, Assign WKO#
Gui, Add, Picture, x%buttonX% y300 w100 h100, icon\sfm_logo.png
Gui, Add, Progress,vLoadProg x%buttonX% y411 x1015 w90 h12
Gui, Add, Text, vLoadText x%buttonX% y410 x900 w100 right
Gui, Default
LV_ModifyCol(1, 50)
LV_ModifyCol(2, 150)
LV_ModifyCol(3, 250)
LV_ModifyCol(4, 70)
LV_ModifyCol(5, 150)
LV_ModifyCol(6, 150)
LV_ModifyCol(7, 150)
LV_ModifyCol(8,0)
Gui, Show, w%guiWidth%, % "Konect Tool - " . currentUser ;show and set the title to include the current user

GoSub GetData
SetTimer, GetData, 1000

; **************** GUI 2 - Changing Supervisor Email ****************

Gui, 2:Add, ListView, w350 gListEmail -LV0x10 -Multi, Role|Email Address ; creates the ui for changing emails for supervisors
Gui, 2:Add, Button, g2Close, Close
Gui, 2:Default
IniRead, em, cfg\emails.cfg, emails, Rural
LV_Add("", "Rural",em)
IniRead, em, cfg\emails.cfg, emails, Urban
LV_Add("", "Urban",em)
IniRead, em, cfg\emails.cfg, emails, Reserves
LV_Add("", "Reserves",em)
LV_ModifyCol(1, 150)
LV_ModifyCol(2, 195)

; **************** GUI 3 - Users ****************

Gui, 3:Add, ListView, w650 gListUsers -LV0x10 -Multi r10, Username|Full Name|Email|Roles|Active
Gui, 3:Add, Button, w75 x665 y5 gNewUser, New User
Gui, 3:Add, Button, w75 x665 y30 gEditUser, Edit User
Gui, 3:Add, Button, w75 x665 y55 gDeleteUser, Delete User
Gui, 3:Default
LV_ModifyCol(1, 100)
LV_ModifyCol(2, 150)
LV_ModifyCol(3, 150)
LV_ModifyCol(4, 200)
LV_ModifyCol(5, 45)
FileRead, FullINI, cfg\users.cfg
Loop, Parse, FullINI,`r,`n
{
	If (RegExMatch(A_LoopField,"O)\[\K.+?(?=])")!=0)
		addUN:= SubStr(A_LoopField,2, StrLen(A_LoopField)-2)
	If InStr(A_LoopField, "username=")
		addFN:= SubStr(A_LoopField,10,StrLen(A_LoopField))
	If InStr(A_LoopField, "active=")
		addAC:= SubStr(A_LoopField,8,StrLen(A_LoopField))
	If InStr(A_LoopField, "email=")
		addEM:= SubStr(A_LoopField,7,StrLen(A_LoopField))
	If InStr(A_LoopField, "roles=")
	{
		addRO:= SubStr(A_LoopField,7,StrLen(A_LoopField))
		LV_Add("", addUn, addFN, addEM, addRO, addAC)
	}
}

; **************** GUI 4 - Edit User ****************

Gui, 4:Add, Text, x5 y5, Username:
Gui, 4:Add, Edit, x100 y5 vSetUN
Gui, 4:Add, Text, x5 y30, Full Name:
Gui, 4:Add, Edit, x100 y30 vSetFN
Gui, 4:Add, Text, x5 y55, Email:
Gui, 4:Add, Edit, x100 y55 vSetEM
Gui, 4:Add, Text, x5 y80, Roles:
Gui, 4:Add, CheckBox, vUrb x100 y80, Urban
Gui, 4:Add, CheckBox, vRur x100 y100, Rural
Gui, 4:Add, CheckBox, vRes x100 y120, Reserves
Gui, 4:Add, CheckBox, vBui x100 y140, Buildings
Gui, 4:Add, Text, x5 y165, Active:
Gui, 4:Add, CheckBox, vAct x100 y165
Gui, 4:Add, Button, Default w50 x130 y190 gSaveUser, Save
Gui, 4:Add, Button, w50 x180 y190 g4Cancel, Cancel
Gui, 4:+toolwindow

return



CopyDesc:
	Gui, Submit, NoHide
	Gui, Default
	If !(SelectedRow := LV_GetNext()) {
		Msgbox, 0, Error, Select an item from the list.
		return
	}
	LV_GetText(grabDesc, SelectedRow, 3)
	clipboard:=grabDesc
	Msgbox Description copied to clipboard!
return

AssignWK:
	If !(SelectedRow := LV_GetNext()) {
		Msgbox, 0, Error, Select an item from the list.
		return
	}
	InputBox, newWKO, Assign Work Order Number, Please enter the work order number as per Asset Master...,,300,150
	if errorlevel
		return
	Gui, Submit, NoHide
	Gui, Default
	If !(SelectedRow := LV_GetNext()) {
		Msgbox, 0, Error, Select an item from the list.
		return
	}
	GuiControl,Text, LoadText, Doing Stuff...
	SetTimer, ProgressBar, 2000
	LV_GetText(grabID, SelectedRow, 8)
	FileAppend, WorkOrderNumber`,_id`n%newWKO%`,%grabID%, exp\upload.csv
	LV_Delete(SelectedRow)
	Uploading:=1
	RunWait, `"%importerDir%`" -pin=834902 -action=update_matching -dataset=`"_Form_Works Order Process`" -file=`"exp\upload.csv`",, Hide
	Uploading:=0
	FileDelete, exp\upload.csv
	Sleep 1200
	GuiControl,Text, LoadText,
	GuiControl,,LoadProg,0
	SetTimer, ProgressBar, Off
return

ListWKO:
	if (A_GuiEvent = "DoubleClick")
		GoSub CopyDesc 
return

ListUsers:
return

4Cancel:
	Gui, 4:Hide
return

impLoc:
if !plsWait {
	if (thisUser = "msnow")
	{
		FileSelectFile, importerDir, 1,,Please Select Import Tool EXE File, Executable Files (*.exe)
		if ((importerDir != "") || (!ErrorLevel))
			IniWrite, %importerDir%, cfg\kon-ie.cfg, IMPORT, TOOLS
	}
	else
		msgbox Please log in as msnow to change this variable.
}
else
	msgbox Please wait for "Doing Stuff" to finish
return

expLoc:
if !plsWait {
	if (thisUser = "msnow")
	{
		FileSelectFile, exporterDir, 1,,Please Select Export Tool EXE File, Executable Files (*.exe)
		if ((exporterDir != "") || (!ErrorLevel))
			IniWrite, %exporterDir%, cfg\kon-ie.cfg, EXPORT, TOOLS
	}
	else
		msgbox Please log in as msnow to change this variable.
}
else
	msgbox Please wait for "Doing Stuff" to finish
return

SaveUser:
	GuiControlGet, getUN, 4:, SetUN
	GuiControlGet, getFN, 4:, SetFN
	GuiControlGet, getEM, 4:, SetEM
	GuiControlGet, getUrb, 4:, Urb
	GuiControlGet, getRur, 4:, Rur
	GuiControlGet, getRes, 4:, Res
	GuiControlGet, getBui, 4:, Bui
	GuiControlGet, getAct, 4:, Act
	getRoles:=""
	if (getUrb == 1)
		getRoles:=getRoles . "Urban | "
	if (getRur == 1)
		getRoles:=getRoles . "Rural | "
	if (getRes == 1)
		getRoles:=getRoles . "Reserves | "
	if (getBui == 1)
		getRoles:=getRoles . "Buildings | "
	if (getAct == 1)
		getAct := "Yes"
	else
		getAct := "No"
	IniWrite, % getFN, cfg\users.cfg, % getUN, username
	IniWrite, % getEM, cfg\users.cfg, % getUN, email
	IniWrite, % getRoles, cfg\users.cfg, % getUN, roles
	IniWrite, % getAct, cfg\users.cfg, % getUN, active
	Gui, 3:Default
	if (Action = "EDIT")
		LV_Modify(SelectedRow, "", getUN, getFN, getEM, getRoles, getAct)
	if (Action = "NEW")
		LV_Add("", getUN, getFN, getEM, getRoles, getAct)
	MsgBox, Details Saved!
	Gui, 4:Submit
return

NewUser:
	Action:="NEW"
	GuiControl, 4:, SetUN, 
	GuiControl, 4:, SetFN,
	GuiControl, 4:, SetEM, @kingborough.tas.gov.au
	GuiControl, 4:, Urb, 0
	GuiControl, 4:, Rur, 0
	GuiControl, 4:, Res, 0
	GuiControl, 4:, Bui, 0
	GuiControl, 4:, Act, 0
	Gui, 4:Show,, User Details
return

EditUser:
	Action:="EDIT"
	Gui, 4:Submit
	If !(SelectedRow := LV_GetNext()) {
		Msgbox, 0, Error, Select an item from the list.
		return
	}
	Gui, 3:Default
	LV_GetText(addUN, SelectedRow, 1)
	LV_GetText(addFN, SelectedRow, 2)
	LV_GetText(addEM, SelectedRow, 3)
	LV_GetText(addRO, SelectedRow, 4)
	LV_GetText(addAC, SelectedRow, 5)
	GuiControl, 4:, SetUN, %addUN%
	GuiControl, 4:, SetFN, %addFN%
	GuiControl, 4:, SetEM, %addEM%
	
	GuiControl, 4:, Urb, 0
	GuiControl, 4:, Rur, 0
	GuiControl, 4:, Res, 0
	GuiControl, 4:, Bui, 0
	GuiControl, 4:, Act, 0
	
	IfInString, addRO, urban
		GuiControl, 4:, Urb, 1
	IfInString, addRO, rural
		GuiControl, 4:, Rur, 1
	IfInString, addRO, reserves
		GuiControl, 4:, Res, 1
	IfInString, addRO, buildings
		GuiControl, 4:, Bui, 1
	IfInString, addAC, Yes
		GuiControl, 4:, Act, 1
	
	Gui, 4:Submit
	Gui, 4:Show,, User Details
return

DeleteUser:
	If !(SelectedRow := LV_GetNext()) {
		Msgbox, 0, Error, Select an item from the list.
		return
	}
	MsgBox, 4, Confirm, Are you sure you want to delete this user?
	IfMsgBox Yes
	{
		Gui, 3:Default
		LV_GetText(addUN, SelectedRow, 1)
		LV_Delete(SelectedRow)
		IniDelete, cfg\users.cfg, % addUN
	}
		
return

ListEmail: ;handles setting/changing email address for supervisor roles
	Gui, 2:Default
	if (A_GuiEvent = "DoubleClick")
	{
		LV_GetText(superV, A_EventInfo, 1)
		LV_GetText(emV, A_EventInfo, 2)
		InputBox, newEm, New Email, Please enter a new email address for the %superV% Supervisor role:,,250,175,,,,,%emV%
		If !ErrorLevel
		{
			IniWrite, % newEm, cfg\emails.cfg, emails, % superV
			msgBox Email for %superV% updated!
			LV_Modify(A_EventInfo,,superV,newEm)
		}
	}
return

changeEmail: ;menu item to open the Gui for changing supervisor emails
if !plsWait {
	Gui, 2:Show,,Double click on a role to change the email address
}
else
	msgbox Please wait for "Doing Stuff" to finish
return

2Close: ;close gui 2
	Gui, 2:Submit
return

sendTest: ;simple test for sending emails
if !plsWait {
	IniRead, myEmailAdd, cfg\users.cfg, % thisUser, email
	MailItem := ComObjCreate("Outlook.Application").CreateItem(0)
	MailItem.Recipients.Add(myEmailAdd)
	MailItem.Subject := "New Work Order!"
	MailItem.body := "You have a new work order to approve in Konect!"
	MailItem.send
	Msgbox Email Sent!
}
else
	msgbox Please wait for "Doing Stuff" to finish
return

runKonect: ;open Konect 
if !plsWait {
	If WinExist("Konect Manager")
		WinActivate
	Else
		if FileExist(KonectDir)
			run, % konectDir
		else
			msgbox % "Konect is not installed on this system."
}
else
	msgbox Please wait for "Doing Stuff" to finish
return

addUser: ;Shows the add/edit user interface
if !plsWait {
	Gui, 3:Show
	Gui, 3:Default
}
else
	msgbox Please wait for "Doing Stuff" to finish
return

myEmail: ;simply changes current users email address
if !plsWait {
	InputBox, myEmailAdd, My Email, Please enter the email address associated with this account (%thisUser%),,300,150
	If !ErrorLevel
	{
		IniWrite, %myEmailAdd%, cfg\users.cfg, % thisUser, email
		MsgBox Success! Email address for %thisUser% saved:`n`n%myEmailAdd%
	}
}
else
	msgbox Please wait for "Doing Stuff" to finish
return

aboutSys: ;information about the program
if !plsWait {
	Msgbox, 262176,About Konect Tool, Tool created by Matthew Snow for Kingborough Council.`n`nThis tool is designed to aid in the development and communication between Konect and Asset Master
}
else
	msgbox Please wait for "Doing Stuff" to finish
return


getData:
if !plsWait {
	plsWait:=true
	Loop Files, exp\*.csv
		FileDelete, %A_LoopFileFullPath%
	Sleep 1000
	Loop Files, exp\*.csv
	{
		Msgbox One of the export files for the Konect Tool is open.  Please close it to prevent further errors.
		return
	}
	RunWait, `"%exporterDir%`" -pin=834902 -dataset="All_Roads_WKO" -format=csv -photos=false -file=`"%A_WorkingDir%\exp\export.csv`",,Hide		
	
	IniRead, urbem, cfg\emails.cfg, emails, Urban
	IniRead, rurem, cfg\emails.cfg, emails, Rural
	
	if FileExist("exp\export.csv")
	{
		SetTimer, ProgressBar, 2000
		; Gui, Default
		; LV_Delete()
		cNWKONo:=0				;== Column Name: WorkOrderNumber
		cNPriority:=0			;== Column Name: DTCase
		cNReqDate:=0			;== Column Name: RequestDate
		cNPersonInspecting:=0	;== Column Name: PersonInspecting
		cNAssetNo:=0			;== Column Name: thisAsset
		cNAssetDesc:=0			;== Column Name: thisAssetDesc
		cNDescription:=0		;== Column Name: CompletionDescription
		cNIssuedBy:=0			;== Column Name: IssuedBy
		cNAssignedTo:=0			;== Column Name: AssignedTo
		cNVar:=0				;== Column Name: Variant
		cNCompletedBy:=0		;== Column Name: CompletedBy
		cNOfficerComment:=0		;== Column Name: OfficerComments
		cNActionTaken:=0		;== Column Name: ActionTakenResolution
		cNReqEmail:=0			;== Column Name: ReqEmailSent
		cNCompEmail:=0			;== Column Name: CompEmailSent
		cNBSOEmail:=0			;== Column Name: BSOEmailSent
		cNRole:=0				;== Column Name: Role
		cNDataTemplate:=0		;== Column Name: DataTemplate
		cNID:=0					;== Column Name: _id
		Loop, read, exp\export.csv
		{
			rowNo:=A_Index
			thisWKONo:=""				;== Column Name: WorkOrderNumber
			thisPriority:=""			;== Column Name: DTCase
			thisReqDate:=""				;== Column Name: RequestDate
			thisPersonInspecting:=""	;== Column Name: PersonInspecting
			thisAssetNo:=""				;== Column Name: thisAsset
			thisAssetDesc:=""			;== Column Name: thisAssetDesc
			thisDescription:=""			;== Column Name: CompletionDescription
			thisIssuedBy:=""			;== Column Name: IssuedBy
			thisAssignedTo:=""			;== Column Name: AssignedTo
			thisVar:=""					;== Column Name: Variant
			thisCompletedBy:=""			;== Column Name: CompletedBy
			thisOfficerComment:=""		;== Column Name: OfficerComments
			thisActionTaken:=""			;== Column Name: ActionTakenResolution
			thisReqEmail:=""			;== Column Name: ReqEmailSent
			thisCompEmail:=""			;== Column Name: CompEmailSent
			thisBSOEmail:=""			;== Column Name: BSOEmailSent
			thisRole:=""				;== Column Name: Role
			thisDataTemplate:=""		;== Column Name: DataTemplate
			thisID:=""					;== Column Name: _id
			
			Loop, parse, A_LoopReadLine, CSV 
			{	;== Loop through each line
				if (rowNo == 1)	
				{	;== Get the row numbers for each field
					if (A_LoopField == "WorkOrderNumber")
						cNWKONo:=A_Index
					if (A_LoopField == "DTCase")
						cNPriority:=A_Index
					if (A_LoopField == "RequestDate")
						cNReqDate:=A_Index
					if (A_LoopField == "PersonInspecting")
						cNPersonInspecting:=A_Index
					if (A_LoopField == "thisAsset")
						cNAssetNo:=A_Index
					if (A_LoopField == "thisAssetDesc")
						cNAssetDesc:=A_Index
					if (A_LoopField == "CompletionDescription")
						cNDescription:=A_Index
					if (A_LoopField == "IssuedBy")
						cNIssuedBy:=A_Index
					if (A_LoopField == "AssignedTo")
						cNAssignedTo:=A_Index
					if (A_LoopField == "Variant")
						cNVar:=A_Index
					if (A_LoopField == "CompletedBy")
						cNCompletedBy:=A_Index
					if (A_LoopField == "OfficerComments")
						cNOfficerComment:=A_Index
					if (A_LoopField == "ActionTakenResolution")
						cNActionTaken:=A_Index
					if (A_LoopField == "ReqEmailSent")
						cNReqEmail:=A_Index
					if (A_LoopField == "CompEmailSent")
						cNCompEmail:=A_Index
					if (A_LoopField == "BSOEmailSent")
						cNBSOEmail:=A_Index
					if (A_LoopField == "Role")
						cNRole:=A_Index
					if (A_LoopField == "DataTemplate")
						cNDataTemplate:=A_Index
					if (A_LoopField == "_id")
						cNID:=A_Index
				}	
				else 
				{	;== Get the data from each field on this line
					if (cNWKONo==A_Index)
						thisWKONo:=A_LoopField
					if (cNPriority==A_Index)
						thisPriority:=A_LoopField
					if (cNReqDate==A_Index)
						thisReqDate:=A_LoopField
					if (cNPersonInspecting==A_Index)
						thisPersonInspecting:=A_LoopField
					if (cNAssetNo==A_Index)
						thisAssetNo:=A_LoopField
					if (cNAssetDesc==A_Index)
						thisAssetDesc:=A_LoopField
					if (cNDescription==A_Index)
						thisDescription:=A_LoopField
					if (cNIssuedBy==A_Index)
						thisIssuedBy:=A_LoopField
					if (cNAssignedTo==A_Index)
						thisAssignedTo:=A_LoopField
					if (cNVar==A_Index)
						thisVar:=A_LoopField 
					if (cNCompletedBy==A_Index)
						thisCompletedBy:=A_LoopField 
					if (cNOfficerComment==A_Index)
						thisOfficerComment:=A_LoopField 
					if (cNActionTaken==A_Index)
						thisActionTaken:=A_LoopField 
					if (cNReqEmail==A_Index)
						thisReqEmail:=A_LoopField
					if (cNCompEmail==A_Index)
						thisCompEmail:=A_LoopField
					if (cNBSOEmail==A_Index)
						thisBSOEmail:=A_LoopField
					if (cNRole==A_Index)
						thisRole:=A_LoopField
					if (cNDataTemplate==A_Index)
						thisDataTemplate:=A_LoopField
					if (cNID==A_Index)
						thisID:=A_LoopField
				
					if ((thisWKONo == "") && (thisPersonInspecting != "") && ((thisReqEmail == "") || (thisReqEmail == 0)) && (thisID != ""))	
					{	;== if Request email hasn't been sent
						FileAppend, ReqEmailSent`,_id`n1`,%thisID%, exp\upload.csv ;=== create upload file with to change ReqEmailSent field from 0 to 1 on this Konect ID
						Sleep 1000 ;== give it 1 second to ensure file has been created
						RunWait, `"%importerDir%`" -pin=834902 -action=update_matching -dataset=`"_Form_Works Order Process`" -file=`"exp\upload.csv`" ,, Hide
						if !ErrorLevel
						{
							emSubject:="You have a new Konect request from " . thisPersonInspecting
							emBody:="Hi,`n`nYou have a new Konect request from " . thisPersonInspecting . "`nLocation: " . thisAssetDesc . "`nDescription: " . thisDescription . "`nDate: " . thisReqDate . "`n`nPlease sync Konect and assess the request for approval`n`nFrom`nMatt's Konect Tool"
							if (thisRole == "Urban")
								notify(urbem, emSubject, emBody)
							if (thisRole == "Rural")
								notify(urbem, emSubject, emBody)
						}
						else
							msgbox There was an error uploading data to Konect
						while FileExist("exp\upload.csv")
							FileDelete, exp\upload.csv
					}
					
					if ((thisCompletedBy != "") && ((thisCompEmail == "") || (thisCompEmail == 0)) && (thisID != "")) 
					{	;== if Completion email hasn't been sent
						FileAppend, CompEmailSent`,_id`n1`,%thisID%, exp\upload.csv ;=== create upload file with to change CompEmailSent field from 0 to 1 on this Konect ID	
						Sleep 1000
						RunWait, `"%importerDir%`" -pin=834902 -action=update_matching -dataset=`"_Form_Works Order Process`" -file=`"exp\upload.csv`" ,, Hide
						if !ErrorLevel
						{
							emSubject:=thisWKONo . " has been completed by " . thisCompletedBy
							emBody:="Hi,`n`nA work order has been completed in Konect and requires your sign off.  Details are below:`n`nWork Order Number: " . thisWKONo . "`nCompleted By: " . thisCompletedBy . "`nOfficer Comments: " . thisOfficerComment . "`nAction Taken/Resolution: " . thisActionTaken . "`n`nPlease sync Konect and assess the job on pictures, or inspect on-site to determine if works are to standard.`n`nFrom`nMatt's Konect Tool"
							if (thisRole == "Urban")
								notify(urbem, emSubject, emBody)
							if (thisRole == "Rural")
								notify(urbem, emSubject, emBody)
						}
						else
							msgbox There was an error uploading data to Konect
						while FileExist("exp\upload.csv")
							FileDelete, exp\upload.csv
					}

					if ((thisWKONo == "") && (thisIssuedBy != "") && ((thisBSOEmail == "") || (thisBSOEmail == 0)) && (thisID != "")) 
					{	;== if BSO hasn't been notified of new WKO
						FileAppend, BSOEmailSent`,_id`n1`,%thisID%, exp\upload.csv ;=== create upload file with to change BSOEmailSent field from 0 to 1 on this Konect ID
						Sleep 1000 ;== give it 1 second to ensure file has been created
						RunWait, `"%importerDir%`" -pin=834902 -action=update_matching -dataset=`"_Form_Works Order Process`" -file=`"exp\upload.csv`" ,, Hide
						if !ErrorLevel
						{
							emSubject:="New work order to create"
							emBody:="Hi,`n`nA supervisor has requested that a work order be created.  Please refer to the Konect Tool for more information"
							IniRead, myEmailAdd, cfg\users.cfg, % thisUser, email
							if ((thisRole == "Urban" || thisRole == "Rural") && ((userRoles = 1) || (userRoles = 0)))
								notify(myEmailAdd, emSubject, emBody)
						}
						else
							msgbox There was an error uploading data to Konect
						while FileExist("exp\upload.csv")
							FileDelete, exp\upload.csv
					}
					
				}				
			}
			Gui, Default
			isNew := 1
			Loop % LV_GetCount()
			{
				LV_GetText(rowID, A_Index, 8)
				if (rowID == thisID)
				{
					isNew := 0
					break
				}
				else
					isNew := 1
			}
			
			if ((thisWKONo == "") && (thisIssuedBy != "") && (isNew == 1))
				LV_Add("",thisPriority,thisDataTemplate,thisDescription,thisAssetNo,thisAssetDesc,thisIssuedBy,thisAssignedTo, thisID)
			
		}
	}
	SetTimer, ProgressBar, Off
	Sleep 1200
	GuiControl,Text, LoadText,
	GuiControl,,LoadProg, 0
	plsWait:=false
}
return




ProgressBar:
	Loop, 100
	{
		GuiControl, , LoadProg, % A_Index
		sleep 50
	}
return

notify(email, subject, body)
{
	MailItem := ComObjCreate("Outlook.Application").CreateItem(0)
	MailItem.Recipients.Add(email)
	MailItem.Subject := subject
	MailItem.body := body
	MailItem.send
}

sendEmail(emad, rw, lno)
{
	if (lno>0)
	{
		MailItem := ComObjCreate("Outlook.Application").CreateItem(0)
		MailItem.Recipients.Add(emad)
		if (rw = "r")
		{
			MailItem.Subject := "New Works Request! (" . lno . ")"
			if (lno > 1)
				MailItem.body := "You have " . lno . " works requests from staff to approve in Konect!"
			else
				MailItem.body := "You have a works request from staff to approve in Konect!"
		}
		else if (rw = "w")
		{
			MailItem.Subject := "Works Complete! (" . lno . ")"
			if (lno > 1)
				MailItem.body := lno . " jobs have been completed in Konect.  Please assess images or site works and approve if the works are to standard."
			else
				MailItem.body := "A job has been completed in Konect.  Please assess images or site work and approve if the works are to standard."
		}
		else
		{
			MailItem.Subject := "New works order to create"
			MailItem.body := "There is a new works order that needs creating in the Konect Tool.  Please action this as soon as possible."
		}
		MailItem.send
	}
}


RL: ;reload client (menu item)
reload

GuiClose:
closeApp: ;close client (menu item)
ExitApp
