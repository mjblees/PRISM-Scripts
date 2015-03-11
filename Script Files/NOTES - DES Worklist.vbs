Option Explicit

DIM beta_agency

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

DIM DES_Worklist_Dialog, New_info, Employer_Name, First_Line, Second_Line, Additional_Information, DES_Combobox, Verified, ButtonPressed, Combo1


BeginDialog DES_Worklist_Dialog, 0, 0, 386, 355, "Des Worklist Dialog"
  ComboBox 25, 25, 140, 35, "Select One"+chr(9)+"Old Information"+chr(9)+"No New Information"+chr(9)+"New Information", DES_Combobox
  EditBox 90, 60, 70, 15, New_info
  EditBox 105, 95, 50, 15, Employer_Name
  DropListBox 105, 130, 60, 45, "Yes"+ chr(9)+ "No", Verified
  EditBox 105, 165, 120, 15, First_Line
  EditBox 105, 185, 120, 15, Second_line
  EditBox 35, 240, 335, 15, Additional_Information
  ButtonGroup ButtonPressed
    OkButton 260, 330, 50, 15
    CancelButton 325, 330, 50, 15
  Text 175, 30, 100, 10, "DES Worklist Reviewed Note"
  Text 30, 60, 60, 10, "New Information."
  Text 30, 100, 60, 10, "Employer Name"
  Text 35, 130, 50, 10, "Verified"
  Text 35, 165, 55, 10, "New Address"
  Text 30, 220, 85, 10, "Additional Information"
  ComboBox -350, 135, 60, 45, "", Combo1
EndDialog


'Connecting to Bluezone
EMConnect "" 

'Checks to make sure we are in PRISM
CALL Check_for_PRISM(true)

Do 

	

dialog DES_Worklist_Dialog
 
If New_info = "" THEN MsgBox "You Must Add New Information"
LOOP UNTIL New_info <> ""

CALL NAVIGATE_TO_PRISM_SCREEN ("CAAD")'GETS TO CAAD
'sets as free mode


PF5 'CREATES A NEW NOTE
EMWriteScreen "A", 3, 29 ' SETS TO ADD NOTE
EMWriteScreen "Free", 4, 54 'ADDS A FREE NOTE
EMSetCursor 16, 4 'PUTS THE CURSOR AT THE BEGINNING OF CAAD

'Writing the CAAD note


Call write_bullet_and_variable_in_CAAD("DES Worklist Review Note", DES_Worklist_Dialog)
Call write_bullet_and_variable_in_CAAD ("New Information", New_info)
Call write_bullet_and_variable_in_CAAD ("Employer Name", Employer_Name)
Call write_bullet_and_variable_in_CAAD ("New Address", First_Line)
Call write_bullet_and_variable_in_CAAD ("New Address", Second_Line)
Call write_bullet_and_variable_in_CAAD ("Additional Information", Additional_Information)

Transmit

StopScript

















