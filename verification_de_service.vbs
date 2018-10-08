'----------------------------------------------------------------------------------------------------------------------------------------
' Fichier: "Verification_de_service.vbs"
' Date : 05/06/2017
' Utilisation: Ce script permet de verifier si un service est activer et s'il n'est pas activer le script l'active
'----------------------------------------------------------------------------------------------------------------------------------------

Set WshShell = WScript.CreateObject("WScript.Shell")
If WScript.Arguments.Length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  WScript.Quit
End if

strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Service Where Name = 'DSAgent'",,48)
For Each objItem in colItems 
	If objItem.State="Running" Then
	Else 
		Dim objFSO, objWMIService, objService, colServiceList
		Dim objReseau, Ordinateur
		Dim Reponse

		Set objReseau = CreateObject("WScript.Network")
		Ordinateur = LCase(objReseau.ComputerName)
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objWMIService = GetObject("winmgmts:" & _
                                     "{impersonationLevel=impersonate}!\\" & Ordinateur & "\root\cimv2")
		Set colServiceList = objWMIService.ExecQuery _
                                    ("Select * from Win32_Service where Name='DSAgent'")
		For Each objService In colServiceList
			If (objService.Name = "DSAgent") Then
				Reponse = objService.StartService()
			End If
		Next
		Set objFSO = Nothing
		Set objReseau = Nothing
	End If 
Next 

Set colItems = Nothing 
Set objWMIService = Nothing 