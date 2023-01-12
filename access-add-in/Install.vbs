'Author: Josef Poetzl

const AddInName = "BetterAccessCharts"
const AddInFileName = "BetterAccessCharts.accda"
const MsgBoxTitle = "Install Better-Access-Charts"

MsgBox "Vor dem Aktualisieren der Add-In-Datei darf das Add-In nicht geladen sein!" & chr(13) & _
       "Zur Sicherheit alle Access-Instanzen schlieﬂen.", , MsgBoxTitle & ": Hinweis"

Select Case MsgBox("Soll das Add-In als accde verwendet werden?" + chr(13) & _
                   "(Add-In wird kompiliert ins Add-In-Verzeichnis kopiert.)", 3, MsgBoxTitle)
   case 6 ' vbYes
      CreateMde GetSourceFileFullName, GetDestFileFullName
	  MsgBox "Add-In wurde kompilliert und in '" + GetAddInLocation + "' abgespeichert", , MsgBoxTitle
   case 7 ' vbNo
      FileCopy GetSourceFileFullName, GetDestFileFullName
	  MsgBox "Add-In wurde in '" + GetAddInLocation + "' abgespeichert", , MsgBoxTitle
   case else
      
End Select





'##################################################
' Hilfsfunktionen:

Function GetSourceFileFullName()
   GetSourceFileFullName = GetScriptLocation & AddInFileName 
End Function

Function GetDestFileFullName()
   GetDestFileFullName = GetAddInLocation & AddInFileName 
End Function

Function GetScriptLocation()

   With WScript
      GetScriptLocation = Replace(.ScriptFullName & ":", .ScriptName & ":", "") 
   End With

End Function

Function GetAddInLocation()

   GetAddInLocation = GetAppDataLocation & "Microsoft\AddIns\"

End Function

Function GetAppDataLocation()

   Set wsShell = CreateObject("WScript.Shell")
   GetAppDataLocation = wsShell.ExpandEnvironmentStrings("%APPDATA%") & "\"

End Function

Function FileCopy(SourceFilePath, DestFilePath)

   set fso = CreateObject("Scripting.FileSystemObject") 
   fso.CopyFile SourceFilePath, DestFilePath

End Function

Function CreateMde(SourceFilePath, DestFilePath)

   Set AccessApp = CreateObject("Access.Application")
   AccessApp.SysCmd 603, (SourceFilePath), (DestFilePath)

End Function
