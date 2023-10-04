Option Explicit

Sub run()

Dim objShell As Object
Dim PythonExe As String
Dim PythonScript As String

Set objShell = VBA.CreateObject("Wscript.Shell")

PythonExe = "C:\Python39\python.exe"

PythonScript = ".\[test.py](http://test.py/)"

objShell.run PythonExe & PythonScript

End Sub

---

Option Explicit

Public WithEvents outApp As Outlook.Application

Sub Intialize_Handler()
Set outApp = Application
End Sub

---

Public WithEvents outApp As Outlook.Application

Sub Application_NewMailEx(ByVal oRequest As MeetingItem)

If oRequest.MessageClass <> "IPM.Schedule.Meeting.Request" Then
End Sub
End If

Shell ("python .\[test.py](http://test.py/)")

End Sub