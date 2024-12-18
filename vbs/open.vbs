Set WshShell = CreateObject("WScript.Shell")

Do
    ' Attempt to run PowerShell as Administrator
    strCommand = "powershell.exe -Command ""Start-Process PowerShell -Verb RunAs"""
    
    ' Ask for permission to run PowerShell as Administrator
    result = WshShell.Run(strCommand, 0, True)
    
    ' If the user canceled, result will be 0 (cancelled)
    If result = 0 Then
        ' Show a message and continue looping
        WshShell.Popup "You canceled the permission request. Trying again...", 2, "Permission Required", 64
    End If
Loop Until result = 1 ' Exit loop when the user allows PowerShell to run
