Option Explicit
Dim app

Call dowork

If not (app Is Nothing) Then
    app.Quit
    Set app = Nothing
End If

Sub dowork()
    On Error Resume Next

    Dim prjName
    Dim prjObject

    '######################################################################################
    'PROJECT PATH FROM SASRun.py
    '######################################################################################
    Dim prjPaths
    prjPaths = Array(Wscript.Arguments(0))	
    For Each prjName in prjPaths
    Wscript.Echo prjName
    '######################################################################################
    '######################################################################################

    Set app = CreateObject("SASEGObjectModel.Application.7.1")
    If Checkerror("CreateObject") = True Then
        Wscript.Echo "[SAS] ERROR OCCURED"
        Exit Sub
    End If

    Wscript.Echo "[SAS] Opening Project File"
    Set prjObject = app.Open(prjName,"")
    If Checkerror("app.Open") = True Then
        Wscript.Echo "[SAS] ERROR OCCURED"
        Exit Sub
    End If

    Wscript.Echo "[SAS] Running Project"
    prjObject.run
    If Checkerror("Project.run") = True Then
        Wscript.Echo "[SAS] ERROR OCCURED"
        Exit Sub
    End If

    Wscript.Echo "[SAS] Saving Project File"
    prjObject.Save
    If Checkerror("Project.Save") = True Then
       Exit Sub
    End If

    Wscript.Echo "[SAS] Closing Project"
    prjObject.Close
    If Checkerror("Project.Close") = True Then
        Wscript.Echo "[SAS] ERROR OCCURED"
        Exit Sub
    End If

    Wscript.Echo "[SAS] COMPLETE"

    Next    
End Sub

Function Checkerror(fnName)
    Checkerror = False
    
    Dim strmsg
    Dim errNum
    
    If Err.Number <> 0 Then
        strmsg = "Error #" & Hex(Err.Number) & vbCrLf & "In Function " & fnName & vbCrLf & Err.Description
        'MsgBox strmsg  'Uncomment this line if you want to be notified via MessageBox of Errors in the script.
        Checkerror = True
    End If
         
End Function
