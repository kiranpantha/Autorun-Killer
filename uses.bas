Attribute VB_Name = "uses"
Option Explicit
Const mnErrDeviceUnavailable = 68
Const mnErrDiskNotReady = 71
Const mnErrDeviceIO = 57
Const mnErrDiskFull = 61
Const mnErrBadFileName = 64
Const mnErrBadFileNameOrNumber = 52
Const mnErrPathDoesNotExist = 76
Const mnErrBadFileMode = 54
Const mnErrFileAlreadyOpen = 55
Const mnErrInputPastEndOfFile = 62

Function autorun(drivelet As String, todel As String)
On Error GoTo errorlist
Dim nepal$
SetAttr todel, vbNormal
nepal$ = "del " & drivelet & todel
MsgBox todel
Open "autoexec.bat" For Output As #1
Print #1, nepal$
Close
Shell ("autoexec.bat")
home.autolist.Caption = "Autorun file lacated in " & todel & " was successfully deleted"
GoTo last
errorlist:
Call FileErrors(Err.Number)
last:
End Function
Sub FileErrors(errornum As Integer)
    Dim intMsgType As Integer
    Dim strMsg As String
    Dim intResponse As Integer
    intMsgType = vbExclamation
    Select Case errornum
        Case mnErrDeviceUnavailable             ' Error 68
            strMsg = "That device appears unavailable."
            intMsgType = vbExclamation + vbOKCancel
           
        Case mnErrDiskNotReady                  ' Error 71
            strMsg = "Insert a disk in the drive and close the door."
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrDeviceIO                      ' Error 57
            strMsg = "Internal disk error."
            intMsgType = vbExclamation + vbOKOnly
        Case mnErrDiskFull                      ' Error 61
            strMsg = "Disk is full. Continue?"
            intMsgType = vbExclamation + vbAbortRetryIgnore
        Case mnErrBadFileName, mnErrBadFileNameOrNumber ' Error 64 & 52
            strMsg = "That filename is illegal."
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrPathDoesNotExist                ' Error 76
            strMsg = "That path doesn't exist."
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrBadFileMode                     ' Error 54
            strMsg = "Can't open your file for that type of access."
        Case mnErrFileAlreadyOpen             ' Error 55
            strMsg = "This file is already open."
            intMsgType = vbExclamation + vbOKOnly
        Case mnErrInputPastEndOfFile              ' Error 62
            strMsg = "This file has a nonstandard end-of-file marker, "
            strMsg = strMsg & "or an attempt was made to read beyond "
            strMsg = strMsg & "the end-of-file marker."
            intMsgType = vbExclamation + vbAbortRetryIgnore
        Case Else
            Exit Sub
    End Select
     MsgBox strMsg, intMsgType
End Sub
