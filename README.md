<div align="center">

## FTP Upload Function \- Simple \- Works

<img src="PIC2003151356317012.jpg">
</div>

### Description

Simple Function designed to upload single/multiple files from VB using just Inet control, Label and the Function.
 
### More Info
 
Single Path+Filename, or comma separated string containing multiple names.

Code for newbies like me - should just work.

The function returns the number of files uploaded. If the value is negative, shows number of files uploaded before a user abort was initiated.

No progress bar - to keep code simple. Label flashes when transfer in progress. Label can be clicked during transfer to abort during multiple transfers.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ken Ashton](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ken-ashton.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ken-ashton-ftp-upload-function-simple-works__1-42204/archive/master.zip)





### Source Code

```
'
'
Add 1 label, Inet control and Function to any form requiring upload
facility. Works extremely reliably when uploading in excess of 100
files, some files up to 8mB in size.
'
When the function is passed Source path+filename/s - it either
1. Recognizes the (local Pc) file pathname as the same structure as
 your web site, and uploads the file to the correct folder, or
2. Recognizes an 'alternate' path (declared as constant in function),
 and uploads the file to the defined 'alternate server path', else
3. Prompts the user for a server path to use for the upload.
'
'
1. Project/Components menu, add Microsoft Internet Transfer Control 6.0
2. On Form1, put Inet1 and Label1 as Single Border, Centered
3. Add the code below
4. Enter your Host, Username, Password and server paths in function
'
Notes: A lot of the stuff is optional, Label1 etc. Inet Control can be a bitch
so take care, key is correct confiuration and correctly 'quoted' command, eg
'
	Inet1.Execute , "c:\samplefolder\somefile.jpg" /images/somefile.jpg
'
If like me, you keep your site in a folder structure that mirrors the server folder
structure (from the web site root up), but still have the odd master file (like a
database) kept somewhere else, you will find the ability to define both the normal
and alternate local folders useful. Should neither appear in the filepath passed to
the function, you may find the 'server folder prompt' useful for ad-hoc uploads.
'
Objective was reliable, and practical functionality with simplicity
Please post any comments with improved code for the benefit of all of us amateurs.
' -- Form Declarations section
'
Option Explicit
Option Compare Text
Dim gState As Integer   ' Used to remember last Inet state
Dim gCount As Integer   ' Used to record 'files Uploaded' count
Dim gCancel As Boolean   ' Used to remember Cancel has been pressed
'
'
' -- Form load event --
'
'
Private Sub Form_Load()
 Label1 = "Click to Upload"
 Label1.ToolTipText = Label1
End Sub
'
'
' -- Label1 click event --
'
'
Private Sub Label1_Click()
 Dim a%, m
 If gState = 0 And Not gCancel Then
  Label1.ToolTipText = "Click to Cancel"
  ' Here is where we make the call FTPFile(File1,File2...etc)
  a = FTPfile("c:\inetpub\wwwroot\mysite\mainindex.htm,c:\inetpub\wwwroot\mysite\towns.htm")
  If gCancel Or a < 0 Then m = vbCrLf & vbCrLf & "** User Aborted **"
  MsgBox "Uploaded " & Str(Abs(a)) & " files" & m
  gState = 0
  gCancel = False
 Else
  Label1 = "User Aborting!"
  Label1.ToolTipText = "Aborting - please be patient!"
  Label1.BackColor = vbRed
  gCancel = True
 End If
End Sub
'
'
' Inet control change statechanged event
'
'
Private Sub Inet1_StateChanged(ByVal State As Integer)
 gState = State
 If State = 12 Then gCount = gCount + 1
End Sub
'
'
' The uploader Function
'
'
Function FTPfile(lf) As Integer
 Dim sf As String
 Dim fn As String
 Dim lastState As Integer
 Dim arrLf As Variant
 Dim tm As Variant
 Dim i As Integer
 Dim tmpS As String
'
' ----------------------- Set your FTP constants -------------------------
'
 Const HostName = "127.127.127.127" ' **** Your Host URL
 Const UserName = "username"   ' **** Your login Username
 Const Password = "password"   ' **** Your loginPpassword
 Const nrmL = "\mysite\"    ' **** Eg if pc path = c:\inetpub\wwwroot\mysite\index.html, use '\mysite\'
 Const altL = "\alternatelocalpath\" ' **** Eg alternate local path
 Const altS = "/alternateserverpath/" ' **** Eg alternate server folder, where
'
' ---------------------- Configure the Inet Control ----------------------
'
 Inet1.AccessType = icDirect
 Inet1.Protocol = icFTP
 Inet1.RemoteHost = HostName
 Inet1.UserName = UserName
 Inet1.Password = Password
'
' --- Extract file list, upload files, return number of files uploaded ---
'
 arrLf = Split(LCase(lf), ",")       ' Get inputted filenames
 Label1 = ""            ' Force label to null
 gCount = 0            ' Set files 'uploaded' to 0
 Screen.MousePointer = vbHourglass      ' Set 'Busy' Hourglass pointer
 For i = LBound(arrLf) To UBound(arrLf)     ' Loop thro' inputted files
 If Len(arrLf(i)) > 0 And Dir(arrLf(i)) <> "" Then  ' Check file exists
  If Label1 = "" Then Label1 = "Connecting"   ' Init Status label
  Label1.BackColor = vbWhite
  fn = Mid(arrLf(i), InStrRev(arrLf(i), "\") + 1)  ' Extract filename
  If InStr(arrLf(i), altL) > 0 Then     ' If its alternate path
   sf = altS & fn         ' construct alt serverpath
  ElseIf InStr(arrLf(i), nrmL) > 0 Then    ' else normal serverpath
   sf = Replace(Mid(arrLf(i), InStr(arrLf(i), nrmL) + 12), "\", "/")
  Else            ' If unknown local path ask
   If Len(tmpS) = 0 Then
    tmpS = "/<folder>/"
    Screen.MousePointer = vbDefault
    tmpS = Trim(Replace(InputBox("Server folder", "Require Server Folder", tmpS), "\", "/"))
    If tmpS = "" Then Exit Function    ' Allow user to quit
    Screen.MousePointer = vbHourglass
    If Right(tmpS, 1) <> "/" Then
     tmpS = tmpS & "/"      ' else use entered serverpath
    End If
   End If
   sf = tmpS & fn
  End If
  sf = "put " & Chr(34) & arrLf(i) & Chr(34) & " " & sf ' Construct upload command
  If Not gCancel Then Inet1.Execute , sf    ' Initiate the Upload
  Do Until Inet1.StillExecuting = 0 ' Hang around
   DoEvents           ' let windows do other stuff
   If tm = 0 Then         ' Set up simple timer.
     tm = Timer        ' Initialize Timer.
   ElseIf Timer > tm + 0.25 Then      ' If timer expires, then
    tm = 0          ' reset timer for next
    tm = Timer         ' interval and update
    If Label1.BackColor = &H80000013 Then   ' 'status label'. Also
     Label1.BackColor = vbWhite    ' 'toggle' backcolor to give
    Else           ' user a simple 'busy'
      Label1.BackColor = &H80000013   ' indication
    End If
    If lastState <> gState Then     ' If the state has changed
     If gState > 4 And gState < 9 Then   ' and is one of the normal
      Label1 = fn & " " & Str(i + 1) & " of " & Str(UBound(arrLf) + 1)
     End If         ' the status label
     lastState = gState      ' Note last state serviced
    End If
   End If
   If gCancel Then      ' If cancel was pressed
    If Inet1.StillExecuting Then ' and Inet still executing
     Inet1.Cancel    ' then issue a cancel to
    End If       ' Inet control. Make upload
    gCount = gCount * -1   ' count Negative to show abort
    Do While Inet1.StillExecuting ' and wait for inet execution
     DoEvents     ' to terminate
    Loop       ' before exiting
    Exit Do       ' the main execution loop
   End If
  Loop
 End If
 If gCancel Then Exit For     ' Don't do any more files
 Next          ' if cancel was pressed
 Inet1.Execute , "Quit"      ' Close down connection
 Do While Inet1.StillExecuting    ' Wait until done
  DoEvents        ' allowing other windows
 Loop          ' events to execute
 Label1 = "Click to Upload"     ' Restore the Status label
 Label1.ToolTipText = Label1
 Label1.BackColor = &H80000013    ' text and backcolor
 Screen.MousePointer = vbDefault    ' Restore nomal pointer
 FTPfile = gCount       ' Return count of files uploaded
            ' (negative indicates aborted)
End Function
```

