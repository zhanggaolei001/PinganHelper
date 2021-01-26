' *******************************************************************************
' PROJECT.......: #vbs Capture Screenshot (beta)
' SCRIPT........: tool.CaptureScreenshot.vbs
' DESCRIPTION...: capture screen shot and save as an image
' REQUIREMENTS..: OS: Windows with Microsoft Paint (画图)
'                 Microsoft Word: tested on 2007/2010
' CREATED.......: 20180208
' AUTHOR........: ion.chen
' NOTES.........:
'                 
' UPDATE........:
'                 20171220 fix: numlock turns off everytime.
'                 20180208 add: check after done, and view it if secceeded.
' TO-DO.........:
'                [ ] Capture Window Only.
'                [ ] Count Down before Capturing.
' *******************************************************************************

' ===============================================================================
'  START OF SCRIPT
' ===============================================================================
'Option Explicit
'On Error Resume Next

    ' ---------------------------------------------------------------------------
    '  Declare Constants
    ' ---------------------------------------------------------------------------
    ' SET_PATH_OF_FOLDER_TO_SAVE_SCREENSHOTS
    Const SCREENSHOT_FOLDER = "D:\"
    ' CAPTURE_FULLSCREEN_OR_ACTIVE_WINDOW_ONLY
    Const SAVE_FULLSCREEN = True
	
'**Start Encode**
    ' ---------------------------------------------------------------------------
    '  Declare Variables
    ' ---------------------------------------------------------------------------

    ' ===============================================================================
    '  SUBROUTINES/FUNCTIONS/CLASSES
    ' ===============================================================================
    Call SaveScreenAsImage

    ' --------------------------------------------------------------------------
    '  SUBROUTINE.....:  SaveScreenAsImage
    '  PURPOSE........:  Save current screen as an image.
    '  EXAMPLE........:  Call SaveScreenAsImage
    ' --------------------------------------------------------------------------

    Sub SaveScreenAsImage 
		set objHTML = CreateObject("htmlfile")
		text = objHTML.ParentWindow.ClipboardData.GetData("text")
		
        ' SET_FILENAME_AS_CURRENT_DATETIME
        Dim sFileName : sFileName = SCREENSHOT_FOLDER & text
		
        Dim WshShell : Set WshShell = WScript.CreateObject("WScript.Shell")
        Dim objFSO : Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
        Dim strCmdOpenFile, strFileExtenstion
        
        ' GO_AROUND_WSH_AND_CAPTURE_SCREEN_(SINCE_IT_HAS_BEEN_DISABLED)
        With CreateObject("Word.Basic") 
            If SAVE_FULLSCREEN _
            Then .Sendkeys "{PrtSc}" _
            Else .Sendkeys "%{PrtSc}"
            .FileQuit
        End With 
        ' RESUME_NUMLOCK_AFTER_PRINT_SCREEN
        WshShell.SendKeys "{NUMLOCK}" 
        ' RUN_MSPAINT
        WshShell.Run "mspaint.exe", 4
        WScript.Sleep 300
        WshShell.AppActivate "画图"
        WshShell.AppActivate "paint"  
        WScript.Sleep 100
        WshShell.SendKeys "^(v)" 
        WshShell.SendKeys "^(s)"
        WScript.Sleep 500
        WshShell.SendKeys sFileName
        WshShell.SendKeys "%(s)"
        ' CLOSE_MSPAINT
        WScript.Sleep 100
        WshShell.SendKeys "%{F4}"
        ' CHECK_WHETHER_IMAGE_FILE_EXISTS
        WScript.Sleep 100
       
        strFileExtenstion = ".png"
        
		WshShell.SendKeys "{NUMLOCK}"
        ' OPEN_IMAGE_FILE
        strCmdOpenFile = "rundll32.exe %WinDir%\System32\shimgvw.dll,ImageView_Fullscreen " & _
            sFileName & strFileExtenstion
        WshShell.Run strCmdOpenFile
		
    End Sub

    ' --------------------------------------------------------------------------
    '  FUNCTION.......:  GetDateTimeNoSeparator(dSpecTime)
    '  PURPOSE........:  Get current date and time formated as YYYYMMDDHHMMSS
    '  EXAMPLE........:  MsgBox GetDateTimeNoSeparator(Now)
    ' --------------------------------------------------------------------------
    Function GetDateTimeNoSeparator(dSpecTime)
        If Not IsDate(dSpecTime) Then dSpecTime = Now()
        GetDateTimeNoSeparator = Year(dSpecTime) _
            & Right("0" & Month(dSpecTime), 2) _
            & Right("0" & Day(dSpecTime), 2) _
            & Right("0" & Hour(dSpecTime), 2) _
            & Right("0" & Minute(dSpecTime), 2) _
            & Right("0" & Second(dSpecTime), 2)
    End Function