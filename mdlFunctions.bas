Attribute VB_Name = "mdlFunctions"
Option Explicit

' Nobody has heard of most of these functions! Amazing!
Declare Function SHChangeIconDialog Lib "Shell32.DLL" Alias "#62" (ByVal hWndOwner As Long, ByVal szInitFilename As String, ByVal dwReserved As Long, lpIconIndex As Long) As Long
Declare Function SHFormatDrive Lib "Shell32.DLL" (ByVal hWndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long
Declare Function SHIsPathExecutable Lib "Shell32.DLL" Alias "#43" (ByVal szPath As String) As Long
Declare Function SHRestartSystemMessageBox Lib "Shell32.DLL" Alias "#59" (ByVal hWndOwner As Long, ByVal szExtraPrompt As String, ByVal uFlags As Long) As Long
Declare Function SHRunDialog Lib "Shell32.DLL" Alias "#61" (ByVal hWndOwner As Long, ByVal dwReserved1 As Long, ByVal dwReserved2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long

' A few more:
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
' Update in this one: The alias was incorrect. (The d in Associated was missing... Bad API viewer file!)
Declare Function ExtractAssociatedIcon Lib "Shell32.DLL" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Declare Function DrawIconEx Lib "User32.DLL" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Declare Function OleTranslateColor Lib "OlePro32.DLL" (ByVal oleColor As OLE_COLOR, ByVal hPalette As Long, pColorRef As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' OLE constant:
Public Const CLR_INVALID = &HFFFFFFFF

' Constants for SHRunDialog:
Public Const SHRD_NOBROWSE = 1 ' If specified, the "Browse" button won't appear
Public Const SHRD_NOSTRING = 2 ' If specified, there won't be an initial string in the dialog

' Constants for SHFormatDrive:
Public Const SHFDCapacityDefault = 0 ' 1.2MB or 1.44MB, depending on drive
Public Const SHFDCapacity360KB = 3 ' 360KB instead of 1.2MB
Public Const SHFDCapacity720KB = 5 ' 720KB instead of 1.44MB

' More constants for SHFormatDrive (these are self-explanatory):
Public Const SHFDTypeQuickFormat = 0
Public Const SHFDTypeFullFormat = 1
Public Const SHFDTypeCopySystemFilesOnly = 2

' Constants for ExitWindowsEx... I found them another use here!
Public Const EWX_LOGOFF = 0 ' Simply log off.
Public Const EWX_SHUTDOWN = 1 ' Shut down.
Public Const EWX_REBOOT = 2 ' Restart.
Public Const EWX_FORCE = 4 ' Do whatever of the others, but FORCE it!
Public Const EWX_POWEROFF = 8 ' (Cute hardware and Win98+) Turn the computer off.

' Update! Constants for DrawIconEx (the reason it didn't work):
Public Const DI_MASK = 1            ' Draw icon using mask
Public Const DI_IMAGE = 2           ' Draw icon using image
Public Const DI_NORMAL = DI_MASK Or DI_IMAGE ' Draw icon using "masked image"

' This function translates any color to RGB, even System Colors like vbButtonFace.
Function TranslateColor(ByVal oleColor As OLE_COLOR, Optional ByVal hPalette As Long = 0) As Long
    If Not OleTranslateColor(oleColor, hPalette, TranslateColor) = 0 Then TranslateColor = CLR_INVALID
End Function

' This function strips StringWithNulls and returns the same string up to the first null.
Sub StripNulls(StringWithNulls As String)
    Dim lPos As Long
    lPos = InStr(StringWithNulls, vbNullChar)
    If lPos > 0 Then StringWithNulls = Left(StringWithNulls, lPos - 1)
End Sub

' This function displays the Change Icon Dialog (like in the Shortcut Properties).
' FileName: The default FileName of the icon file. This may change and the change will be returned (ByRef).
' hWndOwner: The hWnd of the owner of the dialog, if any.
' IconIndex: The default Icon Index of the icon in the file. 0 is 1st, 1 is 2nd, etc. This may change (ByRef).
' GetHandle: If True, on success the function returns an icon handle. If False, on success -1 is returned.
' Returns: Success - dependant on GetHandle. Failure - 0.
Function DisplayChangeIconDialog(FileName As String, Optional ByVal hWndOwner As Long = 0, Optional IconIndex = 0, Optional ByVal GetHandle As Boolean = True) As Long
    If SHChangeIconDialog(hWndOwner, FileName, 0, IconIndex) = 0 Then Exit Function ' Failure? Exit Function
    Call StripNulls(FileName) ' It will probably contain nulls!
    If GetHandle Then
        DisplayChangeIconDialog = ExtractAssociatedIcon(App.hInstance, FileName, IconIndex) ' Extract the icon
        ' Do NOT forget to DeleteObject it when you are done!!!
    Else
        DisplayChangeIconDialog = -1 ' Success but do nothing... FileName and IconIndex are returned ByRef anyway
    End If
End Function

' This function draws the icon in hIcon. (hIcon is returned with DisplayChangeIconDialog when GetHandle = True)
' hDCOwner: The hDC of the Form or PictureBox or object where to draw the icon.
' hIcon: The handle of the icon.
' X, Y: The upper-left corner of the location where drawing is wanted.
' Width, Height: The width and the height of the picture to draw (usually 32x32, for icons).
'                Zero means - use the size of the actual picture.
' BackColor: The color to use for the default background color in drawing the image.
' DeleteAfterDraw: Whether to delete the picture from the memory after drawing.
'                  If you don't delete it, you may use it again, but you must delete it sometime with DeleteObject.
' Returns True on success, False on failure.
' Update! This function didn't work before, and should work now.
' Also see the minor update for the constants in the General Declarations area.
Function DrawExtractedIcon(ByVal hDCOwner As Long, ByVal hIcon As Long, Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0, Optional ByVal Width As Long = 0, Optional ByVal Height As Long = 0, Optional ByVal BackColor As Long = -1, Optional ByVal DeleteAfterDraw As Boolean = True) As Boolean
    Dim hBrush As Long
    BackColor = TranslateColor(BackColor) ' Translate the BackColor (might be a system color)
    If Not BackColor = CLR_INVALID Then hBrush = CreateSolidBrush(BackColor) ' If it's real, make it into a brush
    ' Here's where the update is: (The flag parameter was zero instead of DI_NORMAL)
    DrawExtractedIcon = DrawIconEx(hDCOwner, X, Y, hIcon, Width, Height, 0, hBrush, DI_NORMAL) ' Draw!
    Call DeleteObject(hBrush) ' Delete the brush
    If DeleteAfterDraw Then Call DeleteObject(hIcon) ' Delete the icon, if it is requested
End Function

' This function displays the Format Drive dialog.
' DriveLetter: The drive letter to format. Only the first letter is used.
' hWndOwner: The hWnd of the owner of the dialog, if any.
' FormatType: The type of format to do - quick, full or only copy the system files.
' FormatCapacity: The capacity of the disk - default high-capacity, 360KB or 720KB.
' Returns True on success, False on failure (or on cancel).
Function DisplayFormatDriveDialog(ByVal DriveLetter As String, Optional ByVal hWndOwner As Long = 0, Optional ByVal FormatType As Long = SHFDTypeFullFormat, Optional ByVal FormatCapacity As Long = SHFDCapacityDefault) As Boolean
    Dim iDrive As Integer
    If Len(DriveLetter) = 0 Then Exit Function ' Must format SOMETHING
    iDrive = Asc(UCase(Left(DriveLetter, 1))) - 65 ' Convert first letter to ASCII and decrease by 65 to get A=0, B=1...
    If (iDrive < 0) Or (iDrive > 25) Then Exit Function ' Must be a letter!
    DisplayFormatDriveDialog = (SHFormatDrive(hWndOwner, iDrive, FormatCapacity, FormatType) > 0) ' Result must be > 0
End Function

' Very simple function... Tells you if a FileName can be executed.
' Note - the file DOESN'T have to exist - just have a proper extension.
' PathName: The path to the file in question.
' Returns True if extension is executable, False if it isn't.
Function IsPathExecutable(ByVal PathName As String) As Boolean
    IsPathExecutable = SHIsPathExecutable(PathName) ' As simple as that!
End Function

' This function creates a Windows MsgBox which asks you if you really want to restart or something.
' Operation: What the MsgBox should do, if Yes is clicked - ExitWindowsEx constant(s).
' hWndOwner: The hWnd of the owner of the MsgBox, if any.
' ExtraText: What to say in the beginning.
' Returns the result of the MsgBox (vbYes or vbNo).
Function ExitWindowsMsgBox(Optional ByVal Operation As Long = EWX_SHUTDOWN, Optional ByVal hWndOwner As Long = 0, Optional ByVal ExtraText As String = vbNullString) As VbMsgBoxResult
    ExitWindowsMsgBox = SHRestartSystemMessageBox(hWndOwner, ExtraText, Operation) ' Another simple API call...
End Function


' This function displays the Run Dialog (like in Start -> Run...).
' hWndOwner: The hWnd of the owner of the dialog.
' Caption: "Run" is a bad caption, choose your own!
' Prompt: "Type the name of a program, wait, actually don't!" Finally get to choose what you want to write there.
' BrowseButton: Whether or not you want that Browse... button there.
' InitialSelection: Whether or not you want anything to be written in the ComboBox when started (if False,
'                   a string is retrieved from the Run MRU list in the registry).
' Returns False on failure (though I could never get it to fail) or True on success.
Function DisplayRunDialog(Optional ByVal hWndOwner As Long = 0, Optional ByVal Caption As String = vbNullString, Optional ByVal Prompt As String = vbNullString, Optional ByVal BrowseButton As Boolean = True, Optional ByVal InitialSelection As Boolean = True) As Boolean
    Dim uFlags As Long
    If Not BrowseButton Then uFlags = uFlags Or SHRD_NOBROWSE
    If Not InitialSelection Then uFlags = uFlags Or SHRD_NOSTRING
    DisplayRunDialog = Not CBool(SHRunDialog(hWndOwner, 0, 0, Caption, Prompt, uFlags)) ' No! two Reservedies! "RUN"!
End Function

Sub Main()
    ' One of my favorite examples:
    Call ExitWindowsMsgBox(EWX_REBOOT, , "The mouse has moved, so... ")
End Sub
