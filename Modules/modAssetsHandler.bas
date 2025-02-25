Attribute VB_Name = "modAssetsHandler"
''' modAssetsHandler
''' Handles all the assets.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Win32 API imports.
Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

' Constants.
Public Const IMAGE_FOLDER As String = "Images\"
Public Const DATASHEET_FOLDER As String = "Datasheets\"
Public Const IMAGE_EXT As String = ".bmp"
Public Const DATASHEET_EXT As String = ".pdf"

' Get path to the datasheets directory.
Public Function GetDatasheetsDirectory() As String
    GetDatasheetsDirectory = GetWorkspacePath() & DATASHEET_FOLDER
End Function

' Gets the path of a datasheet given a name.
Public Function GetComponentDatasheetPath(strName As String, _
        blnCheckExists As Boolean) As String
    Dim strPath As String
    
    ' Build datasheet path.
    strPath = GetDatasheetsDirectory() & strName & DATASHEET_EXT
    GetComponentDatasheetPath = strPath
    
    ' Check if the datasheet exists.
    If blnCheckExists Then
        If FileExists(strPath) Then
            GetComponentDatasheetPath = strPath
        Else
            GetComponentDatasheetPath = vbNullString
        End If
    End If
End Function

' Checks if a component has a datasheet.
Public Function ComponentHasDatasheet(strName As String) As Boolean
    If GetComponentDatasheetPath(strName, True) <> vbNullString Then
        ComponentHasDatasheet = True
    Else
        ComponentHasDatasheet = False
    End If
End Function

' Renames the datasheet of a component.
Public Sub RenameComponentDatasheet(strOldName As String, strNewName As String)
    Dim strOldPath As String
    Dim strNewPath As String
    
    ' Check if it actually exists.
    If ComponentHasDatasheet(strOldName) Then
        ' Get new and old paths.
        strOldPath = GetComponentDatasheetPath(strOldName, True)
        strNewPath = GetComponentDatasheetPath(strNewName, False)
        
        ' Check if the new path is available.
        If FileExists(strNewPath) Then
            MsgBox "Cannot rename datasheet from " & strOldName & " to " & _
                strNewName & " becase there's another datasheet with that " & _
                "name already", vbOKOnly + vbCritical, "Datasheet Rename Error"
            Exit Sub
        End If
        
        ' Rename the file.
        'Name strOldPath As strNewPath
        MsgBox "File rename", vbOKOnly + vbInformation, "To Be Implemented"
    End If
End Sub

' Deletes the datasheet of a component.
Public Sub DeleteComponentDatasheet(strName As String)
    ' Check if it actually exists.
    If ComponentHasDatasheet(strName) Then
        DeleteFile GetComponentDatasheetPath(strName, False)
    End If
End Sub

' Opens a component datasheet file.
Public Sub OpenComponentDatasheet(strName As String)
    If ComponentHasDatasheet(strName) Then
        ShellExecute Screen.ActiveForm.hWnd, "open", GetComponentDatasheetPath(strName, True), _
            vbNullString, vbNullString, 1
    End If
End Sub
' Gets the path to the images directory.
Public Function GetImagesDirectory() As String
    GetImagesDirectory = GetWorkspacePath() & IMAGE_FOLDER
End Function

' Gets the path of an image given a name.
Public Function GetComponentImagePath(strName As String, strPackage As String, _
        blnCheckExists As Boolean) As String
    Dim strPath As String
    
    ' Check for image by component name.
    If strName <> vbNullString Then
        ' Build component image path.
        strPath = GetImagesDirectory() & strName & IMAGE_EXT
        GetComponentImagePath = strPath
        
        ' Check if it actually exists.
        If blnCheckExists Then
            If FileExists(strPath) Then
                GetComponentImagePath = strPath
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    ' Check for image by component package.
    If strPackage <> vbNullString Then
        ' Build package image path.
        strPath = GetImagesDirectory() & strPackage & IMAGE_EXT
        GetComponentImagePath = strPath
        
        ' Check if it actually exists.
        If blnCheckExists Then
            If FileExists(strPath) Then
                GetComponentImagePath = strPath
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    GetComponentImagePath = vbNullString
End Function

' Check if a component has an image associated with it.
Public Function ComponentHasImage(strName As String, strPackage As String) As Boolean
    If GetComponentImagePath(strName, strPackage, True) <> vbNullString Then
        ComponentHasImage = True
    Else
        ComponentHasImage = False
    End If
End Function

' Renames the image of a component.
Public Sub RenameComponentImage(strOldName As String, strNewName As String)
    Dim strOldPath As String
    Dim strNewPath As String
    
    ' Check if it actually exists.
    If ComponentHasImage(strOldName, vbNullString) Then
        ' Get new and old paths.
        strOldPath = GetComponentImagePath(strOldName, vbNullString, True)
        strNewPath = GetComponentImagePath(strNewName, vbNullString, False)
        
        ' Check if the new path is available.
        If FileExists(strNewPath) Then
            MsgBox "Cannot rename component image from " & strOldName & " to " & _
                strNewName & " becase there's another image with that " & _
                "name already", vbOKOnly + vbCritical, "Image Rename Error"
            Exit Sub
        End If
        
        ' Rename the file.
        'Name strOldPath As strNewPath
        MsgBox "File rename", vbOKOnly + vbInformation, "To Be Implemented"
    End If
End Sub

' Deletes the image of a component.
Public Sub DeleteComponentImage(strName As String)
    ' Check if it actually exists.
    If ComponentHasImage(strName, vbNullString) Then
        DeleteFile GetComponentImagePath(strName, vbNullString, False)
    End If
End Sub

' Adds a component image.
Public Sub ReplaceComponentImage(strName As String, strPath As String)
    FileCopy strPath, GetComponentImagePath(strName, vbNullString, False)
End Sub

' Checks if a file exists.
Private Function FileExists(strPath As String) As Boolean
    If frmMain.fsMain.Dir(strPath) <> vbNullString Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

' Deletes a file to the recycling bin.
Private Sub DeleteFile(strPath As String)
    MsgBox "File deletion", vbOKOnly + vbInformation, "To Be Implemented"
End Sub
