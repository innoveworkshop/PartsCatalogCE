Attribute VB_Name = "modInterfaceHandler"
''' modInterfaceHandler
''' A module to bridge the user interface gap between the Pocket PC and
''' Handheld PC versions of the app.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Displays an error message.
Public Sub DisplayError(strMessage As String)
    MsgBox strMessage & vbCrLf & cbCrLf & "(" & Err.Number & ") " & Err.Description, _
        vbOKOnly + vbExclamation, "Error"
End Sub

' Category selection changed.
Public Sub CategorySelectionChange(cmbCategories As ComboBox, cmbSubCategories As ComboBox, _
        blnCloseOnExit As Boolean)
    LoadSubCategories cmbCategories.ItemData(cmbCategories.ListIndex), _
        cmbSubCategories, blnCloseOnExit
    cmbSubCategories.ListIndex = 0
End Sub

' Sub-category selection changed.
Public Sub SubCategorySelectionChange(cmbCategories As ComboBox, _
        cmbSubCategories As ComboBox, lstComponents As ListBox, blnCloseOnExit As Boolean)
    LoadComponents cmbCategories.ItemData(cmbCategories.ListIndex), _
        cmbSubCategories.ItemData(cmbSubCategories.ListIndex), lstComponents, blnCloseOnExit
End Sub

' Component selection changed.
Public Sub ComponentSelectionChange(lstComponents As ListBox)
    Dim blnShow As Boolean
    Dim lngID As Long
    
    ' Check if there's anything selected.
    If lstComponents.ListIndex < 0 Then
        Exit Sub
    End If
    
    ' Load component.
    lngID = lstComponents.ItemData(lstComponents.ListIndex)
    blnShow = LoadComponentDetail(lngID)
    
    ' Show the dialog.
    If blnShow Then
        frmComponent.Show
    Else
        MsgBox "Couldn't find any component with the ID of " & lngID, _
            vbOKOnly + vbCritical, "Component Loading Error"
    End If
End Sub

' Populates the component view.
Public Sub PopulateComponentView(rs As ADOCE.Recordset, txtName As TextBox, _
        txtQuantity As TextBox, cmbCategory As ComboBox, cmbSubCategory As ComboBox, _
        cmbPackage As ComboBox, txtNotes As TextBox, grdProperties As GridCtrl)
    Dim intIndex As Integer

    ' Populate text fields.
    frmComponent.SetOriginalName rs.Fields("Name")
    txtName.Text = rs.Fields("Name")
    txtQuantity.Text = rs.Fields("Quantity")
    txtNotes.Text = rs.Fields("Notes")
    
    ' Set the categories.
    cmbSubCategory.Clear
    LoadCategories cmbCategory, False
    If rs.Fields("CategoryID") >= 0 Then
        For intIndex = 0 To cmbCategory.ListCount
            If cmbCategory.ItemData(intIndex) = rs.Fields("CategoryID") Then
                cmbCategory.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    
    ' Load the sub-categories.
    LoadSubCategories rs.Fields("CategoryID"), cmbSubCategory, False
    If rs.Fields("SubCategoryID") >= 0 Then
        For intIndex = 0 To cmbSubCategory.ListCount
            If cmbSubCategory.ItemData(intIndex) = rs.Fields("SubCategoryID") Then
                cmbSubCategory.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    
    ' Set the packages.
    LoadPackages cmbPackage, False
    If rs.Fields("PackageID") >= 0 Then
        For intIndex = 0 To cmbPackage.ListCount
            If cmbPackage.ItemData(intIndex) = rs.Fields("PackageID") Then
                cmbPackage.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    
    ' Populate the properties grid.
    If Not IsNull(rs.Fields("Properties")) Then
        PopulatePropertiesGrid grdProperties, rs.Fields("Properties")
    Else
        grdProperties.Rows = 1
        grdProperties.TextMatrix(0, 0) = ""
        grdProperties.TextMatrix(0, 1) = ""
    End If
End Sub

' Populates the properties grid.
Public Sub PopulatePropertiesGrid(grdProperties As GridCtrl, strProperties As String)
    Dim astrProperties As String
    Dim astrKeyValue As String
    
    ' Split the properties and preparate the grid for the properties.
    astrProperties = Split(strProperties, vbTab)
    If UBound(astrProperties) = 0 Then
        grdProperties.Rows = UBound(astrProperties) + 1
    Else
        grdProperties.Rows = UBound(astrProperties) + 2
    End If
    
    ' Populate the properties.
    Dim intIndex As Integer
    For intIndex = 0 To UBound(astrProperties)
        ' Check if the property is populated.
        If astrProperties(intIndex) <> "" Then
            astrKeyValue = Split(astrProperties(intIndex), ": ")
            grdProperties.TextMatrix(intIndex, 0) = astrKeyValue(0)
            grdProperties.TextMatrix(intIndex, 1) = astrKeyValue(1)
        Else
            grdProperties.TextMatrix(0, 0) = ""
            grdProperties.TextMatrix(0, 1) = ""
        End If
    Next intIndex
End Sub
