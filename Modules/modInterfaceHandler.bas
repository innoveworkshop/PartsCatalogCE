Attribute VB_Name = "modInterfaceHandler"
''' modInterfaceHandler
''' A module to bridge the user interface gap between the Pocket PC and
''' Handheld PC versions of the app.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Category selection changed.
Public Sub CategorySelectionChange(cmbCategories As ComboBox, cmbSubCategories As ComboBox)
    LoadSubCategories cmbCategories.ItemData(cmbCategories.ListIndex), _
        cmbSubCategories, True
    cmbSubCategories.ListIndex = 0
End Sub

' Sub-category selection changed.
Public Sub SubCategorySelectionChange(cmbCategories As ComboBox, _
        cmbSubCategories As ComboBox, lstComponents As ListBox)
    LoadComponents cmbCategories.ItemData(cmbCategories.ListIndex), _
        cmbSubCategories.ItemData(cmbSubCategories.ListIndex), lstComponents, True
End Sub

' Component selection changed.
Public Sub ComponentSelectionChange(lstComponents As ListBox)
    ' Check if there's anything selected.
    If lstComponents.ListIndex < 0 Then
        Exit Sub
    End If
    
    ' Set component and show the dialog.
    frmComponent.SetComponentID lstComponents.ItemData(lstComponents.ListIndex)
    frmComponent.Show
End Sub
