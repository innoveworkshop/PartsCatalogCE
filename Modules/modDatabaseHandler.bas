Attribute VB_Name = "modDatabaseHandler"
''' modDatabaseHandler
''' Handles all of the database operations abstracting this part of the code
''' from the platform specific stuff in the forms.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Public db_strDatabasePath As String
Public db_strWorkspacePath As String
Public db_adoConnection As ADOCE.Connection

' Loads a component by its ID and populates a form.
Public Function LoadComponentDetail(lngID As Long) As Boolean
    Dim rs As ADOCE.Recordset
    Dim strStatement As String
    Set rs = CreateObject("ADOCE.Recordset.3.1")
    
    ' Open the database and query it.
    OpenConnection
    strStatement = "SELECT * FROM Components WHERE ID = " & lngID
    rs.Open strStatement, db_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    If Not rs.EOF Then
        frmComponent.LoadComponent lngID, rs
        LoadComponentDetail = True
    Else
        LoadComponentDetail = False
    End If
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    CloseConnection
End Function

' Loads the categories into a ListBox or ComboBox.
Public Sub LoadCategories(lstBox As Variant, blnCloseExit As Boolean)
    Dim rs As ADOCE.Recordset
    Set rs = CreateObject("ADOCE.Recordset.3.1")
    
    ' Clear the list.
    lstBox.Clear
    
    ' Open the database and query it.
    OpenConnection
    rs.Open "SELECT ID, Name FROM Categories ORDER BY Name ASC", _
        db_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    Do While Not rs.EOF
        lstBox.AddItem rs.Fields("Name")
        lstBox.ItemData(lstBox.NewIndex) = rs.Fields("ID")
        rs.MoveNext
    Loop
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    If blnCloseExit Then
        CloseConnection
    End If
End Sub

' Load sub-categories based on the parent category ID.
Public Sub LoadSubCategories(lngCatID As Long, lstBox As Variant, blnCloseExit As Boolean)
    Dim rs As ADODB.Recordset
    Dim strStatement As String
    Set rs = CreateObject("ADOCE.Recordset.3.1")
    
    ' Clear the list.
    lstBox.Clear
    
    ' Open the database and query it.
    OpenConnection
    strStatement = "SELECT ID, Name FROM SubCategories WHERE ParentID = " & _
        lngCatID & " ORDER BY Name ASC"
    rs.Open strStatement, db_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    Do While Not rs.EOF
        lstBox.AddItem rs.Fields("Name")
        lstBox.ItemData(lstBox.NewIndex) = rs.Fields("ID")
        rs.MoveNext
    Loop
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    If blnCloseExit Then
        CloseConnection
    End If
End Sub

' Load packages.
Public Sub LoadPackages(lstBox As Variant, blnCloseExit As Boolean)
    Dim rs As ADODB.Recordset
    Set rs = CreateObject("ADOCE.Recordset.3.1")
    
    ' Clear the list.
    lstBox.Clear
    
    ' Open the database and query it.
    OpenConnection
    rs.Open "SELECT ID, Name FROM Packages ORDER BY Name ASC", _
        db_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    Do While Not rs.EOF
        lstBox.AddItem rs.Fields("Name")
        lstBox.ItemData(lstBox.NewIndex) = rs.Fields("ID")
        rs.MoveNext
    Loop
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    If blnCloseExit Then
        CloseConnection
    End If
End Sub

' Load components based on their category and sub-category IDs.
Public Sub LoadComponents(lngCatID As Long, lngSubCatID As Long, lstBox As Variant, _
        blnCloseExit As Boolean)
    Dim rs As ADODB.Recordset
    Dim strStatement As String
    Set rs = CreateObject("ADOCE.Recordset.3.1")
    
    ' Clear the list.
    lstBox.Clear
    
    ' Open the database and query it.
    OpenConnection
    strStatement = "SELECT ID, Name FROM Components WHERE CategoryID = " & lngCatID & _
        " AND SubCategoryID = " & lngSubCatID & " ORDER BY Name ASC"
    rs.Open strStatement, db_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    Do While Not rs.EOF
        lstBox.AddItem rs.Fields("Name")
        lstBox.ItemData(lstBox.NewIndex) = rs.Fields("ID")
        rs.MoveNext
    Loop
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    If blnCloseExit Then
        CloseConnection
    End If
End Sub

' Checks if a database is associated.
Public Function IsDatabaseAssociated() As Boolean
    If db_strDatabasePath <> vbNullString Then
        IsDatabaseAssociated = True
    Else
        IsDatabaseAssociated = False
    End If
End Function

' Clears any database association.
Public Sub ClearDatabasePath()
    db_strDatabasePath = vbNullString
    db_strWorkspacePath = vbNullString
End Sub

' Sets the database path.
Public Sub SetDatabasePath(strPath As String)
    db_strDatabasePath = strPath
    db_strWorkspacePath = Left(strPath, InStrRev(strPath, "\"))
End Sub

' Gets the workspace path.
Public Function GetWorkspacePath() As String
    GetWorkspacePath = db_strWorkspacePath
End Function

' Opens a predefined database connection.
Private Sub OpenConnection()
    ' Check if there's a database associated.
    If Not IsDatabaseAssociated Then
        MsgBox "Can't open a connection to the database because there isn't one associated.", _
            vbOKOnly + vbCritical, "Database Connection Error"
    End If
    
    ' Setup connection.
    Set db_adoConnection = CreateObject("ADOCE.Connection.3.1")
    db_adoConnection.ConnectionString = "Data Source = " & db_strDatabasePath
    
    ' Open it.
    db_adoConnection.Open
End Sub

' Closes the default database connection.
Private Sub CloseConnection()
    db_adoConnection.Close
    Set db_adoConnection = Nothing
End Sub
