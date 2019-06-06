VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildPostOfficeCol 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Post Office Collection"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1305
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details"
      Height          =   360
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1305
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   360
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1305
   End
   Begin VB.ListBox lstPostOffices 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
   Begin VB.TextBox txtSearchCriteria 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2412
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1305
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   7
      Top             =   3225
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3731
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "10/07/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "08:23"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFoundNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Found 0 records"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter element of post code :-"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmChildPostOfficeCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrPostOfficeCode() As String

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdDetails_Click()
Dim lstrArray() As String
Dim lstrTimes As String
Dim lintArrInc As Integer
Dim lstrTop As String
Dim lstrPOName As String

    If lstPostOffices.ListIndex = -1 Then
        MsgBox "You must first find a post office!", vbInformation, gconstrTitlPrefix & "Post Office Collection"
        Exit Sub
    End If
    
    GetPostOffice lstrPostOfficeCode(lstPostOffices.ListIndex), lstrArray(), lstrPOName
    lstrTimes = GetPOOpeningTimes(lstrPostOfficeCode(lstPostOffices.ListIndex))

    For lintArrInc = 0 To UBound(lstrArray)
        If Trim$(lstrArray(lintArrInc)) <> "" Then
            lstrTop = lstrTop & lstrArray(lintArrInc) & vbCrLf
        End If
    Next lintArrInc
    
    MsgBox "Office: " & lstrPOName & vbCrLf & vbCrLf & lstrTop & lstrTimes, vbInformation, gconstrTitlPrefix & "Post Office Details"
    
End Sub

Private Sub cmdFind_Click()

    If Trim$(txtSearchCriteria) = "" Then
        MsgBox "You must enter a the first part of the post code e.g. DH8", , gconstrTitlPrefix & "Post Office Collection"
        Exit Sub
    End If
    
    Busy True
    
    FillGenericList lstPostOffices, lstrPostOfficeCode(), _
        "SELECT trim$([Name]) & ', ' & trim$([Add1]) & ', ' & trim$([P_Code]) AS PostOffice, " & _
        "* From " & gtblPADOffice & " WHERE P_Code_S='" & _
        txtSearchCriteria & "' and Org_Unit_Code <> '';", _
        "Org_Unit_Code", "PostOffice", False, "LOCAL"
        
    If lstPostOffices.ListCount <> 0 Then
        lstPostOffices.Selected(0) = True
    End If
    
    lblFoundNumber = "Found " & UBound(lstrPostOfficeCode) & " records."
    Busy False

End Sub

Private Sub cmdSelect_Click()
Dim lstrArray() As String
Dim lstrPOName As String

    If lstPostOffices.ListIndex = -1 Then
        MsgBox "You must first find a post office!", vbInformation, gconstrTitlPrefix & "Post Office Collection"
        Exit Sub
    End If
    
    GetPostOffice lstrPostOfficeCode(lstPostOffices.ListIndex), lstrArray(), lstrPOName
    
    If Len(Trim(lstrPOName)) > 26 Then
        lstrPOName = Left$(lstrPOName, 26)
    End If
    lstrPOName = lstrPOName & " P/O"
    
    With frmAccount
        .txtDeliverAddress1 = lstrPOName '"Post Office Local Collect"
        .txtDeliverAddress2 = lstrArray(0)
        .txtDeliverAddress3 = lstrArray(1)
        .txtDeliverAddress4 = lstrArray(3)
        .txtDeliverAddress5 = lstrArray(4)
        .txtDeliverPostcode = lstrArray(5)
    End With
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
End Sub

Sub GetPostOffice(pstrOrgUnitCode As String, pstrArray() As String, Optional pstrPOName As Variant)
Dim lsnaLists As Recordset
Dim lstrSQL As String

    On Error GoTo ErrHandler
    
    ReDim pstrArray(5)
    
    lstrSQL = "SELECT * From " & gtblPADOffice & " WHERE (((Org_Unit_Code)='" & Trim$(pstrOrgUnitCode) & "'));"
        
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not .EOF Then
            If Not IsMissing(pstrPOName) Then
                pstrPOName = Trim$(.Fields("Name"))
            End If
            pstrArray(0) = Trim$(.Fields("Add1"))
            pstrArray(1) = Trim$(.Fields("Add2"))
            pstrArray(2) = Trim$(.Fields("Add3"))
            pstrArray(3) = Trim$(.Fields("Add4"))
            pstrArray(4) = Trim$(.Fields("Add5"))
            pstrArray(5) = Trim$(.Fields("P_Code"))
        End If
    End With
        
        
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetPostOffice", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub

Private Sub txtSearchCriteria_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cmdFind_Click
    End If
    
End Sub
Function GetPOOpeningTimes(pstrOrgUnitCode As String) As String
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim lstrTimes(5) As String
Dim lstrOpens As String
Dim lintArrInc As Integer

    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT * from " & gtblPADOpeningTimes & _
        " WHERE (((Org_Unit_Code)='" & Trim$(pstrOrgUnitCode) & "'));"
        
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        Do Until .EOF
            lstrOpens = vbTab & Trim$(.Fields("From")) & " " & _
                Trim$(.Fields("To")) & " " & Trim$(.Fields("Lunch_From")) & " " & _
                Trim$(.Fields("Lunch_To")) & vbCrLf
                
            Select Case UCase$(Trim$(.Fields("Weekday")))
            Case "MONDAY"
                lstrTimes(0) = "Monday    " & lstrOpens
            Case "TUESDAY"
                lstrTimes(1) = "Tuesday   " & lstrOpens
            Case "WEDNESDAY"
                lstrTimes(2) = "Wednesday " & lstrOpens
            Case "THURSDAY"
                lstrTimes(3) = "Thursday  " & lstrOpens
            Case "FRIDAY"
                lstrTimes(4) = "Friday       " & lstrOpens
            Case "SATURDAY"
                lstrTimes(5) = "Saturday  " & lstrOpens
            End Select
            .MoveNext
        Loop
    End With
                    
    lsnaLists.Close
    
    GetPOOpeningTimes = vbTab & vbTab & " Normal        Lunch" & vbCrLf
    GetPOOpeningTimes = GetPOOpeningTimes & vbTab & vbTab & "From To       From To" & vbCrLf
    GetPOOpeningTimes = GetPOOpeningTimes & vbTab & vbTab & "==== ====     ==== ====" & vbCrLf
                
    For lintArrInc = 0 To 5
        GetPOOpeningTimes = GetPOOpeningTimes & lstrTimes(lintArrInc)
    Next lintArrInc
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetPOOpeningTimes", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
