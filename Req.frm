VERSION 5.00
Begin VB.Form frmReq 
   Caption         =   "Pre Registrataion"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCompanyName 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   10
      Text            =   "MINDWARP CONSULTANCY LTD"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox txtCompanyTelephoneNum 
      Height          =   285
      Left            =   1500
      MaxLength       =   14
      TabIndex        =   9
      Text            =   "0123 456789"
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1500
      MaxLength       =   40
      TabIndex        =   8
      Text            =   "JULIAN MOORHOUSE"
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox txtKeyCode 
      Height          =   285
      Left            =   1500
      TabIndex        =   7
      Top             =   3000
      Width           =   2115
   End
   Begin VB.CommandButton cmValidate 
      Caption         =   "&Validate"
      Height          =   495
      Left            =   1500
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modem Fax"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate Request Form"
      Height          =   495
      Left            =   2340
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Company name :"
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "generate validate code in form and in text box for use over phone"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Once we have received payment, we will send you a copy of the software and your user license number."
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   6375
   End
   Begin VB.Label Label4 
      Caption         =   "Currently this software is only intended for the UK market."
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   1560
      Width           =   4035
   End
   Begin VB.Label Label3 
      Caption         =   $"Req.frx":0000
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6315
   End
End
Attribute VB_Name = "frmReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmValidate_Click()

    If Len(txtName) < 5 Then
        MsgBox "You must enter a longer name"
        Exit Sub
    End If
    
    If Len(txtCompanyTelephoneNum) < 5 Then
        MsgBox "You must enter a longer phone number"
        Exit Sub
    End If
    
    If Len(txtCompanyName) < 5 Then
        MsgBox "You must enter a longer company name"
        Exit Sub
    End If
    
    With gstrKey
        .strCompanyName = txtCompanyName
        .strCompanyTelephone = txtCompanyTelephoneNum
        .strCompanyContact = txtName
    
        GenerateKey
        txtKeyCode = .strRetVal
    End With
    
End Sub

Private Sub Form_Load()

    With gstrReferenceInfo
        txtCompanyName = UCase$(Trim$(.strCompanyName))
        txtCompanyTelephoneNum = UCase$(Trim$(.strCompanyTelephone))
        txtName = UCase$(Trim$(.strCompanyContact))
    End With
    
End Sub

Private Sub txtCompanyName_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidAlphaNum(KeyAscii)
    
End Sub

Private Sub txtCompanyName_LostFocus()

    txtCompanyName = UCase$(txtCompanyName)
End Sub

Private Sub txtCompanyTelephoneNum_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiTelNum(KeyAscii)
    
End Sub

Private Sub txtCompanyTelephoneNum_LostFocus()

    txtCompanyTelephoneNum = UCase$(txtCompanyTelephoneNum)

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidAlphaNum(KeyAscii)
    
End Sub

Private Sub txtName_LostFocus()

    txtName = UCase$(txtName)
    
End Sub
