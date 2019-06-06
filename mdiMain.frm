VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000F&
   Caption         =   "Mindwarp Mail Order System"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   5400
      Top             =   3960
   End
   Begin VB.PictureBox picListBar 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   8040
      Left            =   0
      ScaleHeight     =   7980
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   0
      Width           =   1395
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   8040
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12753
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "4/16/2019"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:27 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "Print Set&up"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "&1 "
         Index           =   0
         Visible         =   0   'False
         Begin VB.Menu mnuFileHistoryModOrder1 
            Caption         =   "&Modify This Order"
         End
         Begin VB.Menu mnuFileHistoryOrdHistory1 
            Caption         =   "Orders &History"
         End
         Begin VB.Menu mnuFileHistoryPackOrder1 
            Caption         =   "&Pack This Order"
         End
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "&2"
         Index           =   1
         Visible         =   0   'False
         Begin VB.Menu mnuFileHistoryModOrder2 
            Caption         =   "&Modify This Order"
         End
         Begin VB.Menu mnuFileHistoryOrdHistory2 
            Caption         =   "Orders &History"
         End
         Begin VB.Menu mnuFileHistoryPackOrder2 
            Caption         =   "&Pack This Order"
         End
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "&3"
         Index           =   2
         Visible         =   0   'False
         Begin VB.Menu mnuFileHistoryModOrder3 
            Caption         =   "&Modify This Order"
         End
         Begin VB.Menu mnuFileHistoryOrdHistory3 
            Caption         =   "Orders &History"
         End
         Begin VB.Menu mnuFileHistoryPackOrder3 
            Caption         =   "&Pack This Order"
         End
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "&4"
         Index           =   3
         Visible         =   0   'False
         Begin VB.Menu mnuFileHistoryModOrder4 
            Caption         =   "&Modify This Order"
         End
         Begin VB.Menu mnuFileHistoryOrdHistory4 
            Caption         =   "Orders &History"
         End
         Begin VB.Menu mnuFileHistoryPackOrder4 
            Caption         =   "&Pack This Order"
         End
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "&5"
         Index           =   4
         Visible         =   0   'False
         Begin VB.Menu mnuFileHistoryModOrder5 
            Caption         =   "&Modify This Order"
         End
         Begin VB.Menu mnuFileHistoryOrdHistory5 
            Caption         =   "Orders &History"
         End
         Begin VB.Menu mnuFileHistoryPackOrder5 
            Caption         =   "&Pack This Order"
         End
      End
      Begin VB.Menu mnuFileHistorySep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewShowPicBar 
         Caption         =   "Show &Picture Bar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewShowNewFeatures 
         Caption         =   "Show New &Features"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewMaxOnStartup 
         Caption         =   "&Maximize On Startup"
      End
   End
   Begin VB.Menu mnuGo 
      Caption         =   "&Go"
      Begin VB.Menu mnuGoItem1 
         Caption         =   "Item1"
      End
      Begin VB.Menu mnuGoItem2 
         Caption         =   "Item2"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGoItem3 
         Caption         =   "Item3"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGoItem4 
         Caption         =   "Item4"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGoItem5 
         Caption         =   "Item5"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGoItem6 
         Caption         =   "Item6"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsMinder 
         Caption         =   "&Minder Full"
      End
      Begin VB.Menu mnuToolsResetGrid 
         Caption         =   "Reset &Grid(s) Layout"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsConfigureValues 
         Caption         =   "&Configure Values"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsMaintainProducts 
         Caption         =   "Maintain &Products"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsEssentialSettings 
         Caption         =   "Essential &Settings"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsChangePassword 
         Caption         =   "Change Pass&word"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsExternalPrograms 
         Caption         =   "&External Programs"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents and Index	F1"
      End
      Begin VB.Menu mnuHelpWhatsThis 
         Caption         =   "What's This?	Shift + F1"
      End
      Begin VB.Menu mnuHelpTutorial 
         Caption         =   "&Tutorial"
      End
      Begin VB.Menu mnuHelpQuickStart 
         Caption         =   "&Quick Start Sheets"
      End
      Begin VB.Menu mnuHelpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpCFU 
         Caption         =   "Check For &Updates"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lintCurrOrderEntryButton As Integer
Dim lintCurrOrderEnqButton As Integer
Dim lintCurrAcctMaintButton As Integer
Dim lintCurrFinanceButton As Integer
Dim lintCurrPackingButton As Integer
Dim lintCurrOrderMaintButton As Integer

'Scroll Buttons
Private Sub MDIForm_Activate()

    sbStatusBar.Panels(2).Text = gstrGenSysInfo.strUserName
    
End Sub

Private Sub MDIForm_Load()

    MDILoad Me, frmAbout
    
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lintRetVal As Integer
Dim lstrExitMsg As String

    If gintForceAppClose = fcCompleteClose Or gintForceAppClose = fcCloseKeepDB Then
       
    Else
         Select Case gstrButtonRoute
         Case gconstrEntry, gconstrEnquiry, gconstrAccount
             If gstrCurrentLoadedForm.Name <> "frmCustAcctSel" Then
                 lstrExitMsg = "WARNING: closing the system from this screen may result" & vbCrLf & _
                     "in information being lost!" & vbCrLf & vbCrLf
             End If
         
         End Select
        
         lintRetVal = MsgBox(lstrExitMsg & "You are about to logout and close the system! Procced?", _
             vbYesNo + vbDefaultButton1 + vbExclamation, gconstrTitlPrefix & "System Exit")
         
         If lintRetVal = vbNo Then
             Cancel = True
             Exit Sub
         End If
    End If
    
    ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
    
    If gintForceAppClose <> fcCloseKeepDB Then
        Busy True, Me
        gdatCentralDatabase.Close
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing
        Set gdatCentralDatabase = Nothing
    End If
    
    If UCase$(App.ProductName) <> "LITE" Then
        UpdateLoader
    End If
    
    Busy False, Me
    
    If Not DebugVersion Then
        'Stop subclassing.
        Unhook
    End If
    
End Sub

Private Sub MDIForm_Resize()

    If Me.WindowState = vbNormal Then
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2
    End If

End Sub

Private Sub MDIForm_Terminate()

    If Not DebugVersion Then
        'Stop subclassing.
        Unhook
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    Unload frmButtons
    If Not DebugVersion Then
        'Stop subclassing.
        Unhook
    End If
    
End Sub

Private Sub mnuHelpContents_Click()

    StandardMenuOptions mnuHelpContents.Caption
    
End Sub

Private Sub mnuHelpQuickStart_Click()

    StandardMenuOptions mnuHelpQuickStart.Caption
    
End Sub

Private Sub mnuHelpTutorial_Click()

    StandardMenuOptions mnuHelpTutorial.Caption
    
End Sub

Private Sub mnuHelpWhatsThis_Click()

    StandardMenuOptions mnuHelpWhatsThis.Caption
    
End Sub

Private Sub picListBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PicListBarMouseDown Me, Button, Shift, X, Y
    
End Sub
Private Sub picListBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PicListBarMouseMove Me, Button, Shift, X, Y

End Sub
Private Sub picListBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    gbooUIScrollButtonClicked = False
    
End Sub
Sub ButtonSelected(pintButtonIndex As Integer)
Dim lintRetVal As Integer

    Unload frmAbout

    Select Case gstrButtonRoute
    Case gconstrMainMenu
        UnloadLastForm
        Select Case pintButtonIndex
        Case lintCurrOrderEntryButton
            gstrButtonRoute = gconstrEntry
            Set gstrCurrentLoadedForm = frmCustAcctSel
            frmCustAcctSel.Route = gconstrEntry
            DrawButtonSet gstrButtonRoute
            frmCustAcctSel.Show

        Case lintCurrOrderEnqButton
            gstrButtonRoute = gconstrEnquiry
            Set gstrCurrentLoadedForm = frmCustAcctSel
            frmCustAcctSel.Route = gconstrEnquiry
            DrawButtonSet gstrButtonRoute
            frmCustAcctSel.Show

        Case lintCurrAcctMaintButton
            gstrButtonRoute = gconstrAccount
            Set gstrCurrentLoadedForm = frmCustAcctSel
            frmCustAcctSel.Route = gconstrAccount
            DrawButtonSet gstrButtonRoute
            frmCustAcctSel.Show

        Case lintCurrFinanceButton
            gstrButtonRoute = gconstrFinance
            Set gstrCurrentLoadedForm = frmCheque
            frmCheque.Route = gconstrFinance
            DrawButtonSet gstrButtonRoute
            frmCheque.Show

        Case lintCurrPackingButton
            gstrButtonRoute = gconstrPacking
            Set gstrCurrentLoadedForm = frmPackaging
            frmPackaging.Route = gconstrPacking
            DrawButtonSet gstrButtonRoute
            frmPackaging.Show

        Case lintCurrOrderMaintButton
            gstrButtonRoute = gconstrOrdMaint
            Set gstrCurrentLoadedForm = frmQAMisc
            frmQAMisc.Route = gconstrOrdMaint
            DrawButtonSet gstrButtonRoute
            frmQAMisc.Show

        End Select
    Case gconstrEntry
        'Need to work savelocalfields etc if previous form unloaded
        'UnloadLastForm
        Select Case pintButtonIndex
        Case 0 'frmCustAcctSel
            Select Case gstrCurrentLoadedForm.Name
            Case "frmCustAcctSel"
                'Do Nothing
            Case Else
                'loose changes and exit process
                lintRetVal = MsgBox("WARNING: By aborting this order no information will be saved!", vbYesNo + vbExclamation, gconstrTitlPrefix & "Abort Order")
                If lintRetVal <> vbYes Then
                    Exit Sub
                End If
                UnloadLastForm
                Set gstrCurrentLoadedForm = frmCustAcctSel
                'Clear functions here
                ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
                ClearCustomerAcount
                ClearAdviceNote
                ClearGen
                frmCustAcctSel.Route = gconstrEntry
                frmCustAcctSel.Show
            End Select

        Case 1 ' frmAccount selected
            'Decide what to do when access from a certain form
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                MsgBox "You must select a customer account!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmAccount"
                'Do nothing!
            Case "frmOrdDetails"
                'Back
                frmOrdDetails.SaveLocalFields
                Unload frmOrdDetails
                UnloadLastForm
                Set gstrCurrentLoadedForm = frmAccount
                frmAccount.Show
            Case "frmOrder"
                'Back
                frmOrder.SaveLocalFields
                Unload frmOrder
                Set gstrCurrentLoadedForm = frmAccount
                frmAccount.Show
            End Select

        Case 2 ' frmOrderDetails selected
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                MsgBox "You must select a customer account!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmAccount"
                'Forward
                Call frmAccount.cmdNext_Click
            Case "frmOrdDetails"
                'Do Nothing
            Case "frmOrder"
                'Back
                frmOrder.SaveLocalFields
                Unload frmOrder
                Set gstrCurrentLoadedForm = frmOrdDetails
                frmOrdDetails.Show
            End Select

        Case 3 ' frmOrder Selected
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                MsgBox "You must select a customer account!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmAccount"
                'Forward
                If frmAccount.FindFieldError = True Then Exit Sub
                frmAccount.SaveLocalFields
                If FindOrderDetailFieldError = True Then
                    MsgBox "You must first satisfy requirments on the Order Details screen!", , gconstrTitlPrefix & "Sub Screen Selection"
                    Exit Sub
                End If
                Unload frmAccount
                Set gstrCurrentLoadedForm = frmOrder
                frmOrder.Show
            Case "frmOrdDetails"
                'Forward
                Call frmOrdDetails.cmdNext_Click
            Case "frmOrder"
                'Do Nothing
            End Select

        Case 4 'return to menu
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                UnloadLastForm
                frmAbout.Show
            Case Else
                'loose changes and return to menu
                lintRetVal = MsgBox("WARNING: Do you wish to return to the main menu and loose any changes " & vbCrLf & _
                    "you may have made?", vbYesNo + vbExclamation, gconstrTitlPrefix & "Sub Screen Selection")
                If lintRetVal <> vbYes Then
                    Exit Sub
                End If
                'Clear functions here
                ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
                ClearCustomerAcount
                ClearAdviceNote
                ClearGen
                UnloadLastForm
                frmAbout.Show
            End Select
        End Select
        'MsgBox "This feature is not available at this time!", , gconstrTitlPrefix & "Sub Screen Selection"
    Case gconstrEnquiry
        Select Case pintButtonIndex
        Case 0 ' frmCustAcctSel
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                'Do Nothing
            Case Else
                'loose changes and exit process
                lintRetVal = MsgBox("WARNING: Do you wish to stop modifying this order and loose any changes " & vbCrLf & _
                    "you may have made?", vbYesNo + vbExclamation, gconstrTitlPrefix & "Sub Screen Selection")
                If lintRetVal <> vbYes Then
                    Exit Sub
                End If
                UnloadLastForm
                Set gstrCurrentLoadedForm = frmCustAcctSel
                'Clear functions here
                ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
                ClearCustomerAcount
                ClearAdviceNote
                ClearGen
                frmCustAcctSel.Route = gconstrEnquiry
                frmCustAcctSel.Show
            End Select
        Case 1 ' frmOrderHistory selected
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                MsgBox "You must select a customer account!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmOrdHistory"
                'Do Nothing!
            Case Else
                'loose changes and exit modify process
                lintRetVal = MsgBox("WARNING: Do you wish to stop modifying this order and loose any changes " & vbCrLf & _
                    "you may have made?", vbYesNo + vbExclamation, gconstrTitlPrefix & "Sub Screen Selection")
                If lintRetVal <> vbYes Then
                    Exit Sub
                End If
                UnloadLastForm
                ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
                ClearCustomerAcount
                ClearAdviceNote
                ClearGen
                Set gstrCurrentLoadedForm = frmOrdHistory
                frmOrdHistory.Route = gconstrEnquiry
                frmOrdHistory.Show
            End Select
        Case 2 'frmAccount
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                MsgBox "You must select a customer account!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmOrdHistory"
                MsgBox "Must select modify!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmAccount"
                'Do Nothing
            Case "frmOrdDetails"
                frmOrdDetails.SaveLocalFields
                Unload frmOrdDetails
                UnloadLastForm
                Set gstrCurrentLoadedForm = frmAccount
                frmAccount.Route = gconstrOrderModify
                frmAccount.Show
            Case "frmOrder"
                frmOrder.SaveLocalFields
                Unload frmOrder
                Set gstrCurrentLoadedForm = frmAccount
                frmAccount.Route = gconstrOrderModify
                frmAccount.Show
            End Select
        Case 3 'frmOrdDetails
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                MsgBox "You must select a customer account!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmOrdHistory"
                MsgBox "Must select modify!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmAccount"
                Call frmAccount.cmdNext_Click
            Case "frmOrdDetails"
                'Do Nothing
            Case "frmOrder"
                frmOrder.SaveLocalFields
                Unload frmOrder
                Set gstrCurrentLoadedForm = frmOrdDetails
                frmOrdDetails.Route = gconstrOrderModify
                frmOrdDetails.Show
            End Select
        Case 4 ' frmOrder
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                MsgBox "You must select a customer account!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmOrdHistory"
                MsgBox "Must select modify!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmAccount"
                If frmAccount.FindFieldError = True Then Exit Sub
                frmAccount.SaveLocalFields
                If FindOrderDetailFieldError = True Then
                    MsgBox "You must first satisfy requirments on the Order Details screen!", , gconstrTitlPrefix & "Sub Screen Selection"
                    Exit Sub
                End If
                Unload frmAccount
                Set gstrCurrentLoadedForm = frmOrder
                frmOrder.Route = gconstrOrderModify
                frmOrder.Show
            Case "frmOrdDetails"
                Call frmOrdDetails.cmdNext_Click
            Case "frmOrder"
                'Do Nothing
            End Select
        Case 5 'return to menu
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                UnloadLastForm
                frmAbout.Show
            Case Else
                Call frmOrdHistory.cmdBack_Click
            End Select
        End Select
    Case gconstrAccount
        Select Case pintButtonIndex
        Case 0 'frmCustAcctSel
            Select Case gstrCurrentLoadedForm.Name
            Case "frmCustAcctSel"
                'Do nothing!
            Case "frmAccount"
                'loose changes and exit process
                lintRetVal = MsgBox("WARNING: Do you wish to stop modifying this order and loose any changes " & vbCrLf & _
                    "you may have made?", vbYesNo + vbExclamation, gconstrTitlPrefix & "Sub Screen Selection")
                If lintRetVal <> vbYes Then
                    Exit Sub
                End If
                UnloadLastForm
                Set gstrCurrentLoadedForm = frmCustAcctSel
                'Clear functions here
                ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
                ClearCustomerAcount
                ClearGen
                frmCustAcctSel.Route = gconstrAccount
                frmCustAcctSel.Show
            End Select
        Case 1 ' frmAccount
            Select Case gstrCurrentLoadedForm.Name
            Case "frmCustAcctSel"
                MsgBox "You must select a customer account!", , gconstrTitlPrefix & "Sub Screen Selection"
            Case "frmAccount"
                'Do nothing!
            End Select
        Case 2 'return to menu
            Select Case gstrCurrentLoadedForm.Name ' last form
            Case "frmCustAcctSel"
                UnloadLastForm
                frmAbout.Show
            Case Else
                Call frmAccount.cmdNext_Click
            End Select
        End Select
    Case gconstrFinance
        Select Case pintButtonIndex
        Case 1 'return to menu
            Call frmCheque.cmdBack_Click
        End Select
    Case gconstrPacking
        Select Case pintButtonIndex
        Case 1 'return to menu
            Call frmPackaging.cmdBack_Click
        End Select
    Case gconstrOrdMaint
        Select Case pintButtonIndex
        Case 0
            'Do Nothing
        Case 1 'return to menu
            Call frmQAMisc.cmdBack_Click
        End Select
    End Select
    
End Sub
Sub DrawButtonSet(pstrRoute As String, Optional pstrParam As Variant)
Dim llngDownVar As Long

    lintCurrOrderEntryButton = -1
    lintCurrOrderEnqButton = -1
    lintCurrAcctMaintButton = -1
    lintCurrFinanceButton = -1
    lintCurrPackingButton = -1
    lintCurrOrderMaintButton = -1
    
    picListBar.Cls
    
    If IsMissing(pstrParam) Then pstrParam = ""
    
    gstrButtonRoute = pstrRoute
    
    If gstrUILastButtonRoute <> pstrRoute And gstrUILastButtonRoute <> "" Then
        gconUITopPos = gconUIButtonTopPosDefault
    End If
    
    gstrUILastButtonRoute = pstrRoute
    Select Case pstrRoute
    Case gconstrMainMenu
        Select Case gstrGenSysInfo.lngUserLevel
        Case Is < 20 'Distribution
            DrawButton Me, 0, 0, 5, "Packing": lintCurrPackingButton = 0
            gintUINumberofButtonsDraw = 0
        Case Is < 30 'Order Entry
            DrawButton Me, 0, 0, 1, "Order", "Entry":           lintCurrOrderEntryButton = 0
            DrawButton Me, 1, 0, 2, "Order", "Enquiry":         lintCurrOrderEnqButton = 1
            DrawButton Me, 2, 0, 3, "Account", "Maintenence":   lintCurrAcctMaintButton = 2
            DrawButton Me, 3, 0, 4, "Finance":                  lintCurrFinanceButton = 3
            'DrawButton me, 4, 0, 6, "Distribution":             lintCurrDistributionButton = 4
            gintUINumberofButtonsDraw = 3
        Case Is < 40 'Sales
            DrawButton Me, 0, 0, 2, "Order", "Enquiry":         lintCurrOrderEnqButton = 0
            DrawButton Me, 1, 0, 3, "Account", "Maintenence":   lintCurrAcctMaintButton = 1
            gintUINumberofButtonsDraw = 1
        Case Is < 50 'Accounts
            DrawButton Me, 0, 0, 2, "Order", "Enquiry":         lintCurrOrderEnqButton = 0
            DrawButton Me, 1, 0, 3, "Account", "Maintenence":   lintCurrAcctMaintButton = 1
            DrawButton Me, 2, 0, 4, "Finance":                  lintCurrFinanceButton = 2
            gintUINumberofButtonsDraw = 2
        Case Is < 99 ' General Managers
            DrawButton Me, 0, 0, 1, "Order", "Entry":           lintCurrOrderEntryButton = 0
            DrawButton Me, 1, 0, 2, "Order", "Enquiry":         lintCurrOrderEnqButton = 1
            DrawButton Me, 2, 0, 3, "Account", "Maintenence":   lintCurrAcctMaintButton = 2
            DrawButton Me, 3, 0, 4, "Finance":                  lintCurrFinanceButton = 3
            DrawButton Me, 4, 0, 5, "Packing":                  lintCurrPackingButton = 4
            'DrawButton me, 5, 0, 6, "Distribution":             lintCurrDistributionButton = 5
            DrawButton Me, 5, 0, 7, "Order", "Maintenence":     lintCurrOrderMaintButton = 5
            gintUINumberofButtonsDraw = 5
        Case Is < 100 ' Information Systems
            DrawButton Me, 0, 0, 1, "Order", "Entry":           lintCurrOrderEntryButton = 0
            DrawButton Me, 1, 0, 2, "Order", "Enquiry":         lintCurrOrderEnqButton = 1
            DrawButton Me, 2, 0, 3, "Account", "Maintenence":   lintCurrAcctMaintButton = 2
            DrawButton Me, 3, 0, 4, "Finance":                  lintCurrFinanceButton = 3
            DrawButton Me, 4, 0, 5, "Packing":                  lintCurrPackingButton = 4
            'DrawButton me, 5, 0, 6, "Distribution":             lintCurrDistributionButton = 5
            DrawButton Me, 5, 0, 7, "Order", "Maintenence":     lintCurrOrderMaintButton = 5
            gintUINumberofButtonsDraw = 5
        End Select
    Case gconstrEntry
        DrawButton Me, 0, 0, 8, "Customer", "Select"
        DrawButton Me, 1, 0, 3, "Account", "Address"
        DrawButton Me, 2, 0, 10, "Order", "Details"
        DrawButton Me, 3, 0, 1, "Order"
        DrawButton Me, 4, 0, 9, "Back"
        gintUINumberofButtonsDraw = 4
    Case gconstrEnquiry
        DrawButton Me, 0, 0, 8, "Customer", "Select"
        DrawButton Me, 1, 0, 2, "Order", "History"
        DrawButton Me, 2, 0, 3, "Account", "Address"
        DrawButton Me, 3, 0, 10, "Order", "Details"
        DrawButton Me, 4, 0, 1, "Order"
        DrawButton Me, 5, 0, 9, "Back"
        gintUINumberofButtonsDraw = 5
    Case gconstrAccount
        DrawButton Me, 0, 0, 8, "Customer", "Select"
        DrawButton Me, 1, 0, 3, "Account", "Address"
        DrawButton Me, 2, 0, 9, "Back"
        gintUINumberofButtonsDraw = 2
    Case gconstrFinance
        DrawButton Me, 0, 0, 4, "Cash Book"
        DrawButton Me, 1, 0, 9, "Back"
        gintUINumberofButtonsDraw = 1
    Case gconstrPacking
        DrawButton Me, 0, 0, 5, "Packing"
        DrawButton Me, 1, 0, 9, "Back"
        gintUINumberofButtonsDraw = 1
    Case gconstrOrdMaint
        DrawButton Me, 0, 0, 7, "Order", "Maintenence"
        DrawButton Me, 1, 0, 9, "Back"
        gintUINumberofButtonsDraw = 1
    End Select
    
    FinishDrawingButtonSet Me, llngDownVar, pstrParam

End Sub

Private Sub picListBar_Resize()

    gconUITopPos = gconUIButtonTopPosDefault
    DrawButtonSet gstrButtonRoute

End Sub

Private Sub Timer1_Timer()

    CheckActivity

End Sub

Private Sub mnuEditCopy_Click()

    StandardMenuOptions mnuEditCopy.Caption

End Sub

Private Sub mnuEditCut_Click()

    StandardMenuOptions mnuEditCut.Caption

End Sub

Private Sub mnuEditPaste_Click()
    
    StandardMenuOptions mnuEditPaste.Caption

End Sub

Private Sub mnuFileExit_Click()

    StandardMenuOptions mnuFileExit.Caption
    
End Sub

Private Sub mnuFileHistoryModOrder1_Click()

    FileHistOps mnuFileHistoryModOrder1.Caption, 1
    
End Sub

Private Sub mnuFileHistoryModOrder2_Click()

    FileHistOps mnuFileHistoryModOrder2.Caption, 2

End Sub

Private Sub mnuFileHistoryModOrder3_Click()

    FileHistOps mnuFileHistoryModOrder3.Caption, 3

End Sub

Private Sub mnuFileHistoryModOrder4_Click()

    FileHistOps mnuFileHistoryModOrder4.Caption, 4

End Sub

Private Sub mnuFileHistoryModOrder5_Click()

    FileHistOps mnuFileHistoryModOrder5.Caption, 5

End Sub

Private Sub mnuFileHistoryOrdHistory1_Click()

    FileHistOps mnuFileHistoryOrdHistory1.Caption, 1

End Sub

Private Sub mnuFileHistoryOrdHistory2_Click()

    FileHistOps mnuFileHistoryOrdHistory2.Caption, 2

End Sub

Private Sub mnuFileHistoryOrdHistory3_Click()

    FileHistOps mnuFileHistoryOrdHistory3.Caption, 3

End Sub

Private Sub mnuFileHistoryOrdHistory4_Click()

    FileHistOps mnuFileHistoryOrdHistory4.Caption, 4

End Sub

Private Sub mnuFileHistoryOrdHistory5_Click()

    FileHistOps mnuFileHistoryOrdHistory5.Caption, 5

End Sub

Private Sub mnuFileHistoryPackOrder1_Click()

    FileHistOps mnuFileHistoryPackOrder1.Caption, 1

End Sub

Private Sub mnuFileHistoryPackOrder2_Click()

    FileHistOps mnuFileHistoryPackOrder2.Caption, 2
    
End Sub

Private Sub mnuFileHistoryPackOrder3_Click()

    FileHistOps mnuFileHistoryPackOrder3.Caption, 3
    
End Sub

Private Sub mnuFileHistoryPackOrder4_Click()

    FileHistOps mnuFileHistoryPackOrder4.Caption, 4

End Sub

Private Sub mnuFileHistoryPackOrder5_Click()

    FileHistOps mnuFileHistoryPackOrder5.Caption, 5
    
End Sub

Private Sub mnuFilePrintSetup_Click()

    StandardMenuOptions mnuFilePrintSetup.Caption
    
End Sub

Private Sub mnuGoItem1_Click()

    MenuCommands mnuGoItem1.Caption
    
End Sub

Private Sub mnuGoItem2_Click()

    MenuCommands mnuGoItem2.Caption
    
End Sub

Private Sub mnuGoItem3_Click()

    MenuCommands mnuGoItem3.Caption
    
End Sub

Private Sub mnuGoItem4_Click()

    MenuCommands mnuGoItem4.Caption
    
End Sub

Private Sub mnuGoItem5_Click()

    MenuCommands mnuGoItem5.Caption
    
End Sub

Private Sub mnuGoItem6_Click()

    MenuCommands mnuGoItem6.Caption
    
End Sub

Private Sub mnuHelpAbout_Click()

    StandardMenuOptions mnuHelpAbout.Caption
    
End Sub

Private Sub mnuHelpCFU_Click()

    MenuCommands mnuHelpCFU.Caption
    
End Sub

Private Sub mnuToolsChangePassword_Click()

    MenuCommands mnuToolsChangePassword.Caption
    
End Sub

Private Sub mnuToolsConfigureValues_Click()

    MenuCommands mnuToolsConfigureValues.Caption
    
End Sub

Private Sub mnuToolsEssentialSettings_Click()

    MenuCommands mnuToolsEssentialSettings.Caption
    
End Sub

Private Sub mnuToolsExternalPrograms_Click()

    MenuCommands mnuToolsExternalPrograms.Caption
    
End Sub

Private Sub mnuToolsMaintainProducts_Click()

    MenuCommands mnuToolsMaintainProducts.Caption
    
End Sub

Private Sub mnuToolsMinder_Click()

    StandardMenuOptions mnuToolsMinder.Caption
    
End Sub

Private Sub mnuToolsResetGrid_Click()

    StandardMenuOptions mnuToolsResetGrid.Caption
    
End Sub

Private Sub mnuViewMaxOnStartup_Click()

    StandardMenuOptions mnuViewMaxOnStartup.Caption
    
End Sub

Private Sub mnuViewShowNewFeatures_Click()

    StandardMenuOptions mnuViewShowNewFeatures.Caption
    
End Sub

Private Sub mnuViewShowPicBar_Click()

    StandardMenuOptions mnuViewShowPicBar.Caption
    
End Sub

Sub MenuCommands(pstrItem As String)

    Select Case pstrItem
    Case mnuClientGoOrderEntry
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected lintCurrOrderEntryButton
    Case mnuClientGoEnquiry
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected lintCurrOrderEnqButton
    Case mnuClientGoAcctMaint
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected lintCurrAcctMaintButton
    Case mnuClientGoFinance
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected lintCurrFinanceButton
    Case mnuClientGoPacking
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected lintCurrPackingButton
    Case mnuClientGoOrderMaint
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected lintCurrOrderMaintButton
    Case mnuToolsExternalPrograms.Caption
        frmChildToolProgs.Show vbModal
    Case mnuToolsChangePassword.Caption ' not use in lite
        frmChildUserPass.Route = "PASSCHANGE"
        frmChildUserPass.Show vbModal
    End Select

End Sub
Sub FileHistOps(pstrItem As String, plngIndex As Long)
Dim lintRetVal As Integer

    If pstrItem = mnuFileHistoryModOrder1.Caption Then 'Caption not specific to first item
    
        If CheckCanConfirm(mudtUImnuFileHistory(plngIndex - 1).lngOrderNum, "CANCONX") = False Then
            Exit Sub
        End If
        
        lintRetVal = MsgBox("Do you wish to modify Order Number " & _
            mudtUImnuFileHistory(plngIndex - 1).lngOrderNum & _
            " Account Number " & mudtUImnuFileHistory(plngIndex - 1).lngCustNum, vbYesNo, gconstrTitlPrefix & "Modify Order")
            
        RefreshMenu Me
        If lintRetVal = vbYes Then
            gstrCustomerAccount.lngCustNum = mudtUImnuFileHistory(plngIndex - 1).lngCustNum
            gstrAdviceNoteOrder.lngCustNum = mudtUImnuFileHistory(plngIndex - 1).lngCustNum
            gstrAdviceNoteOrder.lngOrderNum = mudtUImnuFileHistory(plngIndex - 1).lngOrderNum
            GetAdviceNote gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum

            gstrOrderEntryOrderStatus = ""

            If gstrReferenceInfo.strDenomination <> gstrAdviceNoteOrder.strDenom Then
                MsgBox "You may not modify this order, as it was entered using a currency that is not in use!", , gconstrTitlPrefix & "Regional Settings"
                ClearCustomerAcount
                ClearAdviceNote
                ClearGen
                RefreshMenu Me
                Exit Sub
            End If
            
            gstrButtonRoute = gconstrEnquiry
            Unload frmAbout
            DrawButtonSet gstrButtonRoute
            Set gstrCurrentLoadedForm = frmAccount
            frmAccount.Route = gconstrOrderModify
            frmAccount.Show
        End If
    End If

    'View Order &History
    If pstrItem = mnuFileHistoryOrdHistory1.Caption Then 'Caption not specific to first item
        GetCustomerAccount mudtUImnuFileHistory(plngIndex - 1).lngCustNum, True
        gstrButtonRoute = gconstrEnquiry
        Unload frmAbout
        DrawButtonSet gstrButtonRoute
        Set gstrCurrentLoadedForm = frmOrdHistory
        frmOrdHistory.Route = gconstrEnquiry
        frmOrdHistory.Show
    End If

    '&Pack This Order
    If pstrItem = mnuFileHistoryPackOrder1.Caption Then 'Caption not specific to first item
        gstrButtonRoute = gconstrPacking
        Set gstrCurrentLoadedForm = frmPackaging
        Unload frmAbout
        frmPackaging.FindOrder = mudtUImnuFileHistory(plngIndex - 1).lngOrderNum
        frmPackaging.Route = gconstrPacking
        DrawButtonSet gstrButtonRoute
        frmPackaging.Show
    End If

End Sub
