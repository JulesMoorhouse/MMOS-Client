Attribute VB_Name = "modMmos"
Option Explicit
Global gbooQAOK As Boolean
Global gbooFreeOrder As Boolean
Sub Main()
Dim lstrErrPosition As String
Dim lstrThisHelpFile As String

    On Error GoTo ErrHandler
  
    gdatSystemStartTime = Now()
    
    gstrSystemRoute = srStandardRoute
    
    lstrThisHelpFile = MCLDebugChoices
        
    SetSystemNames
    
    frmSplash.Show
    frmSplash.Refresh

    lstrErrPosition = "One"
    If InStr(UCase(Command$), "/X") = 0 Then
        MsgBox "You cannot run the program from here!" & vbCrLf & _
            "You must use the Loader program!", , gconstrTitlPrefix & "Startup"
        Unhook
        End
    End If

    Select Case CheckForOtherMMosprog(frmSplash.hwnd)
    Case True
        MsgBox "You may only run one " & gconstrProductFullName & " program at once!", , gconstrTitlPrefix & "Startup"
        Unhook
        End
    Case False
        'MsgBox "no other prog found!"
    End Select
    
    lstrErrPosition = "Two"
    lstrErrPosition = "Three"
    
    InitDb
    lstrErrPosition = "Four"
        
    frmLogin.Show vbModal
    
    lstrErrPosition = "Five"
    If Not frmLogin.OK Then
        gdatCentralDatabase.Close
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing
        Set gdatCentralDatabase = Nothing
        UpdateLoader
        
        Unhook
        End
    End If
    Busy True
   
    ConcurrencyTest
    
    UpdateLists frmSplash, True
    
    UserLicenceCheck
        
    gbooJustPreLoading = True
    
    DoEvents
    
    Unload frmLogin
    DoEvents
    Load frmAccount
    Unload frmAccount
    Load frmCheque
    Unload frmCheque
    Load frmChildCalendar
    Unload frmChildCalendar
    Load frmChildCashbook
    Unload frmChildCashbook
    Load frmChildCuNotes
    Unload frmChildCuNotes
    Load frmChildGenericDropdown
    Unload frmChildGenericDropdown
    Load frmChildNote
    Unload frmChildNote
    Load frmChildOptions
    Unload frmChildOptions
    Load frmChildPrinter
    Unload frmChildPrinter
    Load frmChildProducts
    Unload frmChildProducts
    Load frmCustAcctSel
    Unload frmCustAcctSel
    Load frmMakeIt
    Unload frmMakeIt
    Load frmOrdDetails
    Unload frmOrdDetails
    Load frmOrder
    Unload frmOrder
    Load frmOrdHistory
    Unload frmOrdHistory
    Load frmPackaging
    Unload frmPackaging
    Load frmPrintPreview
    Unload frmPrintPreview
    Load frmQAMisc
    Unload frmQAMisc
    Load frmReportOptions
    Unload frmReportOptions
    
    gbooJustPreLoading = False
    gbooFreeOrder = False
    
    CopyHelpFile lstrThisHelpFile
    Unload frmSplash

    gstrVATRate = gstrReferenceInfo.strVATRate175
    
   'DEVNOTE: 2019 - Removed requirement
'    If gstrGenSysInfo.lngUserLevel = 20 Or _
'        gstrGenSysInfo.lngUserLevel = 50 Then
'
'        gbooQAOK = EstablishQuickAddress
'    Else
    gbooQAOK = False
'    End If
    
    frmAbout.Show
    
    Busy False
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "Main", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub ShowPOCollectForm()

    frmChildPostOfficeCol.Show vbModal
    
End Sub
