' SAP Shipping Automation Tool
' This script automates SAP GUI processes for shipping operations
' Uses pure VBScript with no VBA functions
' Will run continuously until cancelled
' Enhanced to skip invoice validation when CI count = 0
' Enhanced to skip all printing when both PL and CI counts = 0 (BOL only mode)

' Variable declarations
Dim strDeliveryNumber, strWeight, strLength, strWidth, strHeight, strTodayDate, strCommercialInvoice
Dim SapGuiAuto, application, connection, session
Dim blnDataConfirmed, intResponse, strDefault, objShell
Dim intItemCount, i, j, strNodeID, strPaddedDeliveryNum, blnFoundDelivery
Dim strNodeText, strDocType, blnFoundCI
Dim blnContinueAutomation, blnSkipPDFs
Dim strStatusBarText, blnItemsAlreadyPacked
Dim strCurrentUser, objNetwork, strExportPath
Dim strPackingListCount, strCommercialInvoiceCount, strPackingListOutputType
Dim intProcessCount, intMaxProcessesBeforeRefresh

'---------------------------------------------------------------------------
' Memory Management and Resource Cleanup Functions
'---------------------------------------------------------------------------

' Function to clean up COM objects and force garbage collection
Sub CleanupMemory()
    On Error Resume Next
   
    LogMessage "Performing memory cleanup and garbage collection"
   
    ' Force VBScript garbage collection
    If IsObject(CreateObject("ScriptControl")) Then
        Dim objSC
        Set objSC = CreateObject("ScriptControl")
        objSC.Language = "VBScript"
        objSC.ExecuteStatement "CreateObject(""Scripting.Dictionary"")" ' Force object creation/cleanup
        Set objSC = Nothing
    End If
   
    ' Small delay to allow cleanup
    WScript.Sleep 500
   
    ' Log memory cleanup completion
    LogMessage "Memory cleanup completed"
   
    On Error GoTo 0
End Sub

' Function to validate SAP connection and objects
Function ValidateSAPConnection()
    On Error Resume Next
   
    ValidateSAPConnection = False
   
    ' Check if main objects still exist and are responsive
    If Not IsObject(SapGuiAuto) Then
        LogMessage "ERROR: SapGuiAuto object is invalid"
        Exit Function
    End If
   
    If Not IsObject(application) Then
        LogMessage "ERROR: Application object is invalid"
        Exit Function
    End If
   
    If Not IsObject(connection) Then
        LogMessage "ERROR: Connection object is invalid"
        Exit Function
    End If
   
    If Not IsObject(session) Then
        LogMessage "ERROR: Session object is invalid"
        Exit Function
    End If
   
    ' Try to access session properties to ensure it's responsive
    Dim strTest
    strTest = session.Info.SystemName
    If Err.Number <> 0 Then
        LogMessage "ERROR: Session is not responsive - " & Err.Description
        Err.Clear
        Exit Function
    End If
   
    ValidateSAPConnection = True
    LogMessage "SAP connection validation successful"
   
    On Error GoTo 0
End Function

' Function to refresh SAP session and clean up memory
Sub RefreshSAPSession()
    On Error Resume Next
   
    LogMessage "Refreshing SAP session to prevent memory issues"
   
    ' Perform memory cleanup first
    CleanupMemory
   
    ' Try to refresh the session by reconnecting
    If IsObject(session) Then
        ' Navigate to a neutral screen to reset state
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        WScript.Sleep 1000
    End If
   
    ' Re-establish object references if needed
    If Not ValidateSAPConnection() Then
        LogMessage "Re-establishing SAP connection after validation failure"
       
        ' Try to reconnect
        Set SapGuiAuto = GetObject("SAPGUI")
        Set application = SapGuiAuto.GetScriptingEngine
        Set connection = application.Children(0)
        Set session = connection.Children(0)
       
        If IsObject(WScript) Then
            WScript.ConnectObject session, "on"
            WScript.ConnectObject application, "on"
        End If
    End If
   
    LogMessage "SAP session refresh completed"
   
    On Error GoTo 0
End Sub

' Enhanced error handler with memory cleanup
Sub CheckErrorWithCleanup(strStepName)
    If Err.Number <> 0 Then
        LogMessage "ERROR in step: " & strStepName & " - Error #" & Err.Number & ": " & Err.Description
       
        ' Perform memory cleanup before showing error
        CleanupMemory
       
        ' Check if it's a memory-related error
        If InStr(LCase(Err.Description), "memory") > 0 Or _
           InStr(LCase(Err.Description), "resource") > 0 Or _
           Err.Number = -2147024882 Or Err.Number = 462 Then
           
            MsgBox "Memory/Resource Error detected in step: " & strStepName & vbCrLf & _
                   "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
                   "The script will attempt to refresh the SAP session and continue.", _
                   vbExclamation, "Memory Management"
           
            ' Try to refresh the session
            RefreshSAPSession
           
            ' Reset error and continue
            Err.Clear
        Else
            ' Non-memory error - handle normally
            MsgBox "Error in step: " & strStepName & vbCrLf & _
                   "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
                   "Script will be terminated.", vbCritical, "Error"
            WScript.Quit
        End If
    End If
End Sub

'---------------------------------------------------------------------------
' Function for BOL Only Processing (Both PL and CI Count = 0)
'---------------------------------------------------------------------------
Function ProcessDeliveryBOLOnly()
    ' BOL only process when both PL and CI counts are 0 - skip all printing
    On Error Resume Next
   
    LogMessage "Processing delivery for BOL only (both PL and CI count = 0)"
   
    ' Return to delivery change screen with button 18 (Change button)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = strDeliveryNumber
    CheckErrorWithCleanup "Entering delivery number"
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = Len(strDeliveryNumber)
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    CheckErrorWithCleanup "Pressing button 18"
   
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").text = "100"
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").caretPosition = 3
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").setFocus
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").caretPosition = 0
   
    ' Click the "select all" button and check for "no items" message
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT011").press
   
    ' Check status bar for "There are no items that can be selected" message
    If Not CheckStatusBar() Then
        ' User chose to process a different delivery
        LogMessage "User chose to process a different delivery after 'No items' message"
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryBOLOnly = False
        Exit Function
    End If
   
    ' Press the "pack" button
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT001").press
   
    ' Check if serial number popup appears (status bar will be empty)
    strStatusBarText = session.findById("wnd[0]/sbar").text
    If strStatusBarText = "" Then
        ' Serial number popup detected - press check mark button twice to dismiss both dialogs
        LogMessage "Serial number popup detected - automatically handling"
        WScript.Sleep 300
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        WScript.Sleep 300
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
   
    ' Continue with the existing status bar checks for COO or serial number errors
    If Not CheckStatusBar() Then
        ' Critical error requires fixing in SAP first - return to VL02N
        LogMessage "Critical error detected: COO or Serial error. Returning to VL02N."
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryBOLOnly = False
        Exit Function
    End If

    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT008").press
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-GEWEI").text = "lb"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-GEWEI_MAX").text = "lb"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-BRGEW").text = strWeight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-LAENG").text = strLength
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-MEABM").text = "in"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-BREIT").text = strWidth
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").text = strHeight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").setFocus
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").caretPosition = Len(strHeight)
    session.findById("wnd[0]").sendVKey 0
   
    ' Check for any errors after sending the dimensions
    If Not CheckStatusBar() Then
        ' Critical error detected - return to VL02N
        LogMessage "Critical error detected after sending dimensions. Returning to VL02N."
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryBOLOnly = False
        Exit Function
    End If
   
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    ' Skip the printing process entirely and only pack and go to the shipment header details to fill in the BOL only
    session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[3]").select
   
    ' Set commercial invoice to indicate BOL only mode
    strCommercialInvoice = "N/A (BOL Only Mode)"
    LogMessage "Commercial Invoice: " & strCommercialInvoice
   
    ' Display message about BOL only processing
    MsgBox "BOL Only processing complete for delivery " & strDeliveryNumber & vbCrLf & vbCrLf & _
           "Packing completed. You can now proceed to fill in BOL details." & vbCrLf & _
           "No documents will be printed (both PL and CI counts = 0).", vbInformation, strDeliveryNumber & " - BOL Only Mode"
   
    ' No PDF export needed since both counts are 0
    LogMessage "Skipping PDF export - BOL only mode (both PL and CI counts = 0)"
   
    ProcessDeliveryBOLOnly = True
End Function

'---------------------------------------------------------------------------
' Enhanced Function for Packing List Only Processing (PL > 0, CI Count = 0)
'---------------------------------------------------------------------------
Function ProcessDeliveryPackingListOnly()
    ' Streamlined process for packing list only when CI count = 0 but PL count > 0
    On Error Resume Next
   
    LogMessage "Processing delivery for packing list only (PL > 0, CI Count = 0)"
   
    ' Return to delivery change screen with button 18 (Change button)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = strDeliveryNumber
    CheckErrorWithCleanup "Entering delivery number"
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = Len(strDeliveryNumber)
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    CheckErrorWithCleanup "Pressing button 18"
   
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").text = "100"
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").caretPosition = 3
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").setFocus
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").caretPosition = 0
   
    ' Click the "select all" button and check for "no items" message
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT011").press
   
    ' Check status bar for "There are no items that can be selected" message
    If Not CheckStatusBar() Then
        ' User chose to process a different delivery
        LogMessage "User chose to process a different delivery after 'No items' message"
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryPackingListOnly = False
        Exit Function
    End If
   
    ' Press the "pack" button
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT001").press
   
    ' Check if serial number popup appears (status bar will be empty)
    strStatusBarText = session.findById("wnd[0]/sbar").text
    If strStatusBarText = "" Then
        ' Serial number popup detected - press check mark button twice to dismiss both dialogs
        LogMessage "Serial number popup detected - automatically handling"
        WScript.Sleep 300
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        WScript.Sleep 300
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
   
    ' Continue with the existing status bar checks for COO or serial number errors
    If Not CheckStatusBar() Then
        ' Critical error requires fixing in SAP first - return to VL02N
        LogMessage "Critical error detected: COO or Serial error. Returning to VL02N."
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryPackingListOnly = False
        Exit Function
    End If
   
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT008").press
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-GEWEI").text = "lb"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-GEWEI_MAX").text = "lb"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-BRGEW").text = strWeight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-LAENG").text = strLength
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-MEABM").text = "in"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-BREIT").text = strWidth
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").text = strHeight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").setFocus
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").caretPosition = Len(strHeight)
    session.findById("wnd[0]").sendVKey 0
   
    ' Check for any errors after sending the dimensions
    If Not CheckStatusBar() Then
        ' Critical error detected - return to VL02N
        LogMessage "Critical error detected after sending dimensions. Returning to VL02N."
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryPackingListOnly = False
        Exit Function
    End If
   
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[6]").select
   
    ' Select the correct packing list output type based on plant selection
    If SelectOutputTypeRow(strPackingListOutputType) Then
        session.findById("wnd[1]/tbar[0]/btn[6]").press
        session.findById("wnd[2]/usr/txtNAST-ANZAL").text = strPackingListCount
        session.findById("wnd[2]/usr/txtNAST-ANZAL").setFocus
        session.findById("wnd[2]/usr/txtNAST-ANZAL").caretPosition = Len(strPackingListCount)
        session.findById("wnd[2]").sendVKey 0
        session.findById("wnd[1]/tbar[0]/btn[86]").press
    Else
        LogMessage "ERROR: Could not find packing list output type: " & strPackingListOutputType
        MsgBox "Could not find the packing list output type '" & strPackingListOutputType & "'." & vbCrLf & _
               "Please manually select the correct row and continue.", vbExclamation, "Output Type Not Found"
        ProcessDeliveryPackingListOnly = False
        Exit Function
    End If
   
    ' Set commercial invoice to empty since we're not creating one
    strCommercialInvoice = "N/A (Packing List Only)"
    LogMessage "Commercial Invoice: " & strCommercialInvoice
   
    ' Display message about packing list only processing
    MsgBox "Packing List processing complete for delivery " & strDeliveryNumber & vbCrLf & vbCrLf & _
           "No commercial invoice will be generated (count = 0).", vbInformation, strDeliveryNumber & " - Packing List Only"
   
    ' Export PDFs (only packing list, not invoice)
    If ExportPDFs(True) Then  ' Pass True to indicate we're skipping invoice export
        ProcessDeliveryPackingListOnly = True
    Else
        ProcessDeliveryPackingListOnly = False
    End If
End Function

'---------------------------------------------------------------------------
' Function for the N case - no invoice exists
'---------------------------------------------------------------------------
Function ProcessDeliveryN()
    ' This is essentially the same as YY since we need to create a new invoice
    ProcessDeliveryN = ProcessDeliveryYY()
End Function
'---------------------------------------------------------------------------
' Function for the YY case - invoice exists, create new one
'---------------------------------------------------------------------------
Function ProcessDeliveryYY()
    ' Main SAP automation process
    On Error Resume Next
   
    ' Return to delivery change screen with button 18 (Change button)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    CheckErrorWithCleanup "Pressing button 18"
   
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").text = "100" '100
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").caretPosition = 3
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").setFocus
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").caretPosition = 0
   
    ' Click the "select all" button and check for "no items" message
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT011").press
   
    ' Check status bar for "There are no items that can be selected" message
    If Not CheckStatusBar() Then
        ' User chose to process a different delivery
        LogMessage "User chose to process a different delivery after 'No items' message"
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryYY = False
        Exit Function
    End If
   
    ' Press the "pack" button
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT001").press
   
    ' Check if serial number popup appears (status bar will be empty)
    strStatusBarText = session.findById("wnd[0]/sbar").text
    If strStatusBarText = "" Then
        ' Serial number popup detected - press check mark button twice to dismiss both dialogs
        LogMessage "Serial number popup detected - automatically handling"
        WScript.Sleep 300  ' Small delay for stability
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        WScript.Sleep 300  ' Small delay for stability
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
   
    ' Continue with the existing status bar checks for COO or serial number errors
    If Not CheckStatusBar() Then
        ' Critical error requires fixing in SAP first - return to VL02N
        LogMessage "Critical error detected: COO or Serial error. Returning to VL02N."
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryYY = False
        Exit Function
    End If
   
'    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT008").press
   
    ' Clear any existing values and add small delays for data entry stability
''   session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-TARAG").text = "" 'Tare weight
    WScript.Sleep 100
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-GEWEI").text = "lb"
    WScript.Sleep 100
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-GEWEI_MAX").text = "lb"
    WScript.Sleep 100
   
    ' Enter weight with validation
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-BRGEW").text = strWeight
    WScript.Sleep 100
    LogMessage "Entered weight: " & strWeight
   
    ' Enter dimensions with validation
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-LAENG").text = strLength
    WScript.Sleep 100
    LogMessage "Entered length: " & strLength
   
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-MEABM").text = "in"
    WScript.Sleep 100
   
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-BREIT").text = strWidth
    WScript.Sleep 100
    LogMessage "Entered width: " & strWidth
   
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").text = strHeight
    WScript.Sleep 100
    LogMessage "Entered height: " & strHeight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").setFocus
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").caretPosition = Len(strHeight)
    session.findById("wnd[0]").sendVKey 0
   
    ' Check for any errors after sending the dimensions
    If Not CheckStatusBar() Then
        ' Critical error detected - return to VL02N
        LogMessage "Critical error detected after sending dimensions. Returning to VL02N."
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryYY = False
        Exit Function
    End If
   
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[6]").select
   
    ' Select the correct packing list output type based on plant selection
    If SelectOutputTypeRow(strPackingListOutputType) Then
        session.findById("wnd[1]/tbar[0]/btn[6]").press
        session.findById("wnd[2]/usr/txtNAST-ANZAL").text = strPackingListCount
        session.findById("wnd[2]/usr/txtNAST-ANZAL").setFocus
        session.findById("wnd[2]/usr/txtNAST-ANZAL").caretPosition = Len(strPackingListCount)
        session.findById("wnd[2]").sendVKey 0
        session.findById("wnd[1]/tbar[0]/btn[86]").press
    Else
        LogMessage "ERROR: Could not find packing list output type: " & strPackingListOutputType
        MsgBox "Could not find the packing list output type '" & strPackingListOutputType & "'." & vbCrLf & _
               "Please manually select the correct row and continue.", vbExclamation, "Output Type Not Found"
        ProcessDeliveryYY = False
        Exit Function
    End If
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf01"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/cmbRV60A-FKART").key = "ZC1"
    session.findById("wnd[0]/usr/ctxtRV60A-FKDAT").text = strTodayDate
    session.findById("wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,0]").text = strDeliveryNumber
    session.findById("wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,0]").caretPosition = Len(strDeliveryNumber)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf02"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
   
    ' Extract the invoice number from the VF02 screen
    strCommercialInvoice = session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text
    LogMessage "Commercial Invoice Number: " & strCommercialInvoice
    ' Display a message showing the invoice number that was captured
    MsgBox "Commercial Invoice Number: " & strCommercialInvoice, vbInformation, strDeliveryNumber & " - Invoice Generated"
   
    ' Print the invoice
    session.findById("wnd[0]/mbar/menu[0]/menu[11]").select
    session.findById("wnd[1]/tbar[0]/btn[6]").press
    WScript.Sleep 500

    ' Use the enhanced checkbox selection function to ensure the checkbox is selected
    If Not EnsureCheckboxSelected("wnd[2]/usr/chkNAST-DIMME", 5) Then
        LogMessage "WARNING: Failed to select 'Print Immediately' checkbox after multiple attempts"
    End If
   
    ' Set number of commercial invoice copies to user-specified value
    session.findById("wnd[2]/usr/txtNAST-ANZAL").text = strCommercialInvoiceCount
    session.findById("wnd[2]/usr/txtNAST-ANZAL").setFocus
    session.findById("wnd[2]/usr/txtNAST-ANZAL").caretPosition = Len(strCommercialInvoiceCount)
    WScript.Sleep 200  ' Small delay for stability
    session.findById("wnd[2]").sendVKey 0
    session.findById("wnd[1]/tbar[0]/btn[86]").press
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
   
    ' Export PDFs (include both packing list and invoice)
    If ExportPDFs(False) Then  ' Explicitly pass False to export both PL and CI
        ProcessDeliveryYY = True
    Else
        ProcessDeliveryYY = False
    End If
End Function

'---------------------------------------------------------------------------
' Function for the YNY case - invoice exists, don't create new one, but reprint it
'---------------------------------------------------------------------------
Function ProcessDeliveryYNY()
    ' Main SAP automation process for YNY path
    On Error Resume Next
   
    ' Ask user to provide the existing invoice number
    Dim strInvoiceInput
   
    Do
        strInvoiceInput = InputBox("Please enter the existing invoice number for delivery " & strDeliveryNumber & ":", _
                                   "Existing Commercial Invoice for Reprint", "")
       
        If strInvoiceInput = "" Then
            intResponse = MsgBox("No invoice number entered. Would you like to enter one?" & vbCrLf & vbCrLf & _
                                "Click YES to enter an invoice number." & vbCrLf & _
                                "Click NO to return to the main menu.", _
                                vbYesNo + vbQuestion, "Missing Invoice Number")
           
            If intResponse = vbNo Then
                ' User wants to cancel and return to main menu
                ProcessDeliveryYNY = False
                Exit Function
            End If
            ' Otherwise loop continues to prompt again
        Else
            ' Valid input received
            strCommercialInvoice = strInvoiceInput
            Exit Do
        End If
    Loop
   
    ' Return to delivery change screen with button 18 (Change button)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    CheckErrorWithCleanup "Pressing button 18"
   
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").text = "100" '100
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").caretPosition = 3
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").setFocus
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").caretPosition = 0
   
    ' Click the "select all" button and check for "no items" message
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT011").press
   
    ' Check status bar for "There are no items that can be selected" message
    If Not CheckStatusBar() Then
        ' User chose to process a different delivery
        LogMessage "User chose to process a different delivery after 'No items' message"
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryYNY = False
        Exit Function
    End If
   
    ' Press the "pack" button
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT001").press
   
    ' Check if serial number popup appears (status bar will be empty)
    strStatusBarText = session.findById("wnd[0]/sbar").text
    If strStatusBarText = "" Then
        ' Serial number popup detected - press check mark button twice to dismiss both dialogs
        LogMessage "Serial number popup detected - automatically handling"
        WScript.Sleep 300  ' Small delay for stability
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        WScript.Sleep 300  ' Small delay for stability
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
   
    ' Continue with the existing status bar checks for COO or serial number errors
    If Not CheckStatusBar() Then
        ' Critical error requires fixing in SAP first - return to VL02N
        LogMessage "Critical error detected: COO or Serial error. Returning to VL02N."
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryYNY = False
        Exit Function
    End If
   
'    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT008").press
''   session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-TARAG").text = "" 'Tare weight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-GEWEI").text = "lb"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-GEWEI_MAX").text = "lb"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-BRGEW").text = strWeight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-LAENG").text = strLength
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-MEABM").text = "in"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-BREIT").text = strWidth
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").text = strHeight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").setFocus
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").caretPosition = Len(strHeight)
    session.findById("wnd[0]").sendVKey 0
   
    ' Check for any errors after sending the dimensions
    If Not CheckStatusBar() Then
        ' Critical error detected - return to VL02N
        LogMessage "Critical error detected after sending dimensions. Returning to VL02N."
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryYNY = False
        Exit Function
    End If
   
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[6]").select
   
    ' Select the correct packing list output type based on plant selection
    If SelectOutputTypeRow(strPackingListOutputType) Then
        session.findById("wnd[1]/tbar[0]/btn[6]").press
        session.findById("wnd[2]/usr/txtNAST-ANZAL").text = strPackingListCount
        session.findById("wnd[2]/usr/txtNAST-ANZAL").setFocus
        session.findById("wnd[2]/usr/txtNAST-ANZAL").caretPosition = Len(strPackingListCount)
        session.findById("wnd[2]").sendVKey 0
        session.findById("wnd[1]/tbar[0]/btn[86]").press
    Else
        LogMessage "ERROR: Could not find packing list output type: " & strPackingListOutputType
        MsgBox "Could not find the packing list output type '" & strPackingListOutputType & "'." & vbCrLf & _
               "Please manually select the correct row and continue.", vbExclamation, "Output Type Not Found"
        ProcessDeliveryYNY = False
        Exit Function
    End If
   
    LogMessage "Using existing invoice number for reprint: " & strCommercialInvoice
   
    ' Go to VF02 to display the invoice and reprint it
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf02"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
   
    ' Enter the invoice number
    session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text = strCommercialInvoice
    session.findById("wnd[0]/usr/ctxtVBRK-VBELN").caretPosition = Len(strCommercialInvoice)
   
    ' Print the invoice
    session.findById("wnd[0]/mbar/menu[0]/menu[11]").select
    session.findById("wnd[1]/tbar[0]/btn[6]").press
    WScript.Sleep 500

    ' Use the enhanced checkbox selection function to ensure the checkbox is selected
    If Not EnsureCheckboxSelected("wnd[2]/usr/chkNAST-DIMME", 5) Then
        LogMessage "WARNING: Failed to select 'Print Immediately' checkbox after multiple attempts"
    End If
   
    ' Set number of commercial invoice copies to user-specified value
    session.findById("wnd[2]/usr/txtNAST-ANZAL").text = strCommercialInvoiceCount
    session.findById("wnd[2]/usr/txtNAST-ANZAL").setFocus
    session.findById("wnd[2]/usr/txtNAST-ANZAL").caretPosition = Len(strCommercialInvoiceCount)
    WScript.Sleep 200  ' Small delay for stability
    session.findById("wnd[2]").sendVKey 0
    session.findById("wnd[1]/tbar[0]/btn[86]").press
   
    ' Go to VL02N to prepare for PDF export
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
   
    ' Export PDFs (include both packing list and invoice)
    If ExportPDFs(False) Then  ' Explicitly pass False to export both PL and CI
        ProcessDeliveryYNY = True
    Else
        ProcessDeliveryYNY = False
    End If
End Function

'---------------------------------------------------------------------------
' Function for the YNN case - invoice exists, don't create new one, don't reprint
'---------------------------------------------------------------------------
Function ProcessDeliveryYNN()
    ' Main SAP automation process for YNN path
    On Error Resume Next
   
    ' Return to delivery change screen with button 18 (Change button)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    CheckErrorWithCleanup "Pressing button 18"
   
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").text = "100"
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").caretPosition = 3
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").setFocus
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").caretPosition = 0
   
    ' Click the "select all" button and check for "no items" message
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT011").press
   
    ' Check status bar for "There are no items that can be selected" message
    If Not CheckStatusBar() Then
        ' User chose to process a different delivery
        LogMessage "User chose to process a different delivery after 'No items' message"
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryYNN = False
        Exit Function
    End If
   
    ' Press the "pack" button
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT001").press
   
    ' Check if serial number popup appears (status bar will be empty)
    strStatusBarText = session.findById("wnd[0]/sbar").text
    If strStatusBarText = "" Then
        ' Serial number popup detected - press check mark button twice to dismiss both dialogs
        LogMessage "Serial number popup detected - automatically handling"
        WScript.Sleep 300  ' Small delay for stability
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        WScript.Sleep 300  ' Small delay for stability
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
   
    ' Continue with the existing status bar checks for COO or serial number errors
    If Not CheckStatusBar() Then
        ' Critical error requires fixing in SAP first - return to VL02N
        LogMessage "Critical error detected: COO or Serial error. Returning to VL02N."
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryYNN = False
        Exit Function
    End If
   
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT008").press
'   session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-TARAG").text = "" 'Tare weight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-GEWEI").text = "lb"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-GEWEI_MAX").text = "lb"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-BRGEW").text = strWeight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-LAENG").text = strLength
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/ctxtVEKPVB-MEABM").text = "in"
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-BREIT").text = strWidth
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").text = strHeight
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").setFocus
    session.findById("wnd[0]/usr/tabsTS_HU_DET/tabpDETVEKP/ssubTAB:SAPLV51G:6110/txtVEKPVB-HOEHE").caretPosition = Len(strHeight)
    session.findById("wnd[0]").sendVKey 0
   
    ' Check for any errors after sending the dimensions
    If Not CheckStatusBar() Then
        ' Critical error detected - return to VL02N
        LogMessage "Critical error detected after sending dimensions. Returning to VL02N."
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        ProcessDeliveryYNN = False
        Exit Function
    End If
   
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[6]").select
    session.findById("wnd[1]/tbar[0]/btn[6]").press
    session.findById("wnd[2]/usr/txtNAST-ANZAL").text = strPackingListCount
    session.findById("wnd[2]/usr/txtNAST-ANZAL").setFocus
    session.findById("wnd[2]/usr/txtNAST-ANZAL").caretPosition = Len(strPackingListCount)
    session.findById("wnd[2]").sendVKey 0
    session.findById("wnd[1]/tbar[0]/btn[86]").press
   
    ' Export PDFs (only packing list, not invoice)
    If ExportPDFs(True) Then  ' Pass True to indicate we're skipping commercial invoice export
        ProcessDeliveryYNN = True
    Else
        ProcessDeliveryYNN = False
    End If
End Function

' Function to find and select the correct output type row
Function SelectOutputTypeRow(strOutputType)
    On Error Resume Next
    Dim intRow, strCellText, strCellPath, objTable
    Dim blnFound
   
    blnFound = False
    
    LogMessage "Searching for output type: " & strOutputType
    
    ' Get reference to the table object
    Set objTable = session.findById("wnd[1]/usr/tblSAPLVMSGTABCONTROL")
    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not find output type table"
        Err.Clear
        SelectOutputTypeRow = False
        Exit Function
    End If
   
    ' Search through up to 10 rows to find the correct output type
    For intRow = 0 To 9
        strCellPath = "wnd[1]/usr/tblSAPLVMSGTABCONTROL/txtNAST-KSCHL[0," & intRow & "]"
       
        ' Try to get the text from this cell
        strCellText = session.findById(strCellPath).text
       
        ' If we can't find the cell or it's empty, we've reached the end
        If Err.Number <> 0 Or strCellText = "" Then
            Err.Clear
            Exit For
        End If
       
        LogMessage "Row " & intRow & " contains output type: " & strCellText
       
        ' Check if this is the row we're looking for
        If strCellText = strOutputType Then
            ' Found the correct row, select it using the appropriate method
            If intRow = 0 Then
                ' For row 0, use getAbsoluteRow method
                objTable.getAbsoluteRow(0).selected = True
                LogMessage "Selected row 0 using getAbsoluteRow method"
            Else
                ' For other rows, use getAbsoluteRow + setFocus + caretPosition
                objTable.getAbsoluteRow(intRow).selected = True
                WScript.Sleep 100
                session.findById("wnd[1]/usr/tblSAPLVMSGTABCONTROL/txtNAST-KSCHL[0," & intRow & "]").setFocus
                WScript.Sleep 100
                session.findById("wnd[1]/usr/tblSAPLVMSGTABCONTROL/txtNAST-KSCHL[0," & intRow & "]").caretPosition = len(intRow)
                LogMessage "Selected row " & intRow & " using getAbsoluteRow + setFocus method"
            End If
            
            If Err.Number <> 0 Then
                LogMessage "Error selecting row " & intRow & ": " & Err.Description
                Err.Clear
            Else
                LogMessage "Successfully selected row " & intRow & " with output type: " & strOutputType
                blnFound = True
            End If
            Exit For
        End If
    Next
   
    If Not blnFound Then
        LogMessage "WARNING: Could not find output type '" & strOutputType & "' in the available rows"
    End If
   
    SelectOutputTypeRow = blnFound
    On Error GoTo 0
End Function

' SAP automation helper functions
Function ForceWindowFocus(objWindow)
    On Error Resume Next
   
    ' Try multiple focus methods to ensure the window is active
    objWindow.setFocus
    WScript.Sleep 200
   
    ' Get the parent window (if applicable) to ensure proper focus chain
    If TypeName(objWindow.parent) = "GuiFrameWindow" Then
        objWindow.parent.setFocus
        WScript.Sleep 100
        objWindow.setFocus
    End If
   
    On Error GoTo 0
End Function

' Function to safely check if a checkbox was actually selected - Enhanced with memory management
Function EnsureCheckboxSelected(strCheckboxId, intMaxAttempts)
    On Error Resume Next
    Dim i, blnSuccess
   
    ' Set default attempts if not specified
    If intMaxAttempts <= 0 Then intMaxAttempts = 5
   
    ' Try several times to select the checkbox
    blnSuccess = False
    For i = 1 To intMaxAttempts
        ' Validate SAP connection before attempting checkbox operation
        If Not ValidateSAPConnection() Then
            LogMessage "SAP connection validation failed during checkbox operation - attempt " & i
            RefreshSAPSession
            WScript.Sleep 500
        End If
       
        ' First, ensure focus is on the window containing the checkbox
        Dim objWindow
        Set objWindow = session.findById(Left(strCheckboxId, InStr(strCheckboxId, "/usr") - 1))
        ForceWindowFocus objWindow
       
        ' Small delay for UI stability
        WScript.Sleep 300
       
        ' Try to select the checkbox
        session.findById(strCheckboxId).selected = True
        WScript.Sleep 200
       
        ' Verify it was actually selected
        blnSuccess = session.findById(strCheckboxId).selected
       
        ' Clear any error that might have occurred
        If Err.Number <> 0 Then
            LogMessage "Error during checkbox selection attempt " & i & ": " & Err.Description
            Err.Clear
            blnSuccess = False
        End If
       
        If blnSuccess Then
            LogMessage "Checkbox " & strCheckboxId & " successfully selected on attempt " & i
            Exit For
        Else
            LogMessage "Checkbox selection attempt " & i & " for " & strCheckboxId & " failed"
            ' Perform mini cleanup between attempts
            CleanupMemory
            WScript.Sleep 500
        End If
    Next
   
    ' Clean up object reference
    Set objWindow = Nothing
   
    EnsureCheckboxSelected = blnSuccess
    On Error GoTo 0
End Function

' Function to get input with validation and cancel option
Function GetUserInput()
    Dim strInput
   
    ' Reset variables for new run
    strDeliveryNumber = ""
    strWeight = ""
    strLength = ""
    strWidth = ""
    strHeight = ""
    strCommercialInvoice = ""
    strPackingListOutputType = ""
    blnSkipPDFs = False
    blnItemsAlreadyPacked = False
   
    ' Get Delivery Number with validation
    Do
        strInput = InputBox("Enter the Delivery Number:", _
                         "Delivery Information", strDeliveryNumber)
       
        ' Check if user clicked Cancel button (InputBox returns empty string when Cancel is clicked)
        If strInput = "" Then
            ' We need to differentiate between Cancel and empty input with OK
            ' In InputBox, we can't directly distinguish them, so we'll use a simple trick:
            ' If the user enters an empty string and clicks OK, we'll show an error
            ' If the user clicks Cancel, we'll exit the script
           
            ' If we reach here, the user either clicked Cancel or entered empty and clicked OK
            ' Let's ask again with a message box - if they say No, it was likely Cancel
            intResponse = MsgBox("No delivery number entered. Would you like to enter one?", vbOKCancel + vbQuestion, "Missing Input")
            If intResponse = vbCancel Then
                WScript.Quit
            End If
            ' If we get here, the user wants to enter a delivery number, so continue the loop
        Else
            strDeliveryNumber = strInput
            Exit Do
        End If
    Loop
   
    ' Determine the plant type and set packing list output type
    intResponse = MsgBox("Is this an 1812 order?" & vbCrLf & vbCrLf & _
                        "Click YES for 1812" & vbCrLf & _
                        "Click NO for 1814 (VTC)", _
                        vbYesNo + vbQuestion, strDeliveryNumber & " - Plant Selection")
   
    If intResponse = vbYes Then
        ' 1812 plant
        strPackingListOutputType = "ZPL0"
        LogMessage "Plant type selected: 1812 (ZPL0 packing list output type)"
    Else
        ' 1814/VTC plant
        strPackingListOutputType = "YPLA"
        LogMessage "Plant type selected: 1814/VTC (YPLA packing list output type)"
    End If
   
    ' Get Weight with validation
    If strWeight = "" Then
        strDefault = "1"
    Else
        strDefault = strWeight
    End If
   
    Do
        strInput = InputBox("Enter the package weight (lbs):", _
                         "Package Information", strDefault)
       
        ' Check if user clicked Cancel
        If strInput = "" Then
            intResponse = MsgBox("No package weight entered. Would you like to enter one?", vbOKCancel + vbQuestion, strDeliveryNumber & " - Missing Input")
            If intResponse = vbCancel Then
                WScript.Quit
            End If
        Else
            strWeight = strInput
            Exit Do
        End If
    Loop
   
    ' Get Length with validation
    If strLength = "" Then
        strDefault = "13"
    Else
        strDefault = strLength
    End If
   
    Do
        strInput = InputBox("Enter the package length (in):", _
                         "Package Information", strDefault)
       
        ' Check if user clicked Cancel
        If strInput = "" Then
            intResponse = MsgBox("No package length entered. Would you like to enter one?", vbOKCancel + vbQuestion, strDeliveryNumber & " - Missing Input")
            If intResponse = vbCancel Then
                WScript.Quit
            End If
        Else
            strLength = strInput
            Exit Do
        End If
    Loop
   
    ' Get Width with validation
    If strWidth = "" Then
        strDefault = "7"
    Else
        strDefault = strWidth
    End If
   
    Do
        strInput = InputBox("Enter the package width (in):", _
                         "Package Information", strDefault)
       
        ' Check if user clicked Cancel
        If strInput = "" Then
            intResponse = MsgBox("No package width entered. Would you like to enter one?", vbOKCancel + vbQuestion, strDeliveryNumber & " - Missing Input")
            If intResponse = vbCancel Then
                WScript.Quit
            End If
        Else
            strWidth = strInput
            Exit Do
        End If
    Loop
   
    ' Get Height with validation
    If strHeight = "" Then
        strDefault = "7"
    Else
        strDefault = strHeight
    End If
   
    Do
        strInput = InputBox("Enter the package height (in):", _
                         "Package Information", strDefault)
       
        ' Check if user clicked Cancel
        If strInput = "" Then
            intResponse = MsgBox("No package height entered. Would you like to enter one?", vbOKCancel + vbQuestion, strDeliveryNumber & " - Missing Input")
            If intResponse = vbCancel Then
                WScript.Quit
            End If
        Else
            strHeight = strInput
            Exit Do
        End If
    Loop
   
    ' Get Packing List Count with validation
    If strPackingListCount = "" Then
        strDefault = "2"
    Else
        strDefault = strPackingListCount
    End If
   
    Do
        strInput = InputBox("Enter the number of packing lists to print:", _
                         "Document Count", strDefault)
       
        ' Check if user clicked Cancel
        If strInput = "" Then
            intResponse = MsgBox("No packing list count entered. Would you like to enter one?", vbOKCancel + vbQuestion, strDeliveryNumber & " - Missing Input")
            If intResponse = vbCancel Then
                WScript.Quit
            End If
        Else
            strPackingListCount = strInput
            Exit Do
        End If
    Loop
   
    ' Get Commercial Invoice Count with validation
    If strCommercialInvoiceCount = "" Then
        strDefault = "3"
    Else
        strDefault = strCommercialInvoiceCount
    End If
   
    Do
        strInput = InputBox("Enter the number of commercial invoices to print (enter 0 if none):", _
                         "Document Count", strDefault)
       
        ' Check if user clicked Cancel
        If strInput = "" Then
            intResponse = MsgBox("No commercial invoice count entered. Would you like to enter one?", vbOKCancel + vbQuestion, strDeliveryNumber & " - Missing Input")
            If intResponse = vbCancel Then
                WScript.Quit
            End If
        Else
            strCommercialInvoiceCount = strInput
            Exit Do
        End If
    Loop
   
    ' Format today's date in DD.MM.YYYY format
    strTodayDate = Date ' Get current date
    strTodayDate = Right("0" & Day(strTodayDate), 2) & "." & Right("0" & Month(strTodayDate), 2) & "." & Year(strTodayDate)
End Function

' Function to log messages to a file
Sub LogMessage(strMessage)
    On Error Resume Next
   
    Dim objFSO, objLogFile, strLogPath
    strLogPath = "C:\Temp\SAPAutomation.log"
   
    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    ' Create C:\Temp if it doesn't exist
    If Not objFSO.FolderExists("C:\Temp") Then
        objFSO.CreateFolder("C:\Temp")
    End If
   
    ' Create or append to log file
    If objFSO.FileExists(strLogPath) Then
        Set objLogFile = objFSO.OpenTextFile(strLogPath, 8, True) ' 8 = ForAppending
    Else
        Set objLogFile = objFSO.CreateTextFile(strLogPath, True)
    End If
   
    ' Write timestamp and message
    objLogFile.WriteLine Now() & " - " & strMessage
    objLogFile.Close
   
    On Error GoTo 0
End Sub

' Set up an error handler for the entire script
Sub CheckError(strStepName)
    If Err.Number <> 0 Then
        MsgBox "Error in step: " & strStepName & vbCrLf & _
               "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
               "Script will be terminated.", vbCritical, "Error"
        WScript.Quit
    End If
End Sub

' Function to check SAP status bar for error messages
Function CheckStatusBar()
    On Error Resume Next
   
    ' Get the status bar text
    strStatusBarText = session.findById("wnd[0]/sbar").text
   
    ' If status bar has text, check for known error patterns
    If strStatusBarText <> "" Then
        LogMessage "Status bar message: " & strStatusBarText
       
        ' Check for successful packing message
        If InStr(strStatusBarText, "Material was packed") > 0 Then
            ' This is a success message, no need for dialog
            LogMessage "Successful packing detected: " & strStatusBarText
            CheckStatusBar = True  ' Continue normally
            blnItemsAlreadyPacked = False ' Items were just packed, not previously packed
            Exit Function
        End If
       
        ' Check for already packed items message
        If InStr(strStatusBarText, "There are no items that can be selected") > 0 Then
            intResponse = MsgBox("Items are already packed or there are no items to pack." & vbCrLf & vbCrLf & _
                                "Status message: " & strStatusBarText & vbCrLf & vbCrLf & _
                                "Click OK to continue with this delivery, or Cancel to process another delivery.", _
                                vbOKCancel + vbQuestion, strDeliveryNumber & " - Items Already Packed")
           
            If intResponse = vbCancel Then
                CheckStatusBar = False  ' Return to prompt for a new delivery
                Exit Function
            End If
           
            blnItemsAlreadyPacked = True ' Set flag to indicate items are already packed
            CheckStatusBar = True  ' Continue with current delivery
            Exit Function
        End If
       
        ' Check for missing COO message
        If InStr(strStatusBarText, "Please maintain Country of Origin") > 0 Then
            MsgBox "Critical Error: Missing Country of Origin information." & vbCrLf & vbCrLf & _
                   "Status message: " & strStatusBarText & vbCrLf & vbCrLf & _
                   "You need to update the Country of Origin for this part before continuing." & vbCrLf & _
                   "The script will return to the delivery input prompt.", vbExclamation, strDeliveryNumber & " - Missing Country of Origin"
           
            CheckStatusBar = False  ' Return to prompt for a new delivery
            Exit Function
        End If
       
        ' Check for missing serial number message
        If InStr(LCase(strStatusBarText), "serial") > 0 Then
            MsgBox "Critical Error: Missing Serial Number information." & vbCrLf & vbCrLf & _
                   "Status message: " & strStatusBarText & vbCrLf & vbCrLf & _
                   "You need to update the Serial Number for this part before continuing." & vbCrLf & _
                   "The script will return to the delivery input prompt.", vbExclamation, strDeliveryNumber & " - Missing Serial Number"
           
            CheckStatusBar = False  ' Return to prompt for a new delivery
            Exit Function
        End If
    End If
   
    ' No issues found or handled
    CheckStatusBar = True
   
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------
' Enhanced Function to process one delivery with streamlined logic
'---------------------------------------------------------------------------
Function ProcessDelivery()
    ' ENHANCEMENT: Check document counts first to determine routing
    ' Priority: BOL Only > Packing List Only > Full Processing
    
    If strPackingListCount = "0" And strCommercialInvoiceCount = "0" Then
        ' Both counts are 0 - BOL only mode
        LogMessage "Both PL and CI counts are 0 - routing to BOL only process"
        ProcessDelivery = ProcessDeliveryBOLOnly()
        Exit Function
    ElseIf strCommercialInvoiceCount = "0" And strPackingListCount <> "0" Then
        ' CI count is 0 but PL count > 0 - Packing list only mode
        LogMessage "Commercial Invoice count is 0 but PL count > 0 - routing to packing list only process"
        ProcessDelivery = ProcessDeliveryPackingListOnly()
        Exit Function
    End If
    
    ' Original logic for when both counts > 0 (full processing with invoice validation)
    ' Check for existing commercial invoice first
    Dim blnInvoiceExists, blnGenerateNewInvoice, blnReprintExisting
   
    ' Start standard invoice checking process
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = strDeliveryNumber
    CheckErrorWithCleanup "Entering delivery number"

    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = Len(strDeliveryNumber)
    session.findById("wnd[0]/tbar[1]/btn[7]").press ' Document flow button
   
    ' Search for potential invoice entries (starting with "110")
    session.findById("wnd[0]/usr/shell/shellcont[1]/shell[0]").pressButton "&FIND" ' Search button
    session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").text = "110" ' First three numbers of invoice number to search
    session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").caretPosition = 3
    session.findById("wnd[1]/tbar[0]/btn[0]").press ' Check mark button to search the entered string
    session.findById("wnd[1]").close ' To immediately close the pop-up window
   
    ' Ask user if commercial invoice exists
    intResponse = MsgBox("Does a commercial invoice already exist for delivery " & strDeliveryNumber & "?" & vbCrLf & vbCrLf & _
                         "(Check if an invoice line starting with '110' is highlighted in document flow)", _
                         vbYesNo + vbQuestion, "Commercial Invoice Check")
   
    If intResponse = vbYes Then
        ' Invoice exists
        blnInvoiceExists = True
       
        ' Ask if user wants to generate another invoice
        intResponse = MsgBox("Do you want to generate another commercial invoice for this delivery?" & vbCrLf & vbCrLf & _
                            "Click YES to create a new invoice." & vbCrLf & _
                            "Click NO to skip invoice creation and only process PDFs.", _
                            vbYesNo + vbQuestion, "Generate New Invoice?")
       
        If intResponse = vbYes Then
            ' User wants to generate another invoice despite one existing (YY case)
            blnGenerateNewInvoice = True
            LogMessage "User chose to generate new invoice even though one exists (YY path)"
            ProcessDelivery = ProcessDeliveryYY()
        Else
            ' User doesn't want to generate another invoice, ask if they want to reprint the existing one
            blnGenerateNewInvoice = False
           
            intResponse = MsgBox("Would you like to reprint the existing commercial invoice?" & vbCrLf & vbCrLf & _
                                "Click YES to reprint the existing invoice." & vbCrLf & _
                                "Click NO to skip invoice reprinting and only process PDFs.", _
                                vbYesNo + vbQuestion, "Reprint Existing Invoice?")
           
            If intResponse = vbYes Then
                ' User wants to reprint existing invoice (YNY case)
                blnReprintExisting = True
                LogMessage "User chose to reprint existing invoice without generating new one (YNY path)"
                ProcessDelivery = ProcessDeliveryYNY()
            Else
                ' User doesn't want to reprint existing invoice (YNN case)
                blnReprintExisting = False
                LogMessage "User chose not to generate new invoice nor reprint existing one (YNN path)"
                ProcessDelivery = ProcessDeliveryYNN()
            End If
        End If
    Else
        ' No invoice exists (N case)
        blnInvoiceExists = False
        LogMessage "No existing invoice found, proceeding with standard process (N path)"
        ProcessDelivery = ProcessDeliveryN()
    End If
End Function

' Function to handle PDF exports - FIXED with proper Optional parameter
Function ExportPDFs(blnSkipInvoice)
    ' Set default value if parameter is not provided
    If IsEmpty(blnSkipInvoice) Then
        blnSkipInvoice = False
    End If
   
    ' Log parameter value for debugging
    LogMessage "ExportPDFs called with blnSkipInvoice=" & blnSkipInvoice
   
    ' Perform memory cleanup before PDF operations
    CleanupMemory
   
    ' Validate SAP connection before PDF export
    If Not ValidateSAPConnection() Then
        LogMessage "SAP connection validation failed during PDF export - attempting refresh"
        RefreshSAPSession
        If Not ValidateSAPConnection() Then
            LogMessage "Could not restore SAP connection for PDF export"
            ExportPDFs = False
            Exit Function
        End If
    End If
   
    On Error Resume Next
   
    ' Add escape point before PDF export
    intResponse = MsgBox("Ready to export PDFs for delivery " & strDeliveryNumber & "." & vbCrLf & vbCrLf & _
                      "Click OK to export or Cancel to skip.", _
                      vbOKCancel + vbSystemModal + vbQuestion, strDeliveryNumber & " - Continue PDF Export")
                   
    If intResponse = vbCancel Then
        ' Set flag to indicate PDFs were skipped
        blnSkipPDFs = True
       
        ' Display the appropriate completion message
        MsgBox "SAP automation completed successfully!" & vbCrLf & vbCrLf & _
               "Summary:" & vbCrLf & _
                                          "- Delivery: " & strDeliveryNumber & vbCrLf & _
               "- Commercial Invoice: " & strCommercialInvoice & vbCrLf & _
               "- Package: " & strWeight & " lb, " & strLength & " x " & strWidth & " x " & strHeight & " inches" & vbCrLf & _
               "- Packing Lists: " & strPackingListCount & vbCrLf & _
               "- Commercial Invoices: " & strCommercialInvoiceCount & vbCrLf & _
               "- Date: " & strTodayDate & vbCrLf & vbCrLf & _
               "Packing List and Commercial Invoice PDFs were skipped.", vbInformation, strDeliveryNumber & " - Process Complete"
               
        ' Return to VL02N to prepare for next run
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
       
        ' End the function here, returning control to the main loop
        ExportPDFs = True
        Exit Function
    End If
   
    ' Continue with PDF export
    session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[3]").select
    session.findById("wnd[0]/tbar[1]/btn[7]").press
    session.findById("wnd[0]/usr/shell/shellcont[1]/shell[0]").pressButton "&FIND"
    WScript.Sleep 500
    session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").text = strDeliveryNumber
    session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").caretPosition = len(strDeliveryNumber)
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[1]").close
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/titl/shellcont[1]/shell").pressContextButton "%GOS_TOOLBOX"
    session.findById("wnd[0]/titl/shellcont[1]/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
    session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").currentCellColumn = "BITM_DESCR"
    session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").pressToolbarButton "%ATTA_EXPORT"
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = strExportPath
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = strDeliveryNumber & " pl.pdf"
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = Len(strDeliveryNumber) + 7
    session.findById("wnd[2]").sendVKey 0
   
    ' Check if file already exists (status message appears on wnd[0])
    On Error Resume Next
    Dim strFileExistsStatus
    strFileExistsStatus = ""
    WScript.Sleep 500 ' Wait for status message to appear
   
    ' Try to get the status bar text from the main window
    strFileExistsStatus = session.findById("wnd[0]/sbar").text
   
    ' Check if it's a file exists message
    If InStr(strFileExistsStatus, "already exists") > 0 Then
        ' File already exists, ask user what to do
        intResponse = MsgBox("The packing list PDF (" & strDeliveryNumber & " pl.pdf) already exists." & vbCrLf & vbCrLf & _
                           "Click OK to replace the existing file." & vbCrLf & _
                           "Click Cancel to skip generating this document.", _
                           vbOKCancel + vbQuestion, "File Already Exists")
           
        If intResponse = vbOK Then
            ' User chose to replace the file
            session.findById("wnd[2]/tbar[0]/btn[11]").press ' Press Replace button
            LogMessage "User chose to replace existing packing list PDF"
        Else
            ' User chose to skip this file
            session.findById("wnd[2]").close ' Close the dialog
            LogMessage "User chose to skip generating packing list PDF"
        End If
    End If
    On Error GoTo 0
   
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]").sendVKey 3
   
    ' Only export the commercial invoice PDF if we're not skipping it
    If Not blnSkipInvoice Then
        LogMessage "Exporting commercial invoice PDF"
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell[0]").pressButton "&FIND"
        WScript.Sleep 500
        session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").text = strCommercialInvoice
        session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").caretPosition = len(strCommercialInvoice)
        session.findById("wnd[1]").sendVKey 0
        session.findById("wnd[1]").close
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        session.findById("wnd[0]/titl/shellcont[1]/shell").pressContextButton "%GOS_TOOLBOX"
        session.findById("wnd[0]/titl/shellcont[1]/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
        session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").currentCellColumn = "BITM_DESCR"
        session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").pressToolbarButton "%ATTA_EXPORT"
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = strExportPath
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = strDeliveryNumber & " ci.pdf"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = Len(strDeliveryNumber) + 7
        session.findById("wnd[2]").sendVKey 0
       
        ' Check if file already exists (status message appears on wnd[0])
        On Error Resume Next
        strFileExistsStatus = ""
        WScript.Sleep 500 ' Wait for status message to appear
       
        ' Try to get the status bar text from the main window
        strFileExistsStatus = session.findById("wnd[0]/sbar").text
       
        ' Check if it's a file exists message
        If InStr(strFileExistsStatus, "already exists") > 0 Then
            ' File already exists, ask user what to do
            intResponse = MsgBox("The commercial invoice PDF (" & strDeliveryNumber & " ci.pdf) already exists." & vbCrLf & vbCrLf & _
                               "Click OK to replace the existing file." & vbCrLf & _
                               "Click Cancel to skip generating this document.", _
                               vbOKCancel + vbQuestion, "File Already Exists")
               
            If intResponse = vbOK Then
                ' User chose to replace the file
                session.findById("wnd[2]/tbar[0]/btn[11]").press ' Press Replace button
                LogMessage "User chose to replace existing commercial invoice PDF"
            Else
                ' User chose to skip this file
                session.findById("wnd[2]").close ' Close the dialog
                LogMessage "User chose to skip generating commercial invoice PDF"
            End If
        End If
        On Error GoTo 0
       
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    Else
        LogMessage "Skipping commercial invoice PDF export as requested"
    End If
   
    session.findById("wnd[0]").sendVKey 3
    session.findById("wnd[0]").sendVKey 3
   
    ' Display completion message with summary of what was processed
    Dim strDocumentStatus
    If blnSkipInvoice Then
        strDocumentStatus = "Packing List document has been saved to your 'Attachments' folder."
    Else
        strDocumentStatus = "PL & CI documents have been saved to your 'Attachments' folder in your OneDrive."
    End If
   
    MsgBox "SAP automation completed successfully!" & vbCrLf & vbCrLf & _
           "Summary:" & vbCrLf & _
           "- Delivery: " & strDeliveryNumber & vbCrLf & _
           "- Commercial Invoice: " & strCommercialInvoice & vbCrLf & _
           "- Package: " & strWeight & " lb, " & strLength & " x " & strWidth & " x " & strHeight & " inches" & vbCrLf & _
           "- Date: " & strTodayDate & vbCrLf & vbCrLf & _
           strDocumentStatus, vbInformation, strDeliveryNumber & " - Process Complete"
   
    ' Return to VL02N to prepare for next run
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
   
    ' Final cleanup for PDF operations
    CleanupMemory
   
    ' Indicate successful completion
    ExportPDFs = True
End Function

' Connect to SAP GUI with error handling - do this only once
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
CheckErrorWithCleanup "Connecting to SAP GUI"

Set application = SapGuiAuto.GetScriptingEngine
CheckErrorWithCleanup "Getting SAP scripting engine"

If Not IsObject(connection) Then
   Set connection = application.Children(0)
   CheckErrorWithCleanup "Connecting to SAP"
End If

If Not IsObject(session) Then
   Set session = connection.Children(0)
   CheckErrorWithCleanup "Creating SAP session"
End If

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject application, "on"
   CheckErrorWithCleanup "Connecting WScript objects"
End If

' Get the current Windows username
Set objNetwork = CreateObject("WScript.Network")
strCurrentUser = objNetwork.UserName
LogMessage "Current Windows username: " & strCurrentUser

' Set the export path with the current user
strExportPath = "C:\Users\" & strCurrentUser & "\OneDrive - [Company]\Attachments\"
LogMessage "PDF export path set to: " & strExportPath

' Initialize memory management variables
intProcessCount = 0
intMaxProcessesBeforeRefresh = 10  ' Refresh SAP session every 10 deliveries

' Main automation loop - will run continuously until cancelled by user
blnContinueAutomation = True
Do While blnContinueAutomation
    ' Validate SAP connection before processing
    If Not ValidateSAPConnection() Then
        LogMessage "SAP connection validation failed - attempting to refresh"
        RefreshSAPSession
       
        ' If still invalid after refresh, exit
        If Not ValidateSAPConnection() Then
            MsgBox "SAP connection could not be restored. Please restart the script.", vbCritical, "Connection Error"
            Exit Do
        End If
    End If
   
    ' Periodic memory cleanup and session refresh
    If intProcessCount >= intMaxProcessesBeforeRefresh Then
        LogMessage "Performing periodic maintenance after " & intProcessCount & " deliveries"
        RefreshSAPSession
        intProcessCount = 0
    End If
    
    ' Initial data collection
    GetUserInput
   
    ' Confirmation loop - keep showing confirmation until user approves or cancels
    blnDataConfirmed = False
    Do Until blnDataConfirmed
        ' Build confirmation message with all entered data
        Dim strConfirmMessage
        strConfirmMessage = "Please confirm the following information:" & vbCrLf & vbCrLf & _
                           "Delivery Number: " & strDeliveryNumber & vbCrLf & vbCrLf & _
                           "Package Weight: " & strWeight & " lb" & vbCrLf & _
                           "Package Dimensions: " & strLength & " x " & strWidth & " x " & strHeight & " inches" & vbCrLf & _
                           "PL & CI Count: " & strPackingListCount & " PL & " & strCommercialInvoiceCount & " CI" & vbCrLf & _
                           "Today's Date: " & strTodayDate & vbCrLf & vbCrLf & _
                           "Export Path: " & strExportPath & vbCrLf & vbCrLf
        
        ' Add special note if CI count is 0
        If strCommercialInvoiceCount = "0" Then
            strConfirmMessage = strConfirmMessage & _
                               "NOTE: Commercial Invoice count is 0 - only packing list will be processed." & vbCrLf & vbCrLf
        End If
        
        strConfirmMessage = strConfirmMessage & "Is all information correct? (Click OK to continue)"
       
        ' Show confirmation dialog with OK/Cancel options
        intResponse = MsgBox(strConfirmMessage, vbOKCancel + vbQuestion, strDeliveryNumber & " - Confirm Information")
       
        Select Case intResponse
            Case vbOK
                ' User confirmed data is correct
                blnDataConfirmed = True
            Case vbCancel
                ' User wants to edit data
                GetUserInput
        End Select
    Loop
   
    ' Process the delivery with enhanced decision tree approach
    If Not ProcessDelivery() Then
        LogMessage "Returning to input prompt due to error or user choice"
        ' Continue directly to the next iteration, which will prompt for new delivery
    Else
        ' Log the start of a new automation cycle and continue directly
        LogMessage "Process completed successfully, starting new automation cycle"
        ' Increment the process counter for memory management
        intProcessCount = intProcessCount + 1
    End If
   
    ' Perform cleanup after each delivery to prevent memory accumulation
    CleanupMemory
   
    ' Continue to the next delivery - no prompt asking if user wants to continue
Loop

' Final cleanup when exiting
LogMessage "Script ending - performing final cleanup"
CleanupMemory

' Clean up COM objects before exit
On Error Resume Next
Set session = Nothing
Set connection = Nothing
Set application = Nothing
Set SapGuiAuto = Nothing
Set objNetwork = Nothing
LogMessage "Final COM object cleanup completed"