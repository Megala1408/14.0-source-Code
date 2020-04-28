VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form frmAEPresWithEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presented With Details"
   ClientHeight    =   5355
   ClientLeft      =   915
   ClientTop       =   1845
   ClientWidth     =   8115
   Icon            =   "AEPresWithEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   6780
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   5460
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame fraPresentedWith 
      BorderStyle     =   0  'None
      Height          =   4275
      Left            =   180
      TabIndex        =   9
      Top             =   480
      Width           =   7755
      Begin VB.Frame fraSymptoms 
         Caption         =   "Sy&mptoms:"
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   7515
         Begin VB.OptionButton optAlpha 
            Alignment       =   1  'Right Justify
            Caption         =   "&Alphabetical:"
            Height          =   315
            Left            =   4920
            TabIndex        =   3
            Top             =   2160
            Width           =   1215
         End
         Begin VB.OptionButton optSelected 
            Alignment       =   1  'Right Justify
            Caption         =   "&Selected:"
            Height          =   315
            Left            =   6360
            TabIndex        =   4
            Top             =   2160
            Value           =   -1  'True
            Width           =   975
         End
         Begin MSComctlLib.ListView lstPresentedWith 
            Height          =   1905
            Left            =   180
            TabIndex        =   2
            Top             =   240
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   3360
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            OLEDragMode     =   1
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            OLEDragMode     =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "DESC"
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   "CHECK"
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame fraOther 
         Caption         =   "&Other:"
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   2580
         Width           =   7515
         Begin VB.TextBox txtOther 
            Height          =   1215
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   240
            Width           =   7185
         End
      End
   End
   Begin MSComctlLib.TabStrip tabTriage 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   8281
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Presented With"
            Key             =   "PWITH"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Presented With Details"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAEPresWithEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Context.Version" ,"11"
Attribute VB_Ext_KEY = "Context.Name" ,"frmAEPresWithEdit"
Option Explicit
'
Public lnCallingForm As Long
'
Dim mbOKPressed As Boolean
Dim mbIsDirty As Boolean
Dim msPresentedWith As String
Dim mlnPresWithMousePos As Long
Dim mlnsOdpcdRefnos As LongSet
'
' Useful recordsets
Dim mrsPWith As HRS
Dim mrsCurrPWith As HRS
'
' Useful refnos
Dim mlnHeorgRefno As Long
Dim mlnPatntRefno As Long
'
'
Private Sub CustomiseForm()
Rem
Rem Parameters
Rem
Rem Purpose     Customises the form based on call mode
Rem
Rem Comments
Rem
Rem Returns     None
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 28-Aug-97   [100]   KMM     Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.CustomiseForm"  ' ¦Transaction Log¦
    End If
    '
    ' Set the form icon
    Me.Icon = gfrmAppMDI.imlicons16x16.ListImages("PATAE").Picture
    '
End Sub

Private Function bUpdatePresWithSymptoms() As Boolean
Rem
Rem Parameters
Rem
Rem Purpose     Iterates through items in list and adds any new symptoms
Rem             to diagnosis procedures if not there already. Deletes any
Rem             which have been deselected.
Rem
Rem Comments
Rem
Rem Returns     string containing comma separated list of selected symptoms
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 17-Jul-97   [100]   JS      Created
Rem 22-Jul-98   [101]   AES     Database access not handled by form
Rem 02-Jan-07	[102]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.bUpdatePresWithSymptoms"  ' ¦Transaction Log¦
    End If
    '
    Dim itmX As ListItem
    Dim lnOdpcdRefno As Long
    '
    For Each itmX In lstPresentedWith.ListItems
        lnOdpcdRefno = lnExtractRefnoFromKey(itmX.key)
        If itmX.SmallIcon = "CHECK" Then
            '
            mlnsOdpcdRefnos.Add lnOdpcdRefno
            '
        ElseIf itmX.SmallIcon = "UNCHECK" Then
            '
            mlnsOdpcdRefnos.Remove lnOdpcdRefno
            '
        End If
    Next itmX
    '
    bUpdatePresWithSymptoms = True
    '
End Function

Public Function bDisplay(sPresentedWith As String, _
                         lnHeorgRefno As Long, _
                         lnPatntRefno As Long, _
                         lnsOdpcdRefnos As LongSet) As Boolean
Rem
Rem Parameters  sPresentedWith      "presented with" additional symptoms
Rem             lnHeorgRefno        Health Org Refno of AE dept
Rem             lnPatntRefno        Patient refno of AE patient
Rem             lnsOdpcdRefnos      Presented with symptom refnos
Rem
Rem Purpose
Rem
Rem Comments    Triage details now stored on new table ae_attendance_roles
Rem             rather than ae_attendances. Form now GUI only with no database
Rem             access.
Rem
Rem Returns     True if any details have been amended, False otherwise
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 17-Jul-97   [100]   JS      Created
Rem 21-Jul-98   [101]   AES     Form now GUI only
Rem 02-Jan-07	[102]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.bDisplay"  ' ¦Transaction Log¦
    End If
    '
    lnCallingForm = lnGetCallingForm()
    '
    mlnHeorgRefno = lnHeorgRefno
    mlnPatntRefno = lnPatntRefno
    Set mlnsOdpcdRefnos = lnsOdpcdRefnos
    msPresentedWith = sPresentedWith
    '
    Screen.MousePointer = vbHourglass
    '
    Load Me
    mbOKPressed = False
    CentreFormInScreen Me
    '
    Call CustomiseForm
    '
    If bInitForEdit Then
        '
        Call SetHelpContextIDs
        '
        ' Clear all the field dirty flags
        mbIsDirty = False
        '
        Screen.MousePointer = vbDefault
        '
        Me.Show vbModal
    Else
        Screen.MousePointer = vbDefault     ' Reset the screen pointer
    End If
    '
    ' Return string containing other symptoms
    If mbOKPressed Then
        sPresentedWith = txtOther.Text
    End If
    '
    Unload Me
    '
    bDisplay = mbOKPressed
    '
End Function

Private Function bInitForEdit()
Rem
Rem Parameters
Rem
Rem Purpose     Prepares the form for amendment of an AE attendance
Rem             by setting any defaults and pre-populating where necessary
Rem
Rem Comments
Rem
Rem Returns     True if successful, false otherwise
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 30-Jul-97   [100]   KMM     Created
Rem 22-Jul-98   [101]   AES     Presented with symptoms now in set and note in
Rem                             string
Rem 02-Jan-07	[102]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.bInitForEdit"  ' ¦Transaction Log¦
    End If
    '
    bInitForEdit = False
    '
    ' Determine picklist of symptoms for current AE department (health org)
    If Not bSQLSelectHeorgPWithSymptoms(mrsPWith, mlnHeorgRefno) Then
        ' °03975° Sorry, unable to select Health Org 'Presented With' symptoms.
        nMsgBox sTran(3975), vbExclamation
        Exit Function
    End If
    '
    If mlnsOdpcdRefnos.Count > 0 Then
        If Not bSQLSelectPWithSymptoms(mrsCurrPWith, mlnsOdpcdRefnos) Then
            ' °03975° Sorry, unable to select Health Org 'Presented With' symptoms.
            nMsgBox sTran(3975), vbExclamation
            Exit Function
        End If
        '
    End If
    '
    Call Populate_lstPresentedWith
    txtOther.Text = msPresentedWith
    '
    ' OK, if we get this far
    bInitForEdit = True
    '
End Function
Private Function bSQLSelectPWithSymptoms(rsPWith As HRS, _
                                         lnsOdpcdRefnos As LongSet) As Boolean
Rem
Rem Parameters  rsPWith                 ' JIL selectset to populate
Rem             lnOdpcdRefnos           ' set of codes
Rem
Rem Purpose     To populate a selectset with the details of Presented With
Rem             symptoms for the given A+E Attendance episodic triage
Rem
Rem Comments
Rem
Rem Returns     True if successful, false otherwise
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 22-Jul-98   [100]   AES     Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.bSQLSelectPWithSymptoms"  ' ¦Transaction Log¦
    End If
    '
    ' Now build the SQL to populate the selectset
    Dim sSQL As String
    Dim itmX As Variant     ' used to iterate through set
    Dim lnLong As Long
    '
    sSQL = ""
    '
    'Select fields
    AddSqlString sSQL, "SELECT odpcd.odpcd_refno odpcd_odpcd_refno, "
    AddSqlString sSQL, "       odpcd.description odpcd_description "
    '
    'From tables
    AddSqlString sSQL, "FROM   odpcd_codes odpcd,"
    AddSqlString sSQL, "       diagnosis_procedures dgpro"
    '
    'Where criteria
    If lnsOdpcdRefnos.Count = 1 Then
        AddSqlString sSQL, "WHERE  dgpro.odpcd_refno=" & sAddSqlRefno(lnsOdpcdRefnos.Item(1))
    Else
        AddSqlString sSQL, "WHERE dgpro.odpcd_refno IN ("
        For Each itmX In lnsOdpcdRefnos
            '
            ' probably could convert straight to a string
            lnLong = CLng(itmX)
            If lnLong = lnsOdpcdRefnos.Item(lnsOdpcdRefnos.Count) Then
                AddSqlString sSQL, CStr(lnLong) & ")"
            Else
                AddSqlString sSQL, CStr(lnLong) & ","
            End If
        Next
    End If
    '
    AddSqlString sSQL, "AND    dgpro.odpcd_refno=odpcd.odpcd_refno"
    AddSqlString sSQL, "AND    " & sAddSqlArchvTest("diagnosis_procedures", "dgpro")
    AddSqlString sSQL, "AND    " & sAddSqlArchvTest("odpcd_codes", "odpcd")
    '
    JIL_NewSelectSet rsPWith, sSQL
    '
    bSQLSelectPWithSymptoms = Not JIL_bIsFatal()
    '
End Function

Private Sub SetHelpContextIDs()
Rem
Rem Parameters  sTabName as string      Name of tab to show
Rem
Rem Purpose     Sets Help context IDs for this form based on
Rem             the calling mode and the current tab
Rem
Rem Comments    Should be called from ShowTabPicture (after calling
Rem             mode has been determined)
Rem
Rem Returns
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 17-Jul-97   [100]   JS      Created
Rem 09-Sep-98   [101]   MTC     Set new help ID
Rem 15-Sep-98   [102]   MTC     Set new help ID
Rem 02-Jan-07	[103]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.SetHelpContextIDs"  ' ¦Transaction Log¦
    End If
    '
    Me.HelpContextID = 1030
    '
End Sub

Private Sub cmdCancel_Click()
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 25-Jul-97   [100]   JS      Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.cmdCancel_Click"  ' ¦Transaction Log¦
    End If
    '
    ' °03976° 'Presented With' Symptoms
    If bConfirmCancel(sTran(3976), mbIsDirty, , True) Then
        mbOKPressed = False
        HideFrmSetfocus Me
    End If
    '
End Sub

Private Sub cmdOK_Click()
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 25-Jul-97   [100]   JS      Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.cmdOK_Click"  ' ¦Transaction Log¦
    End If
    '
    If bUpdatePresWithSymptoms() Then
        mbOKPressed = True
        HideFrmSetfocus Me
    End If
    '
End Sub

Private Sub Form_Activate()
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 25-Jul-97   [100]   JS      Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.Form_Activate"  ' ¦Transaction Log¦
    End If
    '
    Call ResetTimeoutCount
    '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 25-Jul-97   [100]   JS      Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.Form_KeyDown"  ' ¦Transaction Log¦
    End If
    '
    Call ResetTimeoutCount
    '
End Sub


Private Sub Form_Load()
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 17-Apr-98   [100]   APS     Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.Form_Load"  ' ¦Transaction Log¦
    End If
    '
    TranslateForm
    '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 25-Jul-97   [100]   JS      Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.Form_QueryUnload"  ' ¦Transaction Log¦
    End If
    '
    ' Query user if any change has been made, or if any
    ' episodic controls are currently in edit mode
    If UnloadMode <> vbFormCode And mbIsDirty Then
        ' °03977° Save changes to 'Presented With' symptoms?
        If nMsgBox(sTran(3977), vbExclamation + vbYesNo) = vbYes Then
            Call cmdOK_Click
            Cancel = Not mbOKPressed
        Else
            Call cmdCancel_Click
        End If
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 25-Jul-97   [100]   JS      Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.Form_Unload"  ' ¦Transaction Log¦
    End If
    '
    JIL_Close mrsPWith
    JIL_Close mrsCurrPWith
    '
End Sub

Private Sub lstPresentedWith_Click()
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 06-Aug-97   [100]   JS      Created
Rem 21-Apr-97   [101]   GRS     Handle case where no items exist in the list
Rem 02-Jan-07	[102]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.lstPresentedWith_Click"  ' ¦Transaction Log¦
    End If
    '
    If lstPresentedWith.ListItems.Count > 0 Then
        mbIsDirty = bToggleListItem(lstPresentedWith, mlnPresWithMousePos)
    End If
    '
End Sub

Private Sub lstPresentedWith_KeyPress(KeyAscii As Integer)
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 06-Aug-97   [100]   JS      Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.lstPresentedWith_KeyPress"  ' ¦Transaction Log¦
    End If
    '
    If KeyAscii = vbKeySpace Then
        mlnPresWithMousePos = 0
        Call lstPresentedWith_Click
    End If
    '
End Sub

Private Sub lstPresentedWith_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 06-Aug-97   [100]   JS      Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.lstPresentedWith_MouseDown"  ' ¦Transaction Log¦
    End If
    '
    mlnPresWithMousePos = X
    '
End Sub

Private Sub optAlpha_Click()
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 25-Jul-97   [100]   JS      Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.optAlpha_Click"  ' ¦Transaction Log¦
    End If
    '
    Call SortListView(lstPresentedWith, lstPresentedWith.ColumnHeaders("DESC"))
    '
End Sub


Private Sub optSelected_Click()
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 25-Jul-97   [100]   JS      Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.optSelected_Click"  ' ¦Transaction Log¦
    End If
    '
    Call SortListView(lstPresentedWith, lstPresentedWith.ColumnHeaders("CHECK"))
    '
End Sub

Private Sub Populate_lstPresentedWith()
Rem
Rem Parameters
Rem
Rem Purpose     Populates the 'Presented With' listview with the symptoms for
Rem             the current health org as a picklist. Then initialises the
Rem             picklist with the symptoms for the the current AE attendance.
Rem
Rem Comments
Rem
Rem Returns     None
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 16-Jul-97   [100]   JS      Created
Rem 22-Jul-98   [101]   AES     Symptoms now in set
Rem 02-Jan-07	[102]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.Populate_lstPresentedWith"  ' ¦Transaction Log¦
    End If
    '
    Dim sKey As String          ' Key of listitem to add
    Dim itmX As Variant         ' iterate through the set
    Dim lnLong As Long
    '
    Call ClearListView(lstPresentedWith)
    '
    lstPresentedWith.SmallIcons = frmMDIMenu.imlicons16x16
    '
    Do While Not JIL_bIsEOF(mrsPWith)
        sKey = "ODPCD:" & CStr(JIL_lnGet(mrsPWith, "odpcd_odpcd_refno"))
        Call SetPresWithItem(lstPresentedWith, sKey, mrsPWith, False)
        JIL_MoveNext mrsPWith
    Loop
    '
    If mlnsOdpcdRefnos.Count > 0 Then
        Do While Not JIL_bIsEOF(mrsCurrPWith)
            sKey = "ODPCD:" & CStr(JIL_lnGet(mrsCurrPWith, "odpcd_odpcd_refno"))
            If bItemInList(sKey, lstPresentedWith) Then
                Call SetPresWithItem(lstPresentedWith, sKey, mrsCurrPWith, True)
            End If
            JIL_MoveNext mrsCurrPWith
        Loop
    End If
    '
    Call optSelected_Click
    '
End Sub

Private Sub SetPresWithItem(lstPWith As ListView, _
                            sKey As String, _
                            rsPWith As HRS, _
                            bChecked As Boolean)
Rem
Rem Parameters  lstPWith            Presented With list ListView to update
Rem             sKey                Key of ListItem to update
Rem             rs                  Recordset to read data from
Rem             bChecked            If item should have a 'Checked' icon and status
Rem
Rem Purpose     Sets the subitems of a given ListItem from a Presented
Rem             With symptom according to values in the current row of the
Rem             recordset
Rem
Rem Comments    The Listview MUST contain the following column keys:
Rem                 DESC            Symptom description
Rem                 CHECK           Hidden column used for sorting
Rem             The Symptom description must be the main item
Rem
Rem Returns     None
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 17-Jul-97   [100]   JS      Created
Rem 28-Jul-97   [101]   JS      Initialise 'CHECK' column to allow sorting by
Rem                             Checked/Unchecked items
Rem 02-Jan-07	[102]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.SetPresWithItem"  ' ¦Transaction Log¦
    End If
    '
    Dim itmX As ListItem                ' ListItem to update
    '
    If bItemInList(sKey, lstPWith) Then
        Set itmX = lstPWith.ListItems(sKey)
    Else
        Set itmX = lstPWith.ListItems.Add(, sKey, JIL_sGet(rsPWith, "odpcd_description"))
    End If
    '
    itmX.SmallIcon = IIf(bChecked, "CHECK", "UNCHECK")
    '
    itmX.SubItems(lstPWith.ColumnHeaders("CHECK").Index - 1) = IIf(bChecked, "A", "B")
    '
End Sub

Private Sub txtOther_Change()
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 28-Jul-97   [100]   KMM     Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.txtOther_Change"  ' ¦Transaction Log¦
    End If
    '
    Call NotesMaintain("AEPWI", txtOther)
    '
    mbIsDirty = True
    '
End Sub

Private Sub txtOther_GotFocus()
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 15-Aug-97   [100]   KMM     Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.txtOther_GotFocus"  ' ¦Transaction Log¦
    End If
    '
    DisableDialogFormDefaultButton Me
    '
End Sub

Private Sub txtOther_KeyPress(KeyAscii As Integer)
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 28-Jul-97   [100]   KMM     Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.txtOther_KeyPress"  ' ¦Transaction Log¦
    End If
    '
    If KeyAscii = 14 Then   ' CTRL-N
        Call bNotesEditText("AEPWI", txtOther)
        KeyAscii = 0    ' Prevent any Windows processing
    End If
    '
End Sub

Private Sub txtOther_LostFocus()
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 15-Aug-97   [100]   KMM     Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.txtOther_LostFocus"  ' ¦Transaction Log¦
    End If
    '
    EnableDialogFormDefaultButton Me
    '
End Sub


Private Sub TranslateForm()
Rem
Rem Purpose     Translates form text.
Rem
Rem Comments    Automatically generated by Translation Wizard. Do not modify by hand as
Rem             any changes made may be overwritten.
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 17-Apr-98   [100]   WIZ     Created (new string table)
Rem 26-Jul-98   [101]   WIZ     Modified (string table v1)
Rem 06-Aug-98   [102]   APS     Modified (string table v2)
Rem 26-Nov-98   [103]   WIZ     Modified (string table v4)
Rem 26-Nov-98   [104]   WIZ     Modified (string table v5)
Rem 26-Nov-98   [105]   WIZ     Modified (string table v5)
Rem 26-Nov-98   [106]   WIZ     Modified (string table v5)
Rem 28-Jan-99   [107]   WIZ     Modified (string table v6)
Rem 08-Feb-99   [108]   WIZ     Modified (string table v9)
Rem 02-Jan-07	[109]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.TranslateForm"  ' ¦Transaction Log¦
    End If
    '
'~  Frame1.Caption = sTran(16916)                       ' °16916° Sy&mptoms:
'~  Frame2.Caption = sTran(3981)                        ' °03981° &Other:
    Me.Caption = sTran(3978)                            ' °03978° Presented With Details
    cmdCancel.Caption = sTran(3979)                     ' °03979° Cancel
    cmdOK.Caption = sTran(3980)                         ' °03980° OK
    'lblOther.Caption = sTran(3981)                      ' °03981° &Other:
    fraOther.Caption = sTran(3981)                      ' °03981° &Other:
    fraSymptoms.Caption = sTran(16916)                  ' °16916° Sy&mptoms:
    optAlpha.Caption = sTran(3982)                      ' °03982° &Alphabetical:
    optSelected.Caption = sTran(3983)                   ' °03983° &Selected:
    tabTriage.Tabs(1).Caption = sTran(3984)             ' °03984° Presented With
    tabTriage.Tabs(1).ToolTipText = sTran(3978)         ' °03978° Presented With Details
    '
End Sub

Private Function sTran(lnNativeID As Long, ParamArray pa()) As String
Rem
Rem Parameters  lnNativeID      NativeID
Rem             pa              ParamArray (extra parameters for |1 |2 substitution)
Rem
Rem Purpose     Returns the translated text for string number lnNativeId.
Rem
Rem Comments    This function must be declared Private in each module that uses it.
Rem
Rem Returns     Text in target language.
Rem
Rem Change History
Rem Date        Edit    Author  Comment
Rem -----------+-------+-------+---------------------------------------------
Rem 06-Aug-98   [100]   APS     Created
Rem 02-Jan-07	[101]	DEP  	Added Transaction log RF done from 22.1
Rem -----------+-------+-------+---------------------------------------------
Rem
    If vGetSwitch(FSW_TRANSACTION_LOG) <> ""  Then
       JIL_LogTran "aepreswithedit.sTran"  ' ¦Transaction Log¦
    End If
    '
    Dim vParamArray As Variant
    '
    ' Type-coerce ParamArray to a variant array
    vParamArray = pa
    '
    ' Call the master translation function.
    sTran = sTranMaster(lnNativeID, vParamArray)
    '
End Function

