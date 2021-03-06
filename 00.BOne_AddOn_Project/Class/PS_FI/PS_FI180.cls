VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FI180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_FI180.cls
'//  Module         : FI
'//  Description    : 보조원장
'//  FormType       : PS_FI180
'//  Create Date    : 2011.02.09
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_FI180H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_FI180L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FI180.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_FI180_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_OK_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
'    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), False        '// 취소
    oForm01.EnableMenu ("1293"), False        '// 행삭제
        
    oForm01.Update
    oForm01.Freeze False
    oForm01.Visible = True
    
    Set oXmlDoc01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    If (oForm01 Is Nothing) = False Then
        Set oForm01 = Nothing
    End If
    MDC_Com.MDC_GF_Message "LoadForm_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
'    Set oDS_PS_FI180H = oForm01.DataSources.DBDataSources("@PS_FI180H")
'    Set oDS_PS_FI180L = oForm01.DataSources.DBDataSources("@PS_FI180L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("StrDate", dt_DATE)
    oForm01.Items("StrDate").Specific.DataBind.SetBound True, "", "StrDate"
    oForm01.DataSources.UserDataSources.Item("StrDate").VALUE = Format(Now, "YYYY-MM-DD")

    Call oForm01.DataSources.UserDataSources.Add("EndDate", dt_DATE)
    oForm01.Items("EndDate").Specific.DataBind.SetBound True, "", "EndDate"
    oForm01.DataSources.UserDataSources.Item("EndDate").VALUE = Format(Now, "YYYY-MM-DD")
    
    '//체크박스 처리
    Dim oChkBox As SAPbouiCOM.CheckBox

    Call oForm01.DataSources.UserDataSources.Add("Check01", dt_SHORT_TEXT, 1)
    Set oChkBox = oForm01.Items("ChkBox01").Specific
    oChkBox.ValOn = "Y"
    oChkBox.ValOff = "N"
    oChkBox.DataBind.SetBound True, , "Check01"
    oForm01.DataSources.UserDataSources("Check01").VALUE = "N" '미체크로 값을 주고 폼을 로드
    
    Call oForm01.DataSources.UserDataSources.Add("Check02", dt_SHORT_TEXT, 1)
    Set oChkBox = oForm01.Items("ChkBox02").Specific
    oChkBox.ValOn = "Y"
    oChkBox.ValOff = "N"
    oChkBox.DataBind.SetBound True, , "Check02"
    oForm01.DataSources.UserDataSources("Check02").VALUE = "N" '미체크로 값을 주고 폼을 로드
    
    '기준일자 콤보_S
    Call oForm01.DataSources.UserDataSources.Add("DateCls", dt_SHORT_TEXT, 10)
    Call oForm01.Items("DateCls").Specific.DataBind.SetBound(True, "", "DateCls")
    '기준일자 콤보_E

    Set oChkBox = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId"
    oRecordSet01.DoQuery sQry
    Call oCombo.ValidValues.Add("0", "전체 사업장")
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    Call oCombo.Select(0, psk_Index)
    
    '// 항목선택
    Set oCombo = oForm01.Items("Rpt").Specific
    Call oCombo.ValidValues.Add("0", "전체항목")
    Call oCombo.ValidValues.Add("1", "관리항목 1")
    Call oCombo.ValidValues.Add("2", "관리항목 2")
    Call oCombo.ValidValues.Add("3", "관리항목 3")
    Call oCombo.ValidValues.Add("4", "관리항목 4")
    Call oCombo.ValidValues.Add("5", "관리항목 5")
    Call oCombo.ValidValues.Add("6", "관리항목 6")
    Call oCombo.Select(0, psk_Index)
    
    '기준일자 콤보
    Call oForm01.Items("DateCls").Specific.ValidValues.Add("01", "전기일기준")
    Call oForm01.Items("DateCls").Specific.ValidValues.Add("02", "증빙일기준")
    Call oForm01.Items("DateCls").Specific.Select(0, psk_Index)
    
    oForm01.Items("StrDate").Click ct_Regular
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
'                        If HeaderSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                    End If
                
                ElseIf pval.ItemUID = "Btn01" Then '출력버튼 클릭시
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call Print_Query
                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
                    If pval.ItemUID = "SAcctCode" Then
                        If oForm01.Items("SAcctCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "EAcctCode" Then
                        If oForm01.Items("EAcctCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    
                    If pval.ItemUID = "StrRpt" Then
                        If oForm01.Items("StrRpt").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "EndRpt" Then
                        If oForm01.Items("EndRpt").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    '//라인
'                    If pval.ItemUID = "Mat01" Then
'                        If pval.ColUID = "PP070No" Then
'                            If oMat01.Columns("PP070No").Cells(pval.Row).Specific.Value = "" Then
'                                Sbo_Application.ActivateMenuItem ("7425")
'                                BubbleEvent = False
'                            End If
'                        End If
'                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
        
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
'                If pval.ItemUID = "1" Then
'                    If oForm01.Mode = fm_ADD_MODE Then
'                        oForm01.Mode = fm_OK_MODE
'                        Call Sbo_Application.ActivateMenuItem("1282")
'                    ElseIf oForm01.Mode = fm_OK_MODE Then
'                        FormItemEnabled
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
'                    End If
'                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "Rpt" Then
                    If Trim(oForm01.Items("Rpt").Specific.Selected.VALUE) <> "0" Then
                        If Trim(oForm01.Items("SAcctCode").Specific.VALUE) = "" Then
                            oForm01.Freeze True
                            MDC_Com.MDC_GF_Message "계정과목(시작) 입력 후 항목선택을 하여 주시기 바랍니다.", "E"
                            Call oForm01.Items("Rpt").Specific.Select(0, psk_Index)
                            oForm01.Items("SAcctCode").Click ct_Regular
                            oForm01.Freeze False
                            BubbleEvent = False
                            Exit Sub
                        
                        ElseIf Trim(oForm01.Items("SAcctCode").Specific.VALUE) <> "" Then
                            oForm01.Freeze True
                            Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                            oForm01.Items("StrRpt").Enabled = True
                            oForm01.Items("EndRpt").Enabled = True
                            oForm01.Items("EndRpt").Specific.VALUE = ""
                            oForm01.Items("StrRpt").Specific.VALUE = ""
                            oForm01.Freeze False
                        End If
                    Else
                        oForm01.Freeze True
                        oForm01.Items("Rpttxt").Specific.VALUE = ""
                        oForm01.Items("StrRpt").Specific.VALUE = ""
                        oForm01.Items("EndRpt").Specific.VALUE = ""
                        oForm01.Items("StrDate").Click ct_Regular
                        oForm01.Items("StrRpt").Enabled = False
                        oForm01.Items("EndRpt").Enabled = False
                        oForm01.Freeze False
                    End If
                End If
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "SAcctCode" Then
                        Call oForm01.Items("Rpt").Specific.Select(0, psk_Index)
                        oForm01.Items("Rpttxt").Specific.VALUE = ""
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
'                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    If ErrNum = 101 Then
        ErrNum = 0
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
        BubbleEvent = False
    Else
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1285": '복원
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
        
    '-----------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1285": '복원
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
        End Select
    ElseIf (BusinessObjectInfo.BeforeAction = False) Then '//BeforeAction = False
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormDataEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_FormDataEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim i&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "Rpt"
            If Trim(oForm01.Items("Rpt").Specific.Selected.VALUE) = "1" Then
                sQry = "        Select  TOP 1 "
                sQry = sQry & "         U_Rpttxt01 "
                sQry = sQry & " from    [ZMDC_JDT1]"
                sQry = sQry & " where   AcctCode = '" & Trim(oForm01.Items("SAcctCode").Specific.VALUE) & "'"
                sQry = sQry & "         AND U_Rpttxt01 IS NOT NULL"
            ElseIf Trim(oForm01.Items("Rpt").Specific.Selected.VALUE) = "2" Then
                sQry = "        Select  TOP 1"
                sQry = sQry & "         U_Rpttxt02"
                sQry = sQry & " from    [ZMDC_JDT1]"
                sQry = sQry & " where   AcctCode = '" & Trim(oForm01.Items("SAcctCode").Specific.VALUE) & "'"
                sQry = sQry & "         AND U_Rpttxt02 IS NOT NULL"
            ElseIf Trim(oForm01.Items("Rpt").Specific.Selected.VALUE) = "3" Then
                sQry = "        Select  TOP 1 "
                sQry = sQry & "         U_Rpttxt03 "
                sQry = sQry & " from    [ZMDC_JDT1]"
                sQry = sQry & " where   AcctCode = '" & Trim(oForm01.Items("SAcctCode").Specific.VALUE) & "'"
                sQry = sQry & "         AND U_Rpttxt03 IS NOT NULL"
            ElseIf Trim(oForm01.Items("Rpt").Specific.Selected.VALUE) = "4" Then
                sQry = "        Select  TOP 1 "
                sQry = sQry & "         U_Rpttxt04 "
                sQry = sQry & " from    [ZMDC_JDT1]"
                sQry = sQry & " where   AcctCode = '" & Trim(oForm01.Items("SAcctCode").Specific.VALUE) & "'"
                sQry = sQry & "         AND U_Rpttxt04 IS NOT NULL"
            ElseIf Trim(oForm01.Items("Rpt").Specific.Selected.VALUE) = "5" Then
                sQry = "        Select  TOP 1 "
                sQry = sQry & "         U_Rpttxt05 "
                sQry = sQry & " from    [ZMDC_JDT1]"
                sQry = sQry & " where   AcctCode = '" & Trim(oForm01.Items("SAcctCode").Specific.VALUE) & "'"
                sQry = sQry & "         AND U_Rpttxt05 IS NOT NULL"
            ElseIf Trim(oForm01.Items("Rpt").Specific.Selected.VALUE) = "6" Then
                sQry = "        Select  TOP 1 "
                sQry = sQry & "         U_Rpttxt06 "
                sQry = sQry & " from    [ZMDC_JDT1]"
                sQry = sQry & " where   AcctCode = '" & Trim(oForm01.Items("SAcctCode").Specific.VALUE) & "'"
                sQry = sQry & "         AND U_Rpttxt06 IS NOT NULL"
            End If
            oRecordSet.DoQuery sQry
            oForm01.Items("Rpttxt").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
        
        End Select
    End If
     
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("StrDate").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Len(Trim(oForm01.Items("EndDate").Specific.VALUE)) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "기간 시작일은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "기간 종료일은 필수사항입니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub Print_Query()
On Error GoTo Print_Query_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim Sub_sQry        As String
    
    Dim StrDate         As String
    Dim EndDate         As String
    Dim SAcctCode       As String
    Dim EAcctCode       As String
    Dim BPLID           As String
    Dim Rpt             As String
    Dim Rpttxt          As String
    Dim StrRpt          As String
    Dim EndRpt          As String
    Dim Summary         As String
    Dim Level5          As String
    Dim DateCls As String
        
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    '// 조회조건문
    StrDate = Trim(oForm01.Items("StrDate").Specific.VALUE)
    EndDate = Trim(oForm01.Items("EndDate").Specific.VALUE)
    SAcctCode = Trim(oForm01.Items("SAcctCode").Specific.VALUE)
    EAcctCode = Trim(oForm01.Items("EAcctCode").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    Rpt = Trim(oForm01.Items("Rpt").Specific.Selected.VALUE)
    Rpttxt = Trim(oForm01.Items("Rpttxt").Specific.VALUE)
    StrRpt = Trim(oForm01.Items("StrRpt").Specific.VALUE)
    EndRpt = Trim(oForm01.Items("EndRpt").Specific.VALUE)
    Summary = Trim(oForm01.DataSources.UserDataSources("Check01").VALUE)
    Level5 = Trim(oForm01.DataSources.UserDataSources("Check02").VALUE)
    DateCls = Trim(oForm01.Items("DateCls").Specific.Selected.VALUE)
    
'    If SYYYYMM = "" Then SYYYYMM = "19000101"
'    If EYYYYMM = "" Then EYYYYMM = "21001231"
    If SAcctCode = "" Then SAcctCode = "1"
    If EAcctCode = "" Then EAcctCode = "9999999999"
    If StrRpt = "" Then StrRpt = "!"
    If EndRpt = "" Then EndRpt = "ZZZZZZZZZZ"
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_FI180] 보조원장"
    If Rpt = "0" Then
        ReportName = "PS_FI180_00.RPT"
    ElseIf Rpt = "1" Then
        ReportName = "PS_FI180_01.RPT"
    ElseIf Rpt = "2" Then
        ReportName = "PS_FI180_02.RPT"
    ElseIf Rpt = "3" Then
        ReportName = "PS_FI180_03.RPT"
    ElseIf Rpt = "4" Then
        ReportName = "PS_FI180_04.RPT"
    ElseIf Rpt = "5" Then
        ReportName = "PS_FI180_05.RPT"
    ElseIf Rpt = "6" Then
        ReportName = "PS_FI180_06.RPT"
    End If
    
    If Summary = "Y" Then
        WinTitle = "[PS_FI180] 보조원장 집계표"
        ReportName = "PS_FI180_20.RPT"
    End If
    
    '// Formula 수식필드
    ReDim gRpt_Formula(4)
    ReDim gRpt_Formula_Value(4)
    
    gRpt_Formula(1) = "StrDate"
    gRpt_Formula_Value(1) = IIf(StrDate = "19000101", "All", Format(StrDate, "0000-00-00"))
    gRpt_Formula(2) = "EndDate"
    gRpt_Formula_Value(2) = IIf(EndDate = "21001231", "All", Format(EndDate, "0000-00-00"))
    gRpt_Formula(3) = "BPLId"
    gRpt_Formula_Value(3) = BPLID
    gRpt_Formula(4) = "Rpt"
    gRpt_Formula_Value(4) = Rpt
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    sQry = "EXEC [PS_FI180_01] '" & BPLID & "','" & StrDate & "','" & EndDate & "','" & SAcctCode & "','" & EAcctCode & "','" & Rpt & "','" & Rpttxt & "','" & StrRpt & "','" & EndRpt & "', '" & Summary & "', '" & Level5 & "','" & DateCls & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Query_Error
    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub




