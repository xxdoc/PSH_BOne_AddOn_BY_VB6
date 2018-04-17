VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FI901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_FI901.cls
'//  Module         : FI
'//  Description    : IFRS - 매입 매출 채무 채권 추출
'//  FormType       : PS_FI901
'//  Create Date    : 2009.07.20
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit
Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oGrid01         As SAPbouiCOM.Grid

Private oLast_Item_UID As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row  As Long   '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FI901.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_FI901_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    
    oForm01.Freeze True
    
    CreateItems
    
    oForm01.EnableMenu ("1283"), False        '// 제거
    oForm01.EnableMenu ("1284"), False        '// 취소
    oForm01.EnableMenu ("1287"), False        '// 복원
    oForm01.EnableMenu ("1293"), False        '// 행삭제
   
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    Exit Sub

LoadForm_Error:
    Set oXmlDoc01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
    If (oForm01 Is Nothing) = False Then
        Set oForm01 = Nothing
    End If
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim sReturnValue    As Integer
    
    Dim DocType$, Company$, AcctYear$, AcctMon$, BPLID$
    
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Or pval.ItemUID = "Btn02" Or pval.ItemUID = "Btn03" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    Company = Trim(oForm01.Items("Company").Specific.VALUE)
                    DocType = Trim(oForm01.Items("DocType").Specific.Selected.VALUE)
                    AcctYear = Trim(oForm01.Items("AcctYear").Specific.VALUE)
                    AcctMon = Trim(oForm01.Items("AcctMon").Specific.VALUE)
                    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
                    If BPLID = "" Then BPLID = "%"
                    
                    If pval.ItemUID = "Btn01" Then
                        sQry = "Select * From [ZFI020] Where "
                        sQry = sQry & "Company = '" & Company & "' And "
                        sQry = sQry & "DocType = '" & DocType & "' And "
                        sQry = sQry & "AcctYear = '" & AcctYear & "' And "
                        sQry = sQry & "AcctMon = '" & AcctMon & "'"
                        oRecordSet.DoQuery sQry
                        
                        If oRecordSet.RecordCount > 0 Then
                            sReturnValue = Sbo_Application.MessageBox("해당 조건의 데이터가 존재합니다. 바꾸시겠습니까?", 1, "&확인", "&취소")
                        Else
                            sReturnValue = Sbo_Application.MessageBox("해당 조건의 데이터를 저장하시겠습니까?", 1, "&확인", "&취소")
                        End If
                        
                        Select Case sReturnValue
                            Case 1:
                                If oRecordSet.RecordCount > 0 Then
                                    sQry = "Delete [ZFI020] Where "
                                    sQry = sQry & "Company = '" & Company & "' And "
                                    sQry = sQry & "DocType = '" & DocType & "' And "
                                    sQry = sQry & "AcctYear = '" & AcctYear & "' And "
                                    sQry = sQry & "AcctMon = '" & AcctMon & "'"
                                    oRecordSet.DoQuery sQry
                                End If
                                sQry = "EXEC PS_FI901_01 '" & Company & "', '" & AcctYear & "', '" & AcctMon & "', '" & DocType & "'"
                                oRecordSet.DoQuery sQry
                                MDC_Com.MDC_GF_Message "해당 조건의 데이터가 성공적으로 저장되었습니다. 데이터를 확인해보세요.", "S"
                                
                                sQry = "EXEC PS_FI901_02 '" & Company & "', '" & AcctYear & "', '" & AcctMon & "', '" & DocType & "', '" & BPLID & "'"
                                oForm01.DataSources.DataTables.Item(0).ExecuteQuery (sQry)
                                oGrid01.DataTable = oForm01.DataSources.DataTables.Item("Grid01")
                        
                                DrawGrid
                            Case 2:
                                Sbo_Application.MessageBox ("실행이 취소되었습니다.")
                                BubbleEvent = False
                                Exit Sub
                        End Select
                    ElseIf pval.ItemUID = "Btn02" Then
                        sQry = "EXEC PS_FI901_02 '" & Company & "', '" & AcctYear & "', '" & AcctMon & "', '" & DocType & "', '" & BPLID & "'"
                        oForm01.DataSources.DataTables.Item(0).ExecuteQuery (sQry)
                        oGrid01.DataTable = oForm01.DataSources.DataTables.Item("Grid01")
                        
                        DrawGrid
                    ElseIf pval.ItemUID = "Btn03" Then
                        sQry = "Select * From [ZFI020] Where "
                        sQry = sQry & "Company = '" & Company & "' And "
                        sQry = sQry & "DocType = '" & DocType & "' And "
                        sQry = sQry & "AcctYear = '" & AcctYear & "' And "
                        sQry = sQry & "AcctMon = '" & AcctMon & "'"
                        oRecordSet.DoQuery sQry
                        
                        If oRecordSet.RecordCount > 0 Then
                            sReturnValue = Sbo_Application.MessageBox("해당 조건의 데이터가 존재합니다. 삭제하시겠습니까?", 1, "&확인", "&취소")
                            Select Case sReturnValue
                                Case 1:
                                    sQry = "Delete [ZFI020] Where "
                                    sQry = sQry & "Company = '" & Company & "' And "
                                    sQry = sQry & "DocType = '" & DocType & "' And "
                                    sQry = sQry & "AcctYear = '" & AcctYear & "' And "
                                    sQry = sQry & "AcctMon = '" & AcctMon & "'"
                                    oRecordSet.DoQuery sQry
                                    MDC_Com.MDC_GF_Message "해당 조건의 데이터가 성공적으로 삭제되었습니다. 데이터를 확인해보세요.", "S"
                                    
                                    sQry = "EXEC PS_FI901_02 '" & Company & "', '" & AcctYear & "', '" & AcctMon & "', '" & DocType & "', '" & BPLID & "'"
                                    oForm01.DataSources.DataTables.Item(0).ExecuteQuery (sQry)
                                    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("Grid01")
                        
                                    DrawGrid
                                Case 2:
                                    Sbo_Application.MessageBox ("실행이 취소되었습니다.")
                                    BubbleEvent = False
                                    Exit Sub
                            End Select
                        Else
                            MDC_Com.MDC_GF_Message "해당 조건의 데이터가 존재하지 않습니다. 데이터를 확인해보세요.", "S"
                        End If
                    End If
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
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
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
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oGrid01 = Nothing
        End Select
    End If
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set oRecordSet = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
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

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    Dim oEdit           As SAPbouiCOM.EditText
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim optBtn          As SAPbouiCOM.OptionBtn
    
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 그리드 개체 할당
    Set oGrid01 = oForm01.Items("Grid01").Specific
    oForm01.DataSources.DataTables.Add ("Grid01")
    
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId"
    oRecordSet.DoQuery sQry
    oCombo.ValidValues.Add "", ""
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim(oRecordSet.Fields(0).VALUE), Trim(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop

    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '// DocType
    Set oCombo = oForm01.Items("DocType").Specific
    oCombo.ValidValues.Add "1", "매입"
    oCombo.ValidValues.Add "2", "매출"
    oCombo.ValidValues.Add "3", "채무"
    oCombo.ValidValues.Add "4", "채권"
    
    '// 회사
    oForm01.Items("Company").Specific.VALUE = "PSH"
    
    If Mid(Now, 6, 2) = "01" Then
        oForm01.Items("AcctYear").Specific.VALUE = Left(Now, 4) - 1
        oForm01.Items("AcctMon").Specific.VALUE = "12"
    Else
        oForm01.Items("AcctYear").Specific.VALUE = Left(Now, 4)
        oForm01.Items("AcctMon").Specific.VALUE = Format(Mid(Now, 6, 2) - 1, "00")
    End If
        
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
   Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "CreateItems_Error:" & Err.Description, bmt_Short, True
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String
    
 '/ Check
    Select Case True
        Case oForm01.Items("DocType").Specific.Selected Is Nothing
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oForm01.Items("Company").Specific.VALUE = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oForm01.Items("AcctYear").Specific.VALUE = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case oForm01.Items("AcctYear").Specific.VALUE < 2008 Or Len(oForm01.Items("AcctYear").Specific.VALUE) <> 4
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case oForm01.Items("AcctMon").Specific.VALUE = ""
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
        Case oForm01.Items("AcctMon").Specific.VALUE < 1 Or oForm01.Items("AcctMon").Specific.VALUE > 12 Or Len(oForm01.Items("AcctMon").Specific.VALUE) <> 2
            ErrNum = 6
            GoTo HeaderSpaceLineDel_Error
    End Select
        
    HeaderSpaceLineDel = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
'    If ErrNum = 1 Then
'        ZPSH_SetFunc.PSH_Message "구분은 필수입력사항입니다. 입력하여 주십시오.", "E"
'    ElseIf ErrNum = 2 Then
'        ZPSH_SetFunc.PSH_Message "회사는 필수입력사항입니다. 입력하여 주십시오.", "E"
'    ElseIf ErrNum = 3 Then
'        ZPSH_SetFunc.PSH_Message "회계년도는 필수입력사항입니다. 입력하여 주십시오.", "E"
'    ElseIf ErrNum = 4 Then
'        ZPSH_SetFunc.PSH_Message "회계년도를 정확하게 입력하여 주십시오.", "E"
'    ElseIf ErrNum = 5 Then
'        ZPSH_SetFunc.PSH_Message "회계월은 필수입력사항입니다. 입력하여 주십시오.", "E"
'    ElseIf ErrNum = 6 Then
'        ZPSH_SetFunc.PSH_Message "회계월을 정확하게 입력하여 주십시오.", "E"
'    Else
'        ZPSH_SetFunc.PSH_Message "HeaderSpaceLineDel_Error:" & Err.Description, "E"
'    End If
    HeaderSpaceLineDel = False
End Function

'****************************************************************************************************************
'// Grid 꾸며주기
'****************************************************************************************************************
Private Sub DrawGrid()
On Error GoTo DrawGrid_Error
    Dim i                   As Integer
    Dim sColsTitle          As String
    
    oGrid01.SelectionMode = ms_Auto
    For i = 0 To oGrid01.Columns.Count - 1
        sColsTitle = oGrid01.Columns(i).TitleObject.Caption
    
        If oGrid01.DataTable.Columns(i).Type = ft_Float Then
            oGrid01.Columns(i).RightJustified = True
        End If
    Next i
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
DrawGrid_Error:
'    ZPSH_SetFunc.PSH_Message "DrawGrid_Error:" & Err.Description, "E"
End Sub


