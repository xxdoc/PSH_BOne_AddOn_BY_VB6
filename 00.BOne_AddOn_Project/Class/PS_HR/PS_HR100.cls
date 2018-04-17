VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_HR100.cls
'//  Module : 인사관리>애드온>사규위반 업무처리>조치등록>1.사규위반조치등록
'//  Desc :
'//  FormType : PS_HR100
'//  Create Date : 2011.11.10
'//  Modified Date :
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_HR100H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_HR100L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    BatchNum As String
    Quantity As Long
    OPORNo As Long
    POR1No As Long
    Check As Boolean
    OPDNNo As Long
    PDN1No As Long
End Type
Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long

Private oDocType01 As String
Private oDocEntry01 As String
Private oFormMode01 As SAPbouiCOM.BoFormMode

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR100.srf")
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
    
    oFormUniqueID01 = "PS_HR100_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_HR100_CreateItems
    Call PS_HR100_ComboBox_Setting
    Call PS_HR100_CF_ChooseFromList
    Call PS_HR100_EnableMenus
    Call PS_HR100_SetDocument(oFromDocEntry01)
    Call PS_HR100_FormResize
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            Call Raise_EVENT_ITEM_PRESSED(FormUID, pval, BubbleEvent)
        Case et_KEY_DOWN: '//2
            Call Raise_EVENT_KEY_DOWN(FormUID, pval, BubbleEvent)
        Case et_COMBO_SELECT: '//5
            Call Raise_EVENT_COMBO_SELECT(FormUID, pval, BubbleEvent)
        Case et_CLICK: '//6
            Call Raise_EVENT_CLICK(FormUID, pval, BubbleEvent)
        Case et_DOUBLE_CLICK: '//7
            Call Raise_EVENT_DOUBLE_CLICK(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LINK_PRESSED '//8
            Call Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, pval, BubbleEvent)
        Case et_VALIDATE: '//10
            Call Raise_EVENT_VALIDATE(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LOAD: '//11
            Call Raise_EVENT_MATRIX_LOAD(FormUID, pval, BubbleEvent)
        Case et_FORM_ACTIVATE: '//18
            '//et_FORM_ACTIVATE
        Case et_FORM_DEACTIVATE: '//19
            '//et_FORM_DEACTIVATE
        Case et_FORM_RESIZE '//20
            Call Raise_EVENT_RESIZE(FormUID, pval, BubbleEvent)
        Case et_CHOOSE_FROM_LIST '//27
            Call Raise_EVENT_CHOOSE_FROM_LIST(FormUID, pval, BubbleEvent)
        Case et_GOT_FOCUS: '//3
            Call Raise_EVENT_GOT_FOCUS(FormUID, pval, BubbleEvent)
        Case et_LOST_FOCUS: '//4
            '//et_LOST_FOCUS
        Case et_FORM_UNLOAD: '//17
            Call Raise_EVENT_FORM_UNLOAD(FormUID, pval, BubbleEvent)
    End Select
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_HR100_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_HR100_FormItemEnabled '//UDO방식
                Call PS_HR100_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_HR100_FormItemEnabled
        End Select
    End If
    Exit Sub
Raise_MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
Raise_FormDataEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
    
    End If
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    Else
        oLastItemUID01 = pval.ItemUID
        oLastColUID01 = ""
        oLastColRow01 = 0
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "PS_HR100" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_HR100_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_HR100_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_HR100" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_HR100_FormItemEnabled
                    Call PS_HR100_AddMatrixRow(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_HR100_FormItemEnabled
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ITEM_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error
    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.ColUID = "VioCode" Then
            
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "VioCode")
                
            ElseIf pval.ColUID = "CodeLv1" Then
            
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "CodeLv1")
                
            ElseIf pval.ColUID = "CodeLv2" Then
            
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "CodeLv2")
                
            ElseIf pval.ColUID = "CodeLv3" Then
            
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "CodeLv3")
            
            End If
        
        ElseIf pval.ItemUID = "CrCode" Then
        
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CrCode", "")
        
        End If

    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemChanged = True Then
            
        End If
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error
    If pval.BeforeAction = True Then

        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat01.SelectRow(pval.Row, True, False)
            End If
        Else
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = ""
            oLastColRow01 = 0
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_DOUBLE_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_DOUBLE_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LINK_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LINK_PRESSED_Error
    Dim oTempClass As Variant
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
        
            Call oMat01.FlushToDataSource
            
            If pval.ItemUID = "Mat01" Then
            
                If pval.ColUID = "VioCode" Then
                
                    Call oDS_PS_HR100L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    
                    Query01 = "              EXEC PS_HR100_01 '"
                    Query01 = Query01 & oForm01.Items("BPLId").Specific.VALUE & "','" '사업장
                    Query01 = Query01 & oForm01.Items("GrpCode").Specific.VALUE & "','" '조치구분
                    Query01 = Query01 & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'" '사규위반자 사번
                    
                    Call RecordSet01.DoQuery(Query01)
                    Call oDS_PS_HR100L.setValue("U_VioName", pval.Row - 1, Trim(RecordSet01.Fields("FullName").VALUE))
                    Call oDS_PS_HR100L.setValue("U_ActGd", pval.Row - 1, Trim(RecordSet01.Fields("NextActGd").VALUE))
                    
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_HR100L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_HR100_AddMatrixRow (pval.Row)
                    End If
            
                ElseIf pval.ColUID = "CodeLv1" Then
                    
                    Query01 = "                SELECT   B.U_NameLv1"
                    Query01 = Query01 & " FROM    [@PS_HR000H] AS A"
                    Query01 = Query01 & "             INNER JOIN"
                    Query01 = Query01 & "             [@PS_HR000L] AS B"
                    Query01 = Query01 & "                 ON A.DocEntry = B.DocEntry"
                    Query01 = Query01 & " WHERE   A.Canceled = 'N'"
                    Query01 = Query01 & "             AND B.U_UseYN = 'Y'"
                    Query01 = Query01 & "             AND A.U_GrpCode = '" & oForm01.Items("GrpCode").Specific.VALUE & "'"
                    Query01 = Query01 & "             AND B.U_CodeLv1 = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"
                    
                    Call RecordSet01.DoQuery(Query01)
                    Call oDS_PS_HR100L.setValue("U_NameLv1", pval.Row - 1, Trim(RecordSet01.Fields(0).VALUE))
                    
                ElseIf pval.ColUID = "CodeLv2" Then
                
                    Query01 = "                SELECT    B.U_NameLv2"
                    Query01 = Query01 & " FROM     [@PS_HR001H] AS A"
                    Query01 = Query01 & "              INNER JOIN"
                    Query01 = Query01 & "              [@PS_HR001L] AS B"
                    Query01 = Query01 & "                  ON A.DocEntry = B.DocEntry"
                    Query01 = Query01 & " WHERE    A.Canceled = 'N'"
                    Query01 = Query01 & "              AND B.U_UseYN = 'Y'"
                    Query01 = Query01 & "              AND A.U_GrpCode = '" & oForm01.Items("GrpCode").Specific.VALUE & "'"
                    Query01 = Query01 & "              AND A.U_CodeLv1 = '" & oMat01.Columns("CodeLv1").Cells(pval.Row).Specific.VALUE & "'"
                    Query01 = Query01 & "              AND B.U_CodeLv2 = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"
                
                    Call RecordSet01.DoQuery(Query01)
                    Call oDS_PS_HR100L.setValue("U_NameLv2", pval.Row - 1, Trim(RecordSet01.Fields(0).VALUE))
                
                ElseIf pval.ColUID = "CodeLv3" Then
                
                    Query01 = "                SELECT  B.U_NameLv3"
                    Query01 = Query01 & " FROM    [@PS_HR002H] AS A"
                    Query01 = Query01 & "             INNER JOIN"
                    Query01 = Query01 & "             [@PS_HR002L] AS B"
                    Query01 = Query01 & "                 ON A.DocEntry = B.DocEntry"
                    Query01 = Query01 & " WHERE   A.Canceled = 'N'"
                    Query01 = Query01 & "             AND B.U_UseYN = 'Y'"
                    Query01 = Query01 & "             AND A.U_GrpCode = '" & oForm01.Items("GrpCode").Specific.VALUE & "'"
                    Query01 = Query01 & "             AND A.U_CodeLv1 = '" & oMat01.Columns("CodeLv1").Cells(pval.Row).Specific.VALUE & "'"
                    Query01 = Query01 & "             AND A.U_CodeLv2 = '" & oMat01.Columns("CodeLv2").Cells(pval.Row).Specific.VALUE & "'"
                    Query01 = Query01 & "             AND B.U_CodeLv3 = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"
                
                    Call RecordSet01.DoQuery(Query01)
                    Call oDS_PS_HR100L.setValue("U_NameLv3", pval.Row - 1, Trim(RecordSet01.Fields(0).VALUE))
                
                Else
                    Call oDS_PS_HR100L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
            
            ElseIf pval.ItemUID = "CrCode" Then
            
                Query01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                Call RecordSet01.DoQuery(Query01)
                oForm01.Items("CrName").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
            
            End If
            
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oForm01.Update

        End If
    ElseIf pval.BeforeAction = False Then

    End If
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_HR100_FormItemEnabled
        Call PS_HR100_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        Call oMat01.AutoResizeColumns
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_HR100_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    Else
        oLastItemUID01 = pval.ItemUID
        oLastColUID01 = ""
        oLastColRow01 = 0
    End If
    Exit Sub
Raise_EVENT_GOT_FOCUS_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_GOT_FOCUS_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_FORM_UNLOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_FORM_UNLOAD_Error
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
        RemoveForms oFormUniqueID01
        Set oForm01 = Nothing
        Set oMat01 = Nothing
    End If
    Exit Sub
Raise_EVENT_FORM_UNLOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_UNLOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
'            If (PS_HR100_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_HR100L.RemoveRecord(oDS_PS_HR100L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_HR100_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_HR100L.GetValue("U_CodeLv3", oMat01.RowCount - 1)) <> "" Then
                    Call PS_HR100_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_HR100_CreateItems() As Boolean
On Error GoTo PS_HR100_CreateItems_Error
   
    Set oDS_PS_HR100H = oForm01.DataSources.DBDataSources("@PS_HR100H")
    Set oDS_PS_HR100L = oForm01.DataSources.DBDataSources("@PS_HR100L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.SelectionMode = ms_NotSupported
'    oMat01.AutoResizeColumns
    
    Exit Function
PS_HR100_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR100_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_HR100_ComboBox_Setting()
On Error GoTo PS_HR100_ComboBox_Setting_Error
        
    '//콤보에 기본값설정
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
            
    sQry = "           SELECT     A.U_GrpCode,"
    sQry = sQry & "               A.U_GrpName"
    sQry = sQry & " FROM      [@PS_HR000H] AS A"
    sQry = sQry & " WHERE     A.Canceled = 'N'"
    sQry = sQry & " ORDER BY A.U_GrpCode"
    Call oForm01.Items("GrpCode").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("GrpCode").Specific, sQry, "", False, False)
    Call oForm01.Items("GrpCode").Specific.Select(0, psk_Index)

    '징계양정_S
    sQry = "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'H001'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ActGd"), sQry)
    '징계양정_E

    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Exit Sub
PS_HR100_ComboBox_Setting_Error:
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_HR100_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR100_CF_ChooseFromList()
On Error GoTo PS_HR100_CF_ChooseFromList_Error
'    '//ChooseFromList 설정
'    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
'    Dim oCons               As SAPbouiCOM.Conditions
'    Dim oCon                As SAPbouiCOM.Condition
'    Dim oCFL                As SAPbouiCOM.ChooseFromList
'    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
'    Dim oEdit               As SAPbouiCOM.EditText
'    Dim oColumn             As SAPbouiCOM.Column
'
'    Set oEdit = oForm01.Items("CARDCODE").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_BusinessPartner
'    oCFLCreationParams.uniqueID = "CFLCARDCODE"
'    oCFLCreationParams.MultiSelection = False
''    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLCARDCODE"
'    oEdit.ChooseFromListAlias = "CardCode"
'
'    Set oEdit = oForm01.Items("CARDNAME").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_BusinessPartner
'    oCFLCreationParams.uniqueID = "CFLCARDNAME"
'    oCFLCreationParams.MultiSelection = False
'    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLCARDNAME"
'    oEdit.ChooseFromListAlias = "CardName"
    Exit Sub
PS_HR100_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR100_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR100_FormItemEnabled()
On Error GoTo PS_HR100_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
    
        oForm01.Items("GrpCode").Enabled = True

        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True

        Call PS_HR100_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        
        oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
        oForm01.Items("CrCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
    
        oForm01.Items("GrpCode").Enabled = True
        
        oForm01.Items("DocEntry").Specific.VALUE = ""
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
    
        oForm01.Items("GrpCode").Enabled = False
        
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR100_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR100_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR100_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_HR100_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_HR100L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_HR100L.Offset = oRow
    oDS_PS_HR100L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR100_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR100_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR100_FormClear()
On Error GoTo PS_HR100_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_HR100'", "")
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_HR100_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR100_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR100_EnableMenus()
On Error GoTo PS_HR100_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_HR100_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR100_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR100_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_HR100_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_HR100_FormItemEnabled
        Call PS_HR100_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_HR100_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_HR100_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR100_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_HR100_DataValidCheck() As Boolean
On Error GoTo PS_HR100_DataValidCheck_Error
    PS_HR100_DataValidCheck = False
    Dim i As Long
    Dim j As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_HR100_FormClear
    End If
    
    '조치구분 미입력 시
    If oForm01.Items("GrpCode").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("조치구분이 선택되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("GrpCode").Click(ct_Regular)
        PS_HR100_DataValidCheck = False
        Exit Function
    End If
    
    '등록자 미입력 시
    If oForm01.Items("CrCode").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("등록자가 입력되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("CrCode").Click(ct_Regular)
        PS_HR100_DataValidCheck = False
        Exit Function
    End If
    
    '등록일 미입력 시
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("등록일이 입력되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("DocDate").Click(ct_Regular)
        PS_HR100_DataValidCheck = False
        Exit Function
    End If
    
    '라인정보 미입력 시
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_HR100_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        '위반자사번
        If (oMat01.Columns("VioCode").Cells(i).Specific.VALUE = "") Then
            Call Sbo_Application.SetStatusBarMessage("위반자는 필수입니다.", bmt_Short, True)
            oMat01.Columns("VioCode").Cells(i).Click ct_Regular
            PS_HR100_DataValidCheck = False
            Exit Function
        End If
        
        '장소
        If (oMat01.Columns("Location").Cells(i).Specific.VALUE = "") Then
            Call Sbo_Application.SetStatusBarMessage("장소는 필수입니다.", bmt_Short, True)
            oMat01.Columns("Location").Cells(i).Click ct_Regular
            PS_HR100_DataValidCheck = False
            Exit Function
        End If
        
        '항목
        If (oMat01.Columns("CodeLv1").Cells(i).Specific.VALUE = "") Then
            Call Sbo_Application.SetStatusBarMessage("항목은 필수입니다.", bmt_Short, True)
            oMat01.Columns("CodeLv1").Cells(i).Click ct_Regular
            PS_HR100_DataValidCheck = False
            Exit Function
        End If
        
'        '세부사항
'        If (oMat01.Columns("CodeLv2").Cells(i).Specific.VALUE = "") Then
'            Call Sbo_Application.SetStatusBarMessage("세부사항은 필수입니다.", bmt_Short, True)
'            oMat01.Columns("CodeLv2").Cells(i).Click ct_Regular
'            PS_HR100_DataValidCheck = False
'            Exit Function
'        End If
        
        '조치일자
        If (oMat01.Columns("GrpDate").Cells(i).Specific.VALUE = "") Then
            Call Sbo_Application.SetStatusBarMessage("조치일자는 필수입니다.", bmt_Short, True)
            oMat01.Columns("GrpDate").Cells(i).Click ct_Regular
            PS_HR100_DataValidCheck = False
            Exit Function
        End If
        
    Next
    
    Call oMat01.FlushToDataSource
    Call oDS_PS_HR100L.RemoveRecord(oDS_PS_HR100L.Size - 1)
    Call oMat01.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_HR100_FormClear
    End If
    
    PS_HR100_DataValidCheck = True
    Exit Function
PS_HR100_DataValidCheck_Error:
    PS_HR100_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_HR100_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_HR100_MTX01()
On Error GoTo PS_HR100_MTX01_Error
    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Param01 = Trim(oForm01.Items("Param01").Specific.VALUE)
    Param02 = Trim(oForm01.Items("Param01").Specific.VALUE)
    Param03 = Trim(oForm01.Items("Param01").Specific.VALUE)
    Param04 = Trim(oForm01.Items("Param01").Specific.VALUE)
    
    Query01 = "SELECT 10"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_HR100_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_HR100L.InsertRecord (i)
        End If
        oDS_PS_HR100L.Offset = i
        Call oDS_PS_HR100L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_HR100L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR100_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_HR100_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR100_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_HR100_DI_API() As Boolean
On Error GoTo PS_HR100_DI_API_Error
    PS_HR100_DI_API = True
    Dim i, j As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim LineNumCount As Long
    Dim ResultDocNum As Long
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    For i = 1 To oMat01.VisualRowCount
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.VALUE, "&&&&-&&-&&")
    For i = 0 To ItemInformationCount - 1
        If ItemInformation(i).Check = True Then
            GoTo Continue_First
        End If
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = Trim(oForm01.Items("WhsCode").Specific.VALUE)
        oDIObject.Lines.BaseType = "22"
        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
        For j = i To UBound(ItemInformation)
            If ItemInformation(j).Check = True Then
                GoTo Continue_Second
            End If
            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
                GoTo Continue_Second
            End If
            '//같은것
            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).BatchNum
            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.Add
            ItemInformation(j).PDN1No = LineNumCount
            ItemInformation(j).Check = True
Continue_Second:
        Next
        LineNumCount = LineNumCount + 1
Continue_First:
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To UBound(ItemInformation)
            Call oDS_PS_HR100L.setValue("U_OPDNNo", i, ResultDocNum)
            Call oDS_PS_HR100L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
        Next
    Else
        GoTo PS_HR100_DI_API_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_HR100_DI_API_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_HR100_DI_API = False
    Set oDIObject = Nothing
    Exit Function
PS_HR100_DI_API_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR100_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_HR100_DI_API = False
End Function

Private Sub PS_HR100_FormResize()
On Error GoTo PS_HR100_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_HR100_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR100_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_HR100_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_HR100_Validate_Error
    PS_HR100_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_HR100H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
'        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
'        PS_HR100_Validate = False
'        GoTo PS_HR100_Validate_Exit
'    End If
'
    If ValidateType = "수정" Then
'        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
'        Dim Exist As Boolean
'        Exist = False
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            Exist = False
'            For j = 1 To oMat01.RowCount - 1
'                '//라인번호가 같고, 품목코드가 같으면 존재하는행 , LineNum에 값이 존재하는지 확인필요(행삭제된행인경우 LineNum이 존재하지않음)
'                If Val(RecordSet01.Fields(1).Value) = Val(oMat01.Columns("U_LineNum").Cells(j).Specific.Value) And RecordSet01.Fields(2).Value = oMat01.Columns("1").Cells(j).Specific.Value And oMat01.Columns("U_LineNum").Cells(j).Specific.Value <> "" Then
'                    Exist = True
'                End If
'            Next
'            If (Exist = False) Then '//삭제된 행중
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "삭제된행이 다른사용자에 의해 출하,선출요청되었습니다. 적용할수 없습니다.", "W"
'                    PS_HR100_Validate = False
'                    GoTo PS_HR100_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
'        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
'            If (oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
'            Else
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(oForm01.Items("8").Specific.Value) & "' AND U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 삭제할수 없습니다.", "W"
'                    PS_HR100_Validate = False
'                    GoTo PS_HR100_Validate_Exit
'                End If
'            End If
'        End If
    ElseIf ValidateType = "취소" Then
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할수 없습니다.", "W"
'                PS_HR100_Validate = False
'                GoTo PS_HR100_Validate_Exit
'            End If
'            RecordSet01.MoveNext
'        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_HR100_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_HR100_Validate_Error:
    PS_HR100_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_HR100_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'이하 소스는 참조용으로 보존(2011.11.15 송명규)
Public Sub PS_HR100_GetCodeLv1(ByVal pRow As Integer)
On Error GoTo PS_HR100_GetCodeLv1_Error

    Dim l_QueryString As String
    Dim loopCount As Integer
    
    l_QueryString = "                      SELECT    B.U_CodeLv1,"
    l_QueryString = l_QueryString & "              B.U_NameLv1,"
    l_QueryString = l_QueryString & "              2 AS [Sort]"
    l_QueryString = l_QueryString & " FROM     [@PS_HR000H] AS A"
    l_QueryString = l_QueryString & "              INNER JOIN"
    l_QueryString = l_QueryString & "              [@PS_HR000L] AS B"
    l_QueryString = l_QueryString & "                  ON A.DocEntry = B.DocEntry"
    l_QueryString = l_QueryString & " WHERE    A.Canceled = 'N'"
    l_QueryString = l_QueryString & "              AND B.U_UseYN = 'Y'"
    l_QueryString = l_QueryString & "              AND A.U_GrpCode = '" & oForm01.Items("GrpCode").Specific.VALUE & "'"
    l_QueryString = l_QueryString & " UNION ALL"
    l_QueryString = l_QueryString & " SELECT    '%',"
    l_QueryString = l_QueryString & "              '선택',"
    l_QueryString = l_QueryString & "              1 AS [Sort]"
    l_QueryString = l_QueryString & " ORDER BY Sort"
    
    
    Call oDS_PS_HR100H.setValue("U_GrpCode", 0, oForm01.Items("GrpCode").Specific.VALUE)
    
'    For loopCount = 1 To oMat01.VisualRowCount
    
        Call MDC_SetMod.Remove_ComboList(oMat01.Columns("CodeLv1").Cells(pRow).Specific)
'        Call MDC_SetMod.Remove_ComboList(oMat01.Columns("CodeLv2").Cells(pRow).Specific)
'        Call MDC_SetMod.Remove_ComboList(oMat01.Columns("CodeLv3").Cells(pRow).Specific)
        
        Call MDC_SetMod.Set_ComboList(oMat01.Columns("CodeLv1").Cells(pRow).Specific, l_QueryString, "", False, False)
    
        If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
            Call oMat01.Columns("CodeLv1").Cells(pRow).Specific.Select(0, psk_Index)
        End If
    
'    Next
    
    Exit Sub
    
PS_HR100_GetCodeLv1_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR100_GetCodeLv1_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_HR100_GetCodeLv2(ByVal pRow As Integer)
On Error GoTo PS_HR100_GetCodeLv2_Error

    Dim l_QueryString As String
    Dim loopCount As Integer
    
'    For loopCount = 1 To oMat01.VisualRowCount
    
        l_QueryString = "                      SELECT  B.U_CodeLv2,"
        l_QueryString = l_QueryString & "            B.U_NameLv2,"
        l_QueryString = l_QueryString & "            2 AS [Sort]"
        l_QueryString = l_QueryString & " FROM   [@PS_HR001H] AS A"
        l_QueryString = l_QueryString & "            INNER JOIN"
        l_QueryString = l_QueryString & "            [@PS_HR001L] AS B"
        l_QueryString = l_QueryString & "                ON A.DocEntry = B.DocEntry"
        l_QueryString = l_QueryString & " WHERE  A.Canceled = 'N'"
        l_QueryString = l_QueryString & "            AND B.U_UseYN = 'Y'"
        l_QueryString = l_QueryString & "            AND A.U_GrpCode = '" & oForm01.Items("GrpCode").Specific.VALUE & "'"
        l_QueryString = l_QueryString & "            AND A.U_CodeLv1 = '" & oMat01.Columns("CodeLv1").Cells(pRow).Specific.VALUE & "'"
        l_QueryString = l_QueryString & " UNION ALL"
        l_QueryString = l_QueryString & " SELECT    '%',"
        l_QueryString = l_QueryString & "              '선택',"
        l_QueryString = l_QueryString & "              1 AS [Sort]"
        l_QueryString = l_QueryString & " ORDER BY Sort"
        
        Call MDC_SetMod.Remove_ComboList(oMat01.Columns("CodeLv2").Cells(pRow).Specific)
'        Call MDC_SetMod.Remove_ComboList(oMat01.Columns("CodeLv3").Cells(pRow).Specific)
        
        Call MDC_SetMod.Set_ComboList(oMat01.Columns("CodeLv2").Cells(pRow).Specific, l_QueryString, "", False, False)
    
        If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
            Call oMat01.Columns("CodeLv2").Cells(pRow).Specific.Select(0, psk_Index)
        End If
    
'    Next
    
    Exit Sub
    
PS_HR100_GetCodeLv2_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR100_GetCodeLv2_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_HR100_GetCodeLv3()
On Error GoTo PS_HR100_GetCodeLv3_Error

    Dim l_QueryString As String
    Dim loopCount As Integer
    
    For loopCount = 1 To oMat01.VisualRowCount
    
        l_QueryString = "                      SELECT  B.U_CodeLv3,"
        l_QueryString = l_QueryString & "            B.U_NameLv3,"
        l_QueryString = l_QueryString & "            2 AS [Sort]"
        l_QueryString = l_QueryString & " FROM   [@PS_HR002H] AS A"
        l_QueryString = l_QueryString & "            INNER JOIN"
        l_QueryString = l_QueryString & "            [@PS_HR002L] AS B"
        l_QueryString = l_QueryString & "                ON A.DocEntry = B.DocEntry"
        l_QueryString = l_QueryString & " WHERE  A.Canceled = 'N'"
        l_QueryString = l_QueryString & "            AND B.U_UseYN = 'Y'"
        l_QueryString = l_QueryString & "            AND A.U_GrpCode = '" & oForm01.Items("GrpCode").Specific.VALUE & "'"
        l_QueryString = l_QueryString & "            AND A.U_CodeLv1 = '" & oMat01.Columns("CodeLv1").Cells(loopCount).Specific.VALUE & "'"
        l_QueryString = l_QueryString & "            AND A.U_CodeLv2 = '" & oMat01.Columns("CodeLv2").Cells(loopCount).Specific.VALUE & "'"
        l_QueryString = l_QueryString & " UNION ALL"
        l_QueryString = l_QueryString & " SELECT   '%',"
        l_QueryString = l_QueryString & "             '선택',"
        l_QueryString = l_QueryString & "             1 AS [Sort]"
        l_QueryString = l_QueryString & " ORDER BY Sort"
        
        Call MDC_SetMod.Remove_ComboList(oMat01.Columns("CodeLv3").Cells(loopCount).Specific)
        Call MDC_SetMod.Set_ComboList(oMat01.Columns("CodeLv3").Cells(loopCount).Specific, l_QueryString, "", False, False)
    
        If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
            Call oMat01.Columns("CodeLv3").Cells(loopCount).Specific.Select(0, psk_Index)
        End If
    
    Next
    
    Exit Sub
    
PS_HR100_GetCodeLv3_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR100_GetCodeLv3_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub




