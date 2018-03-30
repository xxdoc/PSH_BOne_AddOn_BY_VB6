VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP071"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP071.cls
'//  Module         : PP
'//  Description    : 작지조회
'//  FormType       : PS_PP071
'//  Create Date    : 2010.12.23
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Public oGrid01              As SAPbouiCOM.Grid
Private oDS_PS_PP071H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_PP071L       As SAPbouiCOM.DBDataSource    '등록라인

Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01  As String
Public oBaseColRow01  As Long
Public oBaseBPLId01   As String
Public oBaseOrdGbn01  As String

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oBPLId02 As String, Optional ByVal oOrdGbn02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP071.srf")
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
    
    oFormUniqueID01 = "PS_PP071_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
'    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    
    Set oBaseForm01 = oForm02
    oBaseItemUID01 = oItemUID02
    oBaseColUID01 = oColUID02
    oBaseColRow01 = oColRow02
    oBaseBPLId01 = oBPLId02
    oBaseOrdGbn01 = oOrdGbn02
        
    Call CreateItems
    Call ComboBox_Setting
'    FormItemEnabled
        
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
'    Set oDS_PS_PP071H = oForm01.DataSources.DBDataSources("@PS_PP071H")
'    Set oDS_PS_PP071L = oForm01.DataSources.DBDataSources("@PS_PP071L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    oGrid01.SelectionMode = ms_Auto
    oForm01.DataSources.DataTables.Add ("ZTEMP")
    
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
    
    Call oForm01.Items("ItemType").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType").Specific, "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", "", False, False)
    Call oForm01.Items("ItemType").Specific.Select(0, psk_Index)

    Call oForm01.Items("Mark").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", "", False, False)
    Call oForm01.Items("Mark").Specific.Select(0, psk_Index)
    
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
                If pval.ItemUID = "Button01" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        Call Search_Grid_Data
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                    ElseIf oForm01.Mode = fm_OK_MODE Then
                    End If
                End If
                If pval.ItemUID = "Button02" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        Call PS_PP071_SetBaseForm '//부모폼에입력
                        Call oForm01.Close
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                    ElseIf oForm01.Mode = fm_OK_MODE Then
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
                    If pval.ItemUID = "Grid01" Then
                        If pval.Row = -1 Then
                            oGrid01.Columns(pval.ColUID).TitleObject.Sortable = True
                        Else
                            If oGrid01.Rows.SelectedRows.Count > 0 Then
                                Call PS_PP071_SetBaseForm '//부모폼에입력
                                Call oForm01.Close
                            Else
                                BubbleEvent = False
                            End If
                        End If
                    End If
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
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oGrid01 = Nothing
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
    Dim i&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
    
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

Public Sub Search_Grid_Data()
    On Error GoTo Search_Grid_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oForm01.Freeze True
    
    Param01 = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    Param02 = oForm01.Items("ItemName").Specific.VALUE
    Param03 = oForm01.Items("Size").Specific.VALUE
    Param04 = oForm01.Items("ItemType").Specific.Selected.VALUE
    Param05 = oForm01.Items("Mark").Specific.Selected.VALUE
    Param06 = oForm01.Items("OrdNum").Specific.VALUE
    
    
    'sQry = "EXEC PS_PP071_01  '" & oBaseBPLId01 & "','" & oBaseOrdGbn01 & "','" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "', '" & Param05 & "'"
    sQry = "EXEC PS_PP070_03  '" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "', '" & Param05 & "', '" & Param06 & "'"
    'sQry = "EXEC PS_PP071_01  '" & 1 & "','" & 101 & "','" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "', '" & Param05 & "'"
        
    '/ Procedure 실행(Grid 사용)
    oForm01.DataSources.DataTables.Item(0).ExecuteQuery (sQry)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP")
    
    Call GridSetting
    
    
    oForm01.Freeze False
    
    Set oRecordSet = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Grid_Data_Error:
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "Search_Grid_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


'****************************************************************************************************************
'// Grid 꾸며주기
'****************************************************************************************************************
Private Sub GridSetting()
On Error GoTo GridSetting_Error
    Dim i           As Integer
    Dim sColsTitle  As String
    Dim sColsLine   As String
    
    oForm01.Freeze True
    
    oGrid01.SelectionMode = ms_Auto
    
'    oGrid01.Columns.Item(0).LinkedObjectType = lf_BusinessPartner
'    oGrid01.Columns.Item(4).LinkedObjectType = lf_Invoice
'    oGrid01.Columns.Item(2).LinkedObjectType = lf_Items
       
    For i = 0 To oGrid01.Columns.Count - 1
        sColsTitle = oGrid01.Columns(i).TitleObject.Caption
        
        oGrid01.Columns(i).Editable = False
        
        If oGrid01.DataTable.Columns(i).Type = ft_Float Then
            oGrid01.Columns(i).RightJustified = True
        End If
        
'        If InStr(1, sColsTitle, "T(mm)") > 0 Or _
'           InStr(1, sColsTitle, "W(mm)") > 0 Or _
'           InStr(1, sColsTitle, "L(mm)") > 0 Or _
'           InStr(1, sColsTitle, "Weight") > 0 Or _
'           InStr(1, sColsTitle, "Qty(Kg)") > 0 Then
'            oGrid01.Columns(i).RightJustified = True
'            oGrid01.Columns(i).BackColor = &HE0E0E0
'        End If
'
'        If InStr(1, sColsTitle, "수량") > 0 Or _
'           InStr(1, sColsTitle, "중량") > 0 Then
'            oGrid01.Columns(i).RightJustified = True
'            oGrid01.Columns(i).BackColor = &HFFC0C0
'        End If
    Next i
        
    oForm01.Freeze False
    
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
GridSetting_Error:
    MDC_Com.MDC_GF_Message "GridSetting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_PP071_SetBaseForm()
On Error GoTo PS_PP071_SetBaseForm_Error
    Dim i As Long
    Dim ItemCode01 As String
    Dim oBaseMat01 As Matrix
    If oBaseForm01 Is Nothing Then
    '//DoNothing
    ElseIf oBaseForm01.TypeEx = "PS_PP070" Then '//사용하려는폼의 폼타입
        Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
        For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
            oBaseMat01.Columns("PP030No").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("문서번호").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
            oBaseColRow01 = oBaseColRow01 + 1
        Next
    ElseIf oBaseForm01.TypeEx = "PS_PP080" Then '//사용하려는폼의 폼타입
        Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
        For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
            oBaseMat01.Columns("PP030No").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("문서번호").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
            oBaseColRow01 = oBaseColRow01 + 1
        Next
'    ElseIf oBaseForm01.TypeEx = "PS_QM010" Then '//사용하려는폼의 폼타입
'        oBaseForm01.Items("ItemCode").Specific.Value = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).Value
'    ElseIf oBaseForm01.TypeEx = "PS_PP077" Then '//사용하려는폼의 폼타입
'        oBaseForm01.Items("ItemCode").Specific.Value = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).Value
'    ElseIf oBaseForm01.TypeEx = "PS_PP078" Then '//사용하려는폼의 폼타입
'        oBaseForm01.Items("ItemCode").Specific.Value = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).Value
    End If
    Exit Sub
PS_PP071_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP071_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub




