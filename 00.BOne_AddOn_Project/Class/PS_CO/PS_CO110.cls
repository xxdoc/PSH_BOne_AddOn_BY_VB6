VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_CO110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_CO110.cls
'//  Module         : CO
'//  Description    : 배부사이클 실행
'//  FormType       : PS_CO110
'//  Create Date    : 2010.11.17
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PS_CO110H  As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_CO110L  As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID  As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row  As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_CO110.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_CO110_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
  
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "Code"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Add_MatrixRow(0, True)
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), False        '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
        
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

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error

    Dim i&
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim Calculate_Weight As Double
    Dim SelectedLineNum&
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_CO111
    
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim Code$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If oForm01.Mode = fm_ADD_MODE Then
                            Code = Trim(oDS_PS_CO110H.GetValue("U_YM", 0)) & Trim(oDS_PS_CO110H.GetValue("U_BPLId", 0)) & Trim(oDS_PS_CO110H.GetValue("U_Type", 0))
                            Call oDS_PS_CO110H.setValue("Code", 0, Code)
                            Call oDS_PS_CO110H.setValue("Name", 0, Code)
                        End If
                        Call Delete_EmptyRow
                    End If
                ElseIf pval.ItemUID = "Btn01" Then
                    If oForm01.Mode = fm_UPDATE_MODE Then
                        oForm01.Items("1").Click ct_Regular
                    Else
'                        Call Sbo_Application.ActivateMenuItem("1280")
'                        oForm01.Items("1").Click ct_Regular
                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "CycleCod" Then
                            If oMat01.Columns("CycleCod").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                    End If
                End If
'et_COMBO_SELECT ////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            
                If pval.ItemUID = "Mat01" Then
                    If pval.Row > 0 Then
        
                        Call oMat01.SelectRow(pval.Row, True, False)
                        
                    End If
                End If
            
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
                Call Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, pval, BubbleEvent)
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
'et_FORM_RESIZE//////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                ElseIf pval.ItemUID = "Btn01" Then
                    sQry = "EXEC PS_CO110_01 '" & Trim(oForm01.Items("Code").Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry
                    MDC_Com.MDC_GF_Message "배부사이클 실행이 완료되었습니다. 데이터를 확인하세요.", "S"
                ElseIf pval.ItemUID = "Btn02" Then
                    SelectedLineNum = 0
                    For i = 0 To oMat01.VisualRowCount - 2
                        If oMat01.IsRowSelected(i + 1) = True Then
                            SelectedLineNum = oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE
                            Exit For
                        End If
                    Next i
                    If SelectedLineNum <> 0 Then
                        Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, SelectedLineNum)
                    Else
                        MDC_Com.MDC_GF_Message "라인이 선택되지 않았습니다. 확인하세요.", "E"
                        BubbleEvent = False
                    End If
                    BubbleEvent = False
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" And pval.Row = "0" And pval.ColUID = "Select" Then
                    Dim SSelect$
                    oForm01.Freeze True
                    oMat01.FlushToDataSource
                    If Trim(oDS_PS_CO110L.GetValue("U_Select", 0)) = "" Or Trim(oDS_PS_CO110L.GetValue("U_Select", 0)) = "N" Then
                        SSelect = "Y"
                    ElseIf Trim(oDS_PS_CO110L.GetValue("U_Select", 0)) = "Y" Then
                        SSelect = "N"
                    End If
                    For i = 0 To oMat01.VisualRowCount - 1
                        oDS_PS_CO110L.setValue "U_Select", i, SSelect
                    Next i
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
'                ElseIf pval.ItemUID = "Mat01" And pval.Row > 0 And pval.ColUID = "LineNum" Then
'                    BubbleEvent = False
                End If

            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "CycleCod" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                    End If
                End If
'et_MATRIX_LOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_MATRIX_LOAD: '//11
                Add_MatrixRow oMat01.RowCount, False
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
                Set oMat01 = Nothing
                Set oDS_PS_CO110H = Nothing
                Set oDS_PS_CO110L = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
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
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                oForm01.Freeze True
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_CO110L.RemoveRecord oDS_PS_CO110L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    
                    If oMat01.Columns("CycleCod").Cells(oMat01.RowCount).Specific.VALUE <> "" Then
                        Call Add_MatrixRow(oMat01.RowCount, False)
                    End If
                End If
                oForm01.Freeze False
            Case "1281": '찾기
                oForm01.Freeze True
                Call FormItemEnabled
'                oForm01.Items("CycleCod").Click ct_Regular
                oForm01.Freeze False
            Case "1282": '추가
                oForm01.Freeze True
                Call FormItemEnabled
                Add_MatrixRow 0, True
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                oForm01.Freeze True
                Call FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("CycleCod").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        Add_MatrixRow oMat01.RowCount, False
'                    End If
'                End If
                oForm01.Freeze False
            Case "1287": '// 복제
                oForm01.Freeze True
                oDS_PS_CO110H.setValue "Code", 0, ""
                oDS_PS_CO110H.setValue "Name", 0, ""
                oDS_PS_CO110H.setValue "U_YM", 0, ""
                oDS_PS_CO110H.setValue "U_BPLId", 0, ""
                oDS_PS_CO110H.setValue "U_Type", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_CO110L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    oForm01.Freeze False
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

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//작업
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
    Set oDS_PS_CO110H = oForm01.DataSources.DBDataSources("@PS_CO110H")
    Set oDS_PS_CO110L = oForm01.DataSources.DBDataSources("@PS_CO110L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
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
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CF_ChooseFromList_Error:
    MDC_Com.MDC_GF_Message "CF_ChooseFromList_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        oForm01.Items("Btn01").Enabled = False
        oForm01.Items("Btn02").Enabled = False
        oMat01.Columns("Select").Editable = False
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("Btn01").Enabled = True
        oForm01.Items("Btn02").Enabled = True
        oMat01.Columns("Select").Editable = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("Btn01").Enabled = True
        oForm01.Items("Btn02").Enabled = True
        oMat01.Columns("Select").Editable = True
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_CO110L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_CO110L.Offset = oRow
    oDS_PS_CO110L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error

    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sRow As Long
    Dim sSeq$
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Select Case oUID
        Case "Mat01"
            If oCol = "CycleCod" Then
                oForm01.Freeze True
                oMat01.FlushToDataSource
                
                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("CycleCod").Cells(oRow).Specific.VALUE) <> "" Then
                    oMat01.FlushToDataSource
                    Call Add_MatrixRow(oMat01.RowCount, False)
                    oMat01.Columns("CycleCod").Cells(oRow).Click ct_Regular
                End If
                
                sQry = "Select Name From [@PS_CO100H] Where Code = '" & Trim(oMat01.Columns(oCol).Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("CycleNam").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                oForm01.Freeze False
            End If
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "구매견적문서가 취소되었거나 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
    Else
        MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error

    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_CO110H.GetValue("U_YM", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_CO110H.GetValue("U_BPLId", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_CO110H.GetValue("U_Type", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "마감년월은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "작업구분은 필수입력사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error

    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 Then
        If oDS_PS_CO110L.GetValue("U_CycleCod", 0) = "" Then
            ErrNum = 2
            GoTo MatrixSpaceLineDel_Error
        End If
    End If
    
    For i = 0 To oMat01.VisualRowCount - 2
'        Select Case True
'            Case oDS_PS_CO110L.GetValue("U_ItemCode", i) = ""
'                ErrNum = 2
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_CO110L.GetValue("U_Qty", i) = "" Or oDS_PS_CO110L.GetValue("U_Qty", i) = 0
'                ErrNum = 3
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_CO110L.GetValue("U_Weight", i) = ""
'                ErrNum = 4
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_CO110L.GetValue("U_Price", i) = 0
'                ErrNum = 5
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_CO110L.GetValue("U_LinTotal", i) = 0
'                ErrNum = 6
'                GoTo MatrixSpaceLineDel_Error
'        End Select
    Next
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "첫라인에 배부사이클 코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "수량은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "중량은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "단가는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "금액은 필수사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error

    Dim i&
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_CO110L.GetValue("U_CycleCod", i)) = "" Then
            oDS_PS_CO110L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
Private Sub Raise_EVENT_MATRIX_LINK_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LINK_PRESSED_Error
    Dim oTempClass As Variant
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "CycleCod" Then
                Set oTempClass = New PS_CO100
                Call oTempClass.LoadForm(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
