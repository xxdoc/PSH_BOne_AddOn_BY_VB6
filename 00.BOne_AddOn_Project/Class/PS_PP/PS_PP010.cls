VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP010.cls
'//  Module         : PP
'//  Description    : 생산의뢰
'//  FormType       : PS_PP010
'//  Create Date    : 2010.10.05
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01  As String
Public oForm01          As SAPbouiCOM.Form
Public oMat01           As SAPbouiCOM.Matrix
Private oDS_PS_PP010H   As SAPbouiCOM.DBDataSource    '등록헤더

Private oLast_Item_UID  As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID   As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row   As Long   '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oLast_RightClick_DocEntry$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP010.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP010_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call Add_MatrixRow(0, True)
    Call LoadCaption
    Call FormItemEnabled
    
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
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim RFC_Sender$
    Dim Seq&
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
    
    Dim sQry          As String
    Dim oRecordSet01  As SAPbobsCOM.Recordset
        
    Dim oCombo        As SAPbouiCOM.ComboBox
    Dim sCount&, sSeq&

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If PS_PP010_AddData() = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
'                        If Add_ProductionRequest(pval) = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                        
                        oMat01.Clear
                        oMat01.FlushToDataSource
                        oMat01.LoadFromDataSource
                        Add_MatrixRow 0, True
                        
'                        Call Delete_EmptyRow
                        oLast_Mode = oForm01.Mode
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If PS_PP010_AddData() = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
'                        If Update_ProductionRequest(pval) = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                        
                        oForm01.Mode = fm_OK_MODE
                        Call LoadCaption
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        oLast_Mode = oForm01.Mode
                    End If
                ElseIf pval.ItemUID = "Btn02" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    oForm01.Mode = fm_OK_MODE
                    Call LoadCaption
                    Call LoadData
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ItemCode" Then
                        If oForm01.Items("ItemCode").Specific.VALUE = "" Then
                            Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "ItemCode" Then
                            If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "RegNum" Then
                            If oMat01.Columns("RegNum").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "ShipCode" Then
                            If oMat01.Columns("ShipCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
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
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "BPLId" Then
                    If oForm01.Items("BPLId").Specific.VALUE = "2" Or oForm01.Items("BPLId").Specific.VALUE = "3" Then
                    
                        sCount = oForm01.Items("ItmBSort").Specific.ValidValues.Count
                        sSeq = sCount
                        For i = 1 To sCount
                            oForm01.Items("ItmBSort").Specific.ValidValues.Remove sSeq - 1, psk_Index
                            sSeq = sSeq - 1
                        Next i
                        
                        Set oCombo = oForm01.Items("ItmBSort").Specific
                        If oForm01.Items("BPLId").Specific.VALUE = "2" Then
                            sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('105', '106') Order by Code"
                        Else
                            sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('108', '109') Order by Code"
                        End If
                        
                        oRecordSet01.DoQuery sQry
                        Do Until oRecordSet01.EOF
                            oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
                            oRecordSet01.MoveNext
                        Loop
                        oCombo.Select 0, psk_Index
                        
                    End If
                End If
                If pval.ItemUID = "Mat01" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                    Else
                        oForm01.Mode = fm_UPDATE_MODE
                        Call LoadCaption
                    End If
                End If
                
                If pval.ItemUID = "WorkGbn" And oForm01.Mode = fm_ADD_MODE Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    
                    Call Add_MatrixRow(0, True)
                    If oForm01.Items("WorkGbn").Specific.Selected.VALUE = "10" Then
                        oMat01.Columns("RegNum").Editable = True
                        oMat01.Columns("ItemCode").Editable = False
                    Else
                        oMat01.Columns("RegNum").Editable = False
                        oMat01.Columns("ItemCode").Editable = True
                    End If
                    oForm01.Freeze False
                End If
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CntcCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "RegNum" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "ItemCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "ShipCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                        Else
                            oForm01.Mode = fm_UPDATE_MODE
                            Call LoadCaption
                        End If
                    End If
                End If
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
                Set oMat01 = Nothing
                Set oDS_PS_PP010H = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    Set oRecordSet01 = Nothing
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
    Dim ReturnValue&
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim ErrNum       As String
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                ReturnValue = Sbo_Application.MessageBox("해당 라인의 생산의뢰접수를 삭제합니다. 삭제 후 복원할 수 없습니다. 삭제하시겠습니까?", 1, "&확인", "&취소")
                Select Case ReturnValue
                    Case 1:
                        '작번등록여부 체크
                        sQry = "        SELECT  COUNT(*) AS [Cnt]"
                        sQry = sQry & " FROM    [@PS_PP020H]"
                        sQry = sQry & " WHERE   U_PP010Doc = '" & oLast_RightClick_DocEntry & "'"
                        sQry = sQry & "         AND [Status] = 'O'"
                        Call oRecordSet01.DoQuery(sQry)
                        
                        If oRecordSet01.Fields("Cnt").VALUE > 0 Then '작번등록이 존재하면
                            ErrNum = "1"
                            BubbleEvent = False
                            GoTo Raise_MenuEvent_Error
                        End If
                    
                        sQry = "Delete [@PS_PP010H] Where DocEntry = '" & oLast_RightClick_DocEntry & "'"
                        oRecordSet01.DoQuery sQry
                        
                        oLast_RightClick_DocEntry = 0
                    Case 2:
                        Sbo_Application.MessageBox ("실행이 취소되었습니다.")
                        BubbleEvent = False
                        Exit Sub
                End Select
                        
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    oForm01.Freeze True
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("DocNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i

                    oMat01.FlushToDataSource
                    oDS_PS_PP010H.RemoveRecord oDS_PS_PP010H.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Mode = fm_OK_MODE
                Call LoadCaption
'                oForm01.Items("OrdType").Click ct_Regular
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                oForm01.Freeze True
                FormItemEnabled
                
                If Trim(oForm01.Items("WorkGbn").Specific.VALUE) = "10" Then
                    oMat01.Columns("RegNum").Editable = True
                    oMat01.Columns("ItemCode").Editable = False
                Else
                    oMat01.Columns("RegNum").Editable = False
                    oMat01.Columns("ItemCode").Editable = True
                End If
                
                oForm01.Mode = fm_ADD_MODE
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Add_MatrixRow 0, True
                Call LoadCaption
                oForm01.Freeze False
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("CGNo").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_PP010H.GetValue("Status", 0) = "O" Then
                            Add_MatrixRow oMat01.RowCount, False
                        End If
                    End If
                End If
  
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = "1" Then
        Call MDC_Com.MDC_GF_Message("작번등록이 존재하는 작번입니다. 삭제할 수 없습니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
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
        If eventInfo.Row > 0 Then
            oLast_RightClick_DocEntry = Trim(oMat01.Columns("DocEntry").Cells(eventInfo.Row).Specific.VALUE)
        End If
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
    Set oDS_PS_PP010H = oForm01.DataSources.DBDataSources("@PS_PP010H")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("RegDateFr", dt_DATE, 8)
    oForm01.Items("RegDateFr").Specific.DataBind.SetBound True, "", "RegDateFr"
'    oForm01.DataSources.UserDataSources.Item("RegDateFr").Value = Format(Now, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("RegDateTo", dt_DATE, 8)
    oForm01.Items("RegDateTo").Specific.DataBind.SetBound True, "", "RegDateTo"
'    oForm01.DataSources.UserDataSources.Item("RegDateTo").Value = Date

    Call oForm01.DataSources.UserDataSources.Add("PuDateFr", dt_DATE, 8)
    oForm01.Items("PuDateFr").Specific.DataBind.SetBound True, "", "PuDateFr"
'    oForm01.DataSources.UserDataSources.Item("PuDateFr").Value = Format(Now, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("PuDateTo", dt_DATE, 8)
    oForm01.Items("PuDateTo").Specific.DataBind.SetBound True, "", "PuDateTo"
'    oForm01.DataSources.UserDataSources.Item("PuDateTo").Value = Date
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo       As SAPbouiCOM.ComboBox
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordset01.Fields(0).Value), Trim(oRecordset01.Fields(1).Value)
        oRecordSet01.MoveNext
    Loop
    
    '//품목대분류
    Set oCombo = oForm01.Items("ItmBSort").Specific
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('105', '106') Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("ItmBSort").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    oCombo.Select 0, psk_Index
    
    '//작업구분
    Set oCombo = oForm01.Items("WorkGbn").Specific
    oCombo.ValidValues.Add "10", "영업"
    oCombo.ValidValues.Add "20", "정비"
    oCombo.ValidValues.Add "30", "멀티"
    oCombo.ValidValues.Add "40", "신동"
    oCombo.ValidValues.Add "50", "R/D"
    oCombo.Select 0, psk_Index
    
    oMat01.Columns("WorkGbn").ValidValues.Add "10", "영업"
    oMat01.Columns("WorkGbn").ValidValues.Add "20", "정비"
    oMat01.Columns("WorkGbn").ValidValues.Add "30", "멀티"
    oMat01.Columns("WorkGbn").ValidValues.Add "40", "신동"
    oMat01.Columns("WorkGbn").ValidValues.Add "50", "R/D"
    
'    '//요청구분
'    oMat01.Columns("ReGbn").ValidValues.Add "10", "계획생산요청"
'    oMat01.Columns("ReGbn").ValidValues.Add "20", "수주생산요청"
    
    '//품목타입
'    Set oCombo = oForm01.Items("ItemType").Specific
'    sQry = "SELECT Code, Name From [@PSH_SHAPE] Order by Code"
'    oRecordset01.DoQuery sQry
'    oCombo.ValidValues.Add "ALL", "ALL"
'    Do Until oRecordset01.EOF
'        oCombo.ValidValues.Add Trim(oRecordset01.Fields(0).Value), Trim(oRecordset01.Fields(1).Value)
'        oRecordset01.MoveNext
'    Loop

    '//사용처
    sQry = "Select PrcCode, PrcName From [OPRC] Where DimCode = '1' Order by PrcCode"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("UseDept").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
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

Sub Initialization()
On Error GoTo Initialization_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//아이디별 사번 세팅
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
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
        oForm01.Items("CardCode").Click ct_Regular
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("ItmBSort").Enabled = True
        oForm01.Items("WorkGbn").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        oForm01.Items("RegDateFr").Enabled = True
        oForm01.Items("RegDateTo").Enabled = True
        oForm01.Items("PuDateFr").Enabled = False
        oForm01.Items("PuDateTo").Enabled = False
        oForm01.Items("Btn02").Enabled = False
        
        oMat01.Columns("RegNum").Editable = True
        oMat01.Columns("ItemCode").Editable = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("CardCode").Click ct_Regular
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("ItmBSort").Enabled = True
        oForm01.Items("WorkGbn").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        oForm01.Items("RegDateFr").Enabled = False
        oForm01.Items("RegDateTo").Enabled = False
        oForm01.Items("PuDateFr").Enabled = True
        oForm01.Items("PuDateTo").Enabled = True
        oForm01.Items("Btn02").Enabled = True
        
        oMat01.Columns("RegNum").Editable = False
        oMat01.Columns("ItemCode").Editable = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP010H.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP010H.Offset = oRow
    oDS_PS_PP010H.setValue "DocNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error

    Dim i&
    Dim ErrNum           As Integer
    Dim sQry             As String
    Dim ItemCode$, Qty&
    Dim Calculate_Weight As Double
    Dim oRecordSet01     As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "ShipCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("ShipCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("ShipName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "ItemCode"
            sQry = "Select ItemName From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            
            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "Mat01"
            If oCol = "RegNum" Then
                Dim RegNum$
                Dim ItmBsort As String '품목대분류 추가(2011.12.01 송명규)
                
                oForm01.Freeze True
                If oForm01.Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("RegNum").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Add_MatrixRow oMat01.RowCount, False
                    End If
                End If
                
                RegNum = Trim(oMat01.Columns("RegNum").Cells(oRow).Specific.VALUE)
                ItmBsort = Trim(oForm01.Items("ItmBsort").Specific.VALUE)
                
                sQry = "        Select  IsNull(U_SjDocLin, '') "
                sQry = sQry & " From    [@PS_SD010H] "
                sQry = sQry & " Where   U_RegNum = '" & RegNum & "'"
                sQry = sQry & "         AND U_ItmBSort = '" & ItmBsort & "'" '품목분류 조건 추가(2011.12.01 송명규 추가)
                
                oRecordSet01.DoQuery sQry
                
                If Trim(oRecordSet01.Fields(0).VALUE) = "" Then
                
                    sQry = "        Select  a.U_ItemCode, "
                    sQry = sQry & "         a.U_ItemName, "
                    sQry = sQry & "         b.U_Material, "
                    sQry = sQry & "         b.SalUnitMsr, "
                    sQry = sQry & "         b.U_Size, "
                    sQry = sQry & "         a.U_ItmBSort, "
                    sQry = sQry & "         e.U_CPNaming, "
                    sQry = sQry & "         a.U_ReWeight "
                    sQry = sQry & " From    [@PS_SD010H] a"
                    sQry = sQry & "         Inner Join "
                    sQry = sQry & "         [OITM] b "
                    sQry = sQry & "             On a.U_ItemCode = b.ItemCode "
                    sQry = sQry & "         Left Join "
                    sQry = sQry & "         [@PSH_ItmMSort] e "
                    sQry = sQry & "             On b.U_ItmBSort = e.U_rCode "
                    sQry = sQry & "             And b.U_ItmMsort = e.U_Code "
                    sQry = sQry & " Where   a.U_RegNum = '" & RegNum & "'"
                    sQry = sQry & "         AND a.U_ItmBSort = '" & ItmBsort & "'" '품목분류 조건 추가(2011.12.01 송명규 추가)
                    oRecordSet01.DoQuery sQry
                Else
                    sQry = "        Select  a.U_ItemCode, "
                    sQry = sQry & "         a.U_ItemName, "
                    sQry = sQry & "         b.U_Material, "
                    sQry = sQry & "         b.SalUnitMsr, "
                    sQry = sQry & "         b.U_Size, "
                    sQry = sQry & "         a.U_ItmBSort, "
                    sQry = sQry & "         e.U_CPNaming, "
                    sQry = sQry & "         a.U_ReWeight, "
                    sQry = sQry & "         Left(a.U_SjDocLin, CharIndex('-', a.U_SjDocLin) - 1) As SjDocNum, "
                    sQry = sQry & "         Right(a.U_SjDocLin, Len(a.U_SjDocLin) - CharIndex('-', a.U_SjDocLin)) As SjLinNum, "
                    sQry = sQry & "         a.U_SjQty, "
                    sQry = sQry & "         a.U_SjWeight, "
                    sQry = sQry & "         c.DocDate, "
                    sQry = sQry & "         d.ShipDate, "
                    sQry = sQry & "         d.LineTotal, "
                    sQry = sQry & "         a.U_CardCode, "
                    sQry = sQry & "         a.U_CardName "
                    sQry = sQry & " From    [@PS_SD010H] a "
                    sQry = sQry & "         Inner Join "
                    sQry = sQry & "         [OITM] b "
                    sQry = sQry & "             On a.U_ItemCode = b.ItemCode "
                    sQry = sQry & "         Inner Join "
                    sQry = sQry & "         [ORDR] c "
                    sQry = sQry & "             On c.DocNum = Left(a.U_SjDocLin, CharIndex('-', a.U_SjDocLin) - 1) "
                    sQry = sQry & "         Inner Join "
                    sQry = sQry & "         [RDR1] d "
                    sQry = sQry & "             On c.DocEntry = d.DocEntry "
                    sQry = sQry & "             And d.LineNum = Right(a.U_SjDocLin, Len(a.U_SjDocLin) - CharIndex('-', a.U_SjDocLin)) "
                    sQry = sQry & "         Left Join "
                    sQry = sQry & "         [@PSH_ItmMSort] e "
                    sQry = sQry & "             On b.U_ItmBSort = e.U_rCode "
                    sQry = sQry & "             And b.U_ItmMsort = e.U_Code "
                    sQry = sQry & " Where   a.U_RegNum = '" & RegNum & "'"
                    sQry = sQry & "         AND a.U_ItmBSort = '" & ItmBsort & "'" '품목분류 조건 추가(2011.12.01 송명규 추가)
                    
                    oRecordSet01.DoQuery sQry
                    
                    oMat01.Columns("SjDocNum").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("SjDocNum").VALUE)
                    oMat01.Columns("SjLinNum").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("SjLinNum").VALUE)
                    oMat01.Columns("SjQty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_SjQty").VALUE)
                    oMat01.Columns("SjWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_SjWeight").VALUE)
                    oMat01.Columns("SjDcDate").Cells(oRow).Specific.VALUE = Format(Trim(oRecordSet01.Fields("DocDate").VALUE), "YYYYMMDD")
                    oMat01.Columns("SjDuDate").Cells(oRow).Specific.VALUE = Format(Trim(oRecordSet01.Fields("ShipDate").VALUE), "YYYYMMDD")
                    oMat01.Columns("SlePrice").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("LineTotal").VALUE)
                    oMat01.Columns("CardCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CardCode").VALUE)
                    oMat01.Columns("CardName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CardName").VALUE)
                End If
                                
                oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemName").VALUE)
                oMat01.Columns("Material").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Material").VALUE)
                oMat01.Columns("Unit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("SalUnitMsr").VALUE)
                oMat01.Columns("Size").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Size").VALUE)
                oMat01.Columns("ItmBSort").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItmBSort").VALUE), psk_ByValue
                oMat01.Columns("CpName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CpNaming").VALUE)
                oMat01.Columns("WorkGbn").Cells(oRow).Specific.Select "10", psk_ByValue
                oMat01.Columns("PuDate").Cells(oRow).Specific.VALUE = Format(Now, "YYYYMMDD")
                oMat01.Columns("ProDate").Cells(oRow).Specific.VALUE = Format(Now, "YYYYMMDD")
                oMat01.Columns("JakName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
                oMat01.Columns("SubNo1").Cells(oRow).Specific.VALUE = "00"
                oMat01.Columns("SubNo2").Cells(oRow).Specific.VALUE = "000"
                oMat01.Columns("ReWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ReWeight").VALUE)
                
                oMat01.Columns("Comments").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "ItemCode" Then
                oForm01.Freeze True
                If oForm01.Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Add_MatrixRow oMat01.RowCount, False
                    End If
                End If
                
                sQry = "        Select  a.ItemCode, "
                sQry = sQry & "         a.FrgnName, "
                sQry = sQry & "         a.U_Material, "
                sQry = sQry & "         a.SalUnitMsr, "
                sQry = sQry & "         a.U_Size, "
                sQry = sQry & "         a.U_ItmBSort, "
                sQry = sQry & "         b.U_CPNaming  "
                sQry = sQry & " From    OITM a "
                sQry = sQry & "         Left Join "
                sQry = sQry & "         [@PSH_ItmMSort] b "
                sQry = sQry & "             On a.U_ItmBSort = b.U_rCode "
                sQry = sQry & "             And a.U_ItmMsort = b.U_Code "
                sQry = sQry & " Where   a.ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                
                oRecordSet01.DoQuery sQry
                
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("FrgnName").VALUE) '품명
                oMat01.Columns("Material").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Material").VALUE) '재질
                oMat01.Columns("Unit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("SalUnitMsr").VALUE) '단위
                oMat01.Columns("Size").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Size").VALUE) '규격
                oMat01.Columns("JakName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("ItemCode").VALUE) '작번
                
                oMat01.Columns("PuDate").Cells(oRow).Specific.VALUE = Format(Date, "YYYYMMDD")
                oMat01.Columns("ItmBSort").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItmBSort").VALUE), psk_ByValue
                                
                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "ShipCode" Then
                sQry = "        Select  CardName "
                sQry = sQry & " From    OCRD "
                sQry = sQry & " Where   CardCode = '" & Trim(oMat01.Columns("ShipCode").Cells(oRow).Specific.VALUE) & "'"
                
                oRecordSet01.DoQuery sQry
                
                oMat01.Columns("ShipName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
        
        Call oMat01.AutoResizeColumns
        
    End Select

'    Set oRecordset = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("CntcCode").Specific.Value) = ""
'            ErrNum = 3
'            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("DeptCode").Specific.Value) = ""
'            ErrNum = 4
'            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
'    ElseIf ErrNum = 2 Then
'        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
'    ElseIf ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "청구인은 필수사항입니다. 확인하세요.", "E"
'    ElseIf ErrNum = 4 Then
'        MDC_Com.MDC_GF_Message "청구부서는 필수사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0

    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If

    If oForm01.Items("WorkGbn").Specific.Selected.VALUE = "10" Then '구분이 영업(10)일 때만 요청번호 체크

        For i = 0 To oMat01.VisualRowCount - 2
            Select Case True
                Case oDS_PS_PP010H.GetValue("U_RegNum", i) = ""
                    ErrNum = 2
                    GoTo MatrixSpaceLineDel_Error
    '            Case oDS_PS_PP010H.GetValue("U_Qty", i) = ""
    '                ErrNum = 3
    '                GoTo MatrixSpaceLineDel_Error
    '            Case oDS_PS_PP010H.GetValue("U_DocDate", i) = ""
    '                ErrNum = 4
    '                GoTo MatrixSpaceLineDel_Error
    '            Case oDS_PS_PP010H.GetValue("U_DueDate", i) = ""
    '                ErrNum = 5
    '                GoTo MatrixSpaceLineDel_Error
            End Select
        Next
    
    End If
    
    oMat01.LoadFromDataSource

    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 생산요청번호가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 수량이 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 4 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 청구일자가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 5 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 납기일자가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error
    Dim i&
    
'    oMat01.FlushToDataSource
'
'    For i = 0 To oMat01.VisualRowCount - 1
'        If Trim(oDS_PS_MM010L.GetValue("U_CGNo", i)) = "" Then
'            oDS_PS_MM010L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
'        End If
'    Next i
'
'    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "추가"
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "갱신"
    End If

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function PS_PP010_AddData()
'******************************************************************************
'Function ID : PS_PP010_AddData()
'해당모듈    : PS_PP010
'기능        : 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
'인수        : 없음
'반환값      : 없음
'특이사항    : 선행프로세스 일자 체크
'******************************************************************************
On Error GoTo PS_PP010_AddData_Error

    Dim loopCount As Integer
    Dim ErrNum    As String
    Dim ErrOrdNum As String '선행프로세스보다 일자가 빨라서 저장되지 않는 작번을 저장
    
    Dim Query01   As String
    Dim Query02   As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry    As String
    Dim BPLID       As String
    Dim RegNum      As String
    Dim ItemCode    As String
    Dim ItemName    As String
    Dim Material    As String
    Dim Unit        As String
    Dim Size        As String
    Dim ItmBsort    As String
    Dim CpName      As String
    Dim SjDocNum    As String
    Dim SjLinNum    As String
    Dim SjQty       As String
    Dim SjWeight    As String
    Dim SjDcDate    As String
    Dim SjDuDate    As String
    Dim SlePrice    As String
    Dim WorkGbn     As String
    Dim CardCode    As String
    Dim CardName    As String
    Dim ShipCode    As String
    Dim ShipName    As String
    Dim PuDate      As String
    Dim ProDate     As String
    Dim Comments    As String
    Dim JakName     As String
    Dim SubNo1      As String
    Dim SubNo2      As String
    Dim UseDept     As String
    Dim ReWeight    As String
    Dim Status      As String
    
    Dim BaseEntry   As String
    Dim BaseLine    As String
    Dim DocType     As String
    Dim CurDocDate  As String
    
'    Dim Entry()    As String
    Dim MinusNum   As Integer '화면모드에 따라 VisualRowCount에서 빼줄 수
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)

    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    
    If oForm01.Mode = fm_ADD_MODE Then
        MinusNum = 2
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        MinusNum = 1
    End If

    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - MinusNum
    
        DocEntry = Trim(oDS_PS_PP010H.GetValue("DocEntry", loopCount)) 'DocEntry는 프로시저에서 처리
        RegNum = Trim(oDS_PS_PP010H.GetValue("U_RegNum", loopCount))
        ItemCode = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", loopCount))
        ItemName = Make_ItemName(Trim(oDS_PS_PP010H.GetValue("U_ItemName", loopCount)))
        Material = Trim(oDS_PS_PP010H.GetValue("U_Material", loopCount))
        Unit = Trim(oDS_PS_PP010H.GetValue("U_Unit", loopCount))
        Size = Trim(oDS_PS_PP010H.GetValue("U_Size", loopCount))
        ItmBsort = Trim(oDS_PS_PP010H.GetValue("U_ItmBSort", loopCount))
        CpName = Trim(oDS_PS_PP010H.GetValue("U_CpName", loopCount))
        SjDocNum = Trim(oDS_PS_PP010H.GetValue("U_SjDocNum", loopCount))
        SjLinNum = Trim(oDS_PS_PP010H.GetValue("U_SjLinNum", loopCount))
        SjQty = IIf(Trim(oDS_PS_PP010H.GetValue("U_SjQty", loopCount)) = "", 0, Trim(oDS_PS_PP010H.GetValue("U_SjQty", loopCount)))
        SjWeight = Trim(oDS_PS_PP010H.GetValue("U_SjWeight", loopCount))
        SjDcDate = Trim(oDS_PS_PP010H.GetValue("U_SjDcDate", loopCount))
        SjDuDate = Trim(oDS_PS_PP010H.GetValue("U_SjDuDate", loopCount))
        SlePrice = Trim(oDS_PS_PP010H.GetValue("U_SlePrice", loopCount))
        CardCode = Trim(oDS_PS_PP010H.GetValue("U_CardCode", loopCount))
        CardName = Trim(oDS_PS_PP010H.GetValue("U_CardName", loopCount))
        ShipCode = Trim(oDS_PS_PP010H.GetValue("U_ShipCode", loopCount))
        ShipName = Trim(oDS_PS_PP010H.GetValue("U_ShipName", loopCount))
        PuDate = Trim(oDS_PS_PP010H.GetValue("U_PuDate", loopCount))
        ProDate = Trim(oDS_PS_PP010H.GetValue("U_ProDate", loopCount))
        Comments = Trim(oDS_PS_PP010H.GetValue("U_Comments", loopCount))
        SubNo1 = IIf(Trim(oDS_PS_PP010H.GetValue("U_SubNo1", loopCount)) = "", "00", Trim(oDS_PS_PP010H.GetValue("U_SubNo1", loopCount)))
        SubNo2 = IIf(Trim(oDS_PS_PP010H.GetValue("U_SubNo2", loopCount)) = "", "000", Trim(oDS_PS_PP010H.GetValue("U_SubNo2", loopCount)))
        UseDept = Trim(oDS_PS_PP010H.GetValue("U_UseDept", loopCount))
        ReWeight = Trim(oDS_PS_PP010H.GetValue("U_ReWeight", loopCount))
        If Left(ItemCode, 1) = "R" Then
            WorkGbn = "50"
            JakName = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", loopCount))
        ElseIf Left(ItemCode, 1) = "E" Then
            WorkGbn = "40"
            JakName = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", loopCount))
        ElseIf Left(ItemCode, 1) = "F" Then
            WorkGbn = "30"
            JakName = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", loopCount))
        ElseIf Left(ItemCode, 1) = "X" Then
            WorkGbn = "20"
            JakName = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", loopCount))
        Else
            WorkGbn = "10"
            JakName = Trim(oDS_PS_PP010H.GetValue("U_JakName", loopCount))
        End If
        Status = "O"
        
        Query01 = "         EXEC PS_PP010_03 '"
        Query01 = Query01 & DocEntry & "','" '문서번호
        Query01 = Query01 & BPLID & "','" '사업장코드
        Query01 = Query01 & RegNum & "','" '생산요청번호
        Query01 = Query01 & ItemCode & "','" '품목코드
        Query01 = Query01 & ItemName & "','" '품목명
        Query01 = Query01 & Material & "','" '재질
        Query01 = Query01 & Unit & "','" '단위
        Query01 = Query01 & Size & "','" '규격
        Query01 = Query01 & ItmBsort & "','" '품목대분류
        Query01 = Query01 & CpName & "','" '생성코드
        Query01 = Query01 & SjDocNum & "','" '수주번호
        Query01 = Query01 & SjLinNum & "','" '수주행
        Query01 = Query01 & SjQty & "','" '수주량
        Query01 = Query01 & SjWeight & "','" '수주중량
        Query01 = Query01 & SjDcDate & "','" '수주일
        Query01 = Query01 & SjDuDate & "','" '납기일
        Query01 = Query01 & SlePrice & "','" '영업금액
        Query01 = Query01 & WorkGbn & "','" '작업구분
        Query01 = Query01 & CardCode & "','" '거래처코드
        Query01 = Query01 & CardName & "','" '거래처명
        Query01 = Query01 & ShipCode & "','" '납품처코드
        Query01 = Query01 & ShipName & "','" '납품처명
        Query01 = Query01 & PuDate & "','" '생산의뢰접수일
        Query01 = Query01 & ProDate & "','" '작업지시일
        Query01 = Query01 & Comments & "','" '비고
        Query01 = Query01 & JakName & "','" '작명
        Query01 = Query01 & SubNo1 & "','" '서브작번1
        Query01 = Query01 & SubNo2 & "','" '서브작번2
        Query01 = Query01 & UseDept & "','" '사용처
        Query01 = Query01 & ReWeight & "','" '생산요청수량
        Query01 = Query01 & Status & "'" '상태
        
        '선행프로세스 대비 일자체크_S
        'Entry = Split(SjDocLin, "-")
        BaseEntry = RegNum
        BaseLine = "0"
        DocType = "PS_PP010"
        CurDocDate = PuDate
     
        Query02 = "         EXEC PS_Z_CHECK_DATE '"
        Query02 = Query02 & BaseEntry & "','"
        Query02 = Query02 & BaseLine & "','"
        Query02 = Query02 & DocType & "','"
        Query02 = Query02 & CurDocDate & "'"
    
        Call RecordSet02.DoQuery(Query02)
        '선행프로세스 대비 일자체크_E
    
        If RecordSet02.Fields("ReturnValue").VALUE = "True" Then
         
            Call RecordSet01.DoQuery(Query01) '등록
         
        Else
         
            ErrOrdNum = ErrOrdNum & " [" & ItemCode & "]"
         
        End If
        
    Next

    If ErrOrdNum <> "" Then '하나라도 선행프로세스 일자가 빠른 작번이 있으면
        ErrNum = "1"
        GoTo PS_PP010_AddData_Error
    End If

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("전체 저장 완료!", "S")
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    
    PS_PP010_AddData = True

    Exit Function

PS_PP010_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PS_PP010_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    
    If ErrNum = "1" Then
    
        Call Sbo_Application.MessageBox("생산의뢰접수일은 생산요청일과 같거나 늦어야합니다. 확인하십시오." & Chr(13) & ErrOrdNum, 1)
        
        '등록되지 않은 작번이 있어도 화면 Clear_S
        Call oMat01.Clear
        Call oMat01.FlushToDataSource
        Call oMat01.LoadFromDataSource
        Call Add_MatrixRow(0, True)
        '등록되지 않은 작번이 있어도 화면 Clear_E
        
    Else
    
        Call MDC_Com.MDC_GF_Message("PS_PP010_AddData_Error:" & Err.Number & " - " & Err.Description, "E")
    
    End If
    
End Function

Public Function Add_ProductionRequest(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
'******************************************************************************
'Function ID : Add_ProductionDemand()
'해당모듈    : PS_PP010
'기능        : 데이터 INSERT
'인수        : 없음
'반환값      : 없음
'특이사항    : 2017.02.17 패치 이후 사용 안함. PS_SD010_AddDate 로 대체 (2017.02.27 송명규)
'******************************************************************************
On Error GoTo Add_ProductionRequest_Error

    Dim i           As Integer
    Dim sQry        As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, BPLID$, RegNum$, ItemCode$, ItemName$, Material$, Unit$, Size$, ItmBsort$, CpName$, SjDocNum$, SjLinNum$
    Dim SjQty&, SjWeight As Currency, SjDcDate$, SjDuDate$, SlePrice As Currency, WorkGbn$, CardCode$, CardName$, ShipCode$, ShipName$
    Dim PuDate$, ProDate$, Comments$, JakName$, SubNo1$, SubNo2$, Temp01$, UseDept$, ReWeight As Currency
    
    oMat01.FlushToDataSource
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
'    WorkGbn = Trim(oForm01.Items("WorkGbn").Specific.Value)
    For i = 0 To oMat01.RowCount - 2
        PuDate = Trim(oDS_PS_PP010H.GetValue("U_PuDate", i))
        RegNum = Trim(oDS_PS_PP010H.GetValue("U_RegNum", i))
        ItemCode = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", i))
        ItemName = Make_ItemName(Trim(oDS_PS_PP010H.GetValue("U_ItemName", i)))
        Material = Trim(oDS_PS_PP010H.GetValue("U_Material", i))
        Unit = Trim(oDS_PS_PP010H.GetValue("U_Unit", i))
        Size = Trim(oDS_PS_PP010H.GetValue("U_Size", i))
        ItmBsort = Trim(oDS_PS_PP010H.GetValue("U_ItmBSort", i))
        
        If ItmBsort = "" Then
            sQry = "Select U_ItmBSort From OITM Where ItemCode = '" & ItemCode & "'"
            RecordSet01.DoQuery sQry
            ItmBsort = Trim(RecordSet01.Fields(0).VALUE)
        End If
        CpName = Trim(oDS_PS_PP010H.GetValue("U_CpName", i))
        SjDocNum = Trim(oDS_PS_PP010H.GetValue("U_SjDocNum", i))
        SjLinNum = Trim(oDS_PS_PP010H.GetValue("U_SjLinNum", i))
        If Trim(oDS_PS_PP010H.GetValue("U_SjQty", i)) = "" Then
            SjQty = 0
        Else
            SjQty = Trim(oDS_PS_PP010H.GetValue("U_SjQty", i))
        End If
        SjWeight = Trim(oDS_PS_PP010H.GetValue("U_SjWeight", i))
        SjDcDate = Trim(oDS_PS_PP010H.GetValue("U_SjDcDate", i))
        SjDuDate = Trim(oDS_PS_PP010H.GetValue("U_SjDuDate", i))
        SlePrice = Trim(oDS_PS_PP010H.GetValue("U_SlePrice", i))
'        WorkGbn = Trim(oDS_PS_PP010H.GetValue("U_WorkGbn", i))
        CardCode = Trim(oDS_PS_PP010H.GetValue("U_CardCode", i))
        CardName = Trim(oDS_PS_PP010H.GetValue("U_CardName", i))
        ShipCode = Trim(oDS_PS_PP010H.GetValue("U_ShipCode", i))
        ShipName = Trim(oDS_PS_PP010H.GetValue("U_ShipName", i))
        PuDate = Trim(oDS_PS_PP010H.GetValue("U_PuDate", i))
        ProDate = Trim(oDS_PS_PP010H.GetValue("U_ProDate", i))
        Comments = Trim(oDS_PS_PP010H.GetValue("U_Comments", i))
        SubNo1 = Trim(oDS_PS_PP010H.GetValue("U_SubNo1", i))
        SubNo2 = Trim(oDS_PS_PP010H.GetValue("U_SubNo2", i))
        UseDept = Trim(oDS_PS_PP010H.GetValue("U_UseDept", i))
        ReWeight = Trim(oDS_PS_PP010H.GetValue("U_ReWeight", i))
        
        If SubNo1 = "" Then SubNo1 = "00"
        If SubNo2 = "" Then SubNo2 = "000"
        
        If Left(ItemCode, 1) = "R" Then
            WorkGbn = "50"
            JakName = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", i))
        ElseIf Left(ItemCode, 1) = "E" Then
            WorkGbn = "40"
            JakName = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", i))
        ElseIf Left(ItemCode, 1) = "F" Then
            WorkGbn = "30"
            JakName = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", i))
        ElseIf Left(ItemCode, 1) = "X" Then
            WorkGbn = "20"
            JakName = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", i))
        Else
            WorkGbn = "10"
            JakName = Trim(oDS_PS_PP010H.GetValue("U_JakName", i))
        End If
        
'        If Left(ItemCode, 1) = "E" Or Left(ItemCode, 1) = "F" Or Left(ItemCode, 1) = "X" Then
'            Temp01 = Left(ItemCode, 1) & CpName
'        Else
'            Temp01 = CpName
'        End If
'        If WorkGbn = "20" Then
'            Temp01 = "E"
'        ElseIf WorkGbn = "30" Then
'            Temp01 = "F"
'        ElseIf WorkGbn = "40" Then
'            Temp01 = "X"
'        End If
'        Temp01 = Temp01 & CpName
        
        '// DocEntry
        sQry = "Select IsNULL(Max(DocEntry), 0) From [@PS_PP010H] "
        RecordSet01.DoQuery sQry
        If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
            DocEntry = 1
        Else
            DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
        End If
        
'        If WorkGbn = "10" Then
'            JakName = Trim(oDS_PS_PP010H.GetValue("U_JakName", i))
'        Else
'            JakName = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", i))
''            sQry = "Select IsNull(Max(Right(U_JakName, 3)), 0) From [@PS_PP010H] "
''            sQry = sQry & "where Left(Convert(char, U_PuDate, 112), 6) = SubString('" & PuDate & "', 2, 6) And Left(U_JakName, 2) = '" & Temp01 & "'"
''            RecordSet01.DoQuery sQry
''            If Trim(RecordSet01.Fields(0).Value) = 0 Then
''                JakName = Temp01 & PuDate & "001"
''            Else
''                JakName = Trim(RecordSet01.Fields(0).Value) + 1
''                JakName = Temp01 & PuDate & JakName
''            End If
'        End If
        
        sQry = "        INSERT INTO [@PS_PP010H]"
        sQry = sQry & " ("
        sQry = sQry & " DocEntry,"
        sQry = sQry & " DocNum,"
        sQry = sQry & " U_BPLId,"
        sQry = sQry & " U_RegNum,"
        sQry = sQry & " U_ItemCode,"
        sQry = sQry & " U_ItemName,"
        sQry = sQry & " U_Material,"
        sQry = sQry & " U_Unit,"
        sQry = sQry & " U_Size,"
        sQry = sQry & " U_ItmBSort,"
        sQry = sQry & " U_CpName,"
        sQry = sQry & " U_SjDocNum,"
        sQry = sQry & " U_SjLinNum,"
        sQry = sQry & " U_SjQty,"
        sQry = sQry & " U_SjWeight,"
        sQry = sQry & " U_SjDcDate,"
        sQry = sQry & " U_SjDuDate,"
        sQry = sQry & " U_SlePrice,"
        sQry = sQry & " U_WorkGbn,"
        sQry = sQry & " U_CardCode,"
        sQry = sQry & " U_CardName,"
        sQry = sQry & " U_ShipCode,"
        sQry = sQry & " U_ShipName,"
        sQry = sQry & " U_PuDate,"
        sQry = sQry & " U_ProDate,"
        sQry = sQry & " U_Comments,"
        sQry = sQry & " U_JakName,"
        sQry = sQry & " U_SubNo1,"
        sQry = sQry & " U_SubNo2,"
        sQry = sQry & " U_UseDept,"
        sQry = sQry & " U_ReWeight,"
        sQry = sQry & " U_Status,"
        sQry = sQry & " CreateDate"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("
        sQry = sQry & DocEntry & ","
        sQry = sQry & DocEntry & ","
        sQry = sQry & "'" & BPLID & "',"
        sQry = sQry & "'" & RegNum & "',"
        sQry = sQry & "'" & ItemCode & "',"
        sQry = sQry & "'" & ItemName & "',"
        sQry = sQry & "'" & Material & "',"
        sQry = sQry & "'" & Unit & "',"
        sQry = sQry & "'" & Size & "',"
        sQry = sQry & "'" & ItmBsort & "',"
        sQry = sQry & "'" & CpName & "',"
        sQry = sQry & "'" & SjDocNum & "',"
        sQry = sQry & "'" & SjLinNum & "',"
        sQry = sQry & "'" & SjQty & "',"
        sQry = sQry & "'" & SjWeight & "',"
        sQry = sQry & "'" & SjDcDate & "',"
        sQry = sQry & "'" & SjDuDate & "',"
        sQry = sQry & "'" & SlePrice & "',"
        sQry = sQry & "'" & WorkGbn & "',"
        sQry = sQry & "'" & CardCode & "',"
        sQry = sQry & "'" & CardName & "',"
        sQry = sQry & "'" & ShipCode & "',"
        sQry = sQry & "'" & ShipName & "',"
        sQry = sQry & "'" & PuDate & "',"
        sQry = sQry & "'" & ProDate & "',"
        sQry = sQry & "'" & Comments & "',"
        sQry = sQry & "'" & JakName & "',"
        sQry = sQry & "'" & SubNo1 & "',"
        sQry = sQry & "'" & SubNo2 & "',"
        sQry = sQry & "'" & UseDept & "',"
        sQry = sQry & "'" & ReWeight & "',"
        sQry = sQry & "'O',"
        sQry = sQry & " GETDATE()"
        sQry = sQry & ")"
        
        RecordSet01.DoQuery sQry
        
        sQry = "Update [@PS_SD010H] Set U_Status = 'C' Where U_RegNum = '" & RegNum & "'"
        RecordSet01.DoQuery sQry
    Next
    
    MDC_Com.MDC_GF_Message "생산의뢰접수 완료!", "S"
    
    Set RecordSet01 = Nothing
    Add_ProductionRequest = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_ProductionRequest_Error:
    Add_ProductionRequest = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Add_ProductionRequest_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function Update_ProductionRequest(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
'******************************************************************************
'Function ID : Update_ProductionRequest()
'해당모듈    : PS_PP010
'기능        : 데이터 UPDATE
'인수        : 없음
'반환값      : 없음
'특이사항    : 2017.02.17 패치 이후 사용 안함. PS_SD010_AddDate 로 대체 (2017.02.17 송명규)
'******************************************************************************
On Error GoTo Update_ProductionRequest_Error
    Dim i As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RegNum$, ItemCode$, ItemName$, Material$, Unit$, Size$, ItmBsort$, CpName$, SjDocNum$, SjLinNum$, DocEntry$
    Dim SjQty&, SjWeight As Currency, SjDcDate$, SjDuDate$, SlePrice As Currency, WorkGbn$, CardCode$, CardName$, ShipCode$, ShipName$
    Dim PuDate$, ProDate$, Comments$, JakName$, SubNo1$, SubNo2$, Temp01$, Status$, UseDept$, ReWeight As Currency
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.RowCount - 1
'        ItemCode = Trim(oDS_PS_PP010H.GetValue("U_ItemCode", i))
'        ItemName = Trim(oDS_PS_PP010H.GetValue("U_ItemName", i))
'        Material = Trim(oDS_PS_PP010H.GetValue("U_Material", i))
'        Unit = Trim(oDS_PS_PP010H.GetValue("U_Unit", i))
'        Size = Trim(oDS_PS_PP010H.GetValue("U_Size", i))
'        ItmBSort = Trim(oDS_PS_PP010H.GetValue("U_ItmBSort", i))
'        CpName = Trim(oDS_PS_PP010H.GetValue("U_CpName", i))
'        SjDocNum = Trim(oDS_PS_PP010H.GetValue("U_SjDocNum", i))
'        SjLinNum = Trim(oDS_PS_PP010H.GetValue("U_SjLinNum", i))
'        If Trim(oDS_PS_PP010H.GetValue("U_SjQty", i)) = "" Then
'            SjQty = 0
'        Else
'            SjQty = Trim(oDS_PS_PP010H.GetValue("U_SjQty", i))
'        End If
'        SjWeight = Trim(oDS_PS_PP010H.GetValue("U_SjWeight", i))
'        SjDcDate = Trim(oDS_PS_PP010H.GetValue("U_SjDcDate", i))
'        SjDuDate = Trim(oDS_PS_PP010H.GetValue("U_SjDuDate", i))
'        SlePrice = Trim(oDS_PS_PP010H.GetValue("U_SlePrice", i))
'        WorkGbn = Trim(oDS_PS_PP010H.GetValue("U_WorkGbn", i))
'        CardCode = Trim(oDS_PS_PP010H.GetValue("U_CardCode", i))
'        CardName = Trim(oDS_PS_PP010H.GetValue("U_CardName", i))
        ShipCode = Trim(oDS_PS_PP010H.GetValue("U_ShipCode", i))
        ShipName = Trim(oDS_PS_PP010H.GetValue("U_ShipName", i))
        PuDate = Trim(oDS_PS_PP010H.GetValue("U_PuDate", i))
        ProDate = Trim(oDS_PS_PP010H.GetValue("U_ProDate", i))
        Comments = Trim(oDS_PS_PP010H.GetValue("U_Comments", i))
        JakName = Trim(oDS_PS_PP010H.GetValue("U_JakName", i))
        Status = Trim(oDS_PS_PP010H.GetValue("U_Status", i))
        DocEntry = Trim(oDS_PS_PP010H.GetValue("DocEntry", i))
        UseDept = Trim(oDS_PS_PP010H.GetValue("U_UseDept", i))
        ReWeight = Trim(oDS_PS_PP010H.GetValue("U_ReWeight", i))
        
        sQry = "UPDATE [@PS_PP010H] "
        sQry = sQry & "SET "
        sQry = sQry & "U_ShipCode = '" & ShipCode & "', "
        sQry = sQry & "U_ShipName = '" & ShipName & "', "
        sQry = sQry & "U_PuDate = '" & PuDate & "', "
        sQry = sQry & "U_ProDate = '" & ProDate & "', "
        sQry = sQry & "U_Comments = '" & Comments & "', "
        sQry = sQry & "U_UseDept = '" & UseDept & "', "
        sQry = sQry & "U_ReWeight = '" & ReWeight & "', "
        sQry = sQry & "U_Status = '" & Status & "' "
        sQry = sQry & "Where DocEntry = '" & DocEntry & "' "
    
        RecordSet01.DoQuery sQry
    Next

    MDC_Com.MDC_GF_Message "생산의뢰 갱신 완료!", "S"
    
    Set RecordSet01 = Nothing
    Update_ProductionRequest = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Update_ProductionRequest_Error:
    Update_ProductionRequest = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Update_ProductionRequest_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RegNum$, ItemCode$, ItemName$, Material$, Unit$, Size$, ItmBsort$, CpName$, SjDocNum$, SjLinNum$
    Dim SjQty&, SjWeight As Currency, SjDcDate$, SjDuDate$, SlePrice As Currency, WorkGbn$, CardCode$, CardName$, ShipCode$, ShipName$
    Dim PuDateFr$, PuDateTo$, Comments$, JakName$, SubNo1$, SubNo2$, Temp01$, Status$, BPLID$, UseDept$, ReWeight As Currency
        
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    WorkGbn = Trim(oForm01.Items("WorkGbn").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    PuDateFr = Trim(oForm01.Items("PuDateFr").Specific.VALUE)
    PuDateTo = Trim(oForm01.Items("PuDateTo").Specific.VALUE)
    JakName = Trim(oForm01.Items("JakName").Specific.VALUE)
    
    If BPLID = "" Then BPLID = "%"
    If CardCode = "" Then CardCode = "%"
    If ItemCode = "" Then ItemCode = "%"
    If PuDateFr = "" Then PuDateFr = "19000101"
    If PuDateTo = "" Then PuDateTo = "20991231"
    If JakName = "" Then JakName = "%"
    
    sQry = "EXEC [PS_PP010_01] '" & BPLID & "','" & ItmBsort & "','" & WorkGbn & "','" & CardCode & "','" & ItemCode & "','" & PuDateFr & "',"
    sQry = sQry & "'" & PuDateTo & "','" & JakName & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_PP010H.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP010H.Size Then
            oDS_PS_PP010H.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_PP010H.Offset = i
        oDS_PS_PP010H.setValue "DocNum", i, i + 1
        oDS_PS_PP010H.setValue "U_RegNum", i, Trim(oRecordSet01.Fields("U_RegNum").VALUE)
        oDS_PS_PP010H.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
        oDS_PS_PP010H.setValue "U_ItemName", i, Trim(oRecordSet01.Fields("U_ItemName").VALUE)
        oDS_PS_PP010H.setValue "U_Material", i, Trim(oRecordSet01.Fields("U_Material").VALUE)
        oDS_PS_PP010H.setValue "U_Unit", i, Trim(oRecordSet01.Fields("U_Unit").VALUE)
        oDS_PS_PP010H.setValue "U_Size", i, Trim(oRecordSet01.Fields("U_Size").VALUE)
        oDS_PS_PP010H.setValue "U_ItmBSort", i, Trim(oRecordSet01.Fields("U_ItmBSort").VALUE)
        oDS_PS_PP010H.setValue "U_CpName", i, Trim(oRecordSet01.Fields("U_CpName").VALUE)
        oDS_PS_PP010H.setValue "U_SjDocNum", i, Trim(oRecordSet01.Fields("U_SjDocNum").VALUE)
        oDS_PS_PP010H.setValue "U_SjLinNum", i, Trim(oRecordSet01.Fields("U_SjLinNum").VALUE)
        oDS_PS_PP010H.setValue "U_SjQty", i, Trim(oRecordSet01.Fields("U_SjQty").VALUE)
        oDS_PS_PP010H.setValue "U_SjWeight", i, Trim(oRecordSet01.Fields("U_SjWeight").VALUE)
        oDS_PS_PP010H.setValue "U_SjDcDate", i, Format(Trim(oRecordSet01.Fields("U_SjDcDate").VALUE), "YYYYMMDD")
        oDS_PS_PP010H.setValue "U_SjDuDate", i, Format(Trim(oRecordSet01.Fields("U_SjDuDate").VALUE), "YYYYMMDD")
        oDS_PS_PP010H.setValue "U_SlePrice", i, Trim(oRecordSet01.Fields("U_SlePrice").VALUE)
        oDS_PS_PP010H.setValue "U_WorkGbn", i, Trim(oRecordSet01.Fields("U_WorkGbn").VALUE)
        oDS_PS_PP010H.setValue "U_CardCode", i, Trim(oRecordSet01.Fields("U_CardCode").VALUE)
        oDS_PS_PP010H.setValue "U_CardName", i, Trim(oRecordSet01.Fields("U_CardName").VALUE)
        oDS_PS_PP010H.setValue "U_ShipCode", i, Trim(oRecordSet01.Fields("U_ShipCode").VALUE)
        oDS_PS_PP010H.setValue "U_ShipName", i, Trim(oRecordSet01.Fields("U_ShipName").VALUE)
        oDS_PS_PP010H.setValue "U_PuDate", i, Format(Trim(oRecordSet01.Fields("U_PuDate").VALUE), "YYYYMMDD")
        oDS_PS_PP010H.setValue "U_ProDate", i, Format(Trim(oRecordSet01.Fields("U_ProDate").VALUE), "YYYYMMDD")
        oDS_PS_PP010H.setValue "U_Comments", i, Trim(oRecordSet01.Fields("U_Comments").VALUE)
        oDS_PS_PP010H.setValue "U_JakName", i, Trim(oRecordSet01.Fields("U_JakName").VALUE)
        oDS_PS_PP010H.setValue "U_SubNo1", i, Trim(oRecordSet01.Fields("U_SubNo1").VALUE)
        oDS_PS_PP010H.setValue "U_SubNo2", i, Trim(oRecordSet01.Fields("U_SubNo2").VALUE)
        oDS_PS_PP010H.setValue "U_UseDept", i, Trim(oRecordSet01.Fields("U_UseDept").VALUE)
        oDS_PS_PP010H.setValue "U_ReWeight", i, Trim(oRecordSet01.Fields("U_ReWeight").VALUE)
        oDS_PS_PP010H.setValue "U_Status", i, Trim(oRecordSet01.Fields("U_Status").VALUE)
        oDS_PS_PP010H.setValue "DocEntry", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)

        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
