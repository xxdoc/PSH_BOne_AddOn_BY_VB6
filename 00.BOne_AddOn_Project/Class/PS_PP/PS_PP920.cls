VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP920"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP920.cls
'//  Module         : PP
'//  Description    : 부문별 P/O 입력
'//  FormType       : PS_PP920
'//  Create Date    : 2011.05.26
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01             As SAPbouiCOM.Matrix
Private oDS_PS_PP920H As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_PP920L As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

Private oYYYYMM$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP920.srf")
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
    
    oFormUniqueID01 = "PS_PP920_" & GetTotalFormsCount
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
    Call Initialization
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), True         '// 삭제
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
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
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
                        
                        Dim Code$, BPLId$, YM$, ItmBsort$
                        BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
                        YM = Trim(Right(oForm01.Items("YM").Specific.VALUE, 4))
                        ItmBsort = Trim(oForm01.Items("ItmBsort").Specific.VALUE)
                        
                        Code = BPLId & YM & ItmBsort
                        Call oDS_PS_PP920H.setValue("Code", 0, Code)
                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "ItmBsort" Then
                        If oForm01.Items("ItmBsort").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
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
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "ItmBsort" Then
                        oForm01.Freeze True
                        
                        sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oForm01.Items("ItmBname").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        Call LoadData
                        oForm01.Freeze False
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "SWeight" Or pval.ColUID = "MWeight" Or pval.ColUID = "Price" Then
                            Dim SWeight As Currency
                            Dim MWeight As Currency
                            Dim Price As Currency
                            Dim SAmt As Currency
                            Dim MAmt As Currency
                            
                            SWeight = oMat01.Columns("SWeight").Cells(pval.Row).Specific.VALUE
                            MWeight = oMat01.Columns("MWeight").Cells(pval.Row).Specific.VALUE
                            Price = oMat01.Columns("Price").Cells(pval.Row).Specific.VALUE
                            oForm01.Freeze True
                            oMat01.Columns("SAmt").Cells(pval.Row).Specific.VALUE = SWeight * Price
                            oMat01.Columns("MAmt").Cells(pval.Row).Specific.VALUE = MWeight * Price
                            
                            oMat01.Columns(pval.ColUID).Cells(pval.Row).Click
                           
                            oForm01.Freeze False
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
                Set oDS_PS_PP920H = Nothing
                Set oDS_PS_PP920L = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
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
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_PP920L.RemoveRecord oDS_PS_PP920L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
            Case "1281": '찾기
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
'                oForm01.Items("CycleCod").Click ct_Regular
                oForm01.Freeze False
            Case "1282": '추가
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
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
                oDS_PS_PP920H.setValue "Code", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_PP920L.setValue "Code", i, ""
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
    Set oDS_PS_PP920H = oForm01.DataSources.DBDataSources("@PS_PP920H")
    Set oDS_PS_PP920L = oForm01.DataSources.DBDataSources("@PS_PP920L")
    
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
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
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

Sub Initialization()
On Error GoTo Initialization_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    oForm01.Items("YM").Specific.VALUE = Left(Format(Now, "YYYYMMDD"), 6)
    
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
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo Add_MatrixRow_Error
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_PP920L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_PP920L.Offset = oRow
'    oDS_PS_PP920L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
'    Exit Sub
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Add_MatrixRow_Error:
'    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
'End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_PP920H.GetValue("U_ItmBsort", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_PP920H.GetValue("U_BPLId", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_PP920H.GetValue("U_YM", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Len(Trim(oForm01.Items("YM").Specific.VALUE)) <> 6
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "대분류는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "년월은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "년월은 6자리 YYYYMM 형식으로 입력해야합니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim i As Integer
    Dim YM As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select Max(U_YM) From [@PS_PP920H] Where U_YM < '" & Trim(oForm01.Items("YM").Specific.VALUE) & "'"
    sQry = sQry & " And U_ItmBsort = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "'"
    oRecordSet01.DoQuery sQry
    
    YM = Trim(oRecordSet01.Fields(0).VALUE)
    
    oForm01.Freeze True
    If Trim(oForm01.Items("ItmBsort").Specific.VALUE) = "104" Then
        sQry = "Select Top 1 ItemCode = '', ItemName = '', a.U_ItmMsort, ItmMname = (select U_CodeName from [@PSH_ITMMSORT] Where U_Code = a.U_ItmMsort), "
        sQry = sQry & " Price = 0,"
        sQry = sQry & " MkCapa = 0 "
        sQry = sQry & " From [OITM] a "
        sQry = sQry & "Where a.U_ItmBsort = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "' Order By a.U_ItmMsort, a.ItemCode"
    Else
    
        sQry = "Select a.ItemCode, a.ItemName, a.U_ItmMsort, ItmMname = (select U_CodeName from [@PSH_ITMMSORT] Where U_Code = a.U_ItmMsort), "
        sQry = sQry & " Price = Isnull((Select U_Price From [@PS_PP920H] t Inner Join [@PS_PP920L] t1 On t.Code = t1.Code And t1.U_ItemCode = a.ItemCode And t.U_YM = '" & YM & "'),0),"
        sQry = sQry & " MkCapa = Isnull((Select U_MkCapa From [@PS_PP920H] t Inner Join [@PS_PP920L] t1 On t.Code = t1.Code And t1.U_ItemCode = a.ItemCode And t.U_YM = '" & YM & "'),0)"
        sQry = sQry & " From [OITM] a "
        sQry = sQry & "Where a.U_ItmBsort = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "' Order By a.U_ItmMsort, a.ItemCode"
    End If
    oRecordSet01.DoQuery sQry
    
    
    oDS_PS_PP920L.Clear
    oMat01.Clear
    oMat01.FlushToDataSource
    
    i = 0
    Do Until oRecordSet01.EOF
        oDS_PS_PP920L.InsertRecord i
        oDS_PS_PP920L.setValue "U_LineNum", i, i + 1
        oDS_PS_PP920L.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields(0).VALUE)
        oDS_PS_PP920L.setValue "U_ItemName", i, Trim(oRecordSet01.Fields(1).VALUE)
        oDS_PS_PP920L.setValue "U_ItmMsort", i, Trim(oRecordSet01.Fields(2).VALUE)
        oDS_PS_PP920L.setValue "U_ItmMName", i, Trim(oRecordSet01.Fields(3).VALUE)
        oDS_PS_PP920L.setValue "U_Price", i, Trim(oRecordSet01.Fields(4).VALUE)
        oDS_PS_PP920L.setValue "U_MkCapa", i, Trim(oRecordSet01.Fields(5).VALUE)
        i = i + 1
        oRecordSet01.MoveNext
    Loop
    
    oMat01.LoadFromDataSource
    oForm01.Freeze False

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


