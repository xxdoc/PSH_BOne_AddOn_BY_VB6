VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM035"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM035.cls
'//  Module         : QM
'//  Description    : M/G원소재수입검자자료등록
'//  FormType       : PS_QM035
'//  Create Date    : 2013.08.21
'//  Modified Date  :
'//  Creator        : NGY
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM035H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_QM035L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private Last_CardCode$, Last_CardName$, Last_InspPrsn$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM035.srf")
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
    
    oFormUniqueID01 = "PS_QM035_" & GetTotalFormsCount
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
    Call FormClear
'    Call Matrix_AddRow(1, 0, True) 'oMat01
    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), True         '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), True         '// 복제
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
    Set oDS_PS_QM035H = oForm01.DataSources.DBDataSources("@PS_QM035H")
'    Set oDS_PS_QM035L = oForm01.DataSources.DBDataSources("@PS_QM035L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific

'    oDS_PS_QM035H.setValue "U_InDate", 0, Format(Now, "YYYYMMDD")
   
'    oForm01.Items("Code").Visible = True 'False

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
    Dim oRecordSet      As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 검사자
'    Set oCombo = oForm01.Items("InspPrsn").Specific
'    sQry = "SELECT DISTINCT U_MSTCOD, lastName+firstName From [OHEM] " '중복제거를 위한 DISTINCT문 추가(2012.06.27 송명규 수정)
''    sQry = sQry & "Where dept='8'"
'    oRecordSet.DoQuery sQry
'    Do Until oRecordSet.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet.Fields(0).VALUE), Trim(oRecordSet.Fields(1).VALUE)
'        oRecordSet.MoveNext
'    Loop
'
'    Call oForm01.Items("TS_Gbn").Specific.ValidValues.Add("10", "Kgf/m2")
'    Call oForm01.Items("TS_Gbn").Specific.ValidValues.Add("20", "N/mm2")
    
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
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
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                        If oForm01.Mode = fm_ADD_MODE Then
                            If Check_Exist = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
'                    If pval.ItemUID = "OrdNum" Then
'                        If oForm01.Items("OrdNum").Specific.VALUE = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
'                        End If
'                    End If
'                    If pval.ItemUID = "CardCode" Then
'                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
'                        End If
'                    End If
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
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    ElseIf oForm01.Mode = fm_OK_MODE Then
                        FormItemEnabled
'                        FlushToItemValue "LotNo"
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//헤더
                    If pval.ItemUID = "LotNo" Then
                        If oForm01.Mode = fm_ADD_MODE Then
                            If oForm01.Items("LotNo").Specific.VALUE <> "" Then
                                FlushToItemValue pval.ItemUID
    '                            FlushToItemValue "LotNo"
                            Else
                                oForm01.Items("ItemCode").Specific.VALUE = ""
                                oForm01.Items("ItemName").Specific.VALUE = ""
                                oForm01.Items("InDate").Specific.VALUE = ""
                                oForm01.Items("InWgt").Specific.VALUE = 0
                            End If
                        End If
                    End If
                
                    '//라인
'                    If pval.ItemUID = "Mat01" And (pval.ColUID = "PP070No") Then
'                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
'                    End If
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
                Set oMat01 = Nothing
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
            Case "1287": '복제
'                oForm01.Freeze True
'                oDS_PS_QM035H.setValue "U_OrdNum", 0, ""
'                oDS_PS_QM035H.setValue "U_ItemCode", 0, ""
'                oDS_PS_QM035H.setValue "U_ItemName", 0, ""
'                oDS_PS_QM035H.setValue "U_LotNo", 0, ""
'                oDS_PS_QM035H.setValue "U_InspDate", 0, Format(Now, "YYYYMMDD")
'                oDS_PS_QM035H.setValue "U_InspPrsn", 0, ""
'                oDS_PS_QM035H.setValue "U_CardCode", 0, ""
'                oDS_PS_QM035H.setValue "U_CardName", 0, ""
'                FormItemEnabled
'                Call FormClear
'                oForm01.Items("OrdNum").Click ct_Regular
'                oForm01.Freeze False
            Case "1293": '행삭제
'                If oMat01.RowCount <> oMat01.VisualRowCount Then
'                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'                    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
'                    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
'                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'                    For i = 0 To oMat01.VisualRowCount - 1
'                        oMat01.Columns("LineId").Cells(i + 1).Specific.Value = i + 1
'                    Next i
'
'                    oMat01.FlushToDataSource
'                    oDS_PS_QM035L.RemoveRecord oDS_PS_QM035L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
'                    oMat01.Clear
'                    oMat01.LoadFromDataSource
'                End If
            
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Items("Tk").Click ct_Regular
                
            Case "1282": '추가
                Call FormItemEnabled
                Call FormClear
                oForm01.Items("C_Tk").Specific.VALUE = ""
                oForm01.Items("C_Rg").Specific.VALUE = ""
                oForm01.Items("LotNo").Click ct_Regular
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                 FormItemEnabled
'                FlushToItemValue "LotNo"
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("PP070No").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        If oDS_PS_QM035H.GetValue("Status", 0) = "O" Then
'                            Matrix_AddRow 1, oMat01.RowCount, False
'                        End If
'                    End If
'                End If
  
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
                oForm01.Items("C_Tk").Specific.VALUE = ""
                oForm01.Items("C_Rg").Specific.VALUE = ""
                oForm01.Items("LotNo").Click ct_Regular
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormDataEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_FormDataEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FormItemEnabled()
    If oForm01.Mode = fm_FIND_MODE Then
           oForm01.Items("LotNo").Enabled = True
           oForm01.Items("Tk").Enabled = False
           oForm01.Items("Rg").Enabled = False
           oForm01.Items("Br").Enabled = False
           oForm01.Items("Lm").Enabled = False
           oForm01.Items("Hd").Enabled = False
           oForm01.Items("Et").Enabled = False
           oForm01.Items("Ts").Enabled = False
           oForm01.Items("El").Enabled = False
        

    ElseIf oForm01.Mode = fm_ADD_MODE Then
           oForm01.Items("LotNo").Enabled = True
           oForm01.Items("Tk").Enabled = True
           oForm01.Items("Rg").Enabled = True
           oForm01.Items("Br").Enabled = True
           oForm01.Items("Lm").Enabled = True
           oForm01.Items("Hd").Enabled = True
           oForm01.Items("Et").Enabled = True
           oForm01.Items("Ts").Enabled = True
           oForm01.Items("El").Enabled = True

    ElseIf oForm01.Mode = fm_OK_MODE Then
           oForm01.Items("LotNo").Enabled = False
           oForm01.Items("Tk").Enabled = True
           oForm01.Items("Rg").Enabled = True
           oForm01.Items("Br").Enabled = True
           oForm01.Items("Lm").Enabled = True
           oForm01.Items("Hd").Enabled = True
           oForm01.Items("Et").Enabled = True
           oForm01.Items("Ts").Enabled = True
           oForm01.Items("El").Enabled = True
    End If

    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM035'", "")
    If DocNum = 0 Then
'        oForm01.Items("Code").Specific.Value = 1
        oDS_PS_QM035H.setValue "Code", 0, 1
    Else
'        oForm01.Items("Code").Specific.Value = DocNum
        oDS_PS_QM035H.setValue "Code", 0, DocNum
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oMat As Integer, ByVal oRow As Long, Optional Insert_YN As Boolean)
On Error GoTo Matrix_AddRow_Error
        
    Select Case oMat
        Case 1: 'oMat01
'            If Insert_YN = False Then
'                oRow = oMat01.RowCount
'                oDS_PS_QM035L.InsertRecord (oRow)
'            End If
'            '수입내역
'            oDS_PS_QM035L.Offset = oRow
'            oDS_PS_QM035L.setValue "LineId", oRow, oRow + 1
'            oDS_PS_QM035L.setValue "U_PP070No", oRow, ""
'            oDS_PS_QM035L.setValue "U_ItemCode", oRow, ""
'            oDS_PS_QM035L.setValue "U_ItemName", oRow, ""
'            oDS_PS_QM035L.setValue "U_Size", oRow, ""
'            oDS_PS_QM035L.setValue "U_Mark", oRow, ""
'            oDS_PS_QM035L.setValue "U_Qty", oRow, ""
'            oDS_PS_QM035L.setValue "U_Weight", oRow, ""
'            oDS_PS_QM035L.setValue "U_DocDate", oRow, ""
'            oMat01.LoadFromDataSource
        
    End Select
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Matrix_AddRow_Error:
    MDC_Com.MDC_GF_Message "Matrix_AddRow_Error:" & Err.Number & " - " & Err.Description, "E"
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
        Case "LotNo"
            sQry = "select Count(*) From [@PS_QM035H] Where U_LotNo = '" & Trim(oDS_PS_QM035H.GetValue("U_LotNo", 0)) & "'"
            oRecordSet.DoQuery sQry
            
            If oRecordSet.Fields(0).VALUE = 0 Then
            
            
                'Lot내역SET
                sQry = "select MAX(a.ItemCode),MAX(d.ItemName),MAX(a.InDate),SUM(B.Quantity),MAX(d.U_Spec1),MAX(U_Spec2) "
                sQry = sQry & "from OBTN a INNER JOIN ITL1 b ON a.ItemCode = b.ItemCode and a.SysNumber = b.SysNumber "
                sQry = sQry & "INNER JOIN OITL c ON c.LogEntry = b.LogEntry "
                sQry = sQry & "INNER JOIN OITM d ON d.ItemCode = a.ItemCode "
                sQry = sQry & "where c.DocType = '59' "
                sQry = sQry & "and left(a.DistNumber,8) Like '" & Trim(oDS_PS_QM035H.GetValue("U_LotNo", 0)) & "' + '%' "
                oRecordSet.DoQuery sQry
                oDS_PS_QM035H.setValue "U_ItemCode", 0, Trim(oRecordSet.Fields(0).VALUE)
                oDS_PS_QM035H.setValue "U_ItemName", 0, Trim(oRecordSet.Fields(1).VALUE)
                oDS_PS_QM035H.setValue "U_InDate", 0, Format(Trim(oRecordSet.Fields(2).VALUE), "YYYYMMDD")
                oDS_PS_QM035H.setValue "U_InWgt", 0, Trim(oRecordSet.Fields(3).VALUE)
                
                oForm01.Items("C_Tk").Specific.VALUE = Trim(oRecordSet.Fields(4).VALUE)
                oForm01.Items("C_Rg").Specific.VALUE = Trim(oRecordSet.Fields(5).VALUE)
                
                '검사기준 SET
                sQry = "SELECT TOP 1 U_S_Tk_P, U_S_Tk_M, U_S_Rg_P, U_S_Rg_M, U_S_Br, U_S_Lm, U_S_Hd, U_S_Et, U_S_Ts_S, U_S_Ts_E, U_S_El "
                sQry = sQry & "From [@PS_QM034H] "
                sQry = sQry & "ORDER BY CODE DESC "
                oRecordSet.DoQuery sQry
                
                oDS_PS_QM035H.setValue "U_S_Tk_P", 0, Trim(oRecordSet.Fields(0).VALUE)
                oDS_PS_QM035H.setValue "U_S_Tk_M", 0, Trim(oRecordSet.Fields(1).VALUE)
                oDS_PS_QM035H.setValue "U_S_Rg_P", 0, Trim(oRecordSet.Fields(2).VALUE)
                oDS_PS_QM035H.setValue "U_S_Rg_M", 0, Trim(oRecordSet.Fields(3).VALUE)
                oDS_PS_QM035H.setValue "U_S_Br", 0, Trim(oRecordSet.Fields(4).VALUE)
                oDS_PS_QM035H.setValue "U_S_Lm", 0, Trim(oRecordSet.Fields(5).VALUE)
                oDS_PS_QM035H.setValue "U_S_Hd", 0, Trim(oRecordSet.Fields(6).VALUE)
                oDS_PS_QM035H.setValue "U_S_Et", 0, Trim(oRecordSet.Fields(7).VALUE)
                oDS_PS_QM035H.setValue "U_S_Ts_S", 0, Trim(oRecordSet.Fields(8).VALUE)
                oDS_PS_QM035H.setValue "U_S_Ts_E", 0, Trim(oRecordSet.Fields(9).VALUE)
                oDS_PS_QM035H.setValue "U_S_El", 0, Trim(oRecordSet.Fields(10).VALUE)
                
                '//기본SET
                oDS_PS_QM035H.setValue "U_Et", 0, "이상없음"
                
            Else

                Sbo_Application.MessageBox ("이미등록된 Lot입니다.")
                oForm01.ActiveItem = "LotNo"
                
            End If
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
    
    Dim SPEC            As Double
    Dim SPEC_MIN        As Double
    Dim SPEC_MAX        As Double
    
    Dim VALUE           As Double
    Dim VALUE_MIN       As Double
    Dim VALUE_MAX       As Double
    
    Dim Tk              As Double
    Dim Rg              As Double
    
    ErrNum = 0
    
    
    Tk = oForm01.Items("C_Tk").Specific.VALUE
    Rg = oForm01.Items("C_Rg").Specific.VALUE
              
    '--------------------------------------------------------------------------------------------
    '//측청치의 유효값 검사
    
    '//두께
    If Val(oDS_PS_QM035H.GetValue("U_Tk", 0)) <> 0 Then
        VALUE_MIN = Tk - Val(oDS_PS_QM035H.GetValue("U_S_Tk_M", 0))
        VALUE_MAX = Tk + Val(oDS_PS_QM035H.GetValue("U_S_Tk_P", 0))
        SPEC = Val(oForm01.Items("Tk").Specific.VALUE)
        If VALUE_MIN > SPEC Or VALUE_MAX < SPEC Then
            ErrNum = 1
            oForm01.Items("Tk").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//폭
    If Val(oDS_PS_QM035H.GetValue("U_Rg", 0)) <> 0 Then
        VALUE_MIN = Rg - Val(oDS_PS_QM035H.GetValue("U_S_Rg_M", 0))
        VALUE_MAX = Rg + Val(oDS_PS_QM035H.GetValue("U_S_Rg_P", 0))
        SPEC = Val(oForm01.Items("Rg").Specific.VALUE)
        If VALUE_MIN > SPEC Or VALUE_MAX < SPEC Then
            ErrNum = 2
            oForm01.Items("Rg").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//Burr
    If Val(oDS_PS_QM035H.GetValue("U_Br", 0)) <> 0 Then
        VALUE = Val(oDS_PS_QM035H.GetValue("U_S_Br", 0))
        SPEC = Val(oForm01.Items("Br").Specific.VALUE)
        If VALUE < SPEC Then
            ErrNum = 3
            oForm01.Items("Br").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//조도
    If Val(oDS_PS_QM035H.GetValue("U_Lm", 0)) <> 0 Then
        VALUE = Val(oDS_PS_QM035H.GetValue("U_S_Lm", 0))
        SPEC = Val(oForm01.Items("Lm").Specific.VALUE)
        If VALUE < SPEC Then
            ErrNum = 4
            oForm01.Items("Lm").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//경도
    If Val(oDS_PS_QM035H.GetValue("U_Hd", 0)) <> 0 Then
        VALUE = Val(oDS_PS_QM035H.GetValue("U_S_Hd", 0))
        SPEC = Val(oForm01.Items("Hd").Specific.VALUE)
        If VALUE < SPEC Then
            ErrNum = 5
            oForm01.Items("Hd").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//인장강도
    If Val(oDS_PS_QM035H.GetValue("U_Ts", 0)) <> 0 Then
        VALUE_MIN = Val(oDS_PS_QM035H.GetValue("U_S_Ts_S", 0))
        VALUE_MAX = Val(oDS_PS_QM035H.GetValue("U_S_Ts_E", 0))
        SPEC = Val(oForm01.Items("Ts").Specific.VALUE)
        If VALUE_MIN > SPEC Or VALUE_MAX < SPEC Then
            ErrNum = 6
            oForm01.Items("Ts").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//연신율
    If Val(oDS_PS_QM035H.GetValue("U_El", 0)) <> 0 Then
        VALUE = Val(oDS_PS_QM035H.GetValue("U_S_El", 0))
        SPEC = Val(oForm01.Items("El").Specific.VALUE)
        If VALUE > SPEC Then
            ErrNum = 7
            oForm01.Items("El").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "두께를 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "폭을 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "Burr를 확인하여 주십시오.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "조도를 확인하여 주십시오.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "경도를 확인하여 주십시오.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "인장강도 확인하여 주십시오.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "연신율을 확인하여 주십시오.", "E"
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
    If oMat01.VisualRowCount = 1 Then '// MAT01에 값이 있는지 확인 (ErrorNumber : 1)
'        ErrNum = 1
'        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat01.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
'        For i = 0 To oMat01.VisualRowCount - 2      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
'            oDS_PS_QM035L.Offset = i
'            If Trim(oDS_PS_QM035L.GetValue("U_PP070No", i)) = "" Then
'                ErrNum = 2
'                oMat01.Columns("PP070No").Cells(i + 1).Click ct_Regular
'                GoTo MatrixSpaceLineDel_Error
'            End If
'        Next i
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
'        oDS_PS_QM035L.RemoveRecord oDS_PS_QM035L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
    End If
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "벌크포장문서 번호는 필수입니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function Check_Exist() As Boolean
On Error GoTo Check_Exist_Error
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    sQry = "select * from [@PS_QM035H] where U_OrdNum = '" & Trim(oDS_PS_QM035H.GetValue("U_OrdNum", 0)) & "'"
'    oRecordSet.DoQuery sQry
'
'    If oRecordSet.RecordCount <> 0 Then
'        ErrNum = 1
'        GoTo Check_Exist_Error
'    End If
'
    Set oRecordSet = Nothing
    Check_Exist = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Check_Exist_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "이미 동일한 작지번호가 존재합니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Check_Exist = False
End Function




