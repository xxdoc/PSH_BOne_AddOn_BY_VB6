VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR406"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_HR406.cls
'//  Module         : HR
'//  Description    : 인사관리>전문직평가
'//  FormType       : PS_HR406
'//  Create Date    : 2012.09.14
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_HR406H As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_HR406L As SAPbouiCOM.DBDataSource    '등록라인

Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oSeq&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR406.srf")
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
    
    oFormUniqueID01 = "PS_HR406_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    
    
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormClear '//UDO방식일때
    Call FormItemEnabled
    Call SetDocument(oFromDocEntry01)
'    Call Add_MatrixRow(0, True) '//UDO방식일때
    
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), True        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제

    oForm01.EnableMenu ("1281"), False        '// 찾기
    
    oForm01.EnableMenu ("1288"), False        '// 레코드이동버튼
    oForm01.EnableMenu ("1289"), False        '// 레코드이동버튼
    oForm01.EnableMenu ("1290"), False        '// 레코드이동버튼
    oForm01.EnableMenu ("1291"), False        '// 레코드이동버튼
     
    
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
    Dim TempForm01 As Variant
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim RFC_Sender$
    Dim Calculate_Weight As Double
    Dim Seq&
    Dim Year$
    
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
'                        Call Delete_EmptyRow
                        oLast_Mode = oForm01.Mode
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        oLast_Mode = oForm01.Mode
                        
                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "MSTCOD" Then
                        If oForm01.Items("MSTCOD").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "EmpNo1" Then
                        If oForm01.Items("EmpNo1").Specific.VALUE = "" Then
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
                        oForm01.Items("BPLId").Specific.Select "4", psk_ByValue
                    ElseIf oLast_Mode = fm_FIND_MODE Then
                         FormItemEnabled
                         oLast_Mode = 100
'
                        
                    End If
                ElseIf pval.ItemUID = "Btn01" Then
                    '//패스워드 확인
                    If PasswordChk(pval) = False Then
                        Sbo_Application.SetStatusBarMessage "패스워드가 틀렸습니다. 확인바랍니다.", bmt_Short, True
                        oForm01.Items("PassWd").Click ct_Regular
                    Else
                        Call LoadData
                    End If
'                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value = "" Then
'                        Set TempForm01 = New PS_SD081
'                        Call TempForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
'                        BubbleEvent = False
'                    End If
                ElseIf pval.ItemUID = "Btn02" Then
                    If PasswordChk(pval) = False Then
                        Sbo_Application.SetStatusBarMessage "패스워드가 틀렸습니다. 확인바랍니다.", bmt_Short, True
                        oForm01.Items("PassWd").Click ct_Regular
                        
                    
                    Else
                        Set TempForm01 = New PS_HR407
                        Year = Trim(oForm01.Items("Year").Specific.VALUE)
                        
                        Call TempForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, Trim(oForm01.Items("BPLId").Specific.Selected.VALUE), Year, Trim(oForm01.Items("MSTCOD").Specific.VALUE), Trim(oForm01.Items("FULLNAME").Specific.VALUE), Trim(oForm01.Items("PassWd").Specific.VALUE), Trim(oForm01.Items("EmpNo1").Specific.VALUE), Trim(oForm01.Items("EmpName1").Specific.VALUE))
                        Set TempForm01 = Nothing
                        
                    End If
                    BubbleEvent = False
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "MSTCOD" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ItemUID = "EmpNo1" Then
                        FlushToItemValue pval.ItemUID
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
                Set oDS_PS_HR406H = Nothing
                Set oDS_PS_HR406L = Nothing
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
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
'[1284:취소] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1284": '취소
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
            Case "1286": '닫기
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_HR406L.RemoveRecord oDS_PS_HR406L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                FormItemEnabled
                FormClear
                oForm01.Items("BPLId").Specific.Select "4", psk_ByValue
'                Add_MatrixRow 0, True
'                oForm01.Items("BPLId").Click ct_Collapsed
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("CGNo").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        If oDS_PS_HR406H.GetValue("Status", 0) = "O" Then
'                            Add_MatrixRow oMat01.RowCount, False
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
    Set oDS_PS_HR406H = oForm01.DataSources.DBDataSources("@PS_HR406H")
    Set oDS_PS_HR406L = oForm01.DataSources.DBDataSources("@PS_HR406L")
    
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
    
    '//사용여부
    oMat01.Columns("Grade").ValidValues.Add "", "선택"
    oMat01.Columns("Grade").ValidValues.Add "S", "S"
    oMat01.Columns("Grade").ValidValues.Add "A", "A"
    oMat01.Columns("Grade").ValidValues.Add "B", "B"
    oMat01.Columns("Grade").ValidValues.Add "C", "C"
    oMat01.Columns("Grade").ValidValues.Add "D", "D"
    
    
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
    
    oForm01.Items("DocDate").Specific.VALUE = Format(Date, "yyyymmdd") 'Left(Date, 4)
    oForm01.Items("Year").Specific.VALUE = Format(Date, "yyyy")
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
    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("Btn01").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("MSTCOD").Enabled = True
        oForm01.Items("EmpNo1").Enabled = True
    ElseIf oForm01.Mode = fm_FIND_MODE Then
        oForm01.Items("DocNum").Enabled = True
        oForm01.Items("Btn01").Enabled = False
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("MSTCOD").Enabled = True
        oForm01.Items("EmpNo1").Enabled = True
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("Btn01").Enabled = False
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("MSTCOD").Enabled = False
        oForm01.Items("EmpNo1").Enabled = False
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_HR406'", "")
    If DocNum = 0 Then
        oForm01.Items("DocNum").Specific.VALUE = 1
    Else
        oForm01.Items("DocNum").Specific.VALUE = DocNum
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_HR406L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_HR406L.Offset = oRow
    oDS_PS_HR406L.setValue "U_LineNum", oRow, oRow + 1
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
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "MSTCOD"
'            sQry = " select U_FullName,"
'            sQry = sQry + " Rspnm = (Select b.U_CodeNm"
'            sQry = sQry + " From [@PS_HR200H] a Inner join [@PS_HR200L] b On a.Code = b.Code"
'            sQry = sQry + " Where a.Code = '2'"
'            sQry = sQry + " And b.U_Code = t.U_RspCode),"
'            sQry = sQry + " JigNm = (Select b.U_CodeNm"
'            sQry = sQry + " From [@PS_HR200H] a Inner join [@PS_HR200L] b On a.Code = b.Code"
'            sQry = sQry + " Where a.Code = 'P129'"
'            sQry = sQry + " And b.U_Code = t.U_JIGCOD),"
'            sQry = sQry + " CallName = (Select b.U_CodeNm"
'            sQry = sQry + " From [@PS_HR200H] a Inner join [@PS_HR200L] b On a.Code = b.Code"
'            sQry = sQry + " Where a.Code = 'P158'"
'            sQry = sQry + " And b.U_Code = t.U_CallName)"
'            sQry = sQry + " from [@PH_PY001A] t Where t.Code = '" & Trim(oDS_PS_HR406H.GetValue("U_MSTCOD", 0)) & "'"
            
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oDS_PS_HR406H.GetValue("U_MSTCOD", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_HR406H.setValue "U_FULLNAME", 0, Trim(oRecordSet01.Fields(0).VALUE)
'            oDS_PS_HR406H.setValue "U_RspNm", 0, Trim(oRecordSet01.Fields(1).VALUE)
'            oDS_PS_HR406H.setValue "U_JigNm", 0, Trim(oRecordSet01.Fields(2).VALUE)
            
        Case "EmpNo1"
'            sQry = " select U_FullName,"
'            sQry = sQry + " Rspnm = (Select b.U_CodeNm"
'            sQry = sQry + " From [@PS_HR200H] a Inner join [@PS_HR200L] b On a.Code = b.Code"
'            sQry = sQry + " Where a.Code = '2'"
'            sQry = sQry + " And b.U_Code = t.U_RspCode),"
'            sQry = sQry + " JigNm = (Select b.U_CodeNm"
'            sQry = sQry + " From [@PS_HR200H] a Inner join [@PS_HR200L] b On a.Code = b.Code"
'            sQry = sQry + " Where a.Code = 'P129'"
'            sQry = sQry + " And b.U_Code = t.U_JIGCOD),"
'            sQry = sQry + " CallName = (Select b.U_CodeNm"
'            sQry = sQry + " From [@PS_HR200H] a Inner join [@PS_HR200L] b On a.Code = b.Code"
'            sQry = sQry + " Where a.Code = 'P158'"
'            sQry = sQry + " And b.U_Code = t.U_CallName)"
'            sQry = sQry + " from [@PH_PY001A] t Where t.Code = '" & Trim(oDS_PS_HR406H.GetValue("U_MSTCOD", 0)) & "'"

            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oDS_PS_HR406H.GetValue("U_EmpNo1", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_HR406H.setValue "U_EmpName1", 0, Trim(oRecordSet01.Fields(0).VALUE)
'            oDS_PS_HR406H.setValue "U_RspNm1", 0, Trim(oRecordSet01.Fields(1).VALUE)
'            oDS_PS_HR406H.setValue "U_JigNm1", 0, Trim(oRecordSet01.Fields(2).VALUE)
'            oDS_PS_HR406H.setValue "U_CallName", 0, Trim(oRecordSet01.Fields(3).VALUE)
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_HR406H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_HR406H.GetValue("U_MSTCOD", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_HR406H.GetValue("U_EmpNo1", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "평가자는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "피평가자는 필수사항입니다. 확인하세요.", "E"
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
'
'    For i = 0 To oMat01.VisualRowCount - 1
'        If oDS_PS_HR406L.GetValue("U_RequestP", i) <= 0 Then
'            ErrNum = 3
'            GoTo MatrixSpaceLineDel_Error
'        End If
'    Next
    
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
        MDC_Com.MDC_GF_Message "초과필요금액은 0 보다 커야 합니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error
    Dim i&
    
    oMat01.FlushToDataSource
    
'    For i = 0 To oMat01.VisualRowCount - 1
'        If Trim(oDS_PS_HR406L.GetValue("U_Grade", i)) = "" Then
'            oDS_PS_HR406L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
'        End If
'    Next i
'
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub



Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim EmpNo1 As String
    Dim JigDiv As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId$, DocDate$
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
         
    EmpNo1 = Trim(oForm01.Items("EmpNo1").Specific.VALUE)
         
    sQry = " Select t.U_position from [@PH_PY001A] t Where t.Code = '" & EmpNo1 & "'"
    oRecordSet01.DoQuery sQry
    
    If Trim(oRecordSet01.Fields(0).VALUE) = "18" Then
        '//반장
        JigDiv = "1"
    Else
        '//사원
        JigDiv = "2"
    End If
    
    sQry = "EXEC [PS_HR406_01] '" & BPLId & "','" & DocDate & "','" & JigDiv & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_HR406L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
                
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        oForm01.Mode = fm_ADD_MODE
        
        Call Add_MatrixRow(0, True)
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_HR406L.Size Then
            oDS_PS_HR406L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_HR406L.Offset = i
        
        oDS_PS_HR406L.setValue "U_LineNum", i, i + 1
        oDS_PS_HR406L.setValue "U_RateCode", i, Trim(oRecordSet01.Fields(0).VALUE) '//RateCode
        oDS_PS_HR406L.setValue "U_RateMNm", i, Trim(oRecordSet01.Fields(1).VALUE) '//RateMNm
        oDS_PS_HR406L.setValue "U_Contents", i, Trim(oRecordSet01.Fields(2).VALUE) '//Contents
'        oDS_PS_HR406L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields(3).VALUE) '//Grade
'        oDS_PS_HR406L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields(4).VALUE) '//Remark
        
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
    
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'// 패스워드 Check
Public Function PasswordChk(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo PasswordChk_Error
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim MSTCOD$, PassWd$
        
    MSTCOD = Trim(oForm01.Items("MSTCOD").Specific.VALUE)
    PassWd = Trim(oForm01.Items("PassWd").Specific.VALUE)
    
    If Trim(MSTCOD) = "" Then
        MDC_Com.MDC_GF_Message "사번이 없습니다. 입력바랍니다!", "E"
        PasswordChk = False
        Exit Function
    End If
        
    sQry = "Select Count(*) From Z_PS_HRPASS Where MSTCOD = '" & Trim(oForm01.Items("MSTCOD").Specific.VALUE) & "'"
    sQry = sQry & " And  BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
    sQry = sQry & " And  PassWd = '" & oForm01.Items("PassWd").Specific.VALUE & "' "
    RecordSet01.DoQuery sQry
                
    If Trim(RecordSet01.Fields(0).VALUE) <= 0 Then
        PasswordChk = False
    Else
        PasswordChk = True
    End If
        
    Set RecordSet01 = Nothing
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PasswordChk_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PasswordChk_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Sub SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call FormItemEnabled
        
    Else
        oForm01.Mode = fm_FIND_MODE
        Call FormItemEnabled
        oForm01.Items("DocNum").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
