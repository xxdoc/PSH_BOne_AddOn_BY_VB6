VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP091"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP091.cls
'//  Module         : PP
'//  Description    : 패킹리스트 조회 및 출력
'//  FormType       : PS_PP091
'//  Create Date    : 2010.11.02
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_PP091L As SAPbouiCOM.DBDataSource '등록헤더

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP091.srf")
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
    
    oFormUniqueID01 = "PS_PP091_" & GetTotalFormsCount
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
'    Call FormItemEnabled
    
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
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim TempForm01 As Variant
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    Call oForm01.Close
                ElseIf pval.ItemUID = "Btn02" Then
                    Call LoadData
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CntcCode" Then
                        If oForm01.Items(pval.ItemUID).Specific.VALUE = "" Then
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
                If pval.ItemUID = "Btn03" Then
                    Call Print_Report01
                ElseIf pval.ItemUID = "Btn04" Then
                    Call Print_Report02
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CntcCode" Then
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
                Set oDS_PS_PP091L = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
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
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP091L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 10)
    oForm01.Items("BPLId").Specific.DataBind.SetBound True, "", "BPLId"
'    oForm01.DataSources.UserDataSources.Item("BPLId").Value = oBaseForm01.Items("BPLId").Specific.Value

    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 8)
    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
'    oForm01.DataSources.UserDataSources.Item("DocDateFr").Value = Date  'Format(Now, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 8)
    oForm01.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
'    oForm01.DataSources.UserDataSources.Item("DocDateTo").Value = Date
        
    Set oRecordSet01 = Nothing
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
    sQry = "SELECT BPLId, BPLName From [OBPL] Where BPLId in ('1', '2')"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
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

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CntcCode"
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId$, CntcCode$, DocDateFr$, DocDateTo$, PackNo$, PackNoTo$
        
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    PackNo = Trim(oForm01.Items("PackNo").Specific.VALUE)
    PackNoTo = Trim(oForm01.Items("PackNoTo").Specific.VALUE)
    
    If BPLId = "" Then BPLId = "%"
    If CntcCode = "" Then CntcCode = "%"
    If DocDateFr = "" Then DocDateFr = "19000101"
    If DocDateTo = "" Then DocDateTo = "20991231"
    If PackNo = "" Then PackNo = "00000000001"
    If PackNoTo = "" Then PackNoTo = "9999999999"
    
    sQry = "EXEC [PS_PP091_01] '" & BPLId & "','" & CntcCode & "','" & DocDateFr & "','" & DocDateTo & "','" & PackNo & "','" & PackNoTo & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_PP091L.Clear
    
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP091L.Size Then
            oDS_PS_PP091L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_PP091L.Offset = i
        oDS_PS_PP091L.setValue "U_LineNum", i, i + 1
        oDS_PS_PP091L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("U_PackNo").VALUE)
        oDS_PS_PP091L.setValue "U_ColNum01", i, Trim(oRecordSet01.Fields("U_Qty").VALUE)
        oDS_PS_PP091L.setValue "U_ColQty01", i, Trim(oRecordSet01.Fields("U_Weight").VALUE)
        oDS_PS_PP091L.setValue "U_ColDt01", i, Format(Trim(oRecordSet01.Fields("U_InDate").VALUE), "YYYYMMDD")
        oDS_PS_PP091L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("U_BPLId").VALUE)
        oDS_PS_PP091L.setValue "U_ColPrc01", i, Trim(oRecordSet01.Fields("U_Comments").VALUE)
        oDS_PS_PP091L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("U_CardName").VALUE)
        oDS_PS_PP091L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("U_Size").VALUE)
        
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
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub Print_Report01()
On Error GoTo Print_Report01_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    
    Dim BPLId$, CntcCode$, DocDateFr$, DocDateTo$, PackNo$
    
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    WinTitle = "[PS_PP091]" & "Packing List"
    ReportName = "PS_PP091_01.RPT"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    '// 조회조건문
    sQry = "Delete [Z_PS_PP091]"
    oRecordSet01.DoQuery sQry
    
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_PP091L.GetValue("U_ColReg01", i)) = "Y" Then
            sQry = "Insert [Z_PS_PP091] values ('" & Trim(oDS_PS_PP091L.GetValue("U_ColReg02", i)) & "')"
            oRecordSet01.DoQuery sQry
        End If
    Next
    
    sQry = "EXEC [PS_PP091_02]"
    oRecordSet01.DoQuery sQry
    If oRecordSet01.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Report01_Error
    End If

    '// Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report01_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub Print_Report02()
On Error GoTo Print_Report02_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    
    Dim BPLId$, CntcCode$, DocDateFr$, DocDateTo$, PackNo$
    
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    WinTitle = "[PS_PP091]" & "Packing List"
    ReportName = "PS_PP091_02.RPT"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    '// 조회조건문
    sQry = "Delete [Z_PS_PP091]"
    oRecordSet01.DoQuery sQry
    
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_PP091L.GetValue("U_ColReg01", i)) = "Y" Then
            sQry = "Insert [Z_PS_PP091] values ('" & Trim(oDS_PS_PP091L.GetValue("U_ColReg02", i)) & "')"
            oRecordSet01.DoQuery sQry
        End If
    Next
    
    sQry = "EXEC [PS_PP091_03]"
    oRecordSet01.DoQuery sQry
    If oRecordSet01.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Report02_Error
    End If

    '// Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report02_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report02_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
