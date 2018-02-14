VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM160.cls
'//  Module         : SD
'//  Description    : 외주가공비집계 및 청구자동등록
'//  FormType       : PS_MM160
'//  Create Date    : 2010.10.25
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_MM160L As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oForm01_Mode&
Private oSeq&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oTradeType02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM160.srf")
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
    
    oFormUniqueID01 = "PS_MM160_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_OK_MODE
    
    oForm01.Freeze True
        
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    
    Call LoadCaption
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False         '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), False         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
        
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
        
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
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
                If pval.ItemUID = "Btn01" Then
                    If oForm01_Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel() = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If Add_MM005(pval) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        oForm01_Mode = fm_OK_MODE
                        oMat01.Clear
                        oDS_PS_MM160L.Clear
                        Call LoadCaption
                    ElseIf oForm01_Mode = fm_OK_MODE Then
                        oForm01.Close
                    End If
                ElseIf pval.ItemUID = "Btn02" Then
                    If HeaderSpaceLineDel() = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call LoadData
                    Call LoadCaption
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "BPLId" Then
                    oMat01.Clear
                    oDS_PS_MM160L.Clear
                    oForm01_Mode = fm_OK_MODE
                    Call LoadCaption
                End If
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CardCode" Or pval.ItemUID = "DateFr" Or pval.ItemUID = "DateTo" Then
                        oForm01.Freeze True
                        oMat01.Clear
                        oDS_PS_MM160L.Clear
                        oForm01_Mode = fm_OK_MODE
                        Call LoadCaption
                        FlushToItemValue pval.ItemUID
                        oForm01.Freeze False
                    ElseIf pval.ItemUID = "CntcCode" Then
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
                Set oDS_PS_MM160L = Nothing
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

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("CardCode").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("DocDate").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("CntcCode").Specific.VALUE) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "외주거래처는 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "요청일은 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "작성자는 필수입력 사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
    Set oDS_PS_MM160L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
        
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("DocDate", dt_DATE, 8)
    oForm01.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    oForm01.DataSources.UserDataSources.Item("DocDate").VALUE = Date
    
    Call oForm01.DataSources.UserDataSources.Add("DateFr", dt_DATE, 8)
    oForm01.Items("DateFr").Specific.DataBind.SetBound True, "", "DateFr"
    oForm01.DataSources.UserDataSources.Item("DateFr").VALUE = Format(Now, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("DateTo", dt_DATE, 8)
    oForm01.Items("DateTo").Specific.DataBind.SetBound True, "", "DateTo"
    oForm01.DataSources.UserDataSources.Item("DateTo").VALUE = Date
    
    Call oForm01.DataSources.UserDataSources.Add("DocTotal", dt_SUM)
    oForm01.Items("DocTotal").Specific.DataBind.SetBound True, "", "DocTotal"
'    oForm01.DataSources.UserDataSources.Item("DocDateFr").Value = 0

    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_SUM)
    oForm01.Items("SumQty").Specific.DataBind.SetBound True, "", "SumQty"
    
    Call oForm01.DataSources.UserDataSources.Add("SumWeight", dt_QUANTITY)
    oForm01.Items("SumWeight").Specific.DataBind.SetBound True, "", "SumWeight"
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
    
    '//대분류
    Set oCombo = oForm01.Items("ItmBSort").Specific
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    oCombo.Select 0, psk_Index
    
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
    Dim ItemCode$, Qty&
    Dim Calculate_Weight As Currency
    Dim Calculate_Qty As Currency
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CntcCode"
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    If oForm01_Mode = fm_ADD_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "추가"
    ElseIf oForm01_Mode = fm_OK_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "확인"
    End If

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId$, CardCode$, CntcCode$, DocDate$, DateFr$, DateTo$
    Dim DocTotal As Currency, SumQty As Long, SumWeight As Currency, ItmBsort$
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    DateFr = Trim(oForm01.Items("DateFr").Specific.VALUE)
    DateTo = Trim(oForm01.Items("DateTo").Specific.VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    
    If CardCode = "" Then CardCode = "%"
    If CntcCode = "" Then CntcCode = "%"
    If DateFr = "" Then DateFr = "18990101"
    If DateTo = "" Then DateTo = "20991231"
        
    sQry = "EXEC [PS_MM160_01] '" & BPLId & "', '" & CardCode & "', '" & CntcCode & "', '" & DocDate & "', '" & DateFr & "', '" & DateTo & "', '" & ItmBsort & "', '1'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_MM160L.Clear
        
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        oForm01_Mode = fm_OK_MODE
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_MM160L.Size Then
            oDS_PS_MM160L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_MM160L.Offset = i
        oDS_PS_MM160L.setValue "U_LineNum", i, i + 1
        oDS_PS_MM160L.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
        oDS_PS_MM160L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("U_ItemName").VALUE)
        oDS_PS_MM160L.setValue "U_ColNum01", i, Trim(oRecordSet01.Fields("U_WorkQty").VALUE)
        oDS_PS_MM160L.setValue "U_ColQty01", i, Trim(oRecordSet01.Fields("U_WorkWt").VALUE)
        oDS_PS_MM160L.setValue "U_ColSum02", i, Trim(oRecordSet01.Fields("Price").VALUE)
        oDS_PS_MM160L.setValue "U_ColSum01", i, Trim(oRecordSet01.Fields("U_Cprice").VALUE)
        oDS_PS_MM160L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("U_CardCode").VALUE)
        oDS_PS_MM160L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("U_CardName").VALUE)
        oDS_PS_MM160L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("CpCode").VALUE)
        oDS_PS_MM160L.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields("CpName").VALUE)

        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    For i = 0 To oMat01.VisualRowCount - 1
        DocTotal = DocTotal + oMat01.Columns("WorkP").Cells(i + 1).Specific.VALUE
        If oMat01.Columns("WorkQty").Cells(i + 1).Specific.VALUE = "" Then
            SumQty = SumQty
        Else
            SumQty = SumQty + oMat01.Columns("WorkQty").Cells(i + 1).Specific.VALUE
        End If
        SumWeight = SumWeight + oMat01.Columns("WorkWt").Cells(i + 1).Specific.VALUE
    Next i
    oForm01.Items("DocTotal").Specific.VALUE = DocTotal
    oForm01.Items("SumQty").Specific.VALUE = SumQty
    oForm01.Items("SumWeight").Specific.VALUE = SumWeight
    
    ProgBar01.Stop
    oForm01.Freeze False
    oForm01_Mode = fm_ADD_MODE
    
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

Public Function Add_MM005(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_MM005_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    Dim BPLId$, DocEntry$, CntcCode$, CntcName$, CardCode$, CardName$, DocDate$, ItemCode$, ItemName$, WorkQty As Long, WorkWt As Currency
    Dim Price As Currency, CPrice As Currency, DateFr$, DateTo$, ItmBsort$
    Dim CpCode$, CpName$
    
    oMat01.FlushToDataSource
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE)
    DateFr = Trim(oForm01.Items("DateFr").Specific.VALUE)
    DateTo = Trim(oForm01.Items("DateTo").Specific.VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    
    Sbo_Company.StartTransaction
    For i = 0 To oMat01.RowCount - 1
'        sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_MM005H] where Left(Convert(char, U_DocDate, 112), 6) = Left('" & DocDate & "', 6)"
        sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_MM005H] where Left(DocEntry, 6) = Left('" & DocDate & "', 6)"
        oRecordSet01.DoQuery sQry
        If Trim(oRecordSet01.Fields(0).VALUE) = 0 Then
            DocEntry = Left(DocDate, 6) + "0001"
        Else
            DocEntry = Trim(oRecordSet01.Fields(0).VALUE) + 1
        End If
        ItemCode = Trim(oDS_PS_MM160L.GetValue("U_ColReg01", i))
        ItemName = Make_ItemName(Trim(oDS_PS_MM160L.GetValue("U_ColReg02", i)))
        If Trim(oDS_PS_MM160L.GetValue("U_ColNum01", i)) = "" Then
            WorkQty = 0
        Else
            WorkQty = Trim(oDS_PS_MM160L.GetValue("U_ColNum01", i))
        End If
        WorkWt = Trim(oDS_PS_MM160L.GetValue("U_ColQty01", i))
        CPrice = Trim(oDS_PS_MM160L.GetValue("U_ColSum01", i))
        CardCode = Trim(oDS_PS_MM160L.GetValue("U_ColReg03", i))
        CardName = Trim(oDS_PS_MM160L.GetValue("U_ColReg04", i))
        CpCode = Trim(oDS_PS_MM160L.GetValue("U_ColReg05", i))
        CpName = Trim(oDS_PS_MM160L.GetValue("U_ColReg06", i))
        
        If WorkWt = 0 Then
            Price = 0
        Else
            Price = Round(CPrice / WorkWt, 2)
        End If
        
        sQry = "INSERT INTO [@PS_MM005H]"
        sQry = sQry & " ("
        sQry = sQry & " DocEntry,"
        sQry = sQry & " DocNum,"
        sQry = sQry & " U_ItemCode,"
        sQry = sQry & " U_ItemName,"
'        sQry = sQry & " U_Qty,"
        sQry = sQry & " U_Weight,"
'        sQry = sQry & " U_Price,"
'        sQry = sQry & " U_LinTotal,"
        sQry = sQry & " U_BPLId,"
        sQry = sQry & " U_CgNum,"
        sQry = sQry & " U_DocDate,"
'        sQry = sQry & " U_DueDate,"
        sQry = sQry & " U_CntcCode,"
        sQry = sQry & " U_CntcName,"
'        sQry = sQry & " U_DeptCode,"
'        sQry = sQry & " U_UseDept,"
        sQry = sQry & " U_Auto,"
        sQry = sQry & " U_QCYN,"
'        sQry = sQry & " U_ReType,"
'        sQry = sQry & " U_Note,"
'        sQry = sQry & " U_IvQty,"
'        sQry = sQry & " U_IvWeight,"
        sQry = sQry & " U_OKYN,"
'        sQry = sQry & " U_OKDate,"
        sQry = sQry & " U_OrdType,"
        sQry = sQry & " U_ProcCode,"
        sQry = sQry & " U_ProcName,"
        sQry = sQry & " U_CardCode,"
        sQry = sQry & " U_CardName,"
        sQry = sQry & " U_Status"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("
        sQry = sQry & DocEntry & ","
        sQry = sQry & DocEntry & ","
        sQry = sQry & "'" & ItemCode & "',"
        sQry = sQry & "'" & ItemName & "',"
        sQry = sQry & "'" & WorkQty & "',"
'        sQry = sQry & "'" & WorkWt & "',"
'        sQry = sQry & "'" & Price & "',"
'        sQry = sQry & "'" & CPrice & "',"
        sQry = sQry & "'" & BPLId & "',"
        sQry = sQry & "'" & DocEntry & "',"
        sQry = sQry & "'" & DocDate & "',"
'        sQry = sQry & "'" & DueDate & "',"
        sQry = sQry & "'" & CntcCode & "',"
        sQry = sQry & "'" & CntcName & "',"
'        sQry = sQry & "'" & DeptCode & "',"
'        sQry = sQry & "'" & UseDept & "',"
        sQry = sQry & "'N',"
        sQry = sQry & "'N',"
'        sQry = sQry & "'10',"
'        sQry = sQry & "'" & Note & "',"
'        sQry = sQry & "'" & IvQty & "',"
'        sQry = sQry & "'" & IvWeight & "',"
        sQry = sQry & "'Y',"
'        sQry = sQry & "'" & OkDate & "',"
        sQry = sQry & "'30',"
        sQry = sQry & "'" & CpCode & "',"
        sQry = sQry & "'" & CpName & "',"
        sQry = sQry & "'" & CardCode & "',"
        sQry = sQry & "'" & CardName & "',"
        sQry = sQry & "'O'"
        sQry = sQry & ")"
        oRecordSet01.DoQuery sQry
    Next
    
    If CardCode = "" Then CardCode = "%"
    If CntcCode = "" Then CntcCode = "%"
    If DateFr = "" Then DateFr = "18990101"
    If DateTo = "" Then DateTo = "20991231"
        
    sQry = "EXEC [PS_MM160_01] '" & BPLId & "', '" & CardCode & "', '" & CntcCode & "', '" & DocDate & "', '" & DateFr & "', '" & DateTo & "', '" & ItmBsort & "', '2'"
    oRecordSet01.DoQuery sQry
    
    Sbo_Company.EndTransaction wf_Commit
    MDC_Com.MDC_GF_Message "청구 자동 등록 완료!", "S"
    Set oRecordSet01 = Nothing
    Add_MM005 = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MM005_Error:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Set oRecordSet01 = Nothing
    Add_MM005 = False
    MDC_Com.MDC_GF_Message "Add_MM005_Error:" & Err.Number & " - " & Err.Description, "E"
End Function
