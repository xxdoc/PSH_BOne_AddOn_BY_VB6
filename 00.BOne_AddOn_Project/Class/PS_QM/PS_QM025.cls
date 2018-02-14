VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM025"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM025.cls
'//  Module         : QM
'//  Description    : 검사성적서 납품처변경
'//  FormType       : PS_QM025
'//  Create Date    : 2011.01.20
'//  Modified Date  :
'//  Creator        : NGY
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_QM025L As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oForm01_Mode&
Private oSeq&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM025.srf")
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
    
    oFormUniqueID01 = "PS_QM025_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    
    oForm01_Mode = fm_ADD_MODE '//모드초기화
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
                    If pval.ItemUID = "Mat01" Then
                        If oMat01.Columns("CardCode").Cells(pval.Row).Specific.VALUE = "" Then
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
                    If oForm01_Mode = fm_UPDATE_MODE Then
                        If CH_QM020(pval) = False Then '//항목 Update
                            BubbleEvent = False
                            Exit Sub
                        End If
                        oForm01_Mode = fm_OK_MODE
                        oMat01.Clear
                        oDS_PS_QM025L.Clear
                        Call LoadData
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
                ElseIf pval.ItemUID = "Btn03" Then
                    If oForm01_Mode = fm_UPDATE_MODE Then
                        If CH_CARDCODE(pval) = False Then    '//납품처 일괄변경
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
                    
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "BPLId" Then
                    oMat01.Clear
                    oDS_PS_QM025L.Clear
                    oForm01_Mode = fm_OK_MODE
                    Call LoadCaption
                End If
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "CardCode" Then
                            oForm01_Mode = fm_UPDATE_MODE
                            Call LoadCaption
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
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
                Set oDS_PS_QM025L = Nothing
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
        Case Trim(oForm01.Items("PackNoF").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("PackNoT").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "PACKING 시작번호는 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "PACKING 종료번호는 필수입력 사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
    Set oDS_PS_QM025L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
        
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
    
'    oForm01.Items("PackNoF").Specific.VALUE = "20100101001"
'    oForm01.Items("PackNoT").Specific.VALUE = "20111231001"
    '//아이디별 사번 세팅
    'oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
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
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim ItemCode$, Qty&
    Dim Calculate_Weight As Currency
    Dim Calculate_Qty As Currency
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim CardCode$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "Mat01"
            If oCol = "CardCode" Then
                CardCode = Trim(oMat01.Columns("CardCode").Cells(oRow).Specific.VALUE)
                sQry = "Select CardName From OCRD Where CardCode = '" & Trim(CardCode) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("CardName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                
'                For i = 1 To oMat01.RowCount
'                    oMat01.Columns("CardCode").Cells(i).Specific.VALUE = CardCode
'                    oMat01.Columns("CardName").Cells(i).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'                Next
            End If
        
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
    If oForm01_Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "수정"
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
    
    Dim BPLId$, PackNoF$, PackNoT$
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    PackNoF = Trim(oForm01.Items("PackNoF").Specific.VALUE)
    PackNoT = Trim(oForm01.Items("PackNoT").Specific.VALUE)
    
    sQry = "EXEC [PS_QM025_01] '" & BPLId & "', '" & PackNoF & "', '" & PackNoT & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_QM025L.Clear
        
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        oForm01_Mode = fm_OK_MODE
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
'    Dim ProgBar01 As SAPbouiCOM.ProgressBar
'    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_QM025L.Size Then
            oDS_PS_QM025L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_QM025L.Offset = i
        oDS_PS_QM025L.setValue "U_LineNum", i, i + 1
        oDS_PS_QM025L.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("U_OrdNum").VALUE)
        oDS_PS_QM025L.setValue "U_ColDt01", i, Format(Trim(oRecordSet01.Fields("U_InspDate").VALUE), "YYYYMMDD")
        oDS_PS_QM025L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("U_CardCode").VALUE)
        oDS_PS_QM025L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("U_CardName").VALUE)
        oDS_PS_QM025L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
        oDS_PS_QM025L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("U_ItemName").VALUE)
        
        oRecordSet01.MoveNext
'        ProgBar01.VALUE = ProgBar01.VALUE + 1
'        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns

    'ProgBar01.Stop
    oForm01.Freeze False
    oForm01_Mode = fm_ADD_MODE
    
'    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
'    ProgBar01.Stop
 '   Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function CH_QM020(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo CH_QM020_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    Dim BPLId$, PackNoF$, PackNoT$
    Dim OrdNum$, CardCode$, CardName$
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    PackNoF = Trim(oForm01.Items("PackNoF").Specific.VALUE)
    PackNoT = Trim(oForm01.Items("PackNoT").Specific.VALUE)
    
    oMat01.FlushToDataSource
    
    
    Sbo_Company.StartTransaction
    
    For i = 0 To oMat01.RowCount - 1
        CardCode = Trim(oDS_PS_QM025L.GetValue("U_ColReg02", i)) '//납품처코드
        sQry = "Select CardName From OCRD Where CardCode = '" & CardCode & "'"
        oRecordSet01.DoQuery sQry
        
        If oRecordSet01.RecordCount = 0 Then
            MDC_Com.MDC_GF_Message "납품처에 잘못된 자료가 있습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
            Set oRecordSet01 = Nothing
            If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
            CH_QM020 = False
            Exit Function
        End If
    Next
    
    For i = 0 To oMat01.RowCount - 1
        
        OrdNum = Trim(oDS_PS_QM025L.GetValue("U_ColReg01", i)) '//작업지시번호
        CardCode = Trim(oDS_PS_QM025L.GetValue("U_ColReg02", i)) '//납품처코드
        sQry = "Select CardName From OCRD Where CardCode = '" & CardCode & "'"
        oRecordSet01.DoQuery sQry
        CardName = Trim(oRecordSet01.Fields(0).VALUE)
        
        
        sQry = "Update [@PS_QM020H]"
        sQry = sQry & " set U_CardCode = '" + CardCode + "',"
        sQry = sQry & " U_CardName = '" + CardName + "'"
        sQry = sQry & " Where U_OrdNum = '" + OrdNum + "'"
        
        oRecordSet01.DoQuery sQry
    Next
    
        
    Sbo_Company.EndTransaction wf_Commit
    

    
    MDC_Com.MDC_GF_Message "검사성적서 납품처수정 완료!", "S"
    Set oRecordSet01 = Nothing
    CH_QM020 = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CH_QM020_Error:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Set oRecordSet01 = Nothing
    CH_QM020 = False
    MDC_Com.MDC_GF_Message "CH_QM020_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function CH_CARDCODE(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo CH_CARDCODE_Error

    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim CardCode$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    CardCode = Trim(oMat01.Columns("CardCode").Cells(1).Specific.VALUE)
    sQry = "Select CardName From OCRD Where CardCode = '" & Trim(CardCode) & "'"
    oRecordSet01.DoQuery sQry
                    
    For i = 1 To oMat01.RowCount
        oMat01.Columns("CardCode").Cells(i).Specific.VALUE = CardCode
        oMat01.Columns("CardName").Cells(i).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    Next
            
    Set oRecordSet01 = Nothing
    CH_CARDCODE = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CH_CARDCODE_Error:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Set oRecordSet01 = Nothing
    CH_CARDCODE = False
    MDC_Com.MDC_GF_Message "CH_CARDCODE_Error:" & Err.Number & " - " & Err.Description, "E"
End Function


