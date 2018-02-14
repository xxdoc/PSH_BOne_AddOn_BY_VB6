VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP580"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP580cls
'//  Module         : PP
'//  Description    : TOOL 관리
'//  FormType       : PS_PP580
'//  Create Date    : 2010.12.22
'//  Creator        : Lee Byong Gak
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_PP580H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP580L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private ItemInformationCount As Long
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP580.srf")
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
    
    oFormUniqueID01 = "PS_PP580_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "Code" '//UDO방식일때

    oForm01.EnableMenu ("1293"), True     '// 행삭제
    oForm01.EnableMenu ("1287"), True     '// 복제
    oForm01.EnableMenu ("1284"), True     '// 취소
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    
    Call AddMatrixRow(0, True)
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
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                  
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                        sQry = "Select ISNULL(MAX(U_Seq),0)+ 1"
                        sQry = sQry & "From [@PS_PP580H] "
                        sQry = sQry & "Where U_ToolType = '" & Trim(oForm01.Items("ToolType").Specific.VALUE) & "' "
                        sQry = sQry & "And U_Year = '" & Trim(oForm01.Items("Year").Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        If Trim(oRecordSet01.Fields(0).VALUE) = 100 Then
                            Sbo_Application.SetStatusBarMessage "순번이 99를 초과할 수 없습니다. 관리자에게 문의하세요.", bmt_Short, True
                        End If
                        oForm01.Items("Seq").Specific.VALUE = Format(Trim(oRecordSet01.Fields(0).VALUE), "00")
                        oForm01.Items("Code").Specific.VALUE = Trim(oForm01.Items("ToolType").Specific.VALUE) + Trim(oForm01.Items("Year").Specific.VALUE) + Format(Trim(oRecordSet01.Fields(0).VALUE), "00")
'                    End If
                    
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                  
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
             If pval.ItemChanged = True Then
               If pval.ColUID = "RqstDate" Then
                  FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
               End If
            End If
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
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    ElseIf oForm01.Mode = fm_OK_MODE Then
                        FormItemEnabled
                        Call AddMatrixRow(0, True)
                    End If
                End If
            If pval.BeforeAction = False And pval.ItemChanged = True Then
               If pval.ColUID = "RqstDate" Then
                  FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
               End If
            End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            If pval.ItemChanged = True Then        '이력일자가 바뀌면 한 줄 더 생기도록
               If pval.ColUID = "RqstDate" Then
                  FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
               End If
            End If
            Case et_MATRIX_LOAD: '//11
                Call AddMatrixRow(pval.Row, False)
                Call FormItemEnabled
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
    
    Set oRecordSet01 = Nothing
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

' 원가요소에 대한 원가요소 이름 뿌려주기 위한 선언부분

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim sSeq$

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
        
' Matrix 필드에 질의 응답 창 띄워주기
   Select Case oCol
        Case "RqstDate"
          oMat01.LoadFromDataSource
          
            '--------------------------------------------------------------------------------------------
            If oRow = oMat01.RowCount And Trim(oDS_PS_PP580L.GetValue("U_RqstDate", oRow - 1)) <> "" Then
            '// 다음 라인 추가
            Call AddMatrixRow(0, False)
            oMat01.Columns("RqstDate").Cells(oRow).Click ct_Regular
            End If
            
   End Select
   
    Set oRecordSet01 = Nothing
    Exit Sub
       
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "이력일자를 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
    Else
        MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo MenuEvent_Error
    
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim i&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                  Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                 oForm01.DataBrowser.BrowseBy = "Code" '//UDO방식일때
            Case "1282": '추가
                 oForm01.DataBrowser.BrowseBy = "ToolType" '//UDO방식일때
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineId").Cells(i).Specific.VALUE = i
                    Next i
                    oMat01.FlushToDataSource        ' DBDataSource에 레코드가 한줄 더 생긴다.
                    Call oDS_PS_PP580L.RemoveRecord(oDS_PS_PP580L.Size - 1)     ' 레코드 한 줄을 지운다.
                    oMat01.LoadFromDataSource       ' DBDataSource를 매트릭스에 올리고
                    If oMat01.RowCount = 0 Then
'                        Call AddMatrixRow(1, 0, True)
                        Call AddMatrixRow(1, 0)
                    Else
                        If Trim(oDS_PS_PP580L.GetValue("U_RqstDate", oMat01.RowCount - 1)) <> "" Then
'                            Call AddMatrixRow(1, oMat01.RowCount, True)
                            Call AddMatrixRow(1, oMat01.RowCount)
                        End If
                    End If
                End If
            Case "1281": '찾기
                 Call AddMatrixRow(0, True) '//UDO방식
                 Call FormItemEnabled
            Case "1287": '복제
                oForm01.Freeze True
                oDS_PS_PP580H.setValue "Code", 0, ""
                oDS_PS_PP580H.setValue "U_Seq", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_PP580L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
            Case "1282": '추가
                '//Call FormItemEnabled '//UDO방식
                 Call AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼             '추가버튼 클릭시 Matrix에 행 추가
                '//Call FormItemEnabled
        End Select
    End If
    Exit Sub
MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function CreateItems() As Boolean
On Error GoTo CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
       
    Set oDS_PS_PP580H = oForm01.DataSources.DBDataSources("@PS_PP580H")
    Set oDS_PS_PP580L = oForm01.DataSources.DBDataSources("@PS_PP580L")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    
'    Call oForm01.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)
'    oForm01.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
'    oForm01.DataSources.UserDataSources.Item("DocDate").Value = Format(Now, "YYYYMMDD")
    
    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
    oForm01.Items("Year").Specific.VALUE = Left(Date, 4)
'    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
   
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error

    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP580", "Mat01", "Condit", "I", "입고")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP580", "Mat01", "Condit", "A", "A")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP580", "Mat01", "Condit", "B", "B")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP580", "Mat01", "Condit", "C", "C")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP580", "Mat01", "Condit", "D", "D")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP580", "Mat01", "Condit", "X", "폐기")

    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Condit"), "PS_PP580", "Mat01", "Condit")

    
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
             
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop

    Set oCombo = oForm01.Items("ToolType").Specific
    oCombo.ValidValues.Add "3", "V금형"
    oCombo.ValidValues.Add "4", "FINAL 압연"

    oCombo.Select 0, psk_Index
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Call oForm01.Freeze(False)
    Exit Sub
ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        'Call FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", True) '//찾기
        oForm01.Items("Code").Enabled = True
        Call oForm01.EnableMenu("1282", True) '//추가
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("Code").Enabled = False
        oForm01.Items("ToolType").Enabled = False
        oForm01.Items("Year").Enabled = False
        oForm01.Items("DocDate").Enabled = False
        
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo AddMatrixRow_Error
    
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oRow = oMat01.RowCount
        oDS_PS_PP580L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP580L.Offset = oRow
    oDS_PS_PP580L.setValue "LineId", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//작업
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineId").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_PP580L.RemoveRecord(oDS_PS_PP580L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call AddMatrixRow(0)
            Else
                If Trim(oDS_PS_PP580L.GetValue("U_RqstDate", oMat01.RowCount - 1)) <> "" Then
                    Call AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i             As Long
    Dim ErrNum        As Integer
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    Dim sQry          As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    ErrNum = 0
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '// 화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.FlushToDataSource
    
    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 Then
        If oDS_PS_PP580L.GetValue("U_RqstDate", 0) = "" Then
            ErrNum = 1
            GoTo MatrixSpaceLineDel_Error
        End If
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '// 맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '// 이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 2
            oDS_PS_PP580L.Offset = i
            Select Case True
                Case oDS_PS_PP580L.GetValue("U_RqstDate", i) = ""
                    ErrNum = 2
                    GoTo MatrixSpaceLineDel_Error
                Case oDS_PS_PP580L.GetValue("U_CompDate", i) = ""
                    ErrNum = 3
                    GoTo MatrixSpaceLineDel_Error
                Case oDS_PS_PP580L.GetValue("U_DiamTop", i) = ""
                    ErrNum = 4
                    GoTo MatrixSpaceLineDel_Error
                Case oDS_PS_PP580L.GetValue("U_DiamBot", i) = ""
                    ErrNum = 5
                Case oDS_PS_PP580L.GetValue("U_Corner", i) = ""
                    ErrNum = 6
                Case oDS_PS_PP580L.GetValue("U_Thick", i) = ""
                    ErrNum = 7
                Case oDS_PS_PP580L.GetValue("U_Condit", i) = ""
                    ErrNum = 8
                Case oDS_PS_PP580L.GetValue("U_WorkQty", i) = ""
                    ErrNum = 9
                    GoTo MatrixSpaceLineDel_Error
            End Select
        Next i
        
        If oDS_PS_PP580L.GetValue("U_RqstDate", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_PP580L.RemoveRecord oMat01.VisualRowCount - 1
        End If
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.LoadFromDataSource
    
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인데이타가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "이력일자는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "완료일자는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "직경(상부)는 필수입력사항입니다. 확인하세요", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "직경(하부)는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "Corner은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "두께는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "상태는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "사용량은 필수입력사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_PP580H.GetValue("U_ToolType", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_PP580H.GetValue("U_Year", 0) = ""
            ErrNum = 2
'        Case oDS_PS_PP580H.GetValue("U_Seq", 0) = ""
'            ErrNum = 3
        Case oDS_PS_PP580H.GetValue("U_Name", 0) = ""
            ErrNum = 4
        Case oDS_PS_PP580H.GetValue("U_Size", 0) = ""
            ErrNum = 5
        Case oDS_PS_PP580H.GetValue("U_DocDate", 0) = ""
            ErrNum = 6
            GoTo HeaderSpaceLineDel_Error
    End Select
   
   HeaderSpaceLineDel = True
   
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "금형종류는 필수입력 사항입니다.", "E"
    End If
    If ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "년도는 필수입력 사항입니다.", "E"
    End If
'    If ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "순번은 필수입력 사항입니다.", "E"
'    End If
    If ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "품명은 필수입력 사항입니다.", "E"
    End If
    If ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "규격은 필수입력 사항입니다.", "E"
    End If
    If ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "등록일자는 필수입력 사항입니다.", "E"
    End If
    If ErrNum = "" Then
        MDC_Com.MDC_GF_Message "정상등록 되었습니다.", "W"
    End If
    HeaderSpaceLineDel = False
End Function

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
