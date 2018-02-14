VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM200.cls
'//  Module         : MM
'//  Description    : 제품재고수불등록(년)
'//  FormType       : PS_MM200
'//  Create Date    : 2015.02.15
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_MM200H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_MM200L As SAPbouiCOM.DBDataSource '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM200.srf")
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
    
    oFormUniqueID01 = "PS_MM200_" & GetTotalFormsCount
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
    
    Call SetDocument(oFromDocEntry01)
    'Call AddMatrixRow(0, True)
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
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
'    Dim ChildForm02 As Variant
'    Set ChildForm02 = New PS_SM030
    
    
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

                        oForm01.Items("Code").Specific.VALUE = Trim(oForm01.Items("YEAR").Specific.VALUE) + Trim(oForm01.Items("BPLId").Specific.VALUE) + Trim(oForm01.Items("Gubun").Specific.VALUE)

                    
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
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "Mat01" Then
                        
                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
'             If pval.ItemChanged = True Then
'               If pval.ColUID = "ItemCode" Then
'                  FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
'               End If
'            End If
            Case et_MATRIX_LOAD: '//11
'                Call AddMatrixRow(oMat01.VisualRowCount, False)
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
                    FormItemEnabled
                    Call AddMatrixRow(0, True)
'                    If oForm01.Mode = fm_ADD_MODE Then
''                        oForm01.Mode = fm_OK_MODE
'                        FormItemEnabled
''                        Call Sbo_Application.ActivateMenuItem("1282")
'                    ElseIf oForm01.Mode = fm_OK_MODE Then
'                        FormItemEnabled
'                        Call AddMatrixRow(0, True)
'                    End If
                ElseIf pval.ItemUID = "Btn01" Then
                     
                     
                ElseIf pval.ItemUID = "Btn02" Then
                     '//엑셀 Upload
                     Call Excel_Upload
                End If
            
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "ItemCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
                Call AddMatrixRow(oMat01.VisualRowCount, False)
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


Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim Calculate_Weight As Currency
    
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    
    Dim sSeq$

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    oMat01.FlushToDataSource
        
' Matrix 필드에 질의 응답 창 띄워주기
   Select Case oUID
        Case ""
            
            
'            oMat01.LoadFromDataSource
        Case "Mat01"
            If oCol = "ItemCode" Then
                oForm01.Freeze True
                
                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) <> "" Then
                    oMat01.FlushToDataSource
                    Call AddMatrixRow(oMat01.RowCount, False)
                    oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
                End If
                
                
                
                sQry = "Select FrgnName, U_Size, InvntryUom From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("FrgnName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                oMat01.Columns("Size").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(1).VALUE)
                oMat01.Columns("Unit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(2).VALUE)

'                oMat01.LoadFromDataSource
                oMat01.FlushToDataSource
                oForm01.Freeze False
            End If
   End Select
   
    Set oRecordSet01 = Nothing
    Exit Sub
       
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "코드를 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
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
                 oForm01.DataBrowser.BrowseBy = "Code" '//UDO방식일때
                 Call AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                    Next i
                    oMat01.FlushToDataSource        ' DBDataSource에 레코드가 한줄 더 생긴다.
                    Call oDS_PS_MM200L.RemoveRecord(oDS_PS_MM200L.Size - 1)     ' 레코드 한 줄을 지운다.
                    oMat01.LoadFromDataSource       ' DBDataSource를 매트릭스에 올리고
                    If oMat01.RowCount = 0 Then
'                        Call AddMatrixRow(1, 0, True)
                        Call AddMatrixRow(1, 0)
                    Else
                        If Trim(oDS_PS_MM200L.GetValue("U_ItemCode", oMat01.RowCount - 1)) <> "" Then
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
                oDS_PS_MM200H.setValue "Code", 0, ""
                oDS_PS_MM200H.setValue "U_YEAR", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_MM200L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
            Case "1282": '추가
                 Call FormItemEnabled '//UDO방식
                 Call AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼             '추가버튼 클릭시 Matrix에 행 추가
                 Call FormItemEnabled
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
       
    Set oDS_PS_MM200H = oForm01.DataSources.DBDataSources("@PS_MM200H")
    Set oDS_PS_MM200L = oForm01.DataSources.DBDataSources("@PS_MM200L")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    
'    Call oForm01.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)
'    oForm01.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
'    oForm01.DataSources.UserDataSources.Item("DocDate").Value = Format(Now, "YYYYMMDD")
    
'    oForm01.Items("DocDate").Specific.VALUE = Format(Date, "YYYYMMDD")
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
    
    
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Set oCombo = oForm01.Items("Gubun").Specific
    oCombo.ValidValues.Add "10", "제품"
    oCombo.ValidValues.Add "20", "제품(임가공)"
    oCombo.ValidValues.Add "30", "상품"
    oCombo.ValidValues.Add "40", "원재료"
    Call oDS_PS_MM200H.setValue("U_Gubun", 0, "10")
    
    
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
        oForm01.Items("Code").Enabled = False
        oForm01.Items("YEAR").Enabled = True
        oForm01.Items("Gubun").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", True) '//찾기
        oForm01.Items("Code").Enabled = False
        oForm01.Items("YEAR").Enabled = True
        oForm01.Items("Gubun").Enabled = True
        Call oForm01.EnableMenu("1282", True) '//추가
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("Code").Enabled = False
        oForm01.Items("YEAR").Enabled = False
        oForm01.Items("Gubun").Enabled = False
        
        Call oForm01.EnableMenu("1282", True) '//추가
      
        
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
        oDS_PS_MM200L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM200L.Offset = oRow
    oDS_PS_MM200L.setValue "LineId", oRow, oRow + 1
    oDS_PS_MM200L.setValue "U_LineNum", oRow, oRow + 1
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
            Call oDS_PS_MM200L.RemoveRecord(oDS_PS_MM200L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call AddMatrixRow(0)
            Else
                If Trim(oDS_PS_MM200L.GetValue("U_ItemCode", oMat01.RowCount - 1)) <> "" Then
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
        If oDS_PS_MM200L.GetValue("U_ItemCode", 0) = "" Then
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
            oDS_PS_MM200L.Offset = i
            Select Case True
                Case oDS_PS_MM200L.GetValue("U_ItemCode", i) = ""
                    ErrNum = 2
                    GoTo MatrixSpaceLineDel_Error
            End Select
        Next i
        
        If oDS_PS_MM200L.GetValue("U_ItemCode", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_MM200L.RemoveRecord oMat01.VisualRowCount - 1
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
        MDC_Com.MDC_GF_Message "품목코드는 필수입력사항입니다. 확인하세요.", "E"
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
        Case oDS_PS_MM200H.GetValue("U_YEAR", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
    End Select
   
   HeaderSpaceLineDel = True
   
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "년도는 필수입력 사항입니다.", "E"
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

Private Sub SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call FormItemEnabled
        Call AddMatrixRow(0, True)
    Else
        oForm01.Mode = fm_FIND_MODE
        Call FormItemEnabled
        oForm01.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Excel_Upload()
    
    Dim i               As Long
    Dim j               As Long
    Dim K               As Long
    Dim CheckLine       As Long
    Dim CheckYN         As Boolean
    Dim sPrice          As String
    Dim sFile           As String
    Dim OneRec          As String
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim xl              As Excel.Application
    Dim xlwb            As Excel.Workbook
    Dim xlsh            As Excel.Worksheet
    
    Dim oEdit           As SAPbouiCOM.EditText
    Dim oForm           As SAPbouiCOM.Form
    
    Dim V_StatusCnt&, oProValue&, TOTCNT&, tRow&  '//progbar
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim Amt01 As Long
    Dim Amt02 As Long
    Dim Amt03 As Long
    Dim Amt04 As Long
    Dim Amt05 As Long
    Dim Amt06 As Long
    Dim Amt07 As Long
    Dim Amt08 As Long
    Dim Amt09 As Long
    Dim Amt10 As Long
    Dim Amt11 As Long
    Dim Amt12 As Long
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    On Error GoTo Err
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Set oForm = Sbo_Application.Forms.ActiveForm
    
    
    
    sFile = FileListBoxForm.OpenDialog(FileListBoxForm, "*.xls", "파일선택", "C:\")
    
    If sFile = "" Then
        Exit Sub
    Else
        If Right(Replace(sFile, Chr$(0), ""), 3) <> "xls" And Right(Replace(sFile, Chr$(0), ""), 4) <> "xlsx" Then
            Sbo_Application.StatusBar.SetText "엑셀파일이 아닙니다.", bmt_Short, smt_Error
            Exit Sub
        End If
    End If
    
    
    '엑셀 Object 연결
    Set xl = CreateObject("excel.application")
    Set xlwb = xl.Workbooks.Open(sFile, , True)
    Set xlsh = xlwb.Worksheets("제품재고")
    
    
    If xlsh.UsedRange.Columns.Count <= 2 Then
        Sbo_Application.StatusBar.SetText "항목이 없습니다.", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 1).VALUE <> "품목코드" Then
        Sbo_Application.StatusBar.SetText "A열 첫번째 행 타이틀은 품목코드", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 2).VALUE <> "품명" Then
        Sbo_Application.StatusBar.SetText "B열 두번째 행 타이틀은 품명, bmt_Short, smt_Error"
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    
    '최대값 구하기
    TOTCNT = xlsh.UsedRange.Rows.Count - 1
   
    V_StatusCnt = Round(TOTCNT / 50, 0)
    oProValue = 1
    tRow = 1
    '/////////////////////////////////////////////////////
    
    For i = 2 To xlsh.UsedRange.Rows.Count
        
        
        oDS_PS_MM200L.InsertRecord (oDS_PS_MM200L.Size)
        oDS_PS_MM200L.Offset = oDS_PS_MM200L.Size - 1
        
       
        
        oMat01.Columns("LineNum").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = i - 1
        oMat01.Columns("ItemCode").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 1)
        oMat01.Columns("iwqty").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 5)  '기초수량
        oMat01.Columns("iwamt").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 6)  '기초금액
        oMat01.Columns("i1qty").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 7)  '입고수량
        oMat01.Columns("i1amt").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 8)  '입고금액
        oMat01.Columns("i2qty").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 9)  '타계정입고수량
        oMat01.Columns("i2amt").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 10)  '타계정입고금액
        oMat01.Columns("o1qty").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 11)  '출고수량
        oMat01.Columns("o1amt").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 12)  '출고금액
        oMat01.Columns("o2qty").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 13)  '타계정출고수량
        oMat01.Columns("o2amt").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 14)  '타계정출고금액
        oMat01.Columns("jgqty").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 15)  '재고수량
        oMat01.Columns("jgamt").Cells(oDS_PS_MM200L.Size - 1).Specific.VALUE = xlsh.Cells(i, 16)  '재고금액
       

        tRow = tRow + 1
        
    Next i
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    
    Sbo_Application.StatusBar.SetText "엑셀을 불러왔습니다.", bmt_Short, smt_Success
    
    '액셀개체 닫음
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    
    Exit Sub
    
Err:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText i & "행의 [" & xlsh.Cells(i, 1) & " ] 사번이 존재하지 않습니다.", bmt_Long, smt_Error
    End If
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    
End Sub




