VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S392"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : S392.cls
'//  Module         :
'//  Description    : 분개
'//  FormType       : 392
'//  Create Date    : 2010.11.23
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Public oFormMode       As String

Public oFormUniqueID02 As String
Public oForm02         As SAPbouiCOM.Form
Public oMat02          As SAPbouiCOM.Matrix

Private oMat01Row      As Long

Public Sub LoadForm(ByVal oFormUniqueID As String)

    oFormUniqueID01 = oFormUniqueID
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    Set oMat01 = oForm01.Items("76").Specific
    oFormMode = oForm01.Mode
    
    Call CreateItems
    oForm01.Mode = oFormMode
    
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error

    Dim i         As Integer
    Dim j         As Integer
    Dim z         As Integer
    Dim ErrNum    As Integer
    Dim BPLID     As String
    Dim DocDateFr As String
    Dim DocDateTo As String
    Dim sQry      As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "76" Then
                        If pval.ColUID = "U_VatBP" Then
                            If oMat01.Columns("U_VatBP").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                    ElseIf pval.ItemUID = "VatBP" Then '거래처
                        If oForm01.Items("VatBP").Specific.VALUE = "" Then
                            Call Sbo_Application.ActivateMenuItem("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "BillCode" Then '법정증빙코드
                        If oForm01.Items("BillCode").Specific.VALUE = "" Then
                            Call Sbo_Application.ActivateMenuItem("7425")
                            BubbleEvent = False
                        End If
                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
                If pval.ItemUID = "76" Then
                    If pval.Row Then
                        oMat01Row = pval.Row
                    End If
                End If
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
            Case et_FORM_LOAD
            Case et_ITEM_PRESSED: '//1

                If pval.ItemUID = "Btn01" Then
                    Call Print_Report01
                    Call UpdateOCRData '배부규칙 업데이트(2017.05.23 송명규)
                ElseIf pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1291")
                    End If
                ElseIf pval.ItemUID = "Btn02" Then '//감가상각비
                    If oForm01.Items("6").Specific.VALUE = "" Then
                        MDC_Com.MDC_GF_Message "전기일자는 필수입니다. 확인하세요.", "E"
                        BubbleEvent = False
                        Exit Sub

                    ElseIf oForm01.Items("BPLId01").Specific.VALUE = "" Then
                        MDC_Com.MDC_GF_Message "사업장은 필수입니다. 확인하세요.", "E"
                        BubbleEvent = False
                        Exit Sub
                    End If

                    BPLID = Trim(oForm01.Items("BPLId01").Specific.VALUE)
                    DocDateFr = Left(Trim(oForm01.Items("6").Specific.VALUE), 6) + "01"
                    DocDateTo = Trim(oForm01.Items("6").Specific.VALUE)
                    sQry = "EXEC [S392_02] '" & BPLID & "', '" & DocDateFr & "', '" & DocDateTo & "'"
                    oRecordSet01.DoQuery sQry

                    If (oRecordSet01.RecordCount = 0) Then
                        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                        Set oRecordSet01 = Nothing
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
'                    oForm01.Freeze True
'                    Dim ProgBar01 As SAPbouiCOM.ProgressBar
'                    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("데이터 가져오기 시작!", oRecordSet01.RecordCount, False)
                        
                    For i = 0 To oRecordSet01.RecordCount - 1
                        oMat01.Columns("1").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        If Trim(oRecordSet01.Fields(2).VALUE) > 0 Then
                            oMat01.Columns("5").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields(2).VALUE)

                            If oMat01.Columns("10002014").Cells(i + 1).Specific.VALUE = "" Then
                                oMat01.Columns("10002014").Cells(i + 1).Click ct_Regular
                                Sbo_Application.SendKeys "{TAB}"

                                Set oForm02 = Nothing
                                Set oMat02 = Nothing
                                Set oForm02 = Sbo_Application.Forms.ActiveForm

                                Set oMat02 = oForm02.Items("10000003").Specific
                                oMat02.Columns("10000003").Cells(1).Specific.Select Trim(oRecordSet01.Fields(1).VALUE)
                                oForm02.Items("10000001").Click ct_Regular
                                oForm02.Items("10000001").Click ct_Regular

                                Set oForm02 = Nothing
                                Set oMat02 = Nothing
                            End If
                        Else
                            oMat01.Columns("6").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields(3).VALUE)
                        End If

                        oRecordSet01.MoveNext
'                        ProgBar01.VALUE = ProgBar01.VALUE + 1
'                        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 데이터 가지고 오는 중...!"
                    Next
                  
'                    oMat01.LoadFromDataSource
'                    oMat01.AutoResizeColumns
'                    ProgBar01.Stop
'                    oForm01.Freeze False
                
'                  Set ProgBar01 = Nothing
                ElseIf pval.ItemUID = "Btn03" Then '//부자재불출
                    If oForm01.Items("6").Specific.VALUE = "" Then
                        MDC_Com.MDC_GF_Message "전기일자는 필수입니다. 확인하세요.", "E"
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Items("BPLId01").Specific.VALUE = "" Then
                        MDC_Com.MDC_GF_Message "사업장은 필수입니다. 확인하세요.", "E"
                        BubbleEvent = False
                        Exit Sub
                    End If

                    BPLID = Trim(oForm01.Items("BPLId01").Specific.VALUE)
                    DocDateFr = Left(Trim(oForm01.Items("6").Specific.VALUE), 6) + "01"
                    DocDateTo = Trim(oForm01.Items("6").Specific.VALUE)
                    sQry = "EXEC [S392_01] '" & BPLID & "', '" & DocDateFr & "', '" & DocDateTo & "'"
                    oRecordSet01.DoQuery sQry

                    If (oRecordSet01.RecordCount = 0) Then
                        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                        Set oRecordSet01 = Nothing
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
'                    oForm01.Freeze True
'                    Dim ProgBar01 As SAPbouiCOM.ProgressBar
'                    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("데이터 가져오기 시작!", oRecordSet01.RecordCount, False)
                        
                    For i = 0 To oRecordSet01.RecordCount - 1
                        oMat01.Columns("1").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        If Trim(oRecordSet01.Fields(2).VALUE) <> 0 Then
                            oMat01.Columns("5").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields(2).VALUE)

                            If oMat01.Columns("10002014").Cells(i + 1).Specific.VALUE = "" Then
                                oMat01.Columns("10002014").Cells(i + 1).Click ct_Regular
                                Sbo_Application.SendKeys "{TAB}"

                                Set oForm02 = Nothing
                                Set oMat02 = Nothing
                                Set oForm02 = Sbo_Application.Forms.ActiveForm

                                Set oMat02 = oForm02.Items("10000003").Specific
                                oMat02.Columns("10000003").Cells(1).Specific.Select Trim(oRecordSet01.Fields(1).VALUE)
                                oForm02.Items("10000001").Click ct_Regular
                                oForm02.Items("10000001").Click ct_Regular

                                Set oForm02 = Nothing
                                Set oMat02 = Nothing
                            End If
                        Else
                            oMat01.Columns("6").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields(3).VALUE)
                        End If

                        oRecordSet01.MoveNext
'                        ProgBar01.VALUE = ProgBar01.VALUE + 1
'                        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 데이터 가지고 오는 중...!"
                    Next
                  
'                    oMat01.LoadFromDataSource
'                    oMat01.AutoResizeColumns
'                    ProgBar01.Stop
'                    oForm01.Freeze False
                
'                  Set ProgBar01 = Nothing

                ElseIf pval.ItemUID = "BtnApply" Then '전체적용

                    Call oForm01.Freeze(True)
                    For i = 1 To oMat01.VisualRowCount
                        If oMat01.Columns("1").Cells(i).Specific.VALUE <> "" Then

                            oMat01.Columns("U_VatBP").Cells(i).Specific.VALUE = oForm01.Items("VatBP").Specific.VALUE '거래처
                            oMat01.Columns("U_VatBPName").Cells(i).Specific.VALUE = oForm01.Items("VatBPName").Specific.VALUE '거래처명
                            oMat01.Columns("U_VatRegN").Cells(i).Specific.VALUE = oForm01.Items("VatRegNo").Specific.VALUE '사업자등록번호

                            oMat01.Columns("U_BillCode").Cells(i).Specific.VALUE = oForm01.Items("BillCode").Specific.VALUE '법정증빙코드
                            oMat01.Columns("U_BillName").Cells(i).Specific.VALUE = oForm01.Items("BillName").Specific.VALUE '법정증빙명
                            oMat01.Columns("U_BillCMT").Cells(i).Specific.VALUE = oForm01.Items("BillCMT").Specific.VALUE '법정증빙비고

                            'Call oMat01.AutoResizeColumns

                        End If
                    Next i
                    Call oForm01.Freeze(False)

                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                If pval.ItemChanged = True Then
                    '//헤더
                    If pval.ItemUID = "BPLId02" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ItemUID = "BPLId01" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    '//라인
                    If pval.ItemUID = "76" And (pval.ColUID = "U_BPLId") Then
                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                    End If
                End If
            Case et_CLICK: '//6
                If pval.ItemUID = "76" And pval.ColUID = "U_BPLId" Then
                    If oMat01.VisualRowCount > 1 And oMat01.Columns("U_BPLId").Cells(pval.Row).Specific.VALUE <> "" Then
                        oForm01.Freeze True
                        oForm01.Items("BPLId02").Specific.Select oMat01.Columns("U_BPLId").Cells(pval.Row).Specific.VALUE
                        oForm01.Freeze False
                    End If
                End If

            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10

                If pval.ItemChanged = True Then

                    If pval.ItemUID = "VatBP" Then

                        sQry = "        SELECT  CardName,"
                        sQry = sQry & "         VATRegNum"
                        sQry = sQry & " FROM    OCRD"
                        sQry = sQry & " WHERE   CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"

                        Call oRecordSet01.DoQuery(sQry)

                        oForm01.Items("VatBPName").Specific.VALUE = oRecordSet01.Fields("CardName").VALUE
                        oForm01.Items("VatRegNo").Specific.VALUE = oRecordSet01.Fields("VATRegNum").VALUE

                    ElseIf pval.ItemUID = "BillCode" Then

                        sQry = "        SELECT  U_CdName"
                        sQry = sQry & " FROM    [@PS_SY001L]"
                        sQry = sQry & " WHERE   Code ='F005'"
                        sQry = sQry & "         AND U_Minor = '" & oForm01.Items("BillCode").Specific.VALUE & "'"

                        Call oRecordSet01.DoQuery(sQry)

                        oForm01.Items("BillName").Specific.VALUE = oRecordSet01.Fields("U_CdName").VALUE

                    End If

                End If

            Case et_MATRIX_LOAD: '//11

                Call oMat01.AutoResizeColumns

            Case et_FORM_ACTIVATE: '//18

                If (pval.ItemUID = "CardCode") Then

                End If

            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
                Call Form_Resize(FormUID, pval, BubbleEvent)
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4

'                Call oForm01.Freeze(True)
'                If pval.ItemUID = "76" Then
'
'                    If pval.ColUID = "10002014" Then
'
'                        If oMat01.VisualRowCount > 1 Then
'
'                            sQry = "        SELECT      PrcName"
'                            sQry = sQry & " FROM        OPRC"
'                            sQry = sQry & " WHERE       PrcCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"
'
'                            Call oForm01.Freeze(True)
'                            Call oRecordSet01.DoQuery(sQry)
'
'                            oMat01.Columns("U_OcrName").Cells(pval.Row).Specific.VALUE = oRecordSet01.Fields("PrcName").VALUE
'
'                            Call oForm01.Freeze(False)
'
'                        End If
'
'                    End If
'
'                    Call oMat01.AutoResizeColumns
'
'                End If
'                Call oForm01.Freeze(False)

            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
'                Set oMat01 = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    
    Exit Sub

Raise_ItemEvent_Error:
    Set oRecordSet01 = Nothing
    Set oForm02 = Nothing
    Set oMat02 = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error

    Dim i As Integer

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

    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1285": '복원
            Case "1287": '복제
            Case "1293": '행삭제
            Case "1281": '찾기
                oForm01.DataSources.UserDataSources("RadioBtn01").VALUE = "0"
            Case "1282": '추가
                oForm01.Freeze True
                Call oForm01.Items("BPLId01").Specific.Select(User_BPLId(), psk_ByValue)
                oForm01.Items("6").Click ct_Regular
                oForm01.DataSources.UserDataSources("RadioBtn01").VALUE = "0"
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                oForm01.DataSources.UserDataSources("RadioBtn01").VALUE = "0"
                Call oMat01.AutoResizeColumns
        End Select
    End If
    
    Exit Sub

Raise_MenuEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error

    If (eventInfo.BeforeAction = True) Then
        If eventInfo.ItemUID = "76" Then
            If eventInfo.Row > 0 Then
                oMat01Row = eventInfo.Row
            End If
        End If
    ElseIf (eventInfo.BeforeAction = False) Then
        
    End If
    
    Exit Sub

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
        End Select
    End If
    
    Exit Sub
    
Raise_FormDataEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_FormDataEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error

    Dim stdItem    As Object
    Dim newItem    As Object
    Dim oEdit      As SAPbouiCOM.EditText
    Dim optBtn     As SAPbouiCOM.OptionBtn
    Dim oCombo01   As SAPbouiCOM.ComboBox
    Dim oCombo02   As SAPbouiCOM.ComboBox
    Dim oStatic    As SAPbouiCOM.StaticText
        
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim sQry       As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RdoBtn01 As Object '담당(경리부서)
    Dim RdoBtn02 As Object '팀장(경리부서)
    Dim RdoBtn03 As Object '사업부장(경리부서)
    Dim RdoBtn04 As Object '전무(경리부서)
    Dim RdoBtn05 As Object '부사장(경리부서)
'    Dim RdoBtn06 As Object '사장(경리부서)

    Dim RdoBtn11 As Object '담당(품의부서)
    Dim RdoBtn12 As Object '팀장(품의부서)
    Dim RdoBtn13 As Object '사업부장(품의부서)
    Dim RdoBtn14 As Object '전무(품의부서)
    Dim RdoBtn15 As Object '부사장(품의부서)
'    Dim RdoBtn16 As Object '사장(품의부서)

    oForm01.Freeze True
    
    Set stdItem = oForm01.Items("2") '기준 아이템(취소버튼)
    
    '//////////경리부서 전결용 라디오버튼//////////
    Call oForm01.DataSources.UserDataSources.Add("RadioBtn01", dt_SHORT_TEXT, 1)
    
    '담당(경리부서) 라디오 버튼_S
    Set RdoBtn01 = oForm01.Items.Add("RdoBtn01", it_OPTION_BUTTON)
    RdoBtn01.Left = stdItem.Left + 85
    RdoBtn01.Top = stdItem.Top - 8
    RdoBtn01.Height = stdItem.Height
    RdoBtn01.Width = 48
    RdoBtn01.Specific.Caption = "담당"
    
    Set optBtn = oForm01.Items("RdoBtn01").Specific
    optBtn.ValOn = "A"
    optBtn.ValOff = "0"
    Call optBtn.DataBind.SetBound(True, , "RadioBtn01")
    optBtn.Selected = True
    '담당(경리부서) 라디오 버튼_E
    
    '팀장(경리부서) 라디오버튼_S
    Set RdoBtn02 = oForm01.Items.Add("RdoBtn02", it_OPTION_BUTTON)
    RdoBtn02.Left = RdoBtn01.Left + RdoBtn01.Width - 3
    RdoBtn02.Top = RdoBtn01.Top
    RdoBtn02.Height = RdoBtn01.Height
    RdoBtn02.Width = 48
    RdoBtn02.Specific.Caption = "팀장"
    
    Set optBtn = oForm01.Items("RdoBtn02").Specific
    optBtn.ValOn = "B"
    optBtn.ValOff = "0"
    Call optBtn.DataBind.SetBound(True, , "RadioBtn01")
    Call optBtn.GroupWith("RdoBtn01")
    '팀장(경리부서) 라디오버튼_E
    
    '사업부장(경리부서) 라디오버튼_S, 상무로 변경(2018.01.02 송명규)
    Set RdoBtn03 = oForm01.Items.Add("RdoBtn03", it_OPTION_BUTTON)
    RdoBtn03.Left = RdoBtn02.Left + RdoBtn02.Width - 3
    RdoBtn03.Top = RdoBtn02.Top
    RdoBtn03.Height = RdoBtn02.Height
    RdoBtn03.Width = 48
    RdoBtn03.Specific.Caption = "상무"
    
    Set optBtn = oForm01.Items("RdoBtn03").Specific
    optBtn.ValOn = "C"
    optBtn.ValOff = "0"
    Call optBtn.DataBind.SetBound(True, , "RadioBtn01")
    Call optBtn.GroupWith("RdoBtn01")
    '사업부장(경리부서) 라디오버튼_E
    
    '전무(경리부서) 라디오버튼_S
    Set RdoBtn04 = oForm01.Items.Add("RdoBtn04", it_OPTION_BUTTON)
    RdoBtn04.Left = RdoBtn03.Left + RdoBtn03.Width - 3
    RdoBtn04.Top = RdoBtn03.Top
    RdoBtn04.Height = RdoBtn03.Height
    RdoBtn04.Width = 48
    RdoBtn04.Specific.Caption = "전무"
    
    Set optBtn = oForm01.Items("RdoBtn04").Specific
    optBtn.ValOn = "D"
    optBtn.ValOff = "0"
    Call optBtn.DataBind.SetBound(True, , "RadioBtn01")
    Call optBtn.GroupWith("RdoBtn01")
    '전무(경리부서) 라디오버튼_E
    
    '부사장 삭제(2017.04.11 송명규, 유주용, 유혜연 요청), 복원(신중현 전무님 승진), 2018.01.02 송명규
    '부사장(경리부서) 라디오버튼_S
    Set RdoBtn05 = oForm01.Items.Add("RdoBtn05", it_OPTION_BUTTON)
    RdoBtn05.Left = RdoBtn04.Left + RdoBtn04.Width - 3
    RdoBtn05.Top = RdoBtn04.Top
    RdoBtn05.Height = RdoBtn04.Height
    RdoBtn05.Width = 58
    RdoBtn05.Specific.Caption = "부사장"
    
    Set optBtn = oForm01.Items("RdoBtn05").Specific
    optBtn.ValOn = "E"
    optBtn.ValOff = "0"
    Call optBtn.DataBind.SetBound(True, , "RadioBtn01")
    Call optBtn.GroupWith("RdoBtn01")
    '부사장(경리부서) 라디오버튼_E
    
    oForm01.DataSources.UserDataSources("RadioBtn01").VALUE = "0"
    '//////////경리부서 전결용 라디오버튼//////////
    
    '//////////품의부서 전결용 라디오버튼//////////
    Call oForm01.DataSources.UserDataSources.Add("RadioBtn11", dt_SHORT_TEXT, 1)
    
    '담당(품의부서) 라디오 버튼_S
    Set RdoBtn11 = oForm01.Items.Add("RdoBtn11", it_OPTION_BUTTON)
    RdoBtn11.Left = stdItem.Left + 85
    RdoBtn11.Top = stdItem.Top + 11
    RdoBtn11.Height = stdItem.Height
    RdoBtn11.Width = 48
    RdoBtn11.Specific.Caption = "담당"
    
    Set optBtn = oForm01.Items("RdoBtn11").Specific
    optBtn.ValOn = "A"
    optBtn.ValOff = "0"
    Call optBtn.DataBind.SetBound(True, , "RadioBtn11")
    optBtn.Selected = True
    '담당(품의부서) 라디오 버튼_E
    
    '팀장(품의부서) 라디오버튼_S
    Set RdoBtn12 = oForm01.Items.Add("RdoBtn12", it_OPTION_BUTTON)
    RdoBtn12.Left = RdoBtn11.Left + RdoBtn11.Width - 3
    RdoBtn12.Top = RdoBtn11.Top
    RdoBtn12.Height = RdoBtn01.Height
    RdoBtn12.Width = 48
    RdoBtn12.Specific.Caption = "팀장"
    
    Set optBtn = oForm01.Items("RdoBtn12").Specific
    optBtn.ValOn = "B"
    optBtn.ValOff = "0"
    Call optBtn.DataBind.SetBound(True, , "RadioBtn11")
    Call optBtn.GroupWith("RdoBtn11")
    '팀장(품의부서) 라디오버튼_E
    
    '사업부장(품의부서) 라디오버튼_S, 상무로 변경(2018.01.02 송명규)
    Set RdoBtn13 = oForm01.Items.Add("RdoBtn13", it_OPTION_BUTTON)
    RdoBtn13.Left = RdoBtn12.Left + RdoBtn12.Width - 3
    RdoBtn13.Top = RdoBtn12.Top
    RdoBtn13.Height = RdoBtn12.Height
    RdoBtn13.Width = 48
    RdoBtn13.Specific.Caption = "상무"
    
    Set optBtn = oForm01.Items("RdoBtn13").Specific
    optBtn.ValOn = "C"
    optBtn.ValOff = "0"
    Call optBtn.DataBind.SetBound(True, , "RadioBtn11")
    Call optBtn.GroupWith("RdoBtn11")
    '사업부장(품의부서) 라디오버튼_E
    
    '전무(품의부서) 라디오버튼_S
    Set RdoBtn14 = oForm01.Items.Add("RdoBtn14", it_OPTION_BUTTON)
    RdoBtn14.Left = RdoBtn13.Left + RdoBtn13.Width - 3
    RdoBtn14.Top = RdoBtn13.Top
    RdoBtn14.Height = RdoBtn13.Height
    RdoBtn14.Width = 48
    RdoBtn14.Specific.Caption = "전무"
    
    Set optBtn = oForm01.Items("RdoBtn14").Specific
    optBtn.ValOn = "D"
    optBtn.ValOff = "0"
    Call optBtn.DataBind.SetBound(True, , "RadioBtn11")
    Call optBtn.GroupWith("RdoBtn11")
    '전무(품의부서) 라디오버튼_E
    
    '부사장 삭제(2017.04.11 송명규, 유주용, 유혜연 요청), 복원(신중현 전무님 승진), 2018.01.02 송명규
    '부사장(품의부서) 라디오버튼_S
    Set RdoBtn15 = oForm01.Items.Add("RdoBtn15", it_OPTION_BUTTON)
    RdoBtn15.Left = RdoBtn14.Left + RdoBtn14.Width - 3
    RdoBtn15.Top = RdoBtn14.Top
    RdoBtn15.Height = RdoBtn14.Height
    RdoBtn15.Width = 58
    RdoBtn15.Specific.Caption = "부사장"
    
    Set optBtn = oForm01.Items("RdoBtn15").Specific
    optBtn.ValOn = "E"
    optBtn.ValOff = "0"
    Call optBtn.DataBind.SetBound(True, , "RadioBtn11")
    Call optBtn.GroupWith("RdoBtn11")
    '부사장(품의부서) 라디오버튼_E
    
    oForm01.DataSources.UserDataSources("RadioBtn11").VALUE = "0"
    '//////////품의부서 전결용 라디오버튼//////////
    
    '//회계전표 버튼
    Set newItem = oForm01.Items.Add("Btn01", it_BUTTON)
    newItem.Left = stdItem.Left + 330
    newItem.Top = stdItem.Top
    newItem.Height = stdItem.Height
    newItem.Width = 70
    newItem.Specific.Caption = "회계 전표"
    
    '//감가상각비 버튼
    Set newItem = oForm01.Items.Add("Btn02", it_BUTTON)
    newItem.Left = stdItem.Left + 425
    newItem.Top = stdItem.Top + 17
    newItem.Height = stdItem.Height
    newItem.Width = 70
    newItem.Specific.Caption = "감가상각비"
    
    '//부자재불출 버튼
    Set newItem = oForm01.Items.Add("Btn03", it_BUTTON)
    newItem.Left = stdItem.Left + 505
    newItem.Top = stdItem.Top + 17
    newItem.Height = stdItem.Height
    newItem.Width = 70
    newItem.Specific.Caption = "부자재불출"

    '//사업장-ComboBox
    Set newItem = oForm01.Items.Add("Static01", it_STATIC)
    Set stdItem = oForm01.Items("2006")
    newItem.Left = stdItem.Left + 93
    newItem.Top = stdItem.Top
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
'    newItem.LinkTo = "TradeType"
    Set oStatic = newItem.Specific
    oStatic.Caption = "사업장"
    
    Set newItem = oForm01.Items.Add("BPLId01", it_COMBO_BOX)
    Set stdItem = oForm01.Items("2007")
    newItem.Left = stdItem.Left + 93
    newItem.Top = stdItem.Top
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width + 40
    newItem.DisplayDesc = True
    Set oCombo01 = newItem.Specific
    Call oCombo01.DataBind.SetBound(True, "OJDT", "U_BPLId")

    sQry = "select BPLId, BPLName from [OBPL] order by BPLId"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo01.ValidValues.Add Trim(oRecordSet.Fields(0).VALUE), Trim(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop
    Call oCombo01.Select(User_BPLId(), psk_ByValue)
    
    '//사업장-ComboBox
    Set newItem = oForm01.Items.Add("Static02", it_STATIC)
    Set stdItem = oForm01.Items("2001")
    newItem.Left = stdItem.Left + 161
    newItem.Top = stdItem.Top
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 2
    newItem.ToPane = 2
'    newItem.LinkTo = "TradeType"
    Set oStatic = newItem.Specific
    oStatic.Caption = "사업장"
    
    Set newItem = oForm01.Items.Add("BPLId02", it_COMBO_BOX)
    Set stdItem = oForm01.Items("2000")
    newItem.Left = stdItem.Left + 161
    newItem.Top = stdItem.Top
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 2
    newItem.ToPane = 2
    newItem.DisplayDesc = True
    Set oCombo02 = newItem.Specific
    Call oCombo02.DataBind.SetBound(True, "JDT1", "U_BPLId")
    
'    sQry = "select BPLId, BPLName from [OBPL] order by BPLId"
'    oRecordset.DoQuery sQry
'    Do Until oRecordset.EOF
'        oCombo02.ValidValues.Add Trim(oRecordset.Fields(0).Value), Trim(oRecordset.Fields(1).Value)
'        oRecordset.MoveNext
'    Loop

    '//////////거래처//////////
    Set newItem = oForm01.Items.Add("Static03", it_STATIC)
    Set stdItem = oForm01.Items("Static01")
    newItem.Left = stdItem.Left + 10
    newItem.Top = stdItem.Top
    newItem.Height = 15
    newItem.Width = 90
    oForm01.Items("Static03").Specific.Caption = "거래처"
    '거래처코드
    Set newItem = oForm01.Items.Add("VatBP", it_EDIT)
    newItem.Left = oForm01.Items("Static01").Left + 90
    newItem.Top = oForm01.Items("Static01").Top
    newItem.Height = 15
    newItem.Width = 80
    '거래처명
    Set newItem = oForm01.Items.Add("VatBPName", it_EDIT)
    newItem.Left = oForm01.Items("VatBP").Left + 80
    newItem.Top = oForm01.Items("VatBP").Top
    newItem.Height = 15
    newItem.Width = 130
    oForm01.Items("VatBPName").Enabled = False
    '사업자번호
    Set newItem = oForm01.Items.Add("VatRegNo", it_EDIT)
    newItem.Left = oForm01.Items("VatBPName").Left + 130
    newItem.Top = oForm01.Items("VatBPName").Top
    newItem.Height = 15
    newItem.Width = 100
    oForm01.Items("VatRegNo").Enabled = False
    
    oForm01.Items("Static03").LinkTo = "VatBP"
    '//////////거래처//////////
    
    '//////////법정지출증빙코드//////////
    Set newItem = oForm01.Items.Add("Static04", it_STATIC)
    Set stdItem = oForm01.Items("BPLId01")
    newItem.Left = stdItem.Left + 10
    newItem.Top = stdItem.Top + 1
    newItem.Height = 15
    newItem.Width = 90
    oForm01.Items("Static04").Specific.Caption = "법정지출증빙"
    '코드
    Set newItem = oForm01.Items.Add("BillCode", it_EDIT)
    newItem.Left = oForm01.Items("BPLId01").Left + 90
    newItem.Top = oForm01.Items("Static04").Top
    newItem.Height = 15
    newItem.Width = 80
    '명
    Set newItem = oForm01.Items.Add("BillName", it_EDIT)
    newItem.Left = oForm01.Items("BillCode").Left + 80
    newItem.Top = oForm01.Items("BillCode").Top
    newItem.Height = 15
    newItem.Width = 130
    oForm01.Items("BillName").Enabled = False
    '비고
    Set newItem = oForm01.Items.Add("BillCMT", it_EDIT)
    newItem.Left = oForm01.Items("BillName").Left + 130
    newItem.Top = oForm01.Items("BillName").Top
    newItem.Height = 15
    newItem.Width = 100
    
    oForm01.Items("Static04").LinkTo = "BillCode"
    '//////////법정지출증빙코드//////////
    
    '적용버튼
    Set newItem = oForm01.Items.Add("BtnApply", it_BUTTON)
    Set stdItem = oForm01.Items("Static04")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top + 20
    newItem.Height = 20
    newItem.Width = 60
    newItem.Specific.Caption = "전체적용"
    
    oForm01.Freeze False
    
    Set oCombo01 = Nothing
    Set oCombo02 = Nothing
    Set oRecordSet = Nothing
    
    Exit Sub
    
CreateItems_Error:
    Set oCombo01 = Nothing
    Set oCombo02 = Nothing
    Set oRecordSet = Nothing
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error

    Dim i       As Integer
    Dim j       As Integer
    Dim DocNum  As String
    Dim LineId  As String
    Dim Counts  As String
    Dim Cnt     As String
    Dim ErrNum  As Integer
    Dim sQry    As String
    Dim oCombo  As SAPbouiCOM.ComboBox
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "BPLId02"
            oForm01.Freeze True
            For i = 1 To oMat01.VisualRowCount
                If oMat01Row = i Then
                    If oMat01.Columns("1").Cells(i).Specific.VALUE <> "" Then
                        oMat01.Columns("U_BPLId").Cells(i).Specific.Select Trim(oForm01.Items("BPLId02").Specific.Selected.VALUE)
                    End If
                End If
            Next i
            oForm01.Freeze False
            
        Case "BPLId01"
            oForm01.Freeze True
            For i = 1 To oMat01.VisualRowCount
                If oMat01.Columns("1").Cells(i).Specific.VALUE <> "" Then
                    oMat01.Columns("U_BPLId").Cells(i).Specific.Select Trim(oForm01.Items("BPLId01").Specific.Selected.VALUE)
                End If
            Next i
            oForm01.Freeze False
        
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "76" Then
        Select Case oCol
            Case "U_BPLId"
                oForm01.Freeze True
                oForm01.Items("BPLId02").Specific.Select oMat01.Columns("U_BPLId").Cells(oRow).Specific.VALUE
                oForm01.Freeze False
        End Select
    End If
     
    Set oRecordSet = Nothing
    
    Exit Sub
    
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub Form_Resize(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Form_Resize_Error
    
    oForm01.Items("Static01").Left = oForm01.Items("2006").Left + 93
    oForm01.Items("BPLId01").Left = oForm01.Items("2007").Left + 93
    oForm01.Items("Static02").Top = oForm01.Items("2001").Top
    oForm01.Items("Static02").Left = oForm01.Items("2001").Left + 161
    oForm01.Items("BPLId02").Top = oForm01.Items("2000").Top
    oForm01.Items("BPLId02").Left = oForm01.Items("2000").Left + 161

    Exit Sub

Form_Resize_Error:
    Sbo_Application.SetStatusBarMessage "Form_Resize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Print_Report01()
On Error GoTo Print_Report01_Error

    Dim i           As Integer
    Dim ErrNum      As Integer
    Dim TransId     As String
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim sQry        As String
    Dim oText(1)    As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    TransId = Trim(oForm01.Items("5").Specific.VALUE)
    
    If TransId = "" Then
        ErrNum = 1
        GoTo Print_Report01_Error
    End If
    
    WinTitle = "회계전표 [PS_FI010]"
    ReportName = "PS_FI010_01.rpt"
       
    '//Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
    gRpt_Formula(1) = "RadioBtn01"
    gRpt_Formula_Value(1) = oForm01.DataSources.UserDataSources("RadioBtn01").VALUE
    
    gRpt_Formula(2) = "RadioBtn11"
    gRpt_Formula_Value(2) = oForm01.DataSources.UserDataSources("RadioBtn11").VALUE
    
    '//SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '//조회조건문
    sQry = "EXEC [PS_FI010_02] '" & TransId & "'"
'    oRecordSet01.DoQuery sQry
'    If oRecordSet01.RecordCount = 0 Then
'        ErrNum = 2
'        GoTo Print_Report01_Error
'    End If
    
    '//CR Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    
    Exit Sub

Print_Report01_Error:
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub UpdateOCRData()
'******************************************************************************
'Function ID : UpdateOCRData()
'해당모듈    : S392
'기능        : 회계전표 클릭 시 배부규칙명 업데이트(2017.05.23 송명규)
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo UpdateOCRData_Error

    Dim QueryString As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim TransId As String
    
    TransId = Trim(oForm01.Items("5").Specific.VALUE)
    
    QueryString = "             EXEC [PS_S392_01] '"
    QueryString = QueryString & TransId & "'" '문서번호
        
    Call RecordSet01.DoQuery(QueryString)

    Set RecordSet01 = Nothing
    
    Exit Sub
UpdateOCRData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "UpdateOCRData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
