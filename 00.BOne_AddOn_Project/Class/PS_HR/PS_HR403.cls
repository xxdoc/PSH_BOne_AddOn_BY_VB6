VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_HR403.cls
'//  Module         : HR
'//  Description    : 고가평정 평가서약
'//  FormType       : PS_HR403
'//  Create Date    : 2012.12.02
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_HR403H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_HR403L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private ItemInformationCount As Long
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oBPLId As String, Optional ByVal oYear As String, Optional ByVal oNumber As String, Optional ByVal oMSTCOD As String, Optional ByVal oFULLNAME As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR403.srf")
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
    
    oFormUniqueID01 = "PS_HR403_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    'oForm01.DataBrowser.BrowseBy = "Code" '//UDO방식일때

    oForm01.EnableMenu ("1293"), False     '// 행삭제
    oForm01.EnableMenu ("1287"), False     '// 복제
    oForm01.EnableMenu ("1284"), False     '// 취소
    oForm01.Freeze True
    
    Call CreateItems
    Call ComboBox_Setting
    
    oForm01.Items("Year").Specific.VALUE = oYear
    oForm01.Items("MSTCOD").Specific.VALUE = oMSTCOD
    oForm01.Items("FULLNAME").Specific.VALUE = oFULLNAME
    
    If oBPLId <> "" Then
        Set oCombo = oForm01.Items("BPLId").Specific
        oCombo.Select oBPLId, psk_ByValue
    End If
    
    If oNumber <> "" Then
        Set oCombo = oForm01.Items("Number").Specific
        oCombo.Select oNumber, psk_ByValue
    End If
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
    Dim Cnt             As Integer
    Dim BPLID           As String
    Dim MSTCOD          As String
    Dim FULLNAME        As String
    Dim Year            As String
    Dim Number          As String
    Dim Check1          As String
    Dim Check2          As String
    Dim Check3          As String
    Dim Check4          As String
    Dim Check5          As String
    Dim Check6          As String
    Dim Check7          As String
    Dim Check8          As String
    Dim Check9          As String
    Dim Check10          As String
    Dim Check11          As String
    Dim vReturnValue$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                       
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        
                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "MSTCOD" Then
                        If oForm01.Items("MSTCOD").Specific.VALUE = "" Then
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
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
                    Year = Trim(oForm01.Items("Year").Specific.VALUE)
                    Number = Trim(oForm01.Items("Number").Specific.VALUE)
                    MSTCOD = oForm01.Items("MSTCOD").Specific.VALUE
                    FULLNAME = oForm01.Items("FULLNAME").Specific.VALUE
                    If oForm01.Items("Check1").Specific.Checked = True Then
                        Check1 = "Y"
                    Else
                        Check1 = "N"
                    End If
                    If oForm01.Items("Check2").Specific.Checked = True Then
                        Check2 = "Y"
                    Else
                        Check2 = "N"
                    End If
                    If oForm01.Items("Check3").Specific.Checked = True Then
                        Check3 = "Y"
                    Else
                        Check3 = "N"
                    End If
                    If oForm01.Items("Check4").Specific.Checked = True Then
                        Check4 = "Y"
                    Else
                        Check4 = "N"
                    End If
                    If oForm01.Items("Check5").Specific.Checked = True Then
                        Check5 = "Y"
                    Else
                        Check5 = "N"
                    End If
                    If oForm01.Items("Check6").Specific.Checked = True Then
                        Check6 = "Y"
                    Else
                        Check6 = "N"
                    End If
                    If oForm01.Items("Check7").Specific.Checked = True Then
                        Check7 = "Y"
                    Else
                        Check7 = "N"
                    End If
                    If oForm01.Items("Check8").Specific.Checked = True Then
                        Check8 = "Y"
                    Else
                        Check8 = "N"
                    End If
                    If oForm01.Items("Check9").Specific.Checked = True Then
                        Check9 = "Y"
                    Else
                        Check9 = "N"
                    End If
                    If oForm01.Items("Check10").Specific.Checked = True Then
                        Check10 = "Y"
                    Else
                        Check10 = "N"
                    End If
                    If oForm01.Items("Check11").Specific.Checked = True Then
                        Check11 = "Y"
                    Else
                        Check11 = "N"
                    End If
                    
                    If Check1 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약1을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    
                    If Check2 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약2을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    
                    If Check3 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약3을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    
                    If Check4 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약4을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    
                    If Check5 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약5을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    If Check6 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약6을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    If Check7 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약7을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    If Check8 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약8을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    If Check9 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약9을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    If Check10 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약10을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    If Check11 = "N" Then
                        Sbo_Application.SetStatusBarMessage "공정서약11을 체크하세요.", bmt_Short, True
                        Exit Sub
                    End If
                    
                    
                    
                    sQry = " Select Count(*) From Z_PS_HR403 Where BPLId = '" & BPLID & "' and Year = '" & Year & "' And Number = '" & Number & "' And MSTCOD = '" & MSTCOD & "'"
                    
                    oRecordSet01.DoQuery sQry
                    
                    Cnt = oRecordSet01.Fields(0).VALUE
                    
                    If Cnt > 0 Then '
                        '//있으면 Update
                        Sbo_Application.SetStatusBarMessage "이미 서약하였습니다.", bmt_Short, True
                    
                    Else
                        '//사원이 없을경우 insert
                        vReturnValue = Sbo_Application.MessageBox("평가서약을 하시겠습니까? ", 1, "&확인", "&취소")
                        If vReturnValue = 1 Then
                            sQry = " insert into Z_PS_HR403 Values ('" & BPLID & "','" & Year & "','" & Number & "','" & MSTCOD & "','" & FULLNAME & "','" & Check1 & "','" & Check2 & "','" & Check3 & "','" & Check4 & "','" & Check5 & "','"
                            sQry = sQry + Check6 & "','" & Check7 & "','" & Check8 & "','" & Check9 & "','" & Check10 & "','" & Check11 & "')"
                            oRecordSet01.DoQuery sQry
                            Sbo_Application.SetStatusBarMessage "평가서약을 완료하였습니다.", bmt_Short, True
                            oForm01.Items("2").Click ct_Regular
                        Else
                            Sbo_Application.SetStatusBarMessage "평가서약을 취소 하였습니다.", bmt_Short, True
                        End If
                    End If
'                    If oForm01.Mode = fm_ADD_MODE Then
''                        oForm01.Mode = fm_OK_MODE
'                        FormItemEnabled
''                        Call Sbo_Application.ActivateMenuItem("1282")
'                    ElseIf oForm01.Mode = fm_OK_MODE Then
'                        FormItemEnabled
'                        Call AddMatrixRow(0, True)
'                    End If
                End If
            
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then        '이력일자가 바뀌면 한 줄 더 생기도록
                   If pval.ItemUID = "MSTCOD" Then
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
    Dim BPLID           As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim sSeq$

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
        
' Matrix 필드에 질의 응답 창 띄워주기
   Select Case oUID
        Case "MSTCOD"
            BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
            sQry = "Select a.U_FULLNAME From OHEM a  "
            sQry = sQry & " Where a.U_MSTCOD = '" & Trim(oForm01.Items("MSTCOD").Specific.VALUE) & "' And a.branch = '" & oForm01.Items("BPLId").Specific.VALUE & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("FULLNAME").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
'            oMat01.LoadFromDataSource
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
            Case "1287": '복제
   
            Case "1282": '추가
   
            Case "1288", "1289", "1290", "1291": '레코드이동버튼             '추가버튼 클릭시 Matrix에 행 추가
   
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
       
'    Set oDS_PS_HR403H = oForm01.DataSources.DBDataSources("@PS_HR403H")
'    Set oDS_PS_HR403L = oForm01.DataSources.DBDataSources("@PS_HR403L")

       
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 1)
    oForm01.Items("BPLId").Specific.DataBind.SetBound True, "", "BPLId"
    
    Call oForm01.DataSources.UserDataSources.Add("Year", dt_SHORT_TEXT, 10)
    oForm01.Items("Year").Specific.DataBind.SetBound True, "", "Year"
    
    Call oForm01.DataSources.UserDataSources.Add("Number", dt_SHORT_TEXT, 10)
    oForm01.Items("Number").Specific.DataBind.SetBound True, "", "Number"
    
    Call oForm01.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    oForm01.Items("MSTCOD").Specific.DataBind.SetBound True, "", "MSTCOD"
    
    Call oForm01.DataSources.UserDataSources.Add("Check1", dt_SHORT_TEXT, 10)
    oForm01.Items("Check1").Specific.DataBind.SetBound True, "", "Check1"
    Call oForm01.DataSources.UserDataSources.Add("Check2", dt_SHORT_TEXT, 10)
    oForm01.Items("Check2").Specific.DataBind.SetBound True, "", "Check2"
    Call oForm01.DataSources.UserDataSources.Add("Check3", dt_SHORT_TEXT, 10)
    oForm01.Items("Check3").Specific.DataBind.SetBound True, "", "Check3"
    Call oForm01.DataSources.UserDataSources.Add("Check4", dt_SHORT_TEXT, 10)
    oForm01.Items("Check4").Specific.DataBind.SetBound True, "", "Check4"
    Call oForm01.DataSources.UserDataSources.Add("Check5", dt_SHORT_TEXT, 10)
    oForm01.Items("Check5").Specific.DataBind.SetBound True, "", "Check5"
    Call oForm01.DataSources.UserDataSources.Add("Check6", dt_SHORT_TEXT, 10)
    oForm01.Items("Check6").Specific.DataBind.SetBound True, "", "Check6"
    Call oForm01.DataSources.UserDataSources.Add("Check7", dt_SHORT_TEXT, 10)
    oForm01.Items("Check7").Specific.DataBind.SetBound True, "", "Check7"
    Call oForm01.DataSources.UserDataSources.Add("Check8", dt_SHORT_TEXT, 10)
    oForm01.Items("Check8").Specific.DataBind.SetBound True, "", "Check8"
    Call oForm01.DataSources.UserDataSources.Add("Check9", dt_SHORT_TEXT, 10)
    oForm01.Items("Check9").Specific.DataBind.SetBound True, "", "Check9"
    Call oForm01.DataSources.UserDataSources.Add("Check10", dt_SHORT_TEXT, 10)
    oForm01.Items("Check10").Specific.DataBind.SetBound True, "", "Check10"
    Call oForm01.DataSources.UserDataSources.Add("Check11", dt_SHORT_TEXT, 10)
    oForm01.Items("Check11").Specific.DataBind.SetBound True, "", "Check11"
    
    
'    oMat01.SelectionMode = ms_NotSupported
       
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
    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Set oCombo = oForm01.Items("Number").Specific
    oCombo.ValidValues.Add 1, "1차"
    oCombo.ValidValues.Add 2, "2차"
    
    
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
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        oForm01.Items("Code").Enabled = False
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        oForm01.Items("Code").Enabled = False
        Call oForm01.EnableMenu("1282", False) '//추가
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("Code").Enabled = False
        Call oForm01.EnableMenu("1282", False) '//추가
      
        
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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


