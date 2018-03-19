VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY402.cls
'//  Module         : 인사관리 > 연말정산관리
'//  Desc           : 정산기초등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oMat             As SAPbouiCOM.Matrix
Public oDS_PH_PY402A    As SAPbouiCOM.DataTable
Private oDS_PH_PY402L   As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY402.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY402_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY402"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.PaneLevel = 1
    oForm.Freeze True
    Call PH_PY402_CreateItems
    Call PH_PY402_FormItemEnabled
    Call PH_PY402_EnableMenus
'    Call PH_PY402_SetDocument(oFromDocEntry01)
'    Call PH_PY402_FormResize
    
    oForm.Update
    oForm.Freeze False
    
    oForm.Visible = True
    Set oXmlDoc = Nothing
    Exit Sub
    
LoadForm_Error:
    oForm.Update
    oForm.Freeze False
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY402_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    Dim CLTCOD      As String
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY402_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY402L = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    Set oGrid1 = oForm.Items("Grid01").Specific
    
    Set oMat = oForm.Items("Mat01").Specific
    oMat.SelectionMode = ms_NotSupported
    
    oForm.DataSources.DataTables.Add ("PH_PY402")
    
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY402")
    Set oDS_PH_PY402A = oForm.DataSources.DataTables.Item("PH_PY402")
    

    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//관계
    
    Set oCombo = oForm.Items("relate").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P121' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    
    
    '장애인코드
    Set oCombo = oForm.Items("hdcode").Specific
    oCombo.ValidValues.Add "", "선택"
    oCombo.ValidValues.Add "1", "장애인복지법에 따른 장애인"
    oCombo.ValidValues.Add "2", "국가유공자등 예우및지원에 관한 법률에 따른 상이자 및 이와 유사한자로서 근로능력이없는자"
    oCombo.ValidValues.Add "3", "그 밖에 항시 치료를 요하는 중증환자"
    oCombo.Select 0, psk_Index
       
'    oCombo.Select 0, psk_Index
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//년도
    Call oForm.DataSources.UserDataSources.Add("Year", dt_SHORT_TEXT, 4)
    oForm.Items("Year").Specific.DataBind.SetBound True, "", "Year"
    
    '//사번
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    oForm.Items("MSTCOD").Specific.DataBind.SetBound True, "", "MSTCOD"
    '//성명
    Call oForm.DataSources.UserDataSources.Add("FullName", dt_SHORT_TEXT, 20)
    oForm.Items("FullName").Specific.DataBind.SetBound True, "", "FullName"
    
    '//공제구분
    Call oForm.DataSources.UserDataSources.Add("div", dt_SHORT_TEXT, 20)
    oForm.Items("div").Specific.DataBind.SetBound True, "", "div"
    '//공제구분명
    Call oForm.DataSources.UserDataSources.Add("divnm", dt_SHORT_TEXT, 50)
    oForm.Items("divnm").Specific.DataBind.SetBound True, "", "divnm"
    '//공제대상
    Call oForm.DataSources.UserDataSources.Add("target", dt_SHORT_TEXT, 20)
    oForm.Items("target").Specific.DataBind.SetBound True, "", "target"
    '//공제대상명
    Call oForm.DataSources.UserDataSources.Add("targetnm", dt_SHORT_TEXT, 100)
    oForm.Items("targetnm").Specific.DataBind.SetBound True, "", "targetnm"
    
'    '//관계
'    Call oForm.DataSources.UserDataSources.Add("relate", dt_SHORT_TEXT, 20)
'    oForm.Items("relate").Specific.DataBind.SetBound True, "", "relate"
    
    '//성명
    Call oForm.DataSources.UserDataSources.Add("kname", dt_SHORT_TEXT, 20)
    oForm.Items("kname").Specific.DataBind.SetBound True, "", "kname"
    '//주민번호
    Call oForm.DataSources.UserDataSources.Add("juminno", dt_SHORT_TEXT, 20)
    oForm.Items("juminno").Specific.DataBind.SetBound True, "", "juminno"
    '//주소
    Call oForm.DataSources.UserDataSources.Add("addr", dt_SHORT_TEXT, 50)
    oForm.Items("addr").Specific.DataBind.SetBound True, "", "addr"
    '//생년월일
    Call oForm.DataSources.UserDataSources.Add("birthymd", dt_SHORT_TEXT, 50)
    oForm.Items("birthymd").Specific.DataBind.SetBound True, "", "birthymd"
    
    '//공제금액(국세청)
    Call oForm.DataSources.UserDataSources.Add("ntsamt", dt_SUM)
    oForm.Items("ntsamt").Specific.DataBind.SetBound True, "", "ntsamt"
    
    '//공제금액(국세청외)
    Call oForm.DataSources.UserDataSources.Add("amt", dt_SUM)
    oForm.Items("amt").Specific.DataBind.SetBound True, "", "amt"
    
    '//한도금액
    Call oForm.DataSources.UserDataSources.Add("handoamt", dt_SUM)
    oForm.Items("handoamt").Specific.DataBind.SetBound True, "", "handoamt"
    
    
    '//공제금액(국세청) 상반기(신용카드공제 입력항목)
    Call oForm.DataSources.UserDataSources.Add("ntsamt24", dt_SUM)
    oForm.Items("ntsamt24").Specific.DataBind.SetBound True, "", "ntsamt24"
    
'    '//공제금액(국세청) 하반기(신용카드공제 입력항목)
'    Call oForm.DataSources.UserDataSources.Add("ntsamt44", dt_SUM)
'    oForm.Items("ntsamt44").Specific.DataBind.SetBound True, "", "ntsamt44"
    
    '//2014년 카드총사용금액
    Call oForm.DataSources.UserDataSources.Add("bcard_t", dt_SUM)
    oForm.Items("bcard_t").Specific.DataBind.SetBound True, "", "bcard_t"
    
'    '//2014년 신용카드외 사용금액
'    Call oForm.DataSources.UserDataSources.Add("bcard44", dt_SUM)
'    oForm.Items("bcard44").Specific.DataBind.SetBound True, "", "bcard44"
    
    '//2013년 카드총사용금액
    Call oForm.DataSources.UserDataSources.Add("bbcard_t", dt_SUM)
    oForm.Items("bbcard_t").Specific.DataBind.SetBound True, "", "bbcard_t"
    
    '//2013년 신용카드외 사용금액
    Call oForm.DataSources.UserDataSources.Add("bbcard44", dt_SUM)
    oForm.Items("bbcard44").Specific.DataBind.SetBound True, "", "bbcard44"
    
    '//전통시장사용분 상반기
    Call oForm.DataSources.UserDataSources.Add("mart24", dt_SUM)
    oForm.Items("mart24").Specific.DataBind.SetBound True, "", "mart24"
'    '//전통시장사용분 하반기
'    Call oForm.DataSources.UserDataSources.Add("mart44", dt_SUM)
'    oForm.Items("mart44").Specific.DataBind.SetBound True, "", "mart44"
    
    '//대중교통사용분 상반기
    Call oForm.DataSources.UserDataSources.Add("trans24", dt_SUM)
    oForm.Items("trans24").Specific.DataBind.SetBound True, "", "trans24"
'    '//대중교통사용분 하반기
'    Call oForm.DataSources.UserDataSources.Add("trans44", dt_SUM)
'    oForm.Items("trans44").Specific.DataBind.SetBound True, "", "trans44"

    '//추가공제율 사용분(상반기)  2016년
    Call oForm.DataSources.UserDataSources.Add("adgong24", dt_SUM)
    oForm.Items("adgong24").Specific.DataBind.SetBound True, "", "adgong24"
    
    oForm.Update
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY402_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY402_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY402_EnableMenus()
    
    On Error GoTo PH_PY402_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY402_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY402_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY402_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY402_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY402_FormItemEnabled
'        Call PH_PY402_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY402_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY402_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY402_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY402_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    
    Dim CLTCOD As String
    Dim sPosDate As String
    
    On Error GoTo PH_PY402_FormItemEnabled_Error
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
    
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)    '//문서추가
        
        If Trim(oForm.Items("Year").Specific.VALUE) = "" Then
            oForm.Items("Year").Specific.VALUE = Format(Now, "YYYY") - 1
        End If
        
        If Trim(oForm.Items("MSTCOD").Specific.VALUE) = "" Then
            oForm.Items("MSTCOD").Specific.VALUE = ""
            oForm.Items("FullName").Specific.VALUE = ""
            oForm.Items("TeamName").Specific.VALUE = ""
            oForm.Items("RspName").Specific.VALUE = ""
            oForm.Items("ClsName").Specific.VALUE = ""
        End If
        
        oForm.DataSources.UserDataSources.Item("div").VALUE = ""
        oForm.DataSources.UserDataSources.Item("divnm").VALUE = ""
        oForm.DataSources.UserDataSources.Item("target").VALUE = ""
        oForm.DataSources.UserDataSources.Item("targetnm").VALUE = ""
        Set oCombo = oForm.Items("relate").Specific
        oCombo.Select "0", psk_Index
        
        Set oCombo = oForm.Items("hdcode").Specific
        oCombo.Select "0", psk_Index
        
        oForm.DataSources.UserDataSources.Item("kname").VALUE = ""
        oForm.DataSources.UserDataSources.Item("juminno").VALUE = ""
        oForm.DataSources.UserDataSources.Item("birthymd").VALUE = ""
        oForm.DataSources.UserDataSources.Item("addr").VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("ntsamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("amt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("handoamt").VALUE = 0
        
        oForm.Items("ntsamt").Enabled = True
        oForm.Items("ntsamt24").Enabled = False
       ' oForm.Items("ntsamt44").Enabled = False
        
        oForm.Items("bcard_t").Enabled = False
        'oForm.Items("bcard44").Enabled = False
        oForm.Items("bbcard_t").Enabled = False
        oForm.Items("bbcard44").Enabled = False
        
        oForm.Items("mart24").Enabled = False
        'oForm.Items("mart44").Enabled = False
        oForm.Items("trans24").Enabled = False
        'oForm.Items("trans44").Enabled = False
        oForm.Items("adgong24").Enabled = False
        
        oForm.DataSources.UserDataSources.Item("ntsamt24").VALUE = 0
        'oForm.DataSources.UserDataSources.Item("ntsamt44").VALUE = 0
        
        oForm.DataSources.UserDataSources.Item("mart24").VALUE = 0
        'oForm.DataSources.UserDataSources.Item("mart44").VALUE = 0
        oForm.DataSources.UserDataSources.Item("trans24").VALUE = 0
        'oForm.DataSources.UserDataSources.Item("trans44").VALUE = 0
        oForm.DataSources.UserDataSources.Item("adgong24").VALUE = 0
        
        
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")


    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    End If
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY402_FormItemEnabled_Error:
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY402_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim tSex        As String
    Dim tBrith      As String
    Dim Day         As String
    Dim ActCode     As String
    Dim CLTCOD      As String
    Dim MSTCOD      As String
    Dim FullName    As String
    Dim Div         As String
    Dim target      As String
    Dim relate      As String
    Dim YYYY        As String
    Dim Result      As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If PH_PY402_DataValidCheck = False Then
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "Btn_ret" Then
                    Call PH_PY402_MTX01
                End If
                    If pval.ItemUID = "Btn01" Then
                
                        CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
                        YYYY = oForm.Items("Year").Specific.VALUE
                        sQry = " Select CODE from [@PSH_PAYBACK] WHERE CODE = '" & YYYY & CLTCOD & "'"
                        oRecordSet.DoQuery sQry
                           
                        Result = oRecordSet.Fields(0).VALUE
                    If Result <> "" Then
                        Call MDC_Com.MDC_GF_Message("등록불가 년도입니다. 담당자에게 문의바랍니다.", "E")
                    End If
                    If Result = "" Then
                        Call PH_PY402_SAVE
                    End If
                    
                    
                End If
                
                                
                If pval.ItemUID = "Btn_del" Then
                    Call PH_PY402_Delete
                    Call PH_PY402_FormItemEnabled
                End If
'                If oForm.Mode = fm_FIND_MODE Then
'                    If pval.ItemUID = "Btn01" Then
'                        Sbo_Application.ActivateMenuItem ("7425")
'                        BubbleEvent = False
'                    End If
'
'                End If
            ElseIf pval.BeforeAction = False Then
                Select Case pval.ItemUID
                Case "1"
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY402_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY402_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY402_FormItemEnabled
                        End If
                    End If
'
                End Select
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True Then
                If pval.CharPressed = 9 Then
'                    If pval.ItemUID = "FullName" Then
'                        If oForm.Items("FullName").Specific.VALUE = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                                BubbleEvent = False
'                        End If
'                    End If
                    
                    If pval.ItemUID = "MSTCOD" Then
                        If oForm.Items("MSTCOD").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                        End If
                    End If
                    
                    If pval.ItemUID = "div" Then
                        If oForm.Items("div").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "target" Then
                        If oForm.Items("target").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                        End If
                    End If
                End If
            End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat01"
                If pval.Row > 0 Then
                    oLastItemUID = pval.ItemUID
                    oLastColUID = pval.ColUID
                    oLastColRow = pval.Row
                End If
            Case Else
                oLastItemUID = pval.ItemUID
                oLastColUID = ""
                oLastColRow = 0
            End Select

'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    '//사업장(헤더)
                    If pval.ItemUID = "relate" Then
                        oMat.Clear
                        oDS_PH_PY402L.Clear
                        
                        MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
                        relate = oForm.Items("relate").Specific.VALUE
                        
                        sQry = "EXEC [PH_PY402_03] '" + MSTCOD + "', '" + relate + "'"
                                                
                        oRecordSet.DoQuery sQry
                        
                        For i = 0 To oRecordSet.RecordCount - 1
                            If i + 1 > oDS_PH_PY402L.Size Then
                                oDS_PH_PY402L.InsertRecord (i)
                            End If
        
                            oMat.AddRow
                            oDS_PH_PY402L.Offset = i
                            
                            oDS_PH_PY402L.setValue "U_LineNum", i, i + 1
                            oDS_PH_PY402L.setValue "U_ColReg01", i, Trim(oRecordSet.Fields("kname").VALUE)
                            oDS_PH_PY402L.setValue "U_ColReg02", i, Trim(oRecordSet.Fields("juminno").VALUE)
                            oDS_PH_PY402L.setValue "U_ColReg03", i, Trim(oRecordSet.Fields("birthymd").VALUE)
                            oDS_PH_PY402L.setValue "U_ColReg04", i, Trim(oRecordSet.Fields("relatenm").VALUE)
                            oDS_PH_PY402L.setValue "U_ColReg05", i, Trim(oRecordSet.Fields("addr").VALUE)
                            oRecordSet.MoveNext
                        Next
                            
                        oMat.LoadFromDataSource
                        oMat.AutoResizeColumns
                        
                        If (oRecordSet.RecordCount = 0) Then
                            oForm.DataSources.UserDataSources.Item("kname").VALUE = ""
                            oForm.DataSources.UserDataSources.Item("juminno").VALUE = ""
                            oForm.DataSources.UserDataSources.Item("birthymd").VALUE = ""
                            oForm.DataSources.UserDataSources.Item("addr").VALUE = ""
                            
'                            oForm.DataSources.UserDataSources.Item("ntsamt").VALUE = 0
'                            oForm.DataSources.UserDataSources.Item("amt").VALUE = 0
'                            oForm.DataSources.UserDataSources.Item("handoamt").VALUE = 0
                        End If
                        
                        If (oRecordSet.RecordCount = 1) Then
                            oForm.Items("kname").Specific.VALUE = oMat.Columns("kname").Cells(1).Specific.VALUE
                            oForm.Items("juminno").Specific.VALUE = oMat.Columns("juminno").Cells(1).Specific.VALUE
                            oForm.Items("birthymd").Specific.VALUE = oMat.Columns("birthymd").Cells(1).Specific.VALUE
                            oForm.Items("addr").Specific.VALUE = oMat.Columns("addr").Cells(1).Specific.VALUE
                        End If
                            
'                        If relate = "01" Then
'                            If oForm.Items("div").Specific.VALUE = "50" And oForm.Items("target").Specific.VALUE = "520" Then
'                                'oForm.Items("bcard_t").Enabled = True '2015년 기준 2014년 총신용카드 사용금액
'                                'oForm.Items("bcard44").Enabled = True '2015년 기준 2014년 신용카드사용분 제외 금액
'                                'oForm.Items("bbcard_t").Enabled = True '2015년 기준 2013년 총신용카드 사용금액
'                                'oForm.Items("bbcard44").Enabled = True '2015년 기준 2013년 신용카드사용분 제외 금액
'
'                                oForm.Items("bcard_t").Enabled = True '2016년 기준 2015년 총신용카드 사용금액
'                               ' oForm.Items("bcard44").Enabled = True '2016년 기준 0
'                                oForm.Items("bbcard_t").Enabled = True '2016년 기준 2014년 총신용카드 사용금액
'                                oForm.Items("bbcard44").Enabled = True '2016년 기준 2014년 신용카드사용분 제외 금액
'
'                                CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
'
'                                sQry = " Select bcard_t = Isnull(Sum(Case When yyyy = '2015' Then Case When target in ('520','540','550','572','574') Then Isnull(amt,0) + Isnull(ntsamt,0) + isnull(mart24,0) + Isnull(trans24,0) + Isnull(mart44,0) + isnull(trans44,0) Else 0 End End), 0),"
'                                sQry = sQry + " bcard44 = Isnull(Sum(Case When yyyy = '2015' Then Case When target in ('520') Then Isnull(mart24,0) + Isnull(trans24,0) + Isnull(mart44,0) + Isnull(trans44,0) Else 0 End End),0) + "
'                                sQry = sQry + " Isnull(Sum(Case When yyyy = '2015' Then Case When target in ('540','550','572','574') Then Isnull(amt,0) + Isnull(ntsamt,0) + Isnull(mart24,0) + Isnull(trans24,0) + Isnull(mart44,0) + Isnull(trans44,0) Else 0 End End),0),"
'                                sQry = sQry + " bbcard_t = Isnull(Sum(Case When yyyy = '2014' Then Case When target in ('520','540','550','572','574') Then Isnull(amt,0) + Isnull(ntsamt,0) + Isnull(mart24,0) + Isnull(trans24,0) + Isnull(mart44,0) + Isnull(trans44,0) Else 0 End End),0),"
'                                sQry = sQry + " bbcard44 = Isnull(Sum(Case When yyyy = '2014' Then Case When target in ('520') Then Isnull(mart24,0) + Isnull(trans24,0) + Isnull(mart44,0) + Isnull(trans44,0) Else 0 End End),0) +"
'                                sQry = sQry + " Isnull(Sum(Case When yyyy = '2014' Then Case When target in ('540','550','572','574') Then Isnull(amt,0) + Isnull(ntsamt,0) + Isnull(mart24,0) + Isnull(trans24,0) + Isnull(mart44,0) + Isnull(trans44,0) Else 0 End End),0)"
'
'                                sQry = sQry + " From p_seoybase "
'                                sQry = sQry + " Where saup = '" & CLTCOD & "'"
'                                sQry = sQry + " and yyyy In ('2014','2015') and sabun = '" & MSTCOD & "' and relate = '01'"
'                                sQry = sQry + " and div = '50' "
'
'
'                                oRecordSet.DoQuery sQry
'
'                                oForm.Items("bcard_t").Specific.VALUE = oRecordSet.Fields("bcard_t").VALUE
'                                'oForm.Items("bcard44").Specific.VALUE = oRecordSet.Fields("bcard44").VALUE
'                                'oForm.Items("bcard44").Specific.VALUE = 0  '2016년에는 없슴
'                                oForm.Items("bbcard_t").Specific.VALUE = oRecordSet.Fields("bbcard_t").VALUE
'                                oForm.Items("bbcard44").Specific.VALUE = oRecordSet.Fields("bbcard44").VALUE
'                            Else
'                                oForm.Items("bcard_t").Enabled = False
'                                'oForm.Items("bcard44").Enabled = False
'                                oForm.Items("bbcard_t").Enabled = False
'                                oForm.Items("bbcard44").Enabled = False
'
'                                oForm.Items("bcard_t").Specific.VALUE = 0
'                                'oForm.Items("bcard44").Specific.VALUE = 0
'                                oForm.Items("bbcard_t").Specific.VALUE = 0
'                                oForm.Items("bbcard44").Specific.VALUE = 0
'                            End If
'                        Else
'                            oForm.Items("bcard_t").Enabled = False
'                            'oForm.Items("bcard44").Enabled = False
'                            oForm.Items("bbcard_t").Enabled = False
'                            oForm.Items("bbcard44").Enabled = False
'
'                            oForm.Items("bcard_t").Specific.VALUE = 0
'                            'oForm.Items("bcard44").Specific.VALUE = 0
'                            oForm.Items("bbcard_t").Specific.VALUE = 0
'                            oForm.Items("bbcard44").Specific.VALUE = 0
'                        End If
                    End If
                    
                End If
            End If

            Call oForm.Freeze(False)
            Set oRecordSet = Nothing
'----------------------------------------------------------
        Case et_CLICK: '//6
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row >= 0 Then
                        Select Case pval.ItemUID
                        Case "Grid01"
                            
                            
                            'Call oMat.SelectRow(pval.Row, True, False)
                            PH_PY402_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                        End Select
                        
                    End If
                Case "Mat01"
                    If pval.Row >= 0 Then
                        Call oMat.SelectRow(pval.Row, True, False)
                    End If
                End Select
                
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row > 0 Then
                        oLastItemUID = pval.ItemUID
                        oLastColUID = pval.ColUID
                        oLastColRow = pval.Row
                    End If
                Case Else
                    oLastItemUID = pval.ItemUID
                    oLastColUID = ""
                    oLastColRow = 0
                End Select
            ElseIf pval.BeforeAction = False Then

            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
            Else
                If pval.ItemUID = "Mat01" Then
                    oForm.Items("kname").Specific.VALUE = oMat.Columns("kname").Cells(pval.Row).Specific.VALUE
                    oForm.Items("juminno").Specific.VALUE = oMat.Columns("juminno").Cells(pval.Row).Specific.VALUE
                    oForm.Items("birthymd").Specific.VALUE = oMat.Columns("birthymd").Cells(pval.Row).Specific.VALUE
                    oForm.Items("addr").Specific.VALUE = oMat.Columns("addr").Cells(pval.Row).Specific.VALUE
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
'            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                If pval.ItemChanged = True Then
                    
                End If
                    
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUID

                        Case "MSTCOD"
                            CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
                            MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
                            
                            sQry = "Select Code,"
                            sQry = sQry + " FullName = U_FullName,"
                            sQry = sQry + " TeamName = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '1'"
                            sQry = sQry + " And U_Code = U_TeamCode),''),"
                            sQry = sQry + " RspName  = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '2'"
                            sQry = sQry + " And U_Code = U_RspCode),''),"
                            sQry = sQry + " ClsName  = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '9'"
                            sQry = sQry + " And U_Code  = U_ClsCode"
                            sQry = sQry + " And U_Char3 = U_CLTCOD),'')"
                            sQry = sQry + " From [@PH_PY001A]"
                            sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'"
                            sQry = sQry + " and Code = '" + MSTCOD + "'"
                              
                            oRecordSet.DoQuery sQry
                            
                            oForm.DataSources.UserDataSources.Item("FullName").VALUE = oRecordSet.Fields("FullName").VALUE
                            
'                            oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
                            oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
                            oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
                            oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE

                        Case "FullName"
                            CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
                            FullName = oForm.Items("FullName").Specific.VALUE
                            
                            sQry = "Select Code,"
                            sQry = sQry + " FullName = U_FullName,"
                            sQry = sQry + " TeamName = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '1'"
                            sQry = sQry + " And U_Code = U_TeamCode),''),"
                            sQry = sQry + " RspName  = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '2'"
                            sQry = sQry + " And U_Code = U_RspCode),''),"
                            sQry = sQry + " ClsName  = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '9'"
                            sQry = sQry + " And U_Code  = U_ClsCode"
                            sQry = sQry + " And U_Char3 = U_CLTCOD),'')"
                            sQry = sQry + " From [@PH_PY001A]"
                            sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'"
                            sQry = sQry + " And U_status <> '5'"      '퇴사자 제외
                            sQry = sQry + " and U_FullName = '" + FullName + "'"
                              
                            oRecordSet.DoQuery sQry
                            
                            oForm.DataSources.UserDataSources.Item("MSTCOD").VALUE = oRecordSet.Fields("Code").VALUE
'                            oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("Code").VALUE
                            oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
                            oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
                            oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
                       Case "div"
                            Div = oForm.Items("div").Specific.VALUE
                            
                            sQry = "Select CodeNm = U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '70'"
                            sQry = sQry + " And U_Code = '" + Div + "'"
                              
                            oRecordSet.DoQuery sQry
                            
                            oForm.Items("divnm").Specific.VALUE = oRecordSet.Fields("CodeNm").VALUE
                          
                       Case "target"
                            target = oForm.Items("target").Specific.VALUE
                            
                            sQry = "Select CodeNm = U_CodeNm, handoamt = Isnull(U_Num1,0)"
                            
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '71'"
                            sQry = sQry + " And U_Code = '" + target + "'"
                              
                            oRecordSet.DoQuery sQry
                            
                            oForm.Items("targetnm").Specific.VALUE = oRecordSet.Fields("CodeNm").VALUE
                            oForm.Items("handoamt").Specific.VALUE = oRecordSet.Fields("handoamt").VALUE
                            
                            If target = "520" Or target = "540" Or target = "550" Or target = "572" Or target = "574" Then
                                oForm.Items("ntsamt24").Enabled = True
                                'oForm.Items("ntsamt44").Enabled = True
                                
                                oForm.Items("mart24").Enabled = True
                                'oForm.Items("mart44").Enabled = True
                                oForm.Items("trans24").Enabled = True
                                'oForm.Items("trans44").Enabled = True
                               ' oForm.Items("adgong24").Enabled = True
                                
                                
                                oForm.Items("ntsamt").Enabled = False
                                oForm.Items("ntsamt24").Specific.VALUE = 0
                                'oForm.Items("ntsamt44").Specific.VALUE = 0
                                
                                oForm.Items("mart24").Specific.VALUE = 0
                                'oForm.Items("mart44").Specific.VALUE = 0
                                oForm.Items("trans24").Specific.VALUE = 0
                                'oForm.Items("trans44").Specific.VALUE = 0
                                oForm.Items("adgong24").Specific.VALUE = 0
                            Else
                                oForm.Items("ntsamt24").Enabled = False
                                'oForm.Items("ntsamt44").Enabled = False
                                
                                oForm.Items("mart24").Enabled = False
                                'oForm.Items("mart44").Enabled = False
                                oForm.Items("trans24").Enabled = False
                                'oForm.Items("trans44").Enabled = False
                                oForm.Items("adgong24").Enabled = False
                                
                                oForm.Items("ntsamt").Enabled = True
                                
                                oForm.Items("ntsamt24").Specific.VALUE = 0
                                'oForm.Items("ntsamt44").Specific.VALUE = 0
                                
                                oForm.Items("mart24").Specific.VALUE = 0
                                'oForm.Items("mart44").Specific.VALUE = 0
                                oForm.Items("trans24").Specific.VALUE = 0
                                'oForm.Items("trans44").Specific.VALUE = 0
                                oForm.Items("adgong24").Specific.VALUE = 0
                            End If
                            
                            
                            Select Case target
                            Case "110"  '본인
                                Set oCombo = oForm.Items("relate").Specific
                                oCombo.Select "01", psk_ByValue
                                oForm.Items("amt").Specific.VALUE = oForm.Items("handoamt").Specific.VALUE
                            Case "120"  '배우자
                                Set oCombo = oForm.Items("relate").Specific
                                oCombo.Select "02", psk_ByValue
                                oForm.Items("amt").Specific.VALUE = oForm.Items("handoamt").Specific.VALUE
                            
                            Case "130"  '부양가족
                                oForm.Items("amt").Specific.VALUE = oForm.Items("handoamt").Specific.VALUE
                                
                            Case Else
                                
                                Set oCombo = oForm.Items("relate").Specific
                                oCombo.Select "", psk_ByValue
                                oMat.Clear
                                
                                oForm.DataSources.UserDataSources.Item("kname").VALUE = ""
                                oForm.DataSources.UserDataSources.Item("juminno").VALUE = ""
                                oForm.DataSources.UserDataSources.Item("birthymd").VALUE = ""
                                oForm.DataSources.UserDataSources.Item("addr").VALUE = ""
                                
                                If oForm.Items("div").Specific.VALUE = "20" Then
                                
                                    If oForm.DataSources.UserDataSources.Item("handoamt").VALUE > 0 Then
                                        oForm.DataSources.UserDataSources.Item("ntsamt").VALUE = 0
                                        oForm.DataSources.UserDataSources.Item("amt").VALUE = oForm.DataSources.UserDataSources.Item("handoamt").VALUE
                                    Else
                                        oForm.DataSources.UserDataSources.Item("ntsamt").VALUE = 0
                                        oForm.DataSources.UserDataSources.Item("amt").VALUE = 0
                                    End If
                                Else
                                    If oForm.DataSources.UserDataSources.Item("handoamt").VALUE > 0 Then
                                        oForm.DataSources.UserDataSources.Item("ntsamt").VALUE = oForm.DataSources.UserDataSources.Item("handoamt").VALUE
                                        oForm.DataSources.UserDataSources.Item("amt").VALUE = 0
                                    Else
                                        oForm.DataSources.UserDataSources.Item("ntsamt").VALUE = 0
                                        oForm.DataSources.UserDataSources.Item("amt").VALUE = 0
                                    End If
                                End If
                                
                            End Select
                        Case "ntsamt"
                        
                            If oForm.Items("handoamt").Specific.VALUE > 0 Then
                                If CDbl(oForm.Items("ntsamt").Specific.VALUE) + CDbl(oForm.Items("amt").Specific.VALUE) > CDbl(oForm.Items("handoamt").Specific.VALUE) Then
                                    oForm.Items("ntsamt").Specific.VALUE = 0
                                    Sbo_Application.MessageBox ("한도금액보다 초과됩니다. 확인하세요")
                                End If
                            End If
                            
                        Case "juminno" '주민번호
                            '주민번호입력시 생년월일 생성
                            If Len(Trim(oForm.Items("juminno").Specific.VALUE)) <> 13 Then
                                oForm.Items("birthymd").Specific.VALUE = ""
                                Sbo_Application.MessageBox ("주민번호자릿수가 틀립니다. 확인하세요.")
                            Else
                                If Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "1" Or Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "2" Then
                                    oForm.Items("birthymd").Specific.VALUE = "19" + Mid(oForm.Items("juminno").Specific.VALUE, 1, 6)
                                ElseIf Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "3" Or Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "4" Then
                                    oForm.Items("birthymd").Specific.VALUE = "20" + Mid(oForm.Items("juminno").Specific.VALUE, 1, 6)
                                ElseIf Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "5" Or Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "6" Then
                                    oForm.Items("birthymd").Specific.VALUE = "19" + Mid(oForm.Items("juminno").Specific.VALUE, 1, 6)
                                ElseIf Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "7" Or Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "8" Then
                                    oForm.Items("birthymd").Specific.VALUE = "20" + Mid(oForm.Items("juminno").Specific.VALUE, 1, 6)
                                End If
                            End If
                        Case "amt"
                            If oForm.Items("handoamt").Specific.VALUE > 0 Then
                                If CDbl(oForm.Items("ntsamt").Specific.VALUE) + CDbl(oForm.Items("amt").Specific.VALUE) > CDbl(oForm.Items("handoamt").Specific.VALUE) Then
                                    oForm.Items("amt").Specific.VALUE = 0
                                    Sbo_Application.MessageBox ("한도금액보다 초과됩니다. 확인하세요")
                                End If
                            End If
                        Case "ntsamt24"
                            'oForm.Items("ntsamt").Specific.VALUE = Val(oForm.Items("ntsamt24").Specific.VALUE) + Val(oForm.Items("ntsamt44").Specific.VALUE)
                            oForm.Items("ntsamt").Specific.VALUE = Val(oForm.Items("ntsamt24").Specific.VALUE)
                        'Case "ntsamt44"
                            'oForm.Items("ntsamt").Specific.VALUE = Val(oForm.Items("ntsamt24").Specific.VALUE) + Val(oForm.Items("ntsamt44").Specific.VALUE)
                    End Select
                 
                End If
            End If
'            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
'                oMat.LoadFromDataSource
'                Call PH_PY402_AddMatrixRow
                
            End If
'----------------------------------------------------------
        Case et_DATASOURCE_LOAD '//12
           
'----------------------------------------------------------
        Case et_FORM_LOAD: '//16
            
'----------------------------------------------------------
        Case et_FORM_UNLOAD: '//17
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_PH_PY402A = Nothing
                
'                Set oMat = Nothing
            End If
'----------------------------------------------------------
        Case et_FORM_ACTIVATE: '//18
            
'----------------------------------------------------------
        Case et_FORM_DEACTIVATE: '//19
            
'----------------------------------------------------------
        Case et_FORM_CLOSE '//20
            
'----------------------------------------------------------
        Case et_FORM_RESIZE '//21
            If pval.BeforeAction = True Then
        
            ElseIf pval.BeforeAction = False Then
                
            End If
'----------------------------------------------------------
        Case et_FORM_KEY_DOWN '//22
            
'----------------------------------------------------------
        Case et_FORM_MENU_HILIGHT '//23
            
'----------------------------------------------------------
        Case et_CHOOSE_FROM_LIST '//27
            If pval.BeforeAction = True Then
            
            ElseIf pval.Before_Action = False Then
'                If pval.ItemUID = "Code" Then
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY402A", "Code")
'                End If
            End If

'----------------------------------------------------------
        Case et_PICKER_CLICKED '//37
            
'----------------------------------------------------------
        Case et_GRID_SORT '//38
            
'----------------------------------------------------------
        Case et_Drag '//39
        
            
    End Select
    
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    oForm.Freeze (False)
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i As Long
    On Error GoTo Raise_FormMenuEvent_Error
    Call oForm.Freeze(True)
    
    If (pval.BeforeAction = True) Then
        Select Case pval.MenuUID
            Case "1283":
                If Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1284":
            Case "1286":
            Case "1293":
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
'                Call PH_PY402_FormItemEnabled
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY402_FormItemEnabled
'                Call PH_PY402_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY402_FormItemEnabled
'                Call PH_PY402_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY402_FormItemEnabled
'                Call PH_PY402_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY402_FormItemEnabled
            Case "1293" '// 행삭제
                
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    Dim i       As Long
    Dim sQry    As String
    Dim oCombo  As SAPbouiCOM.ComboBox
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    On Error GoTo Raise_FormDataEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select

    End If
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
    
Raise_FormDataEvent_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_RightClickEvent_Error
    
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
    End If
    Select Case pval.ItemUID
    Case "Mat01"
        If pval.Row > 0 Then
            oLastItemUID = pval.ItemUID
            oLastColUID = pval.ColUID
            oLastColRow = pval.Row
        End If
    Case Else
        oLastItemUID = pval.ItemUID
        oLastColUID = ""
        oLastColRow = 0
    End Select
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PH_PY402_FormClear()
On Error GoTo PH_PY402_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY402'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY402_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY402_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY402_DataValidCheck() As Boolean
On Error GoTo PH_PY402_DataValidCheck_Error
    PH_PY402_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY402_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY402_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY402_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY402_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim iRow        As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY402_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("CLTCOD").Specific.VALUE
    Param02 = oForm.Items("Year").Specific.VALUE
    Param03 = oForm.Items("MSTCOD").Specific.VALUE
        
    If Trim(Param01) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY402_MTX01_Exit
    End If
        
    If Trim(Param02) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY402_MTX01_Exit
    End If
        
    If Trim(Param03) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY402_MTX01_Exit
    End If
        
        
        
    sQry = "EXEC PH_PY402_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "'"

    Call oDS_PH_PY402A.ExecuteQuery(sQry)
    
    
    
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    Call PH_PY402_TitleSetting(iRow)
    
    oForm.Update
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY402_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY402_MTX01_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY402_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_PY402_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)


    '//그리드 자료를 head에 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim sRow As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY402_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    
    Param01 = oForm.Items("CLTCOD").Specific.VALUE
    Param02 = oDS_PH_PY402A.Columns.Item("Year").Cells(oRow).VALUE
    Param03 = oDS_PH_PY402A.Columns.Item("MSTCOD").Cells(oRow).VALUE
    Param04 = oDS_PH_PY402A.Columns.Item("div").Cells(oRow).VALUE
    Param05 = oDS_PH_PY402A.Columns.Item("target").Cells(oRow).VALUE
    Param06 = oDS_PH_PY402A.Columns.Item("relate").Cells(oRow).VALUE
    Param07 = oDS_PH_PY402A.Columns.Item("juminno").Cells(oRow).VALUE
    
    
    sQry = "EXEC PH_PY402_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        
        oForm.Items("MSTCOD").Specific.VALUE = oDS_PH_PY402A.Columns.Item("MSTCOD").Cells(oRow).VALUE
        oForm.Items("FullName").Specific.VALUE = oDS_PH_PY402A.Columns.Item("FullName").Cells(oRow).VALUE
        
        
        oForm.DataSources.UserDataSources.Item("div").VALUE = ""
        oForm.DataSources.UserDataSources.Item("divnm").VALUE = ""
        oForm.DataSources.UserDataSources.Item("target").VALUE = ""
        oForm.DataSources.UserDataSources.Item("targetnm").VALUE = ""
        
        Set oCombo = oForm.Items("relate").Specific
        oCombo.Select "0", psk_Index
        
        Set oCombo = oForm.Items("hdcode").Specific
        oCombo.Select "0", psk_Index
        
        oForm.DataSources.UserDataSources.Item("kname").VALUE = ""
        oForm.DataSources.UserDataSources.Item("juminno").VALUE = ""
        oForm.DataSources.UserDataSources.Item("birthymd").VALUE = ""
        oForm.DataSources.UserDataSources.Item("addr").VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("ntsamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("amt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("handoamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ntsamt24").VALUE = 0
        'oForm.DataSources.UserDataSources.Item("ntsamt44").VALUE = 0
        oForm.DataSources.UserDataSources.Item("bcard_t").VALUE = 0
        'oForm.DataSources.UserDataSources.Item("bcard44").VALUE = 0
        oForm.DataSources.UserDataSources.Item("bbcard_t").VALUE = 0
        oForm.DataSources.UserDataSources.Item("bbcard44").VALUE = 0
        
        oForm.Items("TeamName").Specific.VALUE = ""
        oForm.Items("RspName").Specific.VALUE = ""
        oForm.Items("ClsName").Specific.VALUE = ""
        
        oForm.Items("bcard_t").Enabled = False
        'oForm.Items("bcard44").Enabled = False
        oForm.Items("bbcard_t").Enabled = False
        oForm.Items("bbcard44").Enabled = False
        
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY402_MTX02_Exit
    End If
    
    
'    oForm.Items("PosDate").Specific.Value = Format(oRecordset.Fields("PosDate").Value, "YYYYMMDD")
        
'    oForm.DataSources.UserDataSources.Item("PosDate").Value = Format(oRecordSet.Fields("PosDate").Value, "YYYYMMDD")
    
    oForm.Items("Year").Specific.VALUE = oRecordSet.Fields("Year").VALUE
    oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("MSTCOD").VALUE
    oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
    
'    '//부서
    oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
    oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
    oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
    
        
    oForm.DataSources.UserDataSources.Item("div").VALUE = oRecordSet.Fields("div").VALUE
    oForm.DataSources.UserDataSources.Item("divnm").VALUE = oRecordSet.Fields("divnm").VALUE
    oForm.DataSources.UserDataSources.Item("target").VALUE = oRecordSet.Fields("target").VALUE
    oForm.DataSources.UserDataSources.Item("targetnm").VALUE = oRecordSet.Fields("targetnm").VALUE
    
    Set oCombo = oForm.Items("relate").Specific
    oCombo.Select oRecordSet.Fields("relate").VALUE, psk_ByValue
    
    Set oCombo = oForm.Items("hdcode").Specific
    oCombo.Select oRecordSet.Fields("hdcode").VALUE, psk_ByValue
    
    oForm.DataSources.UserDataSources.Item("kname").VALUE = oRecordSet.Fields("kname").VALUE
    oForm.DataSources.UserDataSources.Item("juminno").VALUE = oRecordSet.Fields("juminno").VALUE
    oForm.DataSources.UserDataSources.Item("birthymd").VALUE = oRecordSet.Fields("birthymd").VALUE
    oForm.DataSources.UserDataSources.Item("addr").VALUE = oRecordSet.Fields("addr").VALUE
    
    oForm.DataSources.UserDataSources.Item("ntsamt").VALUE = oRecordSet.Fields("ntsamt").VALUE
    oForm.DataSources.UserDataSources.Item("amt").VALUE = oRecordSet.Fields("amt").VALUE
    oForm.DataSources.UserDataSources.Item("handoamt").VALUE = oRecordSet.Fields("handoamt").VALUE
    oForm.DataSources.UserDataSources.Item("ntsamt24").VALUE = oRecordSet.Fields("ntsamt24").VALUE
    'oForm.DataSources.UserDataSources.Item("ntsamt44").VALUE = oRecordSet.Fields("ntsamt44").VALUE
    
    oForm.DataSources.UserDataSources.Item("bcard_t").VALUE = oRecordSet.Fields("bcard_t").VALUE
    'oForm.DataSources.UserDataSources.Item("bcard44").VALUE = oRecordSet.Fields("bcard44").VALUE
    oForm.DataSources.UserDataSources.Item("bbcard_t").VALUE = oRecordSet.Fields("bbcard_t").VALUE
    oForm.DataSources.UserDataSources.Item("bbcard44").VALUE = oRecordSet.Fields("bbcard44").VALUE
    
    oForm.DataSources.UserDataSources.Item("mart24").VALUE = oRecordSet.Fields("mart24").VALUE
    'oForm.DataSources.UserDataSources.Item("mart44").VALUE = oRecordSet.Fields("mart44").VALUE
    oForm.DataSources.UserDataSources.Item("trans24").VALUE = oRecordSet.Fields("trans24").VALUE
    'oForm.DataSources.UserDataSources.Item("trans44").VALUE = oRecordSet.Fields("trans44").VALUE
    oForm.DataSources.UserDataSources.Item("adgong24").VALUE = oRecordSet.Fields("adgong24").VALUE
    
'2018
'    If oForm.Items("div").Specific.VALUE = "50" And oForm.Items("target").Specific.VALUE = "520" And oForm.Items("relate").Specific.VALUE = "01" Then
'        oForm.Items("bcard_t").Enabled = True
'        'oForm.Items("bcard44").Enabled = True
'        oForm.Items("bbcard_t").Enabled = True
'        oForm.Items("bbcard44").Enabled = True
'    Else
'        oForm.Items("bcard_t").Enabled = False
'        'oForm.Items("bcard44").Enabled = False
'        oForm.Items("bbcard_t").Enabled = False
'        oForm.Items("bbcard44").Enabled = False
'    End If
    
    oForm.Update
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY402_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY402_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY402_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY402_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY402_Validate_Error
    PH_PY402_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY402A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY402_Validate = False
        GoTo PH_PY402_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY402_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY402_Validate_Error:
    PH_PY402_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY402_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'//행삭제 (FormUID, pval, BubbleEvent, 매트릭스 이름, 디비데이터소스, 데이터 체크 필드명)
Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean, _
                                                    oMat As SAPbouiCOM.Matrix, DBData As SAPbouiCOM.DBDataSource, CheckField As String)

    Dim i As Long
    
    On Error GoTo Raise_EVENT_ROW_DELETE_Error

    If (oLastColRow > 0) Then
        If pval.BeforeAction = True Then
        
        ElseIf pval.BeforeAction = False Then
            If oMat.RowCount <> oMat.VisualRowCount Then
                oMat.FlushToDataSource
    
                While (i <= DBData.Size - 1)
                    If DBData.GetValue(CheckField, i) = "" Then
                        DBData.RemoveRecord (i)
                        i = 0
                    Else
                        i = i + 1
                    End If
                Wend
                
                For i = 0 To DBData.Size
                    Call DBData.setValue("U_LineNum", i, i + 1)
                Next i
                
                oMat.LoadFromDataSource
            End If
        End If
    End If
Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY402_SAVE()

    '//데이타 저장
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim CLTCOD$, MSTCOD$, FullName$, YEAR$
    
    
        
    Dim Div$, target$, relate$, kname$, juminno$, addr$, birthymd$, hdcode$
    Dim CheckDate1$, CheckDate2$
    
    Dim Amt, ntsamt, ntsamt24, ntsamt44, bcard_t, bcard44, bbcard_t, bbcard44, adgong24, mart24, trans24, mart44, trans44 As Double
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim vReturnValue$
    
    On Error GoTo PH_PY402_SAVE_Error
    
    Div = oForm.Items("div").Specific.VALUE
    target = oForm.Items("target").Specific.VALUE
    relate = oForm.Items("relate").Specific.VALUE
    kname = oForm.Items("kname").Specific.VALUE
    juminno = oForm.Items("juminno").Specific.VALUE
    addr = oForm.Items("addr").Specific.VALUE
    birthymd = oForm.Items("birthymd").Specific.VALUE
    hdcode = oForm.Items("hdcode").Specific.VALUE
    Amt = oForm.Items("amt").Specific.VALUE
    ntsamt = oForm.Items("ntsamt").Specific.VALUE
    ntsamt24 = oForm.Items("ntsamt24").Specific.VALUE
    'ntsamt44 = oForm.Items("ntsamt44").Specific.VALUE
    bcard_t = oForm.Items("bcard_t").Specific.VALUE
    'bcard44 = oForm.Items("bcard44").Specific.VALUE
    bbcard_t = oForm.Items("bbcard_t").Specific.VALUE
    bbcard44 = oForm.Items("bbcard44").Specific.VALUE
    
    mart24 = oForm.Items("mart24").Specific.VALUE
    trans24 = oForm.Items("trans24").Specific.VALUE
    'mart44 = oForm.Items("mart44").Specific.VALUE
    'trans44 = oForm.Items("trans44").Specific.VALUE
    adgong24 = oForm.Items("adgong24").Specific.VALUE
    
    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    FullName = oForm.Items("FullName").Specific.VALUE

    YEAR = oForm.Items("Year").Specific.VALUE
    
    If Trim(YEAR) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY402_SAVE_Exit
    End If
    
    If Trim(CLTCOD) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY402_SAVE_Exit
    End If
    If Trim(MSTCOD) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY402_SAVE_Exit
    End If
            
    If Trim(target) = "220" And Trim(hdcode) = "" Then
        Call MDC_Com.MDC_GF_Message("장애인코드가 없습니다. 장애인인 경우 장애인 코드를 선택바랍니다. 확인바랍니다..", "E")
        GoTo PH_PY402_SAVE_Exit
    End If
    
    If Trim(target) <> "220" And Trim(hdcode) <> "" Then
        hdcode = ""
    End If
        
    bcard_t = oForm.Items("bcard_t").Specific.VALUE
    'bcard44 = oForm.Items("bcard44").Specific.VALUE
    bbcard_t = oForm.Items("bbcard_t").Specific.VALUE
    bbcard44 = oForm.Items("bbcard44").Specific.VALUE
    
    mart24 = oForm.Items("mart24").Specific.VALUE
    trans24 = oForm.Items("trans24").Specific.VALUE
    'mart44 = oForm.Items("mart44").Specific.VALUE
    'trans44 = oForm.Items("trans44").Specific.VALUE
    adgong24 = oForm.Items("adgong24").Specific.VALUE
    
    'If Trim(juminno) = "" Or (Div <> "70" And Val(amt) + Val(ntsamt) + Val(mart24) + Val(trans24) + Val(mart44) + Val(trans44) + Val(bcard_t) + Val(bcard44) + Val(bbcard_t) + Val(bbcard44) = 0) Then
    If Trim(juminno) = "" Or (Div <> "70" And Val(Amt) + Val(ntsamt) + Val(mart24) + Val(trans24) + Val(adgong24) + Val(bcard_t) + Val(bbcard_t) + Val(bbcard44) = 0) Then
        Call MDC_Com.MDC_GF_Message("정상적인 자료가 아닙니다. 확인바랍니다..", "E")
        GoTo PH_PY402_SAVE_Exit
    End If
        
            
    sQry = " Select U_Char2, U_Char3 From [@PS_HR200L] Where Code = '71' And U_Code = '" & target & "' "
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount > 0 Then
    
        CheckDate1 = oRecordSet.Fields(0).VALUE
        CheckDate2 = oRecordSet.Fields(1).VALUE
        
        If Trim(CheckDate1) <> "" Then
            If relate = "05" Or relate = "06" Or relate = "07" Or relate = "12" Or relate = "13" Or relate = "21" Or relate = "22" Then
                If birthymd > CheckDate1 Then
                    vReturnValue = Sbo_Application.MessageBox("기준일자 이후출생자입니다. 계속하겠습니까?", 1, "&확인", "&취소")
                    Select Case vReturnValue
                        Case 1:
                        Case 2:
                            GoTo PH_PY402_SAVE_Exit
                    End Select
                End If
            End If
        End If
        
        If Trim(CheckDate2) <> "" Then
            If relate = "03" Or relate = "04" Or relate = "08" Or relate = "23" Then
                If birthymd <= CheckDate2 Then
                    vReturnValue = Sbo_Application.MessageBox("기준일자 이전출생자입니다. 계속하겠습니까?", 1, "&확인", "&취소")
                    Select Case vReturnValue
                        Case 1:
                        Case 2:
                            GoTo PH_PY402_SAVE_Exit
                    End Select
                End If
            End If
        End If
        
    End If
    
    
    sQry = " Select Count(*) From [p_seoybase] Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "'"
    sQry = sQry + " And div = '" & Div & "' And target = '" & target & "' And relate = '" & relate & "' And juminno = '" & juminno & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.Fields(0).VALUE > 0 Then
        '//갱신
        
        sQry = "Update [p_seoybase] set "
        sQry = sQry + "kname = '" & kname & "',"
        sQry = sQry + "addr = '" & addr & "',"
        sQry = sQry + "birthymd = '" & birthymd & "',"
        sQry = sQry + "hdcode = '" & hdcode & "',"
        sQry = sQry + "amt = " & Amt & ","
        sQry = sQry + "ntsamt =" & ntsamt & ","
        sQry = sQry + "ntsamt24 =" & ntsamt24 & ","
        'sQry = sQry + "ntsamt44 =" & ntsamt44 & ","
        sQry = sQry + "bcard_t =" & bcard_t & ","
        'sQry = sQry + "bcard44 =" & bcard44 & ","
        sQry = sQry + "bbcard_t =" & bbcard_t & ","
        sQry = sQry + "bbcard44 =" & bbcard44 & ","
        sQry = sQry + "mart24 =" & mart24 & ","
        sQry = sQry + "trans24 =" & trans24 & ","
        'sQry = sQry + "mart44 =" & mart44 & ","
        'sQry = sQry + "trans44 =" & trans44
        sQry = sQry + "adgong24 =" & adgong24
        sQry = sQry + " Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "'"
        sQry = sQry + " And div = '" & Div & "' And target = '" & target & "' And relate = '" & relate & "' And juminno = '" & juminno & "'"
        
        oRecordSet.DoQuery sQry
        
    Else
        
        '//신규
        sQry = "INSERT INTO [p_seoybase]"
        sQry = sQry & " ("
        sQry = sQry & "saup,"
        sQry = sQry & "yyyy,"
        sQry = sQry & "sabun,"
        sQry = sQry & "div,"
        sQry = sQry & "target,"
        sQry = sQry & "relate,"
        sQry = sQry & "kname,"
        sQry = sQry & "juminno,"
        sQry = sQry & "addr,"
        sQry = sQry & "birthymd,"
        sQry = sQry & "hdcode,"
        sQry = sQry & "amt,"
        sQry = sQry & "ntsamt,"
        sQry = sQry & "soduk,"
        sQry = sQry & "ntsamt24,"
        'sQry = sQry & "ntsamt44,"
        sQry = sQry & "bcard_t, "
        'sQry = sQry & "bcard44, "
        sQry = sQry & "bbcard_t, "
        sQry = sQry & "bbcard44, "
        sQry = sQry & "mart24, "
        sQry = sQry & "trans24, "
        'sQry = sQry & "mart44, "
        'sQry = sQry & "trans44 )"
        sQry = sQry & "adgong24 )"
        sQry = sQry & " VALUES("

        sQry = sQry & "'" & CLTCOD & "',"
        sQry = sQry & "'" & YEAR & "',"
        sQry = sQry & "'" & MSTCOD & "',"
        sQry = sQry & "'" & Div & "',"
        sQry = sQry & "'" & target & "',"
        sQry = sQry & "'" & relate & "',"
        sQry = sQry & "'" & kname & "',"
        sQry = sQry & "'" & juminno & "',"
        sQry = sQry & "'" & addr & "',"
        sQry = sQry & "'" & birthymd & "',"
        sQry = sQry & "'" & hdcode & "',"
        sQry = sQry & Amt & ","
        sQry = sQry & ntsamt & ", 0 ,"
        
        sQry = sQry & ntsamt24 & ","
        'sQry = sQry & ntsamt44 & ","
        sQry = sQry & bcard_t & ","
        'sQry = sQry & bcard44 & ","
        sQry = sQry & bbcard_t & ","
        sQry = sQry & bbcard44 & ","
        sQry = sQry & mart24 & ","
        sQry = sQry & trans24 & ","
        'sQry = sQry & mart44 & ","
        'sQry = sQry & trans44 & " )"
        sQry = sQry & adgong24 & " )"
        

        oRecordSet.DoQuery sQry
    End If
    
    Call PH_PY402_FormItemEnabled

    Set oRecordSet = Nothing
    Call oForm.Freeze(False)

    Call PH_PY402_MTX01

    Exit Sub
    
PH_PY402_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY402_SAVE_Error:
    Call oForm.Freeze(False)
   
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY402_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY402_Delete()
    '//선택된 자료 삭제
    
    Dim CLTCOD As String
    Dim MSTCOD As String
    Dim YEAR As String
    Dim FullName As String
    
    Dim Div$, target$, relate$, kname$, juminno$, addr$, birthymd$
    
    Dim i As Integer
    Dim cnt As Integer
    
    Dim sQry    As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY402_Delete_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    
    Call oForm.Freeze(True)
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    YEAR = oForm.Items("Year").Specific.VALUE
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    FullName = oForm.Items("kname").Specific.VALUE
    Div = oForm.Items("div").Specific.VALUE
    target = oForm.Items("target").Specific.VALUE
    relate = oForm.Items("relate").Specific.VALUE
    kname = oForm.Items("kname").Specific.VALUE
    juminno = oForm.Items("juminno").Specific.VALUE
    addr = oForm.Items("addr").Specific.VALUE
    birthymd = oForm.Items("birthymd").Specific.VALUE
      
      
    sQry = " Select Count(*) From [p_seoybase] Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "'"
    sQry = sQry + " And div = '" & Div & "' And target = '" & target & "' And relate = '" & relate & "' And juminno = '" & juminno & "'"
    oRecordSet.DoQuery sQry
    
    
    cnt = oRecordSet.Fields(0).VALUE
    If cnt > 0 Then
    
        If Trim(YEAR) = "" Then
            Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY402_Delete_Exit
        End If
        
        If Trim(CLTCOD) = "" Then
            Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY402_Delete_Exit
        End If
        If Trim(MSTCOD) = "" Then
            Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY402_Delete_Exit
        End If
        
        
        
        
        If Sbo_Application.MessageBox(" 선택한대상자('" & FullName & "')을 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
            sQry = "Delete From [p_seoybase] Where saup = '" & CLTCOD & "' AND  yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "'"
            sQry = sQry + " And div = '" & Div & "' And target = '" & target & "' And relate = '" & relate & "' And juminno = '" & juminno & "'"
            oRecordSet.DoQuery sQry
        End If
    End If
    
   
    Call oForm.Freeze(False)
    
        
    Call PH_PY402_MTX01
    
    Set oRecordSet = Nothing
    
    
    Exit Sub
PH_PY402_Delete_Exit:
    Set oRecordSet = Nothing
        
    Call oForm.Freeze(False)
    Exit Sub
PH_PY402_Delete_Error:
    Set oRecordSet = Nothing
   
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY402_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY402_TitleSetting(iRow As Long)
    Dim i               As Long
    Dim j               As Long
    Dim sQry            As String
     
    Dim COLNAM(14)       As String

    Dim oColumn         As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
     
    On Error GoTo Error_Message
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    COLNAM(0) = "년도"
    COLNAM(1) = "사번"
    COLNAM(2) = "공제구분코드"
    COLNAM(3) = "공제구분"
    COLNAM(4) = "공제대상코드"
    COLNAM(5) = "공제대상"
    COLNAM(6) = "관계코드"
    COLNAM(7) = "관계"
    COLNAM(8) = "성명"
    COLNAM(9) = "주민번호"
    COLNAM(10) = "금액(국세청)"
    COLNAM(11) = "금액(국세청외)"
    COLNAM(12) = "합계금액"
    COLNAM(13) = "전통시장"
    COLNAM(14) = "대중교통"

    For i = 0 To UBound(COLNAM)
        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        oGrid1.Columns.Item(i).Editable = False
        If COLNAM(i) = "사번" Or COLNAM(i) = "공제구분코드" Or COLNAM(i) = "공제대상코드" Or COLNAM(i) = "관계코드" Or COLNAM(i) = "주민번호" Then
            oGrid1.Columns.Item(i).Visible = False
        End If
        
        oGrid1.Columns.Item(i).RightJustified = True
        
    Next i
    
    oGrid1.AutoResizeColumns

    oForm.Freeze False

    Set oColumn = Nothing

    Exit Sub

Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY402_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub







