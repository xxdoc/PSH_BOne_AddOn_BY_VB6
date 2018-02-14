VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY411.cls
'//  Module         : 인사관리 > 연말정산관리
'//  Desc           : 연금.저축등소득공제 등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oDS_PH_PY411A     As SAPbouiCOM.DataTable


Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY411.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY411_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY411"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.PaneLevel = 1
    oForm.Freeze True
    Call PH_PY411_CreateItems
    Call PH_PY411_FormItemEnabled
    Call PH_PY411_EnableMenus
'    Call PH_PY411_SetDocument(oFromDocEntry01)
'    Call PH_PY411_FormResize
    
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

Private Function PH_PY411_CreateItems() As Boolean

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
    
    On Error GoTo PH_PY411_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid1 = oForm.Items("Grid01").Specific
    
    oForm.DataSources.DataTables.Add ("PH_PY411")
    
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY411")
    Set oDS_PH_PY411A = oForm.DataSources.DataTables.Item("PH_PY411")
    

    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
    
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
'    oCombo.Select CLTCOD, psk_ByValue
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
    
    '//순번
    Call oForm.DataSources.UserDataSources.Add("seqn", dt_SHORT_TEXT, 3)
    oForm.Items("seqn").Specific.DataBind.SetBound True, "", "seqn"
    
    '//공제구분
    Set oCombo = oForm.Items("gubun").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '77' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oCombo.Select 0, psk_Index
    oForm.Items("gubun").DisplayDesc = True
    
    '//금융기관코드
    Call oForm.DataSources.UserDataSources.Add("bcode", dt_SHORT_TEXT, 10)
    oForm.Items("bcode").Specific.DataBind.SetBound True, "", "bcode"
    
    '//금융기관명
    Call oForm.DataSources.UserDataSources.Add("bname", dt_SHORT_TEXT, 100)
    oForm.Items("bname").Specific.DataBind.SetBound True, "", "bname"
    
    '//계좌번호
    Call oForm.DataSources.UserDataSources.Add("bnum", dt_SHORT_TEXT, 20)
    oForm.Items("bnum").Specific.DataBind.SetBound True, "", "bnum"
    
    '//납입년차
    Call oForm.DataSources.UserDataSources.Add("yuncha", dt_SUM)
    oForm.Items("yuncha").Specific.DataBind.SetBound True, "", "yuncha"
    
    '//불입금액
    Call oForm.DataSources.UserDataSources.Add("amt", dt_SUM)
    oForm.Items("amt").Specific.DataBind.SetBound True, "", "amt"
    
    '//공제금액
    Call oForm.DataSources.UserDataSources.Add("gamt", dt_SUM)
    oForm.Items("gamt").Specific.DataBind.SetBound True, "", "gamt"
    
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
    
PH_PY411_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY411_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY411_EnableMenus()
    
    On Error GoTo PH_PY411_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY411_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY411_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY411_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY411_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY411_FormItemEnabled
'        Call PH_PY411_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY411_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY411_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY411_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY411_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    
    Dim CLTCOD As String
    Dim sPosDate As String
    
    On Error GoTo PH_PY411_FormItemEnabled_Error
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
    
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)    '//문서추가
        
        
        If Trim(oForm.Items("Year").Specific.VALUE) = "" Then
            oForm.Items("Year").Specific.VALUE = Format(Now, "YYYY") - 1
        End If
        'oForm.Items("Year").Specific.VALUE = Format(Now, "YYYY")
        'oForm.Items("MSTCOD").Specific.VALUE = ""
        'oForm.Items("FullName").Specific.VALUE = ""
        'oForm.Items("TeamName").Specific.VALUE = ""
        'oForm.Items("RspName").Specific.VALUE = ""
        'oForm.Items("ClsName").Specific.VALUE = ""
        
        oForm.Items("seqn").Specific.VALUE = ""
        
        Set oCombo = oForm.Items("gubun").Specific
        oCombo.Select 0, psk_Index
    
        oForm.Items("bcode").Specific.VALUE = ""
        oForm.Items("bname").Specific.VALUE = ""
        oForm.Items("bnum").Specific.VALUE = ""
        
        oForm.Items("yuncha").Specific.VALUE = 0
        oForm.Items("amt").Specific.VALUE = 0
        oForm.Items("gamt").Specific.VALUE = 0
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        oForm.Items("amt").Enabled = True

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
    
    '//Key set
    oForm.Items("CLTCOD").Enabled = True
    oForm.Items("Year").Enabled = True
    oForm.Items("MSTCOD").Enabled = True
    
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY411_FormItemEnabled_Error:
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY411_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim tSex        As String
    Dim tBrith      As String
    Dim Day         As String
    Dim ActCode     As String
    Dim CLTCOD      As String
    Dim YEAR        As String
    Dim MSTCOD      As String
    Dim FullName    As String
    Dim seqn        As String
    Dim bcode       As String
    
    Dim amt         As Long
    Dim gamt        As Long
    Dim samt        As Long
    
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
                    If PH_PY411_DataValidCheck = False Then
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "Btn_ret" Then
                    Call PH_PY411_MTX01
                End If
                
                
                
                If pval.ItemUID = "Btn01" Then
                    Call PH_PY411_SAVE
                    
                End If
                
                                
                If pval.ItemUID = "Btn_del" Then
                    Call PH_PY411_Delete
                    Call PH_PY411_FormItemEnabled
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
                            Call PH_PY411_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY411_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY411_FormItemEnabled
                        End If
                    End If
'
                End Select
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True Then
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "MSTCOD" Then
                        If oForm.Items("MSTCOD").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                        End If
                    End If
                    
                    If pval.ItemUID = "FullName" Then
                        If oForm.Items("FullName").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "bcode" Then
                        If oForm.Items("bcode").Specific.VALUE = "" Then
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
                    If pval.ItemUID = "gubun" Then
                        oForm.Items("amt").Specific.VALUE = 0
                        oForm.Items("gamt").Specific.VALUE = 0
                    End If
                    
                End If
            End If

            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row >= 0 Then
                        Select Case pval.ItemUID
                        Case "Grid01"
                            'Call oMat1.SelectRow(pval.Row, True, False)
                            PH_PY411_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                        End Select
                        
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
                            
                            oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
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
                            
                        Case "bcode"
                            bcode = oForm.Items("bcode").Specific.VALUE
                            sQry = "Select Code,"
                            sQry = sQry + " CodeName = U_CodeNm "
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '78'"
                            sQry = sQry + " And U_Code = '" + bcode + "'"
                              
                            oRecordSet.DoQuery sQry
                            
                            oForm.Items("bname").Specific.VALUE = oRecordSet.Fields("CodeName").VALUE
                            
                        Case "amt"
                            CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
                            YEAR = oForm.Items("Year").Specific.VALUE
                            MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
                            seqn = oForm.Items("seqn").Specific.VALUE
                            amt = 0
                            gamt = 0
                            
                            Dim Gubun As String
                            
                            Gubun = oForm.Items("gubun").Specific.VALUE
                            
                            Select Case Gubun
                               Case "11", "12", "22" '11.근로자퇴직급여보장법, '12.과학기술인공제, 22.연금저축
                               
                                   '총급여액계산해서 5,500 이하는 15% 아니면 12%
                                   sQry = "SELECT SUM(gwase) "
                                   sQry = sQry + "FROM( SELECT gwase   = SUM( a.U_GWASEE ) "
                                   sQry = sQry + "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.Code "
                                   sQry = sQry + "       WHERE b.U_CLTCOD = '" + CLTCOD + "' "
                                   sQry = sQry + "         And a.U_CLTCOD = b.U_CLTCOD "
                                   sQry = sQry + "         And a.U_YM     BETWEEN  '" + YEAR + "' + '01' AND '" + YEAR + "' + '12' "
                                   sQry = sQry + "         And isnull(b.Code,'')  = '" + MSTCOD + "' "
                                   sQry = sQry + "      Union All "
                                   sQry = sQry + "      SELECT gwase   = SUM( a.U_GWASEE ) "
                                   sQry = sQry + "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.U_PreCode "
                                   sQry = sQry + "       WHERE b.U_CLTCOD = '" + CLTCOD + "' "
                                   sQry = sQry + "         And a.U_CLTCOD = b.U_CLTCOD "
                                   sQry = sQry + "         And a.U_YM     BETWEEN  '" + YEAR + "' + '01' AND '" + YEAR + "' + '12' "
                                   sQry = sQry + "         And isnull(b.Code,'')  = '" + MSTCOD + "' "
                                   sQry = sQry + "         And Isnull(b.U_PreCode,'') <> '' "
                                   sQry = sQry + "     ) g"
                              
                                   oRecordSet.DoQuery sQry
                                   samt = oRecordSet.Fields(0).VALUE   '총급여액(과세대상)
                               
                                   sQry = " Exec PH_PY411 '" & CLTCOD & "', '" & YEAR & "','" & MSTCOD & "','" & Gubun & "'," & oForm.Items("amt").Specific.VALUE
                                   oRecordSet.DoQuery sQry
                                   gamt = oRecordSet.Fields(0).VALUE
                                   
'                                   sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "' And seqn <> '" & seqn & "' And gubun IN ('11','12','22') "
'                                   oRecordSet.DoQuery sQry
'                                   gamt = oRecordSet.Fields(0).VALUE
                                    
                                    If samt <= 55000000 Then  '5500백기준
                                       amt = Round(gamt * 0.15, 0)   '15%
                                       oForm.Items("gamt").Specific.VALUE = amt
                                       
                                    Else
                                       amt = Round(gamt * 0.12, 0)   '12%
                                       oForm.Items("gamt").Specific.VALUE = amt
                                    End If
                                    oForm.Items("amt").Specific.VALUE = gamt
                               
                               Case "21" '21.개인연금저축
                                    sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "' And seqn <> '" & seqn & "' And gubun = '21'"
                                    oRecordSet.DoQuery sQry
                                    gamt = oRecordSet.Fields(0).VALUE
                                    
                                    amt = Round(oForm.Items("amt").Specific.VALUE * 0.4, 0)
                                    
                                    If gamt + amt > 720000 Then
                                       oForm.Items("gamt").Specific.VALUE = 720000 - gamt
                                    Else
                                       oForm.Items("gamt").Specific.VALUE = amt
                                    End If
                                    
                                    If oForm.Items("gamt").Specific.VALUE < 0 Then
                                       oForm.Items("gamt").Specific.VALUE = 0
                                    End If
                               
                               
                               Case "31", "32", "34"  '31.청약저축, 32.주택청약종합저축, 34.근로자주택마련저축
                                    sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "' And seqn <> '" & seqn & "' And gubun IN ('31','32','34') "
                                    oRecordSet.DoQuery sQry
                                    gamt = oRecordSet.Fields(0).VALUE
                                    
                                    amt = Round(oForm.Items("amt").Specific.VALUE * 0.4, 0)
                                
                                    If gamt + amt > 480000 Then
                                       oForm.Items("gamt").Specific.VALUE = 480000 - gamt
                                    Else
                                       oForm.Items("gamt").Specific.VALUE = amt
                                    End If
                                    
                                    If oForm.Items("gamt").Specific.VALUE < 0 Then
                                       oForm.Items("gamt").Specific.VALUE = 0
                                    End If
                                    
                               Case "51"  '51.장기집합투자증권저축  40% 240만원한도
                                    sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "' And seqn <> '" & seqn & "' And gubun = '51'"
                                    oRecordSet.DoQuery sQry
                                    gamt = oRecordSet.Fields(0).VALUE
                                    
                                    amt = Round(oForm.Items("amt").Specific.VALUE * 0.4, 0)
                                
                                    If gamt + amt > 2400000 Then
                                       oForm.Items("gamt").Specific.VALUE = 2400000 - gamt
                                    Else
                                       oForm.Items("gamt").Specific.VALUE = amt
                                    End If
                                    
                                    If oForm.Items("gamt").Specific.VALUE < 0 Then
                                       oForm.Items("gamt").Specific.VALUE = 0
                                    End If
                               
                         
                            End Select
                    End Select
                 
                End If
            End If
'            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'                Call PH_PY411_AddMatrixRow
                
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
                Set oDS_PH_PY411A = Nothing
                
'                Set oMat1 = Nothing
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY411A", "Code")
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
'                Call PH_PY411_FormItemEnabled
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY411_FormItemEnabled
'                Call PH_PY411_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY411_FormItemEnabled
'                Call PH_PY411_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY411_FormItemEnabled
'                Call PH_PY411_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY411_FormItemEnabled
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


Sub PH_PY411_FormClear()
On Error GoTo PH_PY411_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY411'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY411_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY411_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY411_DataValidCheck() As Boolean
On Error GoTo PH_PY411_DataValidCheck_Error
    PH_PY411_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY411_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY411_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY411_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY411_MTX01()

    '//그리드에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim iRow    As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY411_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("CLTCOD").Specific.VALUE
    Param02 = oForm.Items("Year").Specific.VALUE
    Param03 = oForm.Items("MSTCOD").Specific.VALUE
        
    If Trim(Param01) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY411_MTX01_Exit
    End If
        
    If Trim(Param02) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY411_MTX01_Exit
    End If
    
    If Trim(Param03) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY411_MTX01_Exit
    End If
        
        
        
    sQry = "EXEC PH_PY411_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "'"

    Call oDS_PH_PY411A.ExecuteQuery(sQry)
    
    
    
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    Call PH_PY411_TitleSetting(iRow)
    
    oForm.Update
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY411_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY411_MTX01_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY411_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_PY411_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)


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
    
    On Error GoTo PH_PY411_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    
    Param01 = oDS_PH_PY411A.Columns.Item("사업장").Cells(oRow).VALUE
    Param02 = oDS_PH_PY411A.Columns.Item("년도").Cells(oRow).VALUE
    Param03 = oDS_PH_PY411A.Columns.Item("사번").Cells(oRow).VALUE
    Param04 = oDS_PH_PY411A.Columns.Item("순번").Cells(oRow).VALUE
    
    
    
    sQry = "EXEC PH_PY411_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        
        'oForm.Items("MSTCOD").Specific.VALUE = oDS_PH_PY411A.Columns.Item("MSTCOD").Cells(oRow).VALUE
        'oForm.Items("FullName").Specific.VALUE = oDS_PH_PY411A.Columns.Item("FullName").Cells(oRow).VALUE
        
        oForm.Items("seqn").Specific.VALUE = ""
        
        Set oCombo = oForm.Items("gubun").Specific
        oCombo.Select 0, psk_Index
        
        oForm.Items("bcode").Specific.VALUE = ""
        oForm.Items("bname").Specific.VALUE = ""
        oForm.Items("bnum").Specific.VALUE = ""
        
        oForm.Items("yuncha").Specific.VALUE = 0
        oForm.Items("amt").Specific.VALUE = 0
        oForm.Items("gamt").Specific.VALUE = 0
        
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY411_MTX02_Exit
    End If
    
    oForm.DataSources.UserDataSources.Item("seqn").VALUE = oRecordSet.Fields("seqn").VALUE
    
    Set oCombo = oForm.Items("gubun").Specific
    oCombo.Select oRecordSet.Fields("gubun").VALUE, psk_ByValue
    
    oForm.DataSources.UserDataSources.Item("bcode").VALUE = oRecordSet.Fields("bcode").VALUE
    oForm.DataSources.UserDataSources.Item("bname").VALUE = oRecordSet.Fields("bname").VALUE
    
    oForm.DataSources.UserDataSources.Item("yuncha").VALUE = oRecordSet.Fields("yuncha").VALUE
    oForm.DataSources.UserDataSources.Item("amt").VALUE = oRecordSet.Fields("amt").VALUE
    oForm.DataSources.UserDataSources.Item("gamt").VALUE = oRecordSet.Fields("gamt").VALUE
    oForm.DataSources.UserDataSources.Item("bnum").VALUE = oRecordSet.Fields("bnum").VALUE
'    '//부서
'    oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
'    oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
'    oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
    
    '//Key Disable
    oForm.Items("CLTCOD").Enabled = False
    oForm.Items("Year").Enabled = False
    oForm.Items("MSTCOD").Enabled = False
   ' oForm.Items("amt").Enabled = False
    
    oForm.Update
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY411_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY411_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY411_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY411_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY411_Validate_Error
    PH_PY411_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY411A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY411_Validate = False
        GoTo PH_PY411_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY411_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY411_Validate_Error:
    PH_PY411_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY411_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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


Private Sub PH_PY411_SAVE()

    '//데이타 저장
    
    Dim i       As Long
    Dim seqncom As Integer
    Dim sQry    As String
    
    Dim saup$, yyyy$, sabun$, seqn$, Gubun$, bcode$, bname$, bnum$
    Dim yuncha, amt, gamt As Double
    
    saup = oForm.Items("CLTCOD").Specific.VALUE
    yyyy = oForm.Items("Year").Specific.VALUE
    sabun = Trim(oForm.Items("MSTCOD").Specific.VALUE)
        
    seqn = oForm.Items("seqn").Specific.VALUE
    Gubun = oForm.Items("gubun").Specific.VALUE
    bcode = oForm.Items("bcode").Specific.VALUE
    bname = oForm.Items("bname").Specific.VALUE
    bnum = oForm.Items("bnum").Specific.VALUE
    yuncha = oForm.Items("yuncha").Specific.VALUE
    amt = oForm.Items("amt").Specific.VALUE
    gamt = oForm.Items("gamt").Specific.VALUE
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY411_SAVE_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If Trim(yyyy) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY411_SAVE_Exit
    End If
    
    If Trim(saup) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY411_SAVE_Exit
    End If
    
    If Trim(sabun) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY411_SAVE_Exit
    End If
    
    
    If Trim(Gubun) = "" Or Trim(bcode) = "" Or Trim(bnum) = "" Or amt = 0 Then
        Call MDC_Com.MDC_GF_Message("정상적인 자료가 아닙니다. 확인바랍니다..", "E")
        GoTo PH_PY411_SAVE_Exit
    End If
    
            
    sQry = " Select Count(*) From [p_seoybank] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "' And seqn = '" & seqn & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.Fields(0).VALUE > 0 Then
        '//갱신
        
        sQry = "Update [p_seoybank] set "
        sQry = sQry + "gubun = '" & Gubun & "',"
        sQry = sQry + "bcode = '" & bcode & "',"
        sQry = sQry + "bname = '" & bname & "',"
        sQry = sQry + "bnum = '" & bnum & "',"
        
        sQry = sQry + "yuncha = " & yuncha & ","
        sQry = sQry + "amt = " & amt & ","
        sQry = sQry + "gamt = " & gamt & ""
        
        sQry = sQry + " Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "' And seqn = '" & seqn & "'"
        
        oRecordSet.DoQuery sQry
        
    Else
        '//신규
        '순번 계산
        sQry = " Select Convert(int,Max(seqn)) From [p_seoybank] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
        oRecordSet.DoQuery sQry
        seqncom = oRecordSet.Fields(0).VALUE
        seqncom = seqncom + 1
        seqn = Format(seqncom, "000")
        '------------------------------------------------------
        
        sQry = "INSERT INTO [p_seoybank]"
        sQry = sQry & " ("
        sQry = sQry & "saup,"
        sQry = sQry & "yyyy,"
        sQry = sQry & "sabun,"
        sQry = sQry & "seqn,"
        sQry = sQry & "gubun,"
        sQry = sQry & "bcode,"
        sQry = sQry & "bname,"
        sQry = sQry & "bnum,"
        sQry = sQry & "yuncha,"
        sQry = sQry & "amt,"
        sQry = sQry & "gamt"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("

        sQry = sQry & "'" & saup & "',"
        sQry = sQry & "'" & yyyy & "',"
        sQry = sQry & "'" & sabun & "',"
        sQry = sQry & "'" & seqn & "',"
        sQry = sQry & "'" & Gubun & "',"
        sQry = sQry & "'" & bcode & "',"
        sQry = sQry & "'" & bname & "',"
        sQry = sQry & "'" & bnum & "',"
        
        sQry = sQry & yuncha & ","
        sQry = sQry & amt & ","
        sQry = sQry & gamt & ""
        sQry = sQry & " ) "

        oRecordSet.DoQuery sQry
    End If
    
    Call PH_PY411_FormItemEnabled
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)

    Call PH_PY411_MTX01

    Exit Sub
    
PH_PY411_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY411_SAVE_Error:
    Call oForm.Freeze(False)
   
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY411_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY411_Delete()
    '//선택된 자료 삭제
    
    Dim saup$, yyyy$, sabun$, seqn$
    Dim medcex, ntamt, cont As Double
    
    
    Dim i As Integer
    Dim cnt As Integer
    
    Dim sQry    As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY411_Delete_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    
    
    saup = oForm.Items("CLTCOD").Specific.VALUE
    yyyy = oForm.Items("Year").Specific.VALUE
    sabun = Trim(oForm.Items("MSTCOD").Specific.VALUE)
        
    seqn = oForm.Items("seqn").Specific.VALUE
    
    sQry = " Select Count(*) From [p_seoybank] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "' And seqn = '" & seqn & "'"
    oRecordSet.DoQuery sQry
    
    cnt = oRecordSet.Fields(0).VALUE
    If cnt > 0 Then
    
        If Trim(yyyy) = "" Then
            Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY411_Delete_Exit
        End If
        
        If Trim(saup) = "" Then
            Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY411_Delete_Exit
        End If
        If Trim(sabun) = "" Then
            Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY411_Delete_Exit
        End If
        
        
        If Sbo_Application.MessageBox(" 선택한라인('" & seqn & "')을 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
            sQry = "Delete From [p_seoybank] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "' And seqn = '" & seqn & "'"
            oRecordSet.DoQuery sQry
        End If
    End If
    
   
    Call oForm.Freeze(False)
    
        
    Call PH_PY411_MTX01
    
    Set oRecordSet = Nothing
    
    
    Exit Sub
PH_PY411_Delete_Exit:
    Set oRecordSet = Nothing
        
    Call oForm.Freeze(False)
    Exit Sub
PH_PY411_Delete_Error:
    Set oRecordSet = Nothing
   
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY411_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY411_TitleSetting(iRow As Long)
'    Dim i               As Long
'    Dim j               As Long
'    Dim sQry            As String
'
'    Dim COLNAM(5)       As String
'
'    Dim oColumn         As SAPbouiCOM.EditTextColumn
'    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
'
'    Dim oRecordSet  As SAPbobsCOM.Recordset
'
'    On Error GoTo Error_Message
'
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'    oForm.Freeze True
'
'    COLNAM(0) = "년도"
'    COLNAM(1) = "부서"
'    COLNAM(2) = "담당"
'    COLNAM(3) = "사번"
'    COLNAM(4) = "성명"
'    COLNAM(5) = "직급"
'
'    For i = 0 To UBound(COLNAM)
'        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
'        oGrid1.Columns.Item(i).Editable = False
'
'        oGrid1.Columns.Item(i).RightJustified = True
'
'    Next i
'
'    oGrid1.AutoResizeColumns
'
'    oForm.Freeze False
'
'    Set oColumn = Nothing
'
'    Exit Sub
'
'Error_Message:
'    oForm.Freeze False
'    Set oColumn = Nothing
'    Sbo_Application.SetStatusBarMessage "PH_PY411_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub









