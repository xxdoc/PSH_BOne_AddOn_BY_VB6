VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY407"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY407.cls
'//  Module         : 인사관리 > 연말정산관리
'//  Desc           : 정산기부금등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oMat             As SAPbouiCOM.Matrix
Public oDS_PH_PY407A    As SAPbouiCOM.DataTable
Private oDS_PH_PY407L   As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY407.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY407_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY407"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.PaneLevel = 1
    oForm.Freeze True
    Call PH_PY407_CreateItems
    Call PH_PY407_FormItemEnabled
    Call PH_PY407_EnableMenus
'    Call PH_PY407_SetDocument(oFromDocEntry01)
'    Call PH_PY407_FormResize
    
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

Private Function PH_PY407_CreateItems() As Boolean

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
    
    On Error GoTo PH_PY407_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY407L = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    Set oGrid1 = oForm.Items("Grid01").Specific
    
    Set oMat = oForm.Items("Mat01").Specific
    oMat.SelectionMode = ms_NotSupported
    
    oForm.DataSources.DataTables.Add ("PH_PY407")
    
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY407")
    Set oDS_PH_PY407A = oForm.DataSources.DataTables.Item("PH_PY407")
    

    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
'    oCombo.Select 0, psk_Index
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//년도
    Call oForm.DataSources.UserDataSources.Add("Year", dt_SHORT_TEXT, 4)
    oForm.Items("Year").Specific.DataBind.SetBound True, "", "Year"
    
    '//사번
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    oForm.Items("MSTCOD").Specific.DataBind.SetBound True, "", "MSTCOD"
    
    '//이름
    Call oForm.DataSources.UserDataSources.Add("FullName", dt_SHORT_TEXT, 20)
    oForm.Items("FullName").Specific.DataBind.SetBound True, "", "FullName"
    
    '//관계
    Set oCombo = oForm.Items("rel").Specific
    oCombo.ValidValues.Add "", ""
    oCombo.ValidValues.Add "1", "거주자"
    oCombo.ValidValues.Add "2", "배우자"
    oCombo.ValidValues.Add "3", "직계비속"
    oCombo.ValidValues.Add "4", "직계존속"
    oCombo.ValidValues.Add "5", "형제,자매"
    oCombo.ValidValues.Add "6", "그외"
    oCombo.Select 0, psk_Index
    oForm.Items("rel").DisplayDesc = True
    
    '//성명
    Call oForm.DataSources.UserDataSources.Add("kname", dt_SHORT_TEXT, 20)
    oForm.Items("kname").Specific.DataBind.SetBound True, "", "kname"
    
    '//주민번호
    Call oForm.DataSources.UserDataSources.Add("juminno", dt_SHORT_TEXT, 13)
    oForm.Items("juminno").Specific.DataBind.SetBound True, "", "juminno"
    
    '//기부금코드  73
    Set oCombo = oForm.Items("gibucd").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '73' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    
    '//사업자번호
    Call oForm.DataSources.UserDataSources.Add("saupno", dt_SHORT_TEXT, 13)
    oForm.Items("saupno").Specific.DataBind.SetBound True, "", "saupno"
    
    '//기부처명
    Call oForm.DataSources.UserDataSources.Add("sangho", dt_SHORT_TEXT, 30)
    oForm.Items("sangho").Specific.DataBind.SetBound True, "", "sangho"
    
    '//공제금액(국세청)
    Call oForm.DataSources.UserDataSources.Add("ntamt", dt_SUM)
    oForm.Items("ntamt").Specific.DataBind.SetBound True, "", "ntamt"
    
    '//공제금액(국세청외)
    Call oForm.DataSources.UserDataSources.Add("amt", dt_SUM)
    oForm.Items("amt").Specific.DataBind.SetBound True, "", "amt"
    
    '//기부장려금신청금액 2016
    Call oForm.DataSources.UserDataSources.Add("jamt", dt_SUM)
    oForm.Items("jamt").Specific.DataBind.SetBound True, "", "jamt"
    
    
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
    
PH_PY407_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY407_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY407_EnableMenus()
    
    On Error GoTo PH_PY407_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY407_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY407_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY407_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY407_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY407_FormItemEnabled
'        Call PH_PY407_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY407_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY407_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY407_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY407_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    
    Dim CLTCOD As String
    Dim sPosDate As String
    
    On Error GoTo PH_PY407_FormItemEnabled_Error
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
        
        Set oCombo = oForm.Items("rel").Specific
        oCombo.Select "0", psk_Index
        
        oForm.DataSources.UserDataSources.Item("kname").VALUE = ""
        oForm.DataSources.UserDataSources.Item("juminno").VALUE = ""
        
        Set oCombo = oForm.Items("gibucd").Specific
        oCombo.Select "0", psk_Index
        
        oForm.DataSources.UserDataSources.Item("saupno").VALUE = ""
        oForm.DataSources.UserDataSources.Item("sangho").VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("ntamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("amt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("jamt").VALUE = 0
        
        
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
    
    '//Key set
    oForm.Items("CLTCOD").Enabled = True
    oForm.Items("Year").Enabled = True
    oForm.Items("MSTCOD").Enabled = True
    
    oForm.Items("juminno").Enabled = True
    oForm.Items("saupno").Enabled = True
    oForm.Items("gibucd").Enabled = True
    
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY407_FormItemEnabled_Error:
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY407_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    Dim Div         As String
    Dim target      As String
    Dim relate      As String
    Dim FullName    As String
    Dim YY          As String
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
                    If PH_PY407_DataValidCheck = False Then
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "Btn_ret" Then
                    Call PH_PY407_MTX01
                End If
                
                If pval.ItemUID = "Btn01" Then
                
                        CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
                        YY = oForm.Items("Year").Specific.VALUE
                        sQry = " Select CODE from [@PSH_PAYBACK] WHERE CODE = '" & YY & CLTCOD & "'"
                        oRecordSet.DoQuery sQry
                           
                        Result = oRecordSet.Fields(0).VALUE
                    If Result <> "" Then
                        Call MDC_Com.MDC_GF_Message("등록불가 년도입니다. 담당자에게 문의바랍니다.", "E")
                    End If
                    If Result = "" Then
                        Call PH_PY407_SAVE
                    End If
                End If
                                
                If pval.ItemUID = "Btn_del" Then
                    Call PH_PY407_Delete
                    Call PH_PY407_FormItemEnabled
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
                            Call PH_PY407_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY407_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY407_FormItemEnabled
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
                    If pval.ItemUID = "rel" Then
                        oMat.Clear
                        oDS_PH_PY407L.Clear
                        
                        MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
                        relate = oForm.Items("rel").Specific.VALUE
                        
                        sQry = "EXEC [PH_PY407_03] '" + MSTCOD + "', '" + relate + "'"
                                                
                        oRecordSet.DoQuery sQry
                        
                        For i = 0 To oRecordSet.RecordCount - 1
                            If i + 1 > oDS_PH_PY407L.Size Then
                                oDS_PH_PY407L.InsertRecord (i)
                            End If
        
                            oMat.AddRow
                            oDS_PH_PY407L.Offset = i
                            
                            oDS_PH_PY407L.setValue "U_LineNum", i, i + 1
                            oDS_PH_PY407L.setValue "U_ColReg01", i, Trim(oRecordSet.Fields("kname").VALUE)
                            oDS_PH_PY407L.setValue "U_ColReg02", i, Trim(oRecordSet.Fields("juminno").VALUE)
                            oDS_PH_PY407L.setValue "U_ColReg03", i, Trim(oRecordSet.Fields("birthymd").VALUE)
                            oDS_PH_PY407L.setValue "U_ColReg04", i, Trim(oRecordSet.Fields("relate").VALUE)
                            oDS_PH_PY407L.setValue "U_ColReg05", i, Trim(oRecordSet.Fields("addr").VALUE)
                            oRecordSet.MoveNext
                        Next
                            
                        oMat.LoadFromDataSource
                        oMat.AutoResizeColumns
                        
                        If (oRecordSet.RecordCount = 0) Then
                            oForm.DataSources.UserDataSources.Item("kname").VALUE = ""
                            oForm.DataSources.UserDataSources.Item("juminno").VALUE = ""
                            Set oCombo = oForm.Items("gibucd").Specific
                            oCombo.Select "0", psk_Index
                            oForm.DataSources.UserDataSources.Item("saupno").VALUE = ""
                            oForm.DataSources.UserDataSources.Item("sangho").VALUE = ""
                            
                            oForm.DataSources.UserDataSources.Item("ntamt").VALUE = 0
                            oForm.DataSources.UserDataSources.Item("amt").VALUE = 0
                            oForm.DataSources.UserDataSources.Item("jamt").VALUE = 0
                            
                        End If
                        
                        If (oRecordSet.RecordCount = 1) Then
                            oForm.Items("kname").Specific.VALUE = oMat.Columns("kname").Cells(1).Specific.VALUE
                            oForm.Items("juminno").Specific.VALUE = oMat.Columns("juminno").Cells(1).Specific.VALUE
                        End If
                        
    
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
                            PH_PY407_MTX02 pval.ItemUID, pval.Row, pval.ColUID
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
                            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
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
                            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
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
                        Case "ntamt"
                        
'                            If oForm.Items("handoamt").Specific.VALUE > 0 Then
'                                If CDbl(oForm.Items("ntsamt").Specific.VALUE) + CDbl(oForm.Items("amt").Specific.VALUE) > CDbl(oForm.Items("handoamt").Specific.VALUE) Then
'                                    oForm.Items("ntsamt").Specific.VALUE = 0
'                                    Sbo_Application.MessageBox ("한도금액보다 초과됩니다. 확인하세요")
'                                End If
'                            End If
'
                        Case "juminno" '주민번호
                            '주민번호입력시 생년월일 생성
                            If Len(Trim(oForm.Items("juminno").Specific.VALUE)) <> 13 Then
                               ' oForm.Items("birthymd").Specific.VALUE = ""
                                Sbo_Application.MessageBox ("주민번호자릿수가 틀립니다. 확인하세요.")
                            Else
'                                If Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "1" Or Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "2" Then
'                                    oForm.Items("birthymd").Specific.VALUE = "19" + Mid(oForm.Items("juminno").Specific.VALUE, 1, 6)
'                                ElseIf Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "3" Or Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "4" Then
'                                    oForm.Items("birthymd").Specific.VALUE = "20" + Mid(oForm.Items("juminno").Specific.VALUE, 1, 6)
'                                End If
                            End If
                        Case "amt"
'                            If oForm.Items("handoamt").Specific.VALUE > 0 Then
'                                If CDbl(oForm.Items("ntsamt").Specific.VALUE) + CDbl(oForm.Items("amt").Specific.VALUE) > CDbl(oForm.Items("handoamt").Specific.VALUE) Then
'                                    oForm.Items("amt").Specific.VALUE = 0
'                                    Sbo_Application.MessageBox ("한도금액보다 초과됩니다. 확인하세요")
'                                End If
'                            End If
                    End Select
                 
                End If
            End If
'            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
'                oMat.LoadFromDataSource
'                Call PH_PY407_AddMatrixRow
                
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
                Set oDS_PH_PY407A = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY407A", "Code")
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
'                Call PH_PY407_FormItemEnabled
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY407_FormItemEnabled
'                Call PH_PY407_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY407_FormItemEnabled
'                Call PH_PY407_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY407_FormItemEnabled
'                Call PH_PY407_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY407_FormItemEnabled
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


Sub PH_PY407_FormClear()
On Error GoTo PH_PY407_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY407'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY407_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY407_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY407_DataValidCheck() As Boolean
On Error GoTo PH_PY407_DataValidCheck_Error
    PH_PY407_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY407_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY407_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY407_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY407_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim iRow        As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY407_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = oForm.Items("Year").Specific.VALUE
    Param03 = oForm.Items("MSTCOD").Specific.VALUE
        
    If Trim(Param01) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY407_MTX01_Exit
    End If
        
    If Trim(Param02) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY407_MTX01_Exit
    End If
        
    If Trim(Param03) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY407_MTX01_Exit
    End If
        
        
        
    sQry = "EXEC PH_PY407_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "'"

    Call oDS_PH_PY407A.ExecuteQuery(sQry)
    
    
    
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    Call PH_PY407_TitleSetting(iRow)
    
    oForm.Update
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY407_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY407_MTX01_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY407_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_PY407_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)


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
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY407_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    
    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = oDS_PH_PY407A.Columns.Item("연도").Cells(oRow).VALUE
    Param03 = oDS_PH_PY407A.Columns.Item("사번").Cells(oRow).VALUE
    Param04 = oDS_PH_PY407A.Columns.Item("사업자번호").Cells(oRow).VALUE
    Param05 = oDS_PH_PY407A.Columns.Item("기부금코드").Cells(oRow).VALUE
    Param06 = oDS_PH_PY407A.Columns.Item("주민번호").Cells(oRow).VALUE
    
    
    sQry = "EXEC PH_PY407_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        
      '  oForm.Items("MSTCOD").Specific.VALUE = oDS_PH_PY407A.Columns.Item("MSTCOD").Cells(oRow).VALUE
      '  oForm.Items("FullName").Specific.VALUE = oDS_PH_PY407A.Columns.Item("FullName").Cells(oRow).VALUE
        
        Set oCombo = oForm.Items("rel").Specific
        oCombo.Select "0", psk_Index
        
        oForm.DataSources.UserDataSources.Item("kname").VALUE = ""
        oForm.DataSources.UserDataSources.Item("juminno").VALUE = ""
        
        Set oCombo = oForm.Items("gibucd").Specific
        oCombo.Select "0", psk_Index
        
        oForm.DataSources.UserDataSources.Item("saupno").VALUE = ""
        oForm.DataSources.UserDataSources.Item("sangho").VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("ntamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("amt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("jamt").VALUE = 0
        
        'oForm.Items("TeamName").Specific.VALUE = ""
        'oForm.Items("RspName").Specific.VALUE = ""
        'oForm.Items("ClsName").Specific.VALUE = ""
    
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY407_MTX02_Exit
    End If
    
    'oForm.Items("Year").Specific.VALUE = oRecordSet.Fields("Year").VALUE
    'oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("MSTCOD").VALUE
    'oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
    
'    '//부서
    'oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
    'oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
    'oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
    
     Set oCombo = oForm.Items("gibucd").Specific
    oCombo.Select oRecordSet.Fields("gibucd").VALUE, psk_ByValue
    
    Set oCombo = oForm.Items("rel").Specific
    oCombo.Select oRecordSet.Fields("rel").VALUE, psk_ByValue
    
    oForm.DataSources.UserDataSources.Item("kname").VALUE = oRecordSet.Fields("kname").VALUE
    oForm.DataSources.UserDataSources.Item("juminno").VALUE = oRecordSet.Fields("juminno").VALUE
    
   
    
    oForm.DataSources.UserDataSources.Item("saupno").VALUE = oRecordSet.Fields("saupno").VALUE
    oForm.DataSources.UserDataSources.Item("sangho").VALUE = oRecordSet.Fields("sangho").VALUE
    
    oForm.DataSources.UserDataSources.Item("ntamt").VALUE = oRecordSet.Fields("ntamt").VALUE
    oForm.DataSources.UserDataSources.Item("amt").VALUE = oRecordSet.Fields("amt").VALUE
    oForm.DataSources.UserDataSources.Item("jamt").VALUE = oRecordSet.Fields("jamt").VALUE
    
    oForm.Update
    
    oForm.ActiveItem = "rel"
    
    '//key set
    oForm.Items("CLTCOD").Enabled = False
    oForm.Items("Year").Enabled = False
    oForm.Items("MSTCOD").Enabled = False
    
    
    oForm.Items("gibucd").Enabled = False
    oForm.Items("juminno").Enabled = False
    oForm.Items("saupno").Enabled = False
    
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY407_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY407_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY407_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY407_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY407_Validate_Error
    PH_PY407_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY407A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY407_Validate = False
        GoTo PH_PY407_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY407_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY407_Validate_Error:
    PH_PY407_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY407_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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


Private Sub PH_PY407_SAVE()

    '//데이타 저장
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim saup$, sabun$, FullName$, yyyy$
    Dim rel$, kname$, juminno$, gibucd$, saupno$, sangho$
    
    Dim Amt, ntamt, jamt As Double
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY407_SAVE_Error
    
    saup = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    yyyy = oForm.Items("Year").Specific.VALUE
    sabun = oForm.Items("MSTCOD").Specific.VALUE
    rel = oForm.Items("rel").Specific.VALUE
    kname = oForm.Items("kname").Specific.VALUE
    juminno = oForm.Items("juminno").Specific.VALUE
    gibucd = oForm.Items("gibucd").Specific.VALUE
    saupno = oForm.Items("saupno").Specific.VALUE
    sangho = oForm.Items("sangho").Specific.VALUE
    ntamt = oForm.Items("ntamt").Specific.VALUE
    Amt = oForm.Items("amt").Specific.VALUE
    jamt = oForm.Items("jamt").Specific.VALUE
    
    FullName = oForm.Items("FullName").Specific.VALUE

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    If Trim(yyyy) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY407_SAVE_Exit
    End If
    
    If Trim(saup) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY407_SAVE_Exit
    End If
    If Trim(sabun) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY407_SAVE_Exit
    End If
      
    If Trim(juminno) = "" Or (ntamt = 0 And Amt = 0) Then
        Call MDC_Com.MDC_GF_Message("정상적인 자료가 아닙니다. 확인바랍니다..", "E")
        GoTo PH_PY407_SAVE_Exit
    End If
    
    sQry = " Select Count(*) From [p_seoygibuhis] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
    sQry = sQry + " And saupno = '" & saupno & "' And gibucd = '" & gibucd & "' And  juminno = '" & juminno & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.Fields(0).VALUE > 0 Then
        '//갱신
        sQry = "Update [p_seoygibuhis] set "
        sQry = sQry + "rel = '" & rel & "',"
        sQry = sQry + "kname = '" & kname & "',"
        sQry = sQry + "sangho = '" & sangho & "',"
        sQry = sQry + "ntamt = " & ntamt & ","
        sQry = sQry + "jamt = " & jamt & ","
        sQry = sQry + "amt =" & Amt
        
        sQry = sQry + " Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
        sQry = sQry + " And saupno = '" & saupno & "' And gibucd = '" & gibucd & "' And  juminno = '" & juminno & "'"
        
        oRecordSet.DoQuery sQry
        
    Else
        
        '//신규
        sQry = "INSERT INTO [p_seoygibuhis]"
        sQry = sQry & " ("
        sQry = sQry & "saup,"
        sQry = sQry & "yyyy,"
        sQry = sQry & "sabun,"
        sQry = sQry & "rel,"
        sQry = sQry & "kname,"
        sQry = sQry & "juminno,"
        sQry = sQry & "gibucd,"
        sQry = sQry & "saupno,"
        sQry = sQry & "sangho,"
        sQry = sQry & "ntamt,"
        sQry = sQry & "jamt,"
        sQry = sQry & "amt)"
        
        sQry = sQry & " VALUES("

        sQry = sQry & "'" & saup & "',"
        sQry = sQry & "'" & yyyy & "',"
        sQry = sQry & "'" & sabun & "',"
        sQry = sQry & "'" & rel & "',"
        sQry = sQry & "'" & kname & "',"
        sQry = sQry & "'" & juminno & "',"
        sQry = sQry & "'" & gibucd & "',"
        sQry = sQry & "'" & saupno & "',"
        sQry = sQry & "'" & sangho & "',"

        sQry = sQry & ntamt & ","
        sQry = sQry & jamt & ","
        sQry = sQry & Amt & " )"

        oRecordSet.DoQuery sQry
    End If
    
    
    Call PH_PY407_FormItemEnabled
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)

    Call PH_PY407_MTX01

    Exit Sub
    
PH_PY407_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY407_SAVE_Error:
    Call oForm.Freeze(False)
   
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY407_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY407_Delete()
    '//선택된 자료 삭제
    Dim i       As Long
    Dim sQry    As String
    Dim cnt     As Integer
    
    Dim saup$, sabun$, FullName$, yyyy$
    Dim rel$, kname$, juminno$, gibucd$, saupno$, sangho$
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY407_Delete_Error
    
    saup = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    yyyy = oForm.Items("Year").Specific.VALUE
    sabun = oForm.Items("MSTCOD").Specific.VALUE
    juminno = oForm.Items("juminno").Specific.VALUE
    gibucd = oForm.Items("gibucd").Specific.VALUE
    saupno = oForm.Items("saupno").Specific.VALUE
    
    FullName = oForm.Items("kname").Specific.VALUE
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
      
    sQry = " Select Count(*) From [p_seoygibuhis] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
    sQry = sQry + " And saupno = '" & saupno & "' And gibucd = '" & gibucd & "' And  juminno = '" & juminno & "'"
        
    oRecordSet.DoQuery sQry
    
    cnt = oRecordSet.Fields(0).VALUE
    If cnt > 0 Then
    
        If Trim(yyyy) = "" Then
            Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY407_Delete_Exit
        End If
        
        If Trim(saup) = "" Then
            Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY407_Delete_Exit
        End If
        If Trim(sabun) = "" Then
            Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY407_Delete_Exit
        End If
        
        
        If Sbo_Application.MessageBox(" 선택한대상자('" & FullName & "')을 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
            sQry = "Delete From [p_seoygibuhis] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
            sQry = sQry + " And saupno = '" & saupno & "' And gibucd = '" & gibucd & "' And  juminno = '" & juminno & "'"
            oRecordSet.DoQuery sQry
        End If
    End If
    
   
    Call oForm.Freeze(False)
    
        
    Call PH_PY407_MTX01
    
    Set oRecordSet = Nothing
    
    
    Exit Sub
PH_PY407_Delete_Exit:
    Set oRecordSet = Nothing
        
    Call oForm.Freeze(False)
    Exit Sub
PH_PY407_Delete_Error:
    Set oRecordSet = Nothing
   
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY407_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY407_TitleSetting(iRow As Long)
'    Dim i               As Long
'    Dim j               As Long
'    Dim sQry            As String
'
'    Dim COLNAM(12)       As String
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
'    COLNAM(1) = "사번"
'    COLNAM(2) = "공제구분코드"
'    COLNAM(3) = "공제구분"
'    COLNAM(4) = "공제대상코드"
'    COLNAM(5) = "공제대상"
'    COLNAM(6) = "관계코드"
'    COLNAM(7) = "관계"
'    COLNAM(8) = "대상자성명"
'    COLNAM(9) = "주민번호"
'    COLNAM(10) = "금액(국세청)"
'    COLNAM(11) = "금액(국세청외)"
'    COLNAM(12) = "합계금액"
'
'    For i = 0 To UBound(COLNAM)
'        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
'        oGrid1.Columns.Item(i).Editable = False
'        If COLNAM(i) = "사번" Or COLNAM(i) = "공제구분코드" Or COLNAM(i) = "공제대상코드" Or COLNAM(i) = "관계코드" Or COLNAM(i) = "주민번호" Then
'            oGrid1.Columns.Item(i).Visible = False
'        End If
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
'    Sbo_Application.SetStatusBarMessage "PH_PY407_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub









