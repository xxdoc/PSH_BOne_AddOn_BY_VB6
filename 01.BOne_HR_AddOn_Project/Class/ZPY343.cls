VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY343"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY343.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 월별자료관리
'//  FormType       : 2000060343
'//  Create Date    : 2005.12.10
'//  Modified Date  :
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oCheck          As SAPbouiCOM.CheckBox
Private oDS_ZPY343H     As SAPbouiCOM.DBDataSource    '시스템코드 헤더
Private oDS_ZPY343L     As SAPbouiCOM.DBDataSource    '시스템코드 라인

Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     '클래스에서 선택한 마지막 아이템 Uid값
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional DocNum As String)
'Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY343.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    
    '------------------------------------------------------------------------
    '//여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
    '------------------------------------------------------------------------
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "ZPY343_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    '--------------------------------------------------------------------------------------------------------------
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    '--------------------------------------------------------------------------------------------------------------
    AddForms Me, oFormUniqueID, "ZPY343"
    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm.DataBrowser.BrowseBy = "DocNum"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    CreateItems
    FormItemEnabled
    
    oForm.EnableMenu ("1281"), True  '/ 찾기
    oForm.EnableMenu ("1282"), False '/ 추가
    oForm.EnableMenu ("1284"), False '/ 취소
    oForm.EnableMenu ("1293"), False '/ 행삭제
    
    If DocNum <> "" Then
        Call ShowSource(DocNum)
    End If
    
    oForm.Update
    oForm.Visible = True
    oForm.Freeze False

    Set oXmlDoc = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LoadForm_Error:
    Set oXmlDoc = Nothing
    Sbo_Application.StatusBar.SetText "Form_Load Error:" & Err.Description, bmt_Short, smt_Error
    If (oForm Is Nothing) = False Then
        oForm.Freeze False
        Set oForm = Nothing
    End If
End Sub
'*******************************************************************
'// ItemEventHander
'*******************************************************************
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
'et_ITEM_PRESSED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_ITEM_PRESSED
            If pval.BeforeAction Then
                If pval.ItemUID = "1" Then
                    '------------------------------------------------------------
                    '//추가및 업데이트시에
                    '------------------------------------------------------------
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                ElseIf pval.ItemUID = "CBtn1" Then
                    If oForm.Items("MstCode").Enabled = True Then
                        oForm.Items("MstCode").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            Else
                If pval.ItemUID = "Check1" Then
                   If oCheck.Checked = True Then
                        oForm.ActiveItem = "Focus"
                        oForm.Items("Mat1").Enabled = False
                   Else
                        oForm.Items("Mat1").Enabled = True
                   End If
                End If
            End If
'et_COMBO_SELECT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_COMBO_SELECT
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If pval.ItemUID = "CLTCOD" Then
                    '//기본사항 - 부서1 (사업장에 따른 부서변경)
                    Set oCombo = oForm.Items("DptCode").Specific

                    If oCombo.ValidValues.Count > 0 Then
                        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                            oCombo.ValidValues.Remove i, psk_Index
                        Next i
                        oCombo.ValidValues.Add "", ""
                        oCombo.Select 0, psk_Index
                    End If
                    
                    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                    sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                    sQry = sQry & " ORDER BY U_Code"
                    Call SetReDataCombo(oForm, sQry, oCombo)
                End If
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If pval.ItemUID = "Mat1" And (Left$(pval.ColUID, 3) = "Col") Then
                    Call FlushToItemValue(pval.ColUID, pval.Row)
                ElseIf pval.ItemUID = "MstCode" Or pval.ItemUID = "JnsYear" Then
                    Call FlushToItemValue(pval.ItemUID)
                End If
            End If
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
             If pval.BeforeAction = True And pval.ItemUID = "JsnYear" And pval.CharPressed = 9 Then
                If Len(Trim$(oForm.Items("JsnYear").Specific.String)) = 0 Then
                    Sbo_Application.StatusBar.SetText "정산년도를 입력하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
             ElseIf pval.BeforeAction = True And pval.ItemUID = "MstCode" And pval.CharPressed = 9 Then
                If Len(Trim$(oForm.Items("MstCode").Specific.String)) = 0 Then
                    Sbo_Application.StatusBar.SetText "사원번호를 입력하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
            End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
            If Last_Item = "Mat1" Then
                If pval.Row > 0 Then
                    Last_Item = pval.ItemUID
                End If
            Else
                Last_Item = pval.ItemUID
            End If
'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            '----------------------------------------------------
            '컬렉션에서 삭제및 모든 메모리 제거
            '----------------------------------------------------
            If pval.BeforeAction = False Then
               RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_ZPY343H = Nothing
                Set oDS_ZPY343L = Nothing
                Set oCheck = Nothing
                Set oMat1 = Nothing
            End If
'et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_MATRIX_LOAD:
            If pval.BeforeAction = False Then
                oForm.DataSources.UserDataSources("User").ValueEx = MDC_SetMod.Get_UserName(oDS_ZPY343H.GetValue("UserSign", 0))
'                Call Matrix_TitleSetting
                oMat1.AutoResizeColumns
                FormItemEnabled
                If oDS_ZPY343H.GetValue("U_Check", 0) = "Y" Then
                    oForm.Items("Mat1").Enabled = False
                Else
                    oForm.Items("Mat1").Enabled = True
                End If
            End If
    End Select
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:", bmt_Short, smt_Error
End Sub
'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    
    If pval.BeforeAction = True Then
        Select Case pval.MenuUID
        Case "1283"
            If Trim$(oDS_ZPY343H.GetValue("U_Check", 0)) = "Y" Then
                Sbo_Application.StatusBar.SetText "잠금 자료입니다. 확인하여 주십시오.", bmt_Short, smt_Error
                BubbleEvent = False
                Exit Sub
            Else
                If Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
        Case "1288", "1289", "1290", "1291"
            Call AuthorityCheck(oForm, "CLTCOD", "@ZPY343H", "DocNum")      '//접속자 권한에 따른 사업장 보기
        Case Else
                Exit Sub
        End Select
    Else
        Select Case pval.MenuUID
            Case "1287"             '// 복제
            Case "1283"             '// 제거
                Call FormItemEnabled
            Case "1281", "1282":    '// 찾기, 추가
                Call FormItemEnabled
                Call Matrix_TitleSetting
                Call Clear_UserDb
    '           oForm.Update
                oForm.Items("JsnYear").CLICK ct_Regular
            Case "1288" To "1291":
            Case "1293"
        End Select
    End If
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


'*******************************************************************
'
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oEdit           As SAPbouiCOM.EditText
    Dim sQry            As String

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//디비데이터 소스 개체 할당
    Set oDS_ZPY343H = oForm.DataSources.DBDataSources("@ZPY343H")
    Set oDS_ZPY343L = oForm.DataSources.DBDataSources("@ZPY343L")
    
    Call oForm.DataSources.UserDataSources.Add("User", dt_SHORT_TEXT, 30)
    Set oEdit = oForm.Items("UserSign").Specific
    oEdit.DataBind.SetBound True, "", "User"
    
    '//사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//부서
    oForm.Items("DptCode").DisplayDesc = True

    '// 직책
    Set oCombo = oForm.Items("StpCode").Specific
    sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("StpCode").DisplayDesc = True
    
 '/ 체크버튼(잠금체크 )
    'Call oForm.DataSources.UserDataSources.Add("CheckDS", dt_SHORT_TEXT, 1)

    Set oCheck = oForm.Items("Check1").Specific
    'oCheck.DataBind.SetBound True, "", "CheckDS"
    oCheck.ValOff = "N": oCheck.ValOn = "Y"

'
    Set oMat1 = oForm.Items("Mat1").Specific
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error:" & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim iRow    As Long
'/ 계산시
    Dim GWASEE As Double
    Dim MONPAY As Double
    Dim SUMArr(1 To 51) As Double
    Dim MstInfo         As ZPAY_g_EmpID
    
    If Left$(oUID, 3) <> "Col" Then
        Select Case oUID
            Case "MstCode"
                If oForm.Items(oUID).Specific.String = "" Then
                    oDS_ZPY343H.setValue "U_MstCode", 0, ""
                    oDS_ZPY343H.setValue "U_EmpId", 0, ""
                    oDS_ZPY343H.setValue "U_MstName", 0, ""
                Else
                    oDS_ZPY343H.setValue "U_MstCode", 0, UCase$(oForm.Items(oUID).Specific.String)
                    MstInfo = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                    oDS_ZPY343H.setValue "U_EmpId", 0, MstInfo.EmpID
                    oDS_ZPY343H.setValue "U_MstName", 0, MstInfo.MSTNAM
                End If
                oForm.Items(oUID).Update
                
           Case "JsnYear"
                If oForm.Items(oUID).Specific.String <> "" Then
                    ZPAY_GBL_JSNYER = oForm.Items(oUID).Specific.String
                End If
                Call Matrix_TitleSetting
        End Select
        Exit Sub
    End If
    GWASEE = 0
    MONPAY = 0
    For iRow = 1 To 51
        SUMArr(iRow) = 0
    Next iRow
    
    '------------------------------------------------------
    oMat1.FlushToDataSource
    Select Case oUID
        Case "Col1", "Col2", "Col3", "Col4", "Col6", "Col7", "Col8", "Col9", "Col19", "Col20", "Col21", _
            "Col22", "Col23", "Col24", "Col25", "Col26", "Col27", "Col28", "Col29", "Col30", "Col31", _
            "Col32", "Col33", "Col34", "Col35", "Col36", "Col37", "Col38", "Col39", "Col40", "Col41", _
            "Col42", "Col43", "Col44", "Col45", "Col46", "Col47", "Col48", "Col49", "Col50", "Col51", "Col52"
            
            oMat1.FlushToDataSource
            
            oDS_ZPY343L.Offset = oRow - 1
           '/ 과세총계
            GWASEE = Val(oDS_ZPY343L.GetValue("U_GwaPay", oRow - 1))
            GWASEE = GWASEE + Val(oDS_ZPY343L.GetValue("U_GwaBns", oRow - 1))
            GWASEE = GWASEE + Val(oDS_ZPY343L.GetValue("U_InJBns", oRow - 1))
            GWASEE = GWASEE + Val(oDS_ZPY343L.GetValue("U_JUSBNS", oRow - 1))  '/2007주식행사이익추가
            GWASEE = GWASEE + Val(oDS_ZPY343L.GetValue("U_URIBNS", oRow - 1))  '/2009우리사주조합인출추가
            Call oDS_ZPY343L.setValue("U_GwaSee", oRow - 1, GWASEE)
            '/ 월지급총계
            MONPAY = GWASEE
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BiGwa01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BiGwa02", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BiGwa03", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BiGwa05", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BiGwa06", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BiGwu03", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BiGwa04", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BiGwa07", oRow - 1))
            
            '/ 2009년추가.
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGG01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGH01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGH05", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGH06", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGH07", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGH08", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGH09", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGH10", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGH11", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGH12", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGH13", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGI01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGK01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGM01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGM02", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGM03", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGO01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGQ01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGS01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGT01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGX01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGY01", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGY02", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGY03", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGY20", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGY21", oRow - 1))
            MONPAY = MONPAY + Val(oDS_ZPY343L.GetValue("U_BIGZ01", oRow - 1))
            
            Call oDS_ZPY343L.setValue("U_JigTotal", oRow - 1, MONPAY)
            oMat1.SetLineData oRow
    End Select
    '/ 항목별 합계
    For iRow = 1 To oMat1.VisualRowCount - 1
        oDS_ZPY343L.Offset = iRow - 1
        SUMArr(1) = SUMArr(1) + Val(oDS_ZPY343L.GetValue("U_GwaPay", iRow - 1)) '/ 과세급여
        SUMArr(2) = SUMArr(2) + Val(oDS_ZPY343L.GetValue("U_GwaBns", iRow - 1)) '/ 과세상여
        SUMArr(3) = SUMArr(3) + Val(oDS_ZPY343L.GetValue("U_InJBns", iRow - 1)) '/ 인정상여
        SUMArr(4) = SUMArr(4) + Val(oDS_ZPY343L.GetValue("U_GwaSee", iRow - 1)) '/ 과세총계
        SUMArr(5) = SUMArr(5) + Val(oDS_ZPY343L.GetValue("U_BiGwa01", iRow - 1)) '/ 생산비과
        SUMArr(6) = SUMArr(6) + Val(oDS_ZPY343L.GetValue("U_BiGwa02", iRow - 1)) '/ 기타비과
        SUMArr(7) = SUMArr(7) + Val(oDS_ZPY343L.GetValue("U_BiGwa03", iRow - 1)) '/ 국외비과(\)
        SUMArr(8) = SUMArr(8) + Val(oDS_ZPY343L.GetValue("U_BiGwa04", iRow - 1)) '/ 국외비과($)
        SUMArr(9) = SUMArr(9) + Val(oDS_ZPY343L.GetValue("U_JiGTotal", iRow - 1)) '/ 월지급총액
        SUMArr(10) = SUMArr(10) + Val(oDS_ZPY343L.GetValue("U_KukAmt", iRow - 1)) '/ 국민연금
        SUMArr(11) = SUMArr(11) + Val(oDS_ZPY343L.GetValue("U_MedAmt", iRow - 1)) '/ 건강보험
        SUMArr(12) = SUMArr(12) + Val(oDS_ZPY343L.GetValue("U_GBHAmt", iRow - 1)) '/ 고용보험
        SUMArr(13) = SUMArr(13) + Val(oDS_ZPY343L.GetValue("U_GabGun", iRow - 1)) '/ 소득세
        SUMArr(14) = SUMArr(14) + Val(oDS_ZPY343L.GetValue("U_Jumin", iRow - 1)) '/ 주민세
        SUMArr(15) = SUMArr(15) + Val(oDS_ZPY343L.GetValue("U_NonTK", iRow - 1)) '/ 농특세
        'SUMArr(16) = SUMArr(16) + Val(oDS_ZPY343L.GetValue("U_GamSe", iRow - 1)) '/ 감면세액
        SUMArr(17) = SUMArr(17) + Val(oDS_ZPY343L.GetValue("U_GbuAmt", iRow - 1)) '/ 기부금
        SUMArr(18) = SUMArr(18) + Val(oDS_ZPY343L.GetValue("U_JUSBNS", iRow - 1)) '/ 주식행사이익
        SUMArr(19) = SUMArr(19) + Val(oDS_ZPY343L.GetValue("U_BiGwa05", iRow - 1)) '/ 기타비과
        SUMArr(20) = SUMArr(20) + Val(oDS_ZPY343L.GetValue("U_BiGwa06", iRow - 1)) '/ 연구비과세
        '/ 2009년추가
        SUMArr(21) = SUMArr(21) + Val(oDS_ZPY343L.GetValue("U_BIGG01", iRow - 1))
        SUMArr(22) = SUMArr(22) + Val(oDS_ZPY343L.GetValue("U_BIGH01", iRow - 1))
        SUMArr(23) = SUMArr(23) + Val(oDS_ZPY343L.GetValue("U_BIGH05", iRow - 1))
        SUMArr(24) = SUMArr(24) + Val(oDS_ZPY343L.GetValue("U_BIGH06", iRow - 1))
        SUMArr(25) = SUMArr(25) + Val(oDS_ZPY343L.GetValue("U_BIGH07", iRow - 1))
        SUMArr(26) = SUMArr(26) + Val(oDS_ZPY343L.GetValue("U_BIGH08", iRow - 1))
        SUMArr(27) = SUMArr(27) + Val(oDS_ZPY343L.GetValue("U_BIGH09", iRow - 1))
        SUMArr(28) = SUMArr(28) + Val(oDS_ZPY343L.GetValue("U_BIGH10", iRow - 1))
        SUMArr(29) = SUMArr(29) + Val(oDS_ZPY343L.GetValue("U_BIGH11", iRow - 1))
        SUMArr(30) = SUMArr(30) + Val(oDS_ZPY343L.GetValue("U_BIGH12", iRow - 1))
        SUMArr(31) = SUMArr(31) + Val(oDS_ZPY343L.GetValue("U_BIGH13", iRow - 1))
        SUMArr(32) = SUMArr(32) + Val(oDS_ZPY343L.GetValue("U_BIGI01", iRow - 1))
        SUMArr(33) = SUMArr(33) + Val(oDS_ZPY343L.GetValue("U_BIGK01", iRow - 1))
        SUMArr(34) = SUMArr(34) + Val(oDS_ZPY343L.GetValue("U_BIGM01", iRow - 1))
        SUMArr(35) = SUMArr(35) + Val(oDS_ZPY343L.GetValue("U_BIGM02", iRow - 1))
        SUMArr(36) = SUMArr(36) + Val(oDS_ZPY343L.GetValue("U_BIGM03", iRow - 1))
        SUMArr(37) = SUMArr(37) + Val(oDS_ZPY343L.GetValue("U_BIGO01", iRow - 1))
        SUMArr(38) = SUMArr(38) + Val(oDS_ZPY343L.GetValue("U_BIGQ01", iRow - 1))
        SUMArr(39) = SUMArr(39) + Val(oDS_ZPY343L.GetValue("U_BIGS01", iRow - 1))
        SUMArr(40) = SUMArr(40) + Val(oDS_ZPY343L.GetValue("U_BIGT01", iRow - 1))
        SUMArr(41) = SUMArr(41) + Val(oDS_ZPY343L.GetValue("U_BIGX01", iRow - 1))
        SUMArr(42) = SUMArr(42) + Val(oDS_ZPY343L.GetValue("U_BIGY01", iRow - 1))
        SUMArr(43) = SUMArr(43) + Val(oDS_ZPY343L.GetValue("U_BIGY02", iRow - 1))
        SUMArr(44) = SUMArr(44) + Val(oDS_ZPY343L.GetValue("U_BIGY03", iRow - 1))
        SUMArr(45) = SUMArr(45) + Val(oDS_ZPY343L.GetValue("U_BIGY20", iRow - 1))
        SUMArr(46) = SUMArr(46) + Val(oDS_ZPY343L.GetValue("U_BIGZ01", iRow - 1))
        
        SUMArr(47) = SUMArr(47) + Val(oDS_ZPY343L.GetValue("U_URIBNS", iRow - 1))
        SUMArr(48) = SUMArr(48) + Val(oDS_ZPY343L.GetValue("U_BiGwu03", iRow - 1))
        SUMArr(49) = SUMArr(49) + Val(oDS_ZPY343L.GetValue("U_BiGwa07", iRow - 1))
        SUMArr(50) = SUMArr(50) + Val(oDS_ZPY343L.GetValue("U_NGYAMT", iRow - 1))
        
        '// 2010년 추가
        SUMArr(51) = SUMArr(51) + Val(oDS_ZPY343L.GetValue("U_BIGY21", iRow - 1))
    
    Next iRow
    iRow = 12
    oDS_ZPY343L.Offset = iRow
    Call oDS_ZPY343L.setValue("U_GwaPay", iRow, SUMArr(1))
    Call oDS_ZPY343L.setValue("U_GwaBns", iRow, SUMArr(2))
    Call oDS_ZPY343L.setValue("U_InJBns", iRow, SUMArr(3))
    Call oDS_ZPY343L.setValue("U_JUSBNS", iRow, SUMArr(18))
    Call oDS_ZPY343L.setValue("U_GwaSee", iRow, SUMArr(4))
    Call oDS_ZPY343L.setValue("U_BiGwa01", iRow, SUMArr(5))
    Call oDS_ZPY343L.setValue("U_BiGwa02", iRow, SUMArr(6))
    Call oDS_ZPY343L.setValue("U_BiGwa03", iRow, SUMArr(7))
    Call oDS_ZPY343L.setValue("U_BiGwa04", iRow, SUMArr(8))
    Call oDS_ZPY343L.setValue("U_BiGwa05", iRow, SUMArr(19))
    Call oDS_ZPY343L.setValue("U_BiGwa06", iRow, SUMArr(20))
    
    Call oDS_ZPY343L.setValue("U_JiGTotal", iRow, SUMArr(9))
    Call oDS_ZPY343L.setValue("U_KukAmt", iRow, SUMArr(10))
    Call oDS_ZPY343L.setValue("U_MedAmt", iRow, SUMArr(11))
    Call oDS_ZPY343L.setValue("U_GBHAmt", iRow, SUMArr(12))
    Call oDS_ZPY343L.setValue("U_GabGun", iRow, SUMArr(13))
    Call oDS_ZPY343L.setValue("U_Jumin", iRow, SUMArr(14))
    Call oDS_ZPY343L.setValue("U_NonTK", iRow, SUMArr(15))
    'Call oDS_ZPY343L.SetValue("U_GamSe", iRow, SUMArr(16))
    Call oDS_ZPY343L.setValue("U_GbuAmt", iRow, SUMArr(17))
    '/ 2009년추가
    Call oDS_ZPY343L.setValue("U_BIGG01", iRow, SUMArr(21))
    Call oDS_ZPY343L.setValue("U_BIGH01", iRow, SUMArr(22))
    Call oDS_ZPY343L.setValue("U_BIGH05", iRow, SUMArr(23))
    Call oDS_ZPY343L.setValue("U_BIGH06", iRow, SUMArr(24))
    Call oDS_ZPY343L.setValue("U_BIGH07", iRow, SUMArr(25))
    Call oDS_ZPY343L.setValue("U_BIGH08", iRow, SUMArr(26))
    Call oDS_ZPY343L.setValue("U_BIGH09", iRow, SUMArr(27))
    Call oDS_ZPY343L.setValue("U_BIGH10", iRow, SUMArr(28))
    Call oDS_ZPY343L.setValue("U_BIGH11", iRow, SUMArr(29))
    Call oDS_ZPY343L.setValue("U_BIGH12", iRow, SUMArr(30))
    Call oDS_ZPY343L.setValue("U_BIGH13", iRow, SUMArr(31))
    Call oDS_ZPY343L.setValue("U_BIGI01", iRow, SUMArr(32))
    Call oDS_ZPY343L.setValue("U_BIGK01", iRow, SUMArr(33))
    Call oDS_ZPY343L.setValue("U_BIGM01", iRow, SUMArr(34))
    Call oDS_ZPY343L.setValue("U_BIGM02", iRow, SUMArr(35))
    Call oDS_ZPY343L.setValue("U_BIGM03", iRow, SUMArr(36))
    Call oDS_ZPY343L.setValue("U_BIGO01", iRow, SUMArr(37))
    Call oDS_ZPY343L.setValue("U_BIGQ01", iRow, SUMArr(38))
    Call oDS_ZPY343L.setValue("U_BIGS01", iRow, SUMArr(39))
    Call oDS_ZPY343L.setValue("U_BIGT01", iRow, SUMArr(40))
    Call oDS_ZPY343L.setValue("U_BIGX01", iRow, SUMArr(41))
    Call oDS_ZPY343L.setValue("U_BIGY01", iRow, SUMArr(42))
    Call oDS_ZPY343L.setValue("U_BIGY02", iRow, SUMArr(43))
    Call oDS_ZPY343L.setValue("U_BIGY03", iRow, SUMArr(44))
    Call oDS_ZPY343L.setValue("U_BIGY20", iRow, SUMArr(45))
    Call oDS_ZPY343L.setValue("U_BIGZ01", iRow, SUMArr(46))
    
    Call oDS_ZPY343L.setValue("U_URIBNS", iRow, SUMArr(47))
    Call oDS_ZPY343L.setValue("U_BiGwu03", iRow, SUMArr(48))
    Call oDS_ZPY343L.setValue("U_BiGwa07", iRow, SUMArr(49))
    Call oDS_ZPY343L.setValue("U_NGYAMT", iRow, SUMArr(50))
    
    Call oDS_ZPY343L.setValue("U_BIGY21", iRow, SUMArr(51))
    
    oMat1.SetLineData iRow + 1
    oForm.Update
End Sub

Private Sub FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim optBtn      As SAPbouiCOM.OptionBtn
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    If oForm.Mode = fm_FIND_MODE Then
        oForm.Items("DocNum").Enabled = True
        oForm.Items("JsnYear").Enabled = True
        oForm.Items("MstCode").Enabled = True
        oForm.Items("MstName").Enabled = True
        oForm.Items("CLTCOD").Enabled = True
        If Len(ZPAY_GBL_JSNYER) > 0 Then
            oForm.Items("JsnYear").Specific.VALUE = ZPAY_GBL_JSNYER
        End If
        
         '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '//기본사항 - 부서 (사업장에 따른 부서변경)
        Set oCombo = oForm.Items("DptCode").Specific

        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", ""
            oCombo.Select 0, psk_Index
        End If
        
        If oDS_ZPY343H.GetValue("U_CLTCOD", 0) <> "" Then
            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
            sQry = sQry & " ORDER BY U_Code"
            Call SetReDataCombo(oForm, sQry, oCombo)
        End If
        
    ElseIf oForm.Mode = fm_ADD_MODE Then
        oForm.Items("DocNum").Enabled = True
        oForm.Items("JsnYear").Enabled = True
        oForm.Items("MstCode").Enabled = True
        oForm.Items("MstName").Enabled = False
        oForm.Items("CLTCOD").Enabled = True
        If Len(ZPAY_GBL_JSNYER) > 0 Then
            oForm.Items("JsnYear").Specific.VALUE = ZPAY_GBL_JSNYER
        End If
        
         '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '//부서
        Set oCombo = oForm.Items("TeamCode").Specific
        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", "-"
            oCombo.Select 0, psk_Index
        End If
        
    ElseIf oForm.Mode = fm_OK_MODE Then
        oForm.Items("DocNum").Enabled = False
        oForm.Items("JsnYear").Enabled = False
        oForm.Items("MstCode").Enabled = False
        oForm.Items("MstName").Enabled = False
        oForm.Items("CLTCOD").Enabled = False
        
         '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
    End If
End Sub

Private Function MatrixSpaceLineDel() As Boolean
'----------------------------------------------------------------------------
'저장할 데이터의 유효성을 점검한다
'----------------------------------------------------------------------------
On Error GoTo Error_Message
    Dim i           As Long
    Dim k           As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    
    ErrNum = 0
    '/ 헤더부분 체크
    Select Case True
      Case oDS_ZPY343H.GetValue("U_JsnYear", 0) = ""
           ErrNum = 4
           GoTo Error_Message
      Case oDS_ZPY343H.GetValue("U_MstCode", 0) = ""
           ErrNum = 5
           GoTo Error_Message
      Case oDS_ZPY343H.GetValue("U_CLTCOD", 0) = ""
           ErrNum = 7
           GoTo Error_Message
    End Select
    
    '----------------------------------------------------------------------------
    '화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
    '----------------------------------------------------------------------------
    oMat1.FlushToDataSource
        
    If oMat1.RowCount = 1 Then          '// Mat1에 값이 있는지 확인 (ErrorNumber : 1)
        ErrNum = 1
        GoTo Error_Message
    End If
    
    '----------------------------------------------------------------------------
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat1.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    '----------------------------------------------------------------------------
    For i = 0 To oMat1.VisualRowCount - 2      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 3)
        oDS_ZPY343L.Offset = i
        If Val(oDS_ZPY343L.GetValue("U_JIGTotal", i)) <> 0 And Trim$(oDS_ZPY343L.GetValue("U_JIGDate", i)) = "" Then
            ErrNum = 2
            oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY343L.GetValue("U_JIGDate", i)) <> "" And IsDate(Format$(oDS_ZPY343L.GetValue("U_JIGDate", i), "0000-00-00")) = False Then
            ErrNum = 6
            oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular
            GoTo Error_Message
        Else
            '------------------------------------------------
            '중복체크작업
            '------------------------------------------------
            Chk_Data = Trim$(oDS_ZPY343L.GetValue("U_JIGDate", i))
            For k = i + 1 To oMat1.VisualRowCount - 2
                oDS_ZPY343L.Offset = k
                If Trim$(Chk_Data) = Trim$(oDS_ZPY343L.GetValue("U_JIGDate", k)) And Val(oDS_ZPY343L.GetValue("U_JIGTotal", i)) <> 0 Then
                    ErrNum = 3
                    oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular
                    GoTo Error_Message
                End If
            Next k
        End If
    Next i
    
    '--------------------------------------------------------------------------------------------
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    '--------------------------------------------------------------------------------------------
  '  oDS_ZPY343L.RemoveRecord oDS_ZPY343L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
    
    '--------------------------------------------------------------------------------------------
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    '--------------------------------------------------------------------------------------------
    oMat1.LoadFromDataSource
    
    MatrixSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "입력할 데이터가 없습니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "지급일자가 없습니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "지급일자가 중복입력되었습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "정산년도는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "사원번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "지급일자를 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "자사코드는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    End If
    MatrixSpaceLineDel = False
End Function
Private Sub ShowSource(DocNum$)
    oForm.Items("DocNum").Enabled = True
    oForm.Items("DocNum").Specific.VALUE = DocNum
    
    oForm.Items("1").CLICK ct_Regular
End Sub
Private Sub Clear_UserDb()
    oForm.DataSources.UserDataSources("User").VALUE = ""
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Matrix_TitleSetting
' DateTime  : 2009-12-29 11:08
' Author    : Choi Dong Kwon
' Purpose   : 비과세 코드 설정의 데이터를 읽어와 Matrix Title의 비과세 컬럼에 대하여 표시여부와 타이틀을 적용한다
'             단, 찾기모드에서는 전체컬럼을 표시한다
'---------------------------------------------------------------------------------------
'
Private Sub Matrix_TitleSetting()
On Error GoTo Error_Message

    Dim JSNYER      As String
    Dim sQry        As String
    Dim iCol        As Long
    Dim COLNAM      As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    JSNYER = Trim$(oDS_ZPY343H.GetValue("U_JsnYear", 0))
    
    If (IsNumeric(JSNYER) = False Or JSNYER = "") = False Then
        If Val(JSNYER) <= 2008 Then             '// 2008년 이전
            sQry = "SELECT  U_BTXCOD, U_BTXNAM, ISNULL(U_MONCHK,'Y') AS U_MONCHK " & _
                   "FROM    [@ZPY117L] T0 " & _
                   "WHERE   T0.CODE = (SELECT MAX(CODE) FROM [@ZPY117L] T1 WHERE CODE <= '" & JSNYER & "') " & _
                   "AND     T0.CODE <= '2008' " & _
                   "ORDER   BY U_BTXCOD "
        Else                                    '// 2009년 이후
            sQry = "SELECT  U_BTXCOD, U_BTXNAM, ISNULL(U_MONCHK,'Y') AS U_MONCHK  " & _
                   "FROM    [@ZPY117L] T0 " & _
                   "WHERE   T0.CODE = (SELECT MAX(CODE) FROM [@ZPY117L] T1 WHERE CODE <= '" & JSNYER & "') " & _
                   "AND     T0.CODE >= '2009' " & _
                   "ORDER   BY U_BTXCOD "
        End If
        oRecordSet.DoQuery sQry
    End If
    
    oForm.Freeze True
    If oForm.Mode = fm_FIND_MODE Or IsNumeric(JSNYER) = False Or JSNYER = "" Or oRecordSet.RecordCount = 0 Then
        oMat1.Columns("Col6").TitleObject.Caption = "비과세-생산"
        oMat1.Columns("Col7").TitleObject.Caption = "비과세-식대.차량"
        oMat1.Columns("Col8").TitleObject.Caption = "비과세-국외"
        oMat1.Columns("Col23").TitleObject.Caption = "비과세-외국인"
        oMat1.Columns("Col21").TitleObject.Caption = "비과세-보육"
        oMat1.Columns("Col20").TitleObject.Caption = "비과세-연구"
        oMat1.Columns("Col9").TitleObject.Caption = "비과세-기타(제출)"
        oMat1.Columns("Col24").TitleObject.Caption = "비과세-기타(미제출)"
        
        oMat1.Columns("Col26").TitleObject.Caption = "비과세(G01)"
        oMat1.Columns("Col27").TitleObject.Caption = "비과세(H01)"
        oMat1.Columns("Col28").TitleObject.Caption = "비과세(H05)"
        oMat1.Columns("Col29").TitleObject.Caption = "비과세(H06)"
        oMat1.Columns("Col30").TitleObject.Caption = "비과세(H07)"
        oMat1.Columns("Col31").TitleObject.Caption = "비과세(H08)"
        oMat1.Columns("Col32").TitleObject.Caption = "비과세(H09)"
        oMat1.Columns("Col33").TitleObject.Caption = "비과세(H10)"
        oMat1.Columns("Col34").TitleObject.Caption = "비과세(H11)"
        oMat1.Columns("Col35").TitleObject.Caption = "비과세(H12)"
        oMat1.Columns("Col36").TitleObject.Caption = "비과세(H13)"
        oMat1.Columns("Col37").TitleObject.Caption = "비과세(I01)"
        oMat1.Columns("Col38").TitleObject.Caption = "비과세(K01)"
        oMat1.Columns("Col39").TitleObject.Caption = "비과세(M01)"
        oMat1.Columns("Col40").TitleObject.Caption = "비과세(M02)"
        oMat1.Columns("Col41").TitleObject.Caption = "비과세(M03)"
        oMat1.Columns("Col42").TitleObject.Caption = "비과세(O01)"
        oMat1.Columns("Col43").TitleObject.Caption = "비과세(Q01)"
        oMat1.Columns("Col44").TitleObject.Caption = "비과세(S01)"
        oMat1.Columns("Col45").TitleObject.Caption = "비과세(T01)"
        oMat1.Columns("Col46").TitleObject.Caption = "비과세(X01)"
        oMat1.Columns("Col47").TitleObject.Caption = "비과세(R10,Y01)"
        oMat1.Columns("Col48").TitleObject.Caption = "비과세(Y02)"
        oMat1.Columns("Col49").TitleObject.Caption = "비과세(Y03)"
        oMat1.Columns("Col50").TitleObject.Caption = "비과세(Y22,Y20)"
        oMat1.Columns("Col51").TitleObject.Caption = "비과세(Z01)"
        oMat1.Columns("Col52").TitleObject.Caption = "비과세(Y21)"
        
        '// 비과세 컬럼 전체 표시
        For iCol = 6 To 52
            If (iCol >= 6 And iCol <= 9) Or iCol = 20 Or iCol = 21 Or iCol = 23 Or iCol = 24 Or _
               (iCol >= 26 And iCol <= 52) Then
                oMat1.Columns("Col" & CStr(iCol)).Visible = False
            End If
        Next iCol
    
    ElseIf Val(JSNYER) <= 2008 Then
        Do Until oRecordSet.EOF
            '// 비과세 코드에 따라 컬럼UID 확인
            Select Case oRecordSet.Fields("U_BTXCOD").VALUE
                Case "01":      COLNAM = "Col6"     '/ 비과세-생산
                Case "02":      COLNAM = "Col7"     '/ 비과세-식대.차량
                Case "03":      COLNAM = "Col8"     '/ 비과세-국외
                Case "04":      COLNAM = "Col23"    '/ 비과세-외국인
                Case "05":      COLNAM = "Col9"     '/ 비과세-기타(제출)
                Case "06":      COLNAM = "Col20"    '/ 비과세-연구
                Case "07":      COLNAM = "Col21"    '/ 비과세-보육
                Case "08":      COLNAM = "Col24"    '/ 비과세-기타(미제출)
            End Select
            
            '// 컬럼명, 화면표시 여부 적용
            With oMat1.Columns(COLNAM)
                .TitleObject.Caption = oRecordSet.Fields("U_BTXNAM").VALUE
                If oRecordSet.Fields("U_MONCHK").VALUE = "N" Then
                    .Visible = False
                Else
                    .Visible = True
                End If
            End With
            
            oRecordSet.MoveNext
        Loop
        
        '// 2009년 이후 생성 컬럼은 화면 표시 안함
        For iCol = 26 To 52
            oMat1.Columns("Col" & CStr(iCol)).Visible = False
        Next iCol
    Else
        Do Until oRecordSet.EOF
            '// 비과세 코드에 따라 컬럼UID 확인
            Select Case oRecordSet.Fields("U_BTXCOD").VALUE
                Case "G01":     COLNAM = "Col26"    '// 비과세(G01)
                Case "H01":     COLNAM = "Col27"    '// 비과세(H01)
                Case "H05":     COLNAM = "Col28"    '// 비과세(H05)
                Case "H06":     COLNAM = "Col29"    '// 비과세(H06)
                Case "H07":     COLNAM = "Col30"    '// 비과세(H07)
                Case "H08":     COLNAM = "Col31"    '// 비과세(H08)
                Case "H09":     COLNAM = "Col32"    '// 비과세(H09)
                Case "H10":     COLNAM = "Col33"    '// 비과세(H10)
                Case "H11":     COLNAM = "Col34"    '// 비과세(H11)
                Case "H12":     COLNAM = "Col35"    '// 비과세(H12)
                Case "H13":     COLNAM = "Col36"    '// 비과세(H13)
                Case "I01":     COLNAM = "Col37"    '// 비과세(I01)
                Case "K01":     COLNAM = "Col38"    '// 비과세(K01)
                Case "M01":     COLNAM = "Col39"    '// 비과세(M01)
                Case "M02":     COLNAM = "Col40"    '// 비과세(M02)
                Case "M03":     COLNAM = "Col41"    '// 비과세(M03)
                Case "O01":     COLNAM = "Col42"    '// 비과세(O01)
                Case "Q01":     COLNAM = "Col43"    '// 비과세(Q01)
                Case "S01":     COLNAM = "Col44"    '// 비과세(S01)
                Case "T01":     COLNAM = "Col45"    '// 비과세(T01)
                Case "X01":     COLNAM = "Col46"    '// 비과세(X01)
                Case "Y01", "R10":
                                COLNAM = "Col47"    '// 비과세(Y01)
                Case "Y02":     COLNAM = "Col48"    '// 비과세(Y02)
                Case "Y03":     COLNAM = "Col49"    '// 비과세(Y03)
                Case "Y20", "Y22":
                                COLNAM = "Col50"    '// 비과세(Y20)
                Case "Y21":     COLNAM = "Col52"    '// 비과세(Y21)
                Case "Z01":     COLNAM = "Col51"    '// 비과세(Z01)
            End Select
            
            '// 컬럼명, 화면표시 여부 적용
            With oMat1.Columns(COLNAM)
                .TitleObject.Caption = oRecordSet.Fields("U_BTXNAM").VALUE
                If oRecordSet.Fields("U_MONCHK").VALUE = "Y" Then
                    .Visible = True
                Else
                    .Visible = False
                End If
            End With
            
            oRecordSet.MoveNext
        Loop
        oMat1.Columns("Col6").Visible = False       '비과세-생산
        oMat1.Columns("Col7").Visible = True        '비과세-식대차량
        oMat1.Columns("Col8").Visible = False       '비과세-국외
        oMat1.Columns("Col23").Visible = False      '비과세-외국인
        oMat1.Columns("Col21").Visible = False      '비과세-보육
        oMat1.Columns("Col20").Visible = False      '비과세-연구
        oMat1.Columns("Col9").Visible = True        '비과세-기타(제출)
        oMat1.Columns("Col24").Visible = True       '비과세-기타(미제출)
        
    End If
    
    oForm.Freeze False
    Set oRecordSet = Nothing
    Exit Sub
Error_Message:
    oForm.Freeze False
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "Matrix_TitleSetting 실행 중 오류가 발생하였습니다.", bmt_Short, smt_Error
End Sub


