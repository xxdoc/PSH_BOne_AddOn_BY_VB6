VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY501"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY501.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 소득공제항목 등록
'//  FormType       : 2010110501
'//  Create Date    : 2006.01.19
'//  Modified Date  :
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oDS_ZPY501H     As SAPbouiCOM.DBDataSource    '시스템코드 헤더
Private oDS_ZPY501L     As SAPbouiCOM.DBDataSource    '시스템코드 라인
Private oMat1           As SAPbouiCOM.Matrix
Private oCheck          As SAPbouiCOM.CheckBox
Private MsterChk        As Boolean
Private oOLDCHK         As String
Private BOHAL1          As Double   '/ 종전전액보험료
Private BOHAL2          As Double   '/ 주현전액보험료


Private Last_Item       As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private Col_Last_Uid    As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private Col_Last_Row    As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional JSNYER$, Optional MSTCOD$, Optional CLTCOD$)
'Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY501.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "ZPY501_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "ZPY501"
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
    oForm.Items("Folder1").CLICK ct_Regular
    
    oForm.EnableMenu ("1293"), True  '/ 행삭제
    oForm.EnableMenu ("1284"), False  '/ 취소
    
    If Trim$(JSNYER) <> "" Then
        Call ShowSource(JSNYER, MSTCOD, CLTCOD)
    End If
    
    oForm.Freeze False
    oForm.Update
    'oForm.Visible = True
    

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
'// RaiseDataEvent
'*******************************************************************
Public Sub RaiseDataEvent(ByVal pval As SAPbouiCOM.IBusinessObjectInfo, BubbleEvent As Boolean)
On Error GoTo RaiseDataEvent_Error
    
    If oForm Is Nothing Then
        Exit Sub
    End If
    
    Select Case pval.EventType
'et_FORM_DATA_ADD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_DATA_ADD
            If pval.BeforeAction = False Then
                If pval.ActionSuccess = True Then
                    If MsterChk = True Then
                        Call MasterUpdate
                    End If
                End If
            End If
'et_FORM_DATA_UPDATE ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_DATA_UPDATE
            If pval.BeforeAction = False Then
                If pval.ActionSuccess = True Then
                    If MsterChk = True Then
                        Call MasterUpdate
                    End If
                End If
            End If
'et_FORM_DATA_DELETE ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_DATA_DELETE
'et_FORM_DATA_LOAD ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_DATA_LOAD
    
    End Select
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
RaiseDataEvent_Error:
    Sbo_Application.StatusBar.SetText "RaiseDataEvent_Error:", bmt_Short, smt_Error
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
                        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                        '//추가및 업데이시에
                        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                        If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                            '//추가및 업데이시에
                            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                            If HeaderSpaceLineDel = True Then
                                If MatrixSpaceLineDel = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Else
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                ElseIf pval.ItemUID = "CBtn1" And oForm.Items("MSTCOD").Enabled = True Then   '/ ChooseBtn사원리스트
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUID = "Folder1" Then
                    oForm.PaneLevel = 1
                ElseIf pval.ItemUID = "Folder2" Then
                    oForm.PaneLevel = 2
                ElseIf pval.ItemUID = "Folder3" Then
                    oForm.PaneLevel = 3
                ElseIf pval.ItemUID = "Btn1" Then '/ 가족사항등록 창 띄우기
                    Call CallSource(0)
                End If
            Else
                If pval.ItemUID = "1" And pval.ActionSuccess = True And oForm.Mode = fm_ADD_MODE Then
                    oForm.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")
                    oForm.Items("Folder1").CLICK ct_Regular
                ElseIf pval.ItemUID = "1" And pval.ActionSuccess = True And oForm.Mode = fm_OK_MODE Then
                    Call FormItemEnabled
                ElseIf pval.ItemUID = "Btn2" Then '/ 기존부양가족 정보가져오기
                    If oForm.Items("JSNYER").Specific.VALUE = "" Then
'                    If oDS_ZPY501H.GetValue("U_JSNYER", 0) = "" Then
                        oForm.Items("JSNYER").CLICK ct_Regular
                        Sbo_Application.StatusBar.SetText "정산년도는 필수입니다. 확인하여 주십시오.", bmt_Short, smt_Error
                    ElseIf oDS_ZPY501H.GetValue("U_CLTCOD", 0) = "" Then
                        Sbo_Application.StatusBar.SetText "자사코드는 필수입니다. 확인하여 주십시오.", bmt_Short, smt_Error
                    Else
                        Call Family_Data_Display
                    End If
                ElseIf pval.ItemUID = "Btn3" Then   '/ 부양가족명세 공제항목에 적용
                    Call Family_Total("All")
                ElseIf pval.ItemUID = "Btn4" Then   '/ 연금.저축 명세 가져오기
                    Call SavingAmount_Data_Display
                End If
            End If
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
            If pval.BeforeAction = True And pval.ItemUID <> "2" And oForm.Mode <> fm_FIND_MODE Then
                If Last_Item = "MSTCOD" Then
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(Last_Item).Specific.String & "'", "") = True _
                    And oForm.Items(Last_Item).Specific.String <> "" And Last_Item <> pval.ItemUID Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
            If pval.FormUID = oForm.uniqueID And pval.BeforeAction = True And Last_Item = "Mat1" And Col_Last_Uid = "Col1" And Col_Last_Row > 0 _
              And (Col_Last_Uid <> pval.ColUID Or Col_Last_Row <> pval.Row) And pval.ItemUID <> "2" Then
                If Col_Last_Row > oMat1.VisualRowCount Then
                    Exit Sub
                End If
            ElseIf pval.BeforeAction = False And pval.ItemUID = "Mat1" And pval.Row > 0 And pval.ColUID = "Col0" Then
                oForm.DataSources.UserDataSources("FAMNAM").ValueEx = oDS_ZPY501L.GetValue("U_FamNam", pval.Row - 1)
                oForm.DataSources.UserDataSources("FAMPER").ValueEx = oDS_ZPY501L.GetValue("U_FamPer", pval.Row - 1)
                oForm.Items("FAMNAM").Update
                oForm.Items("FAMPER").Update
            ElseIf pval.FormUID = oForm.uniqueID And pval.BeforeAction = False And pval.ItemUID = "Mat1" And _
                (pval.ColUID = "Col5" Or pval.ColUID = "Col6" Or pval.ColUID = "Col7" Or _
                 pval.ColUID = "Col8" Or pval.ColUID = "Col9" Or pval.ColUID = "Col10") And pval.Row > 0 Then
                FlushToItemValue pval.ColUID, pval.Row
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True And _
            (pval.ItemUID = "MSTCOD" Or pval.ItemUID = "JSNYER" Or pval.ItemUID = "BAEWOO" Or _
            pval.ItemUID = "BUYN20" Or pval.ItemUID = "BUYN60" Or pval.ItemUID = "GYNGLO" Or _
            pval.ItemUID = "GYNGL2" Or pval.ItemUID = "JANGAE" Or pval.ItemUID = "MZBURI" Or _
            pval.ItemUID = "BUYN06" Or pval.ItemUID = "BONSCH" Or pval.ItemUID = "CSHSAV" Or _
            pval.ItemUID = "LAWGBU" Or pval.ItemUID = "POCGBU" Or pval.ItemUID = "SP1GBU" Or _
            pval.ItemUID = "SP2GBU" Or pval.ItemUID = "USJGBU" Or pval.ItemUID = "JIJGBU" Or _
            pval.ItemUID = "BONCSH1" Or pval.ItemUID = "BONEDC1") Then
                FlushToItemValue pval.ItemUID
            ElseIf pval.BeforeAction = False And pval.ItemChanged = True And pval.ItemUID = "Mat1" _
            And (pval.ColUID = "Col1" Or pval.ColUID = "Col2" Or pval.ColUID = "Col5" Or pval.ColUID = "Col6" Or _
                pval.ColUID = "Col7" Or pval.ColUID = "Col8" Or pval.ColUID = "Col9" Or pval.ColUID = "Col23" Or _
                pval.ColUID = "Col10") Then
                FlushToItemValue pval.ColUID, pval.Row
            End If
'et_COMBO_SELECT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_COMBO_SELECT
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If pval.ItemUID = "MZBURI" Or pval.ItemUID = "BAEWOO" Or pval.ItemUID = "HUSMAN" Then
                    MsterChk = True
                ElseIf pval.ItemUID = "Mat1" And (pval.ColUID = "Col3" Or pval.ColUID = "Col4") And pval.Row > 0 Then
                    FlushToItemValue pval.ColUID, pval.Row
                End If
            End If
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '//추가모드에서 코드이벤트가 코드에서 일어 났을때
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
             If pval.BeforeAction = True And pval.ItemUID = "MSTCOD" And pval.CharPressed = 9 And pval.FormMode <> fm_FIND_MODE Then
                If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(pval.ItemUID).Specific.String & "'", "") = True Then
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                Else
'                    If oMat1.RowCount > 0 Then
'                        oMat1.Columns("Col1").Cells(oMat1.VisualRowCount).Click ct_Regular
'                        BubbleEvent = False
'                    End If
                End If
             ElseIf pval.BeforeAction = True And pval.ItemUID = "JSNYER" And pval.CharPressed = 9 And pval.FormMode <> fm_FIND_MODE Then
                If Trim$(oForm.Items(pval.ItemUID).Specific.String) = "" Then
                    oForm.Items("JSNYER").CLICK ct_Regular
                    Sbo_Application.StatusBar.SetText "정산년도는 필수입니다. 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
             End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
            If Last_Item = "Mat1" Then
                If pval.Row > 0 Then
                    Last_Item = pval.ItemUID
                    Col_Last_Row = pval.Row
                    Col_Last_Uid = pval.ColUID
                End If
            Else
                Last_Item = pval.ItemUID
                Col_Last_Row = 0
                Col_Last_Uid = ""
            End If
'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '컬렉션에서 삭제및 모든 메모리 제거
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            If pval.BeforeAction = False Then
               RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_ZPY501H = Nothing
                Set oDS_ZPY501L = Nothing
                Set oCheck = Nothing
                Set oMat1 = Nothing
            End If
'et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_MATRIX_LOAD:
            If pval.BeforeAction = False Then
                FormItemEnabled
                Call Matrix_AddRow(oMat1.VisualRowCount)
            End If
'et_FORM_RESIZE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_RESIZE
            If pval.BeforeAction = False Then
                oForm.Freeze True
                'oForm.Items("Reg1").Width = oForm.Width - 26
                'oForm.Items("Reg1").Height = oForm.Height - 358
                oMat1.AutoResizeColumns
                oForm.Freeze False
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
    Dim i&
    
    
    If pval.BeforeAction = True Then
        Select Case pval.MenuUID
        Case "1283"
            If Trim$(oDS_ZPY501H.GetValue("U_Check", 0)) = "Y" Then
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
            Call AuthorityCheck(oForm, "CLTCOD", "@ZPY501H", "DocNum")      '//접속자 권한에 따른 사업장 보기
            
        Case Else
                Exit Sub
        End Select
    Else
        Select Case pval.MenuUID
            Case "1287" '/ 복제
               ' oForm.Items("Btn1").Visible = True
            Case "1283" '/ 제거
                 FormItemEnabled
            Case "1281", "1282":
                FormItemEnabled
                If pval.MenuUID = "1282" Then
                    FormClear
                    Call Matrix_AddRow(0, True)
                    oForm.Items("JSNYER").CLICK ct_Regular
                End If
            Case "1288" To "1291":
                FormItemEnabled
            Case "1293"
                If oMat1.RowCount <> oMat1.VisualRowCount Then
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
                    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    For i = 0 To oMat1.VisualRowCount - 1
                        oMat1.Columns("Col0").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    oMat1.FlushToDataSource
                    oDS_ZPY501L.RemoveRecord oDS_ZPY501L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
                    oMat1.Clear
                    oMat1.LoadFromDataSource
                    Call Family_Total("All")
                End If
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

Private Sub CallSource(oRow&)
On Error GoTo Error_Message
    Dim oTmpObject  As Variant
    
'    If oForm.Mode <> fm_FIND_MODE Then
'        Set oTmpObject = New ZPY121     'ZPY121(2010110121) : 가족사항등록
'        Call oTmpObject.LoadForm(oForm.uniqueID, oForm.Items("MSTCOD").Specific.String, oForm.Items("JSNYER").Specific.String)
'        Sbo_Application.Forms.ActiveForm.Select
'    End If
        If oForm.Mode <> fm_FIND_MODE Then
        Set oTmpObject = New PH_PY001
        Call oTmpObject.LoadForm
        Sbo_Application.Forms.ActiveForm.Select
    End If
    Set oTmpObject = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oTmpObject = Nothing
    Sbo_Application.StatusBar.SetText "CallSource Error:" & Err.Description, bmt_Short, smt_Error
End Sub
'*******************************************************************
'
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oEdit           As SAPbouiCOM.EditText
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim oFolder         As SAPbouiCOM.Folder
    Dim oColumn         As SAPbouiCOM.Column
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//디비데이터 소스 개체 할당
    Set oDS_ZPY501H = oForm.DataSources.DBDataSources("@ZPY501H")
    Set oDS_ZPY501L = oForm.DataSources.DBDataSources("@ZPY501L")
    '/ UserDataSource
    Call oForm.DataSources.UserDataSources.Add("FAMNAM", dt_SHORT_TEXT, 30)
    Call oForm.DataSources.UserDataSources.Add("FAMPER", dt_SHORT_TEXT, 40)
    Set oEdit = oForm.Items("FAMNAM").Specific
    oEdit.DataBind.SetBound True, "", "FAMNAM"
    Set oEdit = oForm.Items("FAMPER").Specific
    oEdit.DataBind.SetBound True, "", "FAMPER"
    
    oForm.PaneLevel = 1
    
    Set oMat1 = oForm.Items("Mat1").Specific
    
  '/ Folder /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    oForm.DataSources.UserDataSources.Add "FolderDS", dt_SHORT_TEXT, 1
'    For i = 1 To 2
'        Set oFolder = oForm.Items("Folder" & i).Specific
'        oFolder.DataBind.SetBound True, "", "FolderDS"
'        If i = 1 Then
'            oFolder.Select
'        Else
'            oFolder.GroupWith ("Folder" & i - 1)
'        End If
'    Next i
    oForm.Items("Folder1").AffectsFormMode = False
    oForm.Items("Folder2").AffectsFormMode = False
    oForm.Items("Folder3").AffectsFormMode = False
    oForm.Items("Folder1").Enabled = True
    oForm.Items("Folder2").Enabled = True
    oForm.Items("Folder3").Enabled = True
  
  '/ Check 버튼
    Set oCheck = oForm.Items("Check1").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
'    Set oCheck = oForm.Items("BONJAN").Specific
'    oCheck.ValOn = "1": oCheck.ValOff = "0"

    Set oCombo = oForm.Items("BAEWOO").Specific
    oCombo.ValidValues.Add "1", "YES"
    oCombo.ValidValues.Add "0", "NO"
    Set oCombo = oForm.Items("MZBURI").Specific
    oCombo.ValidValues.Add "1", "YES"
    oCombo.ValidValues.Add "0", "NO"
    Set oCombo = oForm.Items("HUSMAN").Specific
    oCombo.ValidValues.Add "1", "세대주"
    oCombo.ValidValues.Add "2", "세대원"
    
    '// 사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT Code, Name FROM [@PH_PY005A]"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True


    '// 교육대상구분
    Set oColumn = oMat1.Columns("Col16")
    oColumn.DisplayDesc = True
    oColumn.ValidValues.Add "0", "본인"
    oColumn.ValidValues.Add "1", "영유아"
    oColumn.ValidValues.Add "2", "초중고"
    oColumn.ValidValues.Add "3", "대학생"
    oColumn.ValidValues.Add "4", "대학원생"
    oColumn.ValidValues.Add "5", "특수교육"
    '// 관계코드
    Set oColumn = oMat1.Columns("Col3")
    oColumn.ValidValues.Add "0", "본인"
    oColumn.ValidValues.Add "1", "소득자의 직계존속"
    oColumn.ValidValues.Add "2", "배우자의 직계존속"
    oColumn.ValidValues.Add "3", "배우자"
    oColumn.ValidValues.Add "4", "직계비속 자녀.입양자"
'    oColumn.ValidValues.Add "5", "형제자매"
'    oColumn.ValidValues.Add "6", "기타"
    '/ (2010.01.19) 2009년귀속 관계코드변경
    oColumn.ValidValues.Add "5", "직계비속 자녀.입양자외"
    oColumn.ValidValues.Add "6", "형제자매"
    'oColumn.ValidValues.Add "7", "기타"
    '/ (2010.01.04) 2010년귀속 관계코드변경
    oColumn.ValidValues.Add "7", "수급자"
    oColumn.ValidValues.Add "8", "위탁아동"

    '// 내외국인
    Set oColumn = oMat1.Columns("Col4")
    oColumn.ValidValues.Add "1", "내국인"
    oColumn.ValidValues.Add "9", "외국인"
    '// 보험료구분
    Set oColumn = oMat1.Columns("Col11")
    oColumn.ValidValues.Add "1", "일반인보험"
    oColumn.ValidValues.Add "2", "장애인보험"
    '// 부양가족 인적공제내용
    Set oColumn = oMat1.Columns("Col5") '/ 기본
    oColumn.ValOn = "1": oColumn.ValOff = "0"
    Set oColumn = oMat1.Columns("Col6") '/ 부녀자
    oColumn.ValOn = "1": oColumn.ValOff = "0"
    Set oColumn = oMat1.Columns("Col7") '/ 장애인
    oColumn.ValOn = "1": oColumn.ValOff = "0"
    Set oColumn = oMat1.Columns("Col8") '/ 경로우대
    oColumn.ValOn = "1": oColumn.ValOff = "0"
    Set oColumn = oMat1.Columns("Col9") '/ 양육
    oColumn.ValOn = "1": oColumn.ValOff = "0"
    Set oColumn = oMat1.Columns("Col10") '/ 다자녀
    oColumn.ValOn = "1": oColumn.ValOff = "0"
    oColumn.Editable = False
    Set oColumn = oMat1.Columns("Col23") '/ 출산입양
    oColumn.ValOn = "1": oColumn.ValOff = "0"
    
    Set oEdit = Nothing
    Set oFolder = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oFolder = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error:" & Err.Description, bmt_Short, smt_Error
End Sub

Private Function Exist_YN(JOBYER$, MSTCOD$, CLTCOD$) As String
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'저장할 데이터의 기존데이터가 있는지 확인한다
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "SELECT Top 1 T1.DocNum FROM [@ZPY501H] T1 "
    sQry = sQry & " WHERE T1.U_JSNYER = N'" & Trim$(JOBYER) & "'"
    sQry = sQry & "   AND T1.U_MSTCOD = N'" & Trim$(MSTCOD) & "'"
    sQry = sQry & "   AND T1.U_CLTCOD = N'" & Trim$(CLTCOD) & "'"
    oRecordSet.DoQuery sQry
    
    Do Until oRecordSet.EOF
        Exist_YN = oRecordSet(0).VALUE
        oRecordSet.MoveNext
    Loop
    
    If Trim$(Exist_YN) = "" Then
        Exist_YN = ""
        Exit Function
    End If
    
    Set oRecordSet = Nothing
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim oMast           As ZPAY_g_EmpID
    Dim INTGBN$
    
    Select Case oUID
      Case "JSNYER"
            If Trim$(oForm.Items(oUID).Specific.String) <> "" Then
                ZPAY_GBL_JSNYER = oForm.Items(oUID).Specific.String
            Else
                oDS_ZPY501H.setValue "U_JSNYER", 0, ZPAY_GBL_JSNYER
            End If
            oForm.Items(oUID).Update
      Case "MSTCOD"
            If oForm.Items(oUID).Specific.String = "" Then
                oDS_ZPY501H.setValue "U_MSTCOD", 0, ""
                oDS_ZPY501H.setValue "U_MSTNAM", 0, ""
                oDS_ZPY501H.setValue "U_EmpID", 0, ""
                oDS_ZPY501H.setValue "U_CLTCOD", 0, ""
            Else
                oDS_ZPY501H.setValue "U_MSTCOD", 0, UCase$(oForm.Items(oUID).Specific.String)
                oMast = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oDS_ZPY501H.setValue "U_MSTNAM", 0, oMast.MSTNAM
                oDS_ZPY501H.setValue "U_EmpID", 0, oMast.EmpID
                oDS_ZPY501H.setValue "U_CLTCOD", 0, oMast.CLTCOD
            End If
         '/ 추가모드일경우 인적공제 뿌려줌
            If oForm.Mode = fm_ADD_MODE Then
            '/ 사원마스터의 부양가족수
                Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
                sQry = "SELECT U_BAEWOO, U_BUYNSU, U_BUYN20, U_BUYN60, U_GYNGLO, U_GYNGL2, U_JANGAE, U_MZBURI, U_BUYN06, U_DAGYSU, U_INTGBN "
                sQry = sQry & " FROM [@PH_PY001A]  WHERE Code = N'" & oDS_ZPY501H.GetValue("U_MSTCOD", 0) & "'"
                oRecordSet.DoQuery sQry
                If oRecordSet.RecordCount > 0 Then
                    oDS_ZPY501H.setValue "U_BAEWOO", 0, oRecordSet.Fields("U_BAEWOO").VALUE
                    oDS_ZPY501H.setValue "U_BUYN20", 0, oRecordSet.Fields("U_BUYN20").VALUE
                    oDS_ZPY501H.setValue "U_BUYN60", 0, oRecordSet.Fields("U_BUYN60").VALUE
                    oDS_ZPY501H.setValue "U_GYNGLO", 0, oRecordSet.Fields("U_GYNGLO").VALUE
                    oDS_ZPY501H.setValue "U_GYNGL2", 0, oRecordSet.Fields("U_GYNGL2").VALUE
                    oDS_ZPY501H.setValue "U_JANGAE", 0, oRecordSet.Fields("U_JANGAE").VALUE
                    oDS_ZPY501H.setValue "U_MZBURI", 0, oRecordSet.Fields("U_MZBURI").VALUE
                    oDS_ZPY501H.setValue "U_BUYN06", 0, oRecordSet.Fields("U_BUYN06").VALUE
                    oDS_ZPY501H.setValue "U_DAGYSU", 0, oRecordSet.Fields("U_DAGYSU").VALUE
                    oDS_ZPY501H.setValue "U_CHLSAN", 0, 0
                    oForm.Items("BAEWOO").Update
                    oForm.Items("BUYN20").Update
                    oForm.Items("BUYN60").Update
                    oForm.Items("GYNGLO").Update
                    oForm.Items("GYNGL2").Update
                    oForm.Items("JANGAE").Update
                    oForm.Items("MZBURI").Update
                    oForm.Items("BUYN06").Update
                    oForm.Items("DAGYSU").Update
                    oForm.Items("CHLSAN").Update
                    INTGBN = oRecordSet.Fields("U_INTGBN").VALUE
                    MsterChk = False
                End If
            End If
            oForm.Items("MSTNAM").Update
            oForm.Items("EmpID").Update
            oForm.Items(oUID).Update
        Case "BAEWOO", "BUYN20", "GYNGLO", "GYNGL2", "JANGAE", "MZBURI", "BUYN06", "HUSMAN"
            MsterChk = True
    End Select
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If Left$(oUID, 3) = "Col" Then
        oMat1.FlushToDataSource
        oDS_ZPY501L.Offset = oRow - 1
        Select Case oUID
        Case "Col1", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7", "Col8", "Col9", "Col10", "Col23"
             MsterChk = True
        End Select
        oDS_ZPY501L.Offset = oRow - 1
     '/
        If oRow = oMat1.VisualRowCount And Trim$(oDS_ZPY501L.GetValue("U_FamNam", oRow - 1)) <> "" Then
            Matrix_AddRow oRow
            oMat1.Columns("Col1").Cells(oRow).CLICK ct_Regular
        End If
    End If
End Sub

Private Sub FormClear()
    Dim DocNum&
    
    DocNum = MDC_SetMod.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'ZPY501'", "")
    
    If DocNum = 0 Then
        oForm.Items("DocNum").Specific.String = 1
    Else
        oForm.Items("DocNum").Specific.String = DocNum
    End If
    Call FlushToItemValue("JSNYER")
    
End Sub

Private Sub FormItemEnabled()
    If oForm.Mode = fm_FIND_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = True
        oForm.Items("DocNum").Enabled = True
        oForm.Items("CLTCOD").Enabled = True
        If Len(ZPAY_GBL_JSNYER) > 0 Then
            oForm.Items("JSNYER").Specific.VALUE = ZPAY_GBL_JSNYER
        End If
        Call CLTCOD_Select(oForm, "CLTCOD")
        
    ElseIf oForm.Mode = fm_ADD_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = False
        oForm.Items("CLTCOD").Enabled = True
        If Len(ZPAY_GBL_JSNYER) > 0 Then
            oForm.Items("JSNYER").Specific.VALUE = ZPAY_GBL_JSNYER
        End If
        Call CLTCOD_Select(oForm, "CLTCOD")
    ElseIf oForm.Mode = fm_OK_MODE Then
        oForm.Items("JSNYER").Enabled = False
        oForm.Items("MSTCOD").Enabled = False
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = False
        oForm.Items("CLTCOD").Enabled = False
        oOLDCHK = oDS_ZPY501H.GetValue("U_Check", 0)
        
    End If
    oForm.Items("Folder1").Enabled = True
    oForm.Items("Folder2").Enabled = True
    oForm.Items("Folder3").Enabled = True
    
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim DocNum          As String
    
    ErrNum = 0
    
    Call Family_Total("ErrorCheck")
    
 '/ Check
    Select Case True
    Case oDS_ZPY501H.GetValue("U_JSNYER", 0) = ""
        ErrNum = 1
        GoTo Error_Message
    Case oDS_ZPY501H.GetValue("U_MSTCOD", 0) = ""
        ErrNum = 2
        GoTo Error_Message
    Case (Val(oDS_ZPY501H.GetValue("U_BOHAMT", 0)) + Val(oDS_ZPY501H.GetValue("U_JGABOA", 0)) + Val(oDS_ZPY501H.GetValue("U_BOHAL1", 0)) + Val(oDS_ZPY501H.GetValue("U_BOHAL2", 0))) <> _
         (Val(oDS_ZPY501H.GetValue("U_BONBOH1", 0)) + Val(oDS_ZPY501H.GetValue("U_BONBOH2", 0)))
        ErrNum = 3
        GoTo Error_Message
    Case (Val(oDS_ZPY501H.GetValue("U_JGAMED", 0)) + Val(oDS_ZPY501H.GetValue("U_GENMED", 0))) <> _
         (Val(oDS_ZPY501H.GetValue("U_BONMED1", 0)) + Val(oDS_ZPY501H.GetValue("U_BONMED2", 0)))
        ErrNum = 4
        GoTo Error_Message
    Case (Val(oDS_ZPY501H.GetValue("U_BONSCH", 0)) + Val(oDS_ZPY501H.GetValue("U_JGASCH", 0)) + Val(oDS_ZPY501H.GetValue("U_JICSCH", 0))) > _
         (Val(oDS_ZPY501H.GetValue("U_BONEDC1", 0)) + Val(oDS_ZPY501H.GetValue("U_BONEDC2", 0)))
        ErrNum = 5
        GoTo Error_Message
    Case (Val(oDS_ZPY501H.GetValue("U_CADSAV", 0)) + Val(oDS_ZPY501H.GetValue("U_GIRSAV", 0))) <> _
         (Val(oDS_ZPY501H.GetValue("U_BONCAD1", 0)) + Val(oDS_ZPY501H.GetValue("U_BONCAD2", 0)))
        ErrNum = 6
        GoTo Error_Message
    Case Val(oDS_ZPY501H.GetValue("U_CA1SAV", 0)) <> _
         (Val(oDS_ZPY501H.GetValue("U_BONCA11", 0)) + Val(oDS_ZPY501H.GetValue("U_BONCA12", 0)))
        ErrNum = 13
        GoTo Error_Message
    Case Val(oDS_ZPY501H.GetValue("U_CSHSAV", 0)) <> Val(oDS_ZPY501H.GetValue("U_BONCSH1", 0))
        ErrNum = 7
        GoTo Error_Message
    Case (Val(oDS_ZPY501H.GetValue("U_LAWGBU", 0)) + Val(oDS_ZPY501H.GetValue("U_POCGBU", 0)) + _
          Val(oDS_ZPY501H.GetValue("U_SP1GBU", 0)) + Val(oDS_ZPY501H.GetValue("U_SP2GBU", 0)) + _
          Val(oDS_ZPY501H.GetValue("U_USJGBU", 0)) + Val(oDS_ZPY501H.GetValue("U_JIJGBU", 0)) + _
          Val(oDS_ZPY501H.GetValue("U_JI2GBU", 0)) + Val(oDS_ZPY501H.GetValue("U_JNGGBU", 0))) <> (Val(oDS_ZPY501H.GetValue("U_BONGBU1", 0)) + Val(oDS_ZPY501H.GetValue("U_BONGBU2", 0)))
        ErrNum = 8
        GoTo Error_Message
    Case Trim$(oDS_ZPY501H.GetValue("U_Check", 0)) = "Y" And Trim$(oOLDCHK) = "Y"
        ErrNum = 9
        GoTo Error_Message
    Case Val(oDS_ZPY501H.GetValue("U_JHEJA1", 0)) <> 0 And Val(oDS_ZPY501H.GetValue("U_JHEJA2", 0)) <> 0
        ErrNum = 10
        GoTo Error_Message
    Case oDS_ZPY501H.GetValue("U_CLTCOD", 0) = ""
        ErrNum = 11
        GoTo Error_Message
    End Select
    
    DocNum = Exist_YN(oDS_ZPY501H.GetValue("U_JSNYER", 0), oDS_ZPY501H.GetValue("U_MSTCOD", 0), oDS_ZPY501H.GetValue("U_CLTCOD", 0))
    If Trim$(DocNum) <> "" And Trim$(oDS_ZPY501H.GetValue("DocNum", 0)) <> Trim$(DocNum) Then
        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        '같은데이터가 존재하는데 자기 자신이 현재 자기자신이 아니라면(같은월에는 취소한거 아니면 하나만 존재해야함)
        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        ErrNum = 12
        GoTo Error_Message
    End If
    
    HeaderSpaceLineDel = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "귀속 연도는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "사원번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "보험료금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "의료비금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "교육비금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "신용카드금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "현금영수증금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 8 Then
        Sbo_Application.StatusBar.SetText "기부금금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 9 Then
        Sbo_Application.StatusBar.SetText "잠금 자료입니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 10 Then
        Sbo_Application.StatusBar.SetText "장기주택저당차입금이자상환액 600만원한도금액 또는 1000만원 한도 금액 중 하나만 입력가능합니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 11 Then
        Sbo_Application.StatusBar.SetText "자사코드는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 12 Then
        Sbo_Application.StatusBar.SetText "문서번호" & DocNum & " 와(과) 데이터가 일치합니다. 저장되지 않습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 13 Then
        Sbo_Application.StatusBar.SetText "직불카드 금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "HeaderSpaceLineDel 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub MasterUpdate()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
'/ Question
   If Sbo_Application.MessageBox("변경된 인적공제내역을 사원마스터-급여기본사항에 반영하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
      Exit Sub
   End If
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "UPDATE [@PH_PY001A] SET U_BAEWOO = '" & Val(oDS_ZPY501H.GetValue("U_BAEWOO", 0)) & "'"
    sQry = sQry & "         , U_BUYN20 = '" & Val(oDS_ZPY501H.GetValue("U_BUYN20", 0)) & "'"
    sQry = sQry & "         , U_BUYN60 = '" & Val(oDS_ZPY501H.GetValue("U_BUYN60", 0)) & "'"
    sQry = sQry & "         , U_BUYNSU = '" & Val(oDS_ZPY501H.GetValue("U_BUYN20", 0)) + Val(oDS_ZPY501H.GetValue("U_BUYN60", 0)) & "'"
    sQry = sQry & "         , U_GYNGLO = '" & Val(oDS_ZPY501H.GetValue("U_GYNGLO", 0)) & "'"
    sQry = sQry & "         , U_GYNGL2 = '" & Val(oDS_ZPY501H.GetValue("U_GYNGL2", 0)) & "'"
    sQry = sQry & "         , U_JANGAE = '" & Val(oDS_ZPY501H.GetValue("U_JANGAE", 0)) & "'"
    sQry = sQry & "         , U_MZBURI = '" & Val(oDS_ZPY501H.GetValue("U_MZBURI", 0)) & "'"
    sQry = sQry & "         , U_BUYN06 = '" & Val(oDS_ZPY501H.GetValue("U_BUYN06", 0)) & "'"
    sQry = sQry & "         , U_DAGYSU = '" & Val(oDS_ZPY501H.GetValue("U_DAGYSU", 0)) & "'"
    sQry = sQry & "         , U_HUSMAN = '" & Val(oDS_ZPY501H.GetValue("U_HUSMAN", 0)) & "'"
    sQry = sQry & "   WHERE Code = N'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'"
    oRecordSet.DoQuery sQry
    '/ 본인주민번호도 수정
    If oDS_ZPY501L.GetValue("U_FamPer", 0) <> "" And Trim$(oDS_ZPY501L.GetValue("U_ChkCod", 0)) = "0" Then
        '/ 인사마스터
        sQry = "UPDATE [@PH_PY001A] SET GOVID = '" & oDS_ZPY501L.GetValue("U_FamPer", 0) & "'"
        sQry = sQry & "   WHERE Code = N'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'"
        oRecordSet.DoQuery sQry
        '/ 급여기본등록
        sQry = "UPDATE [@PH_PY001A] SET  U_INTGBN = '" & oDS_ZPY501L.GetValue("U_ChkInt", 0) & "'"  '/ 내외국인
        sQry = sQry & "         ,      U_BJNGAE = '" & oDS_ZPY501L.GetValue("U_ChkJan", 0) & "'"  '/ 본인장애유무
        sQry = sQry & "   WHERE Code = N'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'"
        oRecordSet.DoQuery sQry
    End If
    
    '/ 가족사항등록화면에 부양가족정보 적용
    If Sbo_Application.MessageBox("변경된 인적공제내역을 가족사항등록에 반영하시겠습니까?", 2, "&Yes!", "&No") = 1 Then
        sQry = "EXEC ZPY501_1 '" & oDS_ZPY501H.GetValue("U_JSNYER", 0) & "', '" & oDS_ZPY501H.GetValue("U_CLTCOD", 0) & "', '" & oDS_ZPY501H.GetValue("U_MSTCOD", 0) & "', '" & oCompany.UserSignature & "'"
        oRecordSet.DoQuery sQry
    End If
    
    '/ 메세지
  ' SBO_Application.StatusBar.SetText "마스터에 적용하였습니다.", bmt_Short, smt_Success
    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "MasterUpdate Error:" & Space$(10) & Err.Description, bmt_Short, smt_Error
    
End Sub

Private Function MatrixSpaceLineDel() As Boolean
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'저장할 데이터의 유효성을 점검한다
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
On Error GoTo Error_Message
    Dim i           As Long
    Dim k           As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    
    Dim U_BAEWOO As Integer, U_GYNGLO As Integer
    Dim U_JANGAE As Integer, U_MZBURI As Integer, U_BUYN06 As Integer, U_DAGYSU As Integer, U_INJTOT As Integer
    Dim GovidChk As String
    Dim CLTCOD   As String
    
    U_BAEWOO = 0: U_GYNGLO = 0: U_JANGAE = 0: U_MZBURI = 0
    U_BUYN06 = 0: U_DAGYSU = 0: U_INJTOT = 0
    
    ErrNum = 0
    CLTCOD = MDC_SetMod.Get_ReData("ISNULL(U_CLTCOD,'')", "Code", "[@PH_PY001A]", "'" & oDS_ZPY501H.GetValue("U_MSTCOD", 0) & "'", "")
    GovidChk = MDC_SetMod.Get_ReData("ISNULL(U_GovIDChk,'N')", "Code", "[@PH_PY005A]", "'" & CLTCOD & "'", "")

    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat1.FlushToDataSource
        
    If oMat1.RowCount = 0 Then          '// Mat1에 값이 있는지 확인 (ErrorNumber : 1)
        ErrNum = 1
        GoTo Error_Message
        Exit Function
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat1.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    For i = 0 To oMat1.VisualRowCount - 1      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 3)
        oDS_ZPY501L.Offset = i
        '// 가족성명이 있을 경우만
        If Trim$(oDS_ZPY501L.GetValue("U_FamNam", i)) <> "" Then
            If Trim$(oDS_ZPY501L.GetValue("U_ChkBas", i)) = "1" Then
                U_INJTOT = U_INJTOT + Val(oDS_ZPY501L.GetValue("U_ChkBas", i))
                U_BAEWOO = U_BAEWOO + IIf(Trim$(oDS_ZPY501L.GetValue("U_ChkCod", i)) = "3", 1, 0)
                U_GYNGLO = U_GYNGLO + Val(oDS_ZPY501L.GetValue("U_ChkJeL", i))
                U_MZBURI = U_MZBURI + Val(oDS_ZPY501L.GetValue("U_ChkbuY", i))
                U_DAGYSU = U_DAGYSU + Val(oDS_ZPY501L.GetValue("U_ChkDaJ", i))
            End If
            U_JANGAE = U_JANGAE + Val(oDS_ZPY501L.GetValue("U_ChkJan", i))
            U_BUYN06 = U_BUYN06 + Val(oDS_ZPY501L.GetValue("U_ChkChl", i))
        
            If Trim$(oDS_ZPY501L.GetValue("U_FamPer", i)) = "" Then
                ErrNum = 5
                oMat1.Columns("Col2").Cells(i + 1).CLICK ct_Regular
                GoTo Error_Message
            ElseIf Trim$(oDS_ZPY501L.GetValue("U_ChkCod", i)) = "" Then
                ErrNum = 6
                oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular
                GoTo Error_Message
            ElseIf Trim$(oDS_ZPY501L.GetValue("U_ChkInt", i)) = "" Then
                ErrNum = 7
                oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular
                GoTo Error_Message
             '/ 교육비공제는 기본공제대상자가 아니여도 가능함.(기본공제대상자에서 나이제한으로 기본공제대상자가 안될경우도 포함됨. 20세이상자녀대학교)
    '        ElseIf Trim$(oDS_ZPY501L.GetValue("U_ChkBas", i)) <> "1" And (Val(oDS_ZPY501L.GetValue("U_EDCAMT1", i)) + Val(oDS_ZPY501L.GetValue("U_EDCAMT1", i))) > 0 Then
    '            ErrNum = 15
    '            oMat1.Columns("Col1").Cells(i + 1).Click ct_Regular
    '            GoTo Error_Message
            
            ElseIf (Val(oDS_ZPY501L.GetValue("U_EDCAMT1", i)) > 0 Or Val(oDS_ZPY501L.GetValue("U_EDCAMT2", i)) > 0) And _
                   Trim$(oDS_ZPY501L.GetValue("U_SCHGBN", i)) = "" Then
                ErrNum = 17
                oMat1.Columns("Col17").Cells(i + 1).CLICK ct_Regular
                GoTo Error_Message
            Else
                'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                '중복체크작업
                'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                Chk_Data = Trim$(oDS_ZPY501L.GetValue("U_FamNam", i))
                For k = i + 1 To oMat1.VisualRowCount - 1
                    oDS_ZPY501L.Offset = k
                    If Trim$(Chk_Data) = Trim$(oDS_ZPY501L.GetValue("U_FamNam", k)) Then
                        ErrNum = 3
                        oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular
                        GoTo Error_Message
                    End If
                Next k
            End If
            '// 6.주민번호 오류 체크
            If Trim$(GovidChk) = "Y" And Len(oDS_ZPY501L.GetValue("U_FamPer", i)) > 0 Then
                If GovIDCheck(oDS_ZPY501L.GetValue("U_FamPer", i)) = False Then
                    ErrNum = 16
                    oMat1.Columns("Col2").Cells(i + 1).CLICK ct_Regular
                    GoTo Error_Message
                End If
            End If
        
        '// 가족성명이 없을경우
        Else
            '/ 주민등록번호는 있는데 가족성명이 안들어오는 경우
             If Trim$(oDS_ZPY501L.GetValue("U_FamNam", i)) = "" And Trim$(oDS_ZPY501L.GetValue("U_FamPer", i)) <> "" Then
                ErrNum = 4
                oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular
                GoTo Error_Message
             End If
        End If

    Next i
 '/
    Select Case True
    Case U_INJTOT <> (Val(oDS_ZPY501H.GetValue("U_BAEWOO", 0)) + _
          Val(oDS_ZPY501H.GetValue("U_BUYN20", 0)) + Val(oDS_ZPY501H.GetValue("U_BUYN60", 0)) + 1)
          ErrNum = 8
          GoTo Error_Message
    Case U_BAEWOO <> Val(oDS_ZPY501H.GetValue("U_BAEWOO", 0))
          ErrNum = 9
          GoTo Error_Message
    Case U_GYNGLO <> (Val(oDS_ZPY501H.GetValue("U_GYNGLO", 0)) + Val(oDS_ZPY501H.GetValue("U_GYNGL2", 0)))
          ErrNum = 10
          GoTo Error_Message
    Case U_JANGAE <> Val(oDS_ZPY501H.GetValue("U_JANGAE", 0))
          ErrNum = 11
          GoTo Error_Message
    Case U_MZBURI <> Val(oDS_ZPY501H.GetValue("U_MZBURI", 0))
          ErrNum = 12
          GoTo Error_Message
    Case U_BUYN06 <> Val(oDS_ZPY501H.GetValue("U_BUYN06", 0))
          ErrNum = 13
          GoTo Error_Message
    Case U_DAGYSU <> Val(oDS_ZPY501H.GetValue("U_DAGYSU", 0))
          ErrNum = 14
          GoTo Error_Message
    End Select

    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If Trim$(oDS_ZPY501L.GetValue("U_FamNam", oDS_ZPY501L.Size - 1)) = "" Then
        oDS_ZPY501L.RemoveRecord oDS_ZPY501L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
    End If
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat1.LoadFromDataSource
    
    MatrixSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "부양가족명세에 데이터가 없습니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "귀속 연도는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "중복된 자료가 있습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "부양가족 성명은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "부양가족 주민번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "부양가족 관계코드는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "부양가족 내외국인구분은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 8 Then
        Sbo_Application.StatusBar.SetText "부양가족수가 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 9 Then
        Sbo_Application.StatusBar.SetText "배우자가 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 10 Then
        Sbo_Application.StatusBar.SetText "경로우대인원수가 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 11 Then
        Sbo_Application.StatusBar.SetText "장애인인원수가 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 12 Then
        Sbo_Application.StatusBar.SetText "부녀자인원수가 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 13 Then
        Sbo_Application.StatusBar.SetText "자녀양육인원수가 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 14 Then
        Sbo_Application.StatusBar.SetText "다자녀공제인원수가 부양가족명세와 다릅니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 15 Then
        Sbo_Application.StatusBar.SetText "교육비 공제는 기본공제대상자만 가능합니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 16 Then
        Sbo_Application.StatusBar.SetText "주민등록번호를 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 17 Then
        Sbo_Application.StatusBar.SetText "교육비가 입력된 경우 교육대상 구분은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Error_Message:" & Err.Number & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    MatrixSpaceLineDel = False
End Function

Private Sub ShowSource(JSNYER$, MSTCOD$, CLTCOD$)
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim DocEntry$
    Dim oMast           As ZPAY_g_EmpID
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    sQry = "SELECT DocEntry FROM [@ZPY501H]"
    sQry = sQry & "   WHERE U_JSNYER = N'" & JSNYER & "'"
    sQry = sQry & "   AND   U_MSTCOD = N'" & MSTCOD & "'"
    sQry = sQry & "   AND   U_CLTCOD = N'" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            DocEntry = oRecordSet.Fields(0).VALUE
            oRecordSet.MoveNext
        Loop
        oForm.Items("JSNYER").Specific.VALUE = JSNYER
        oForm.Items("MSTCOD").Specific.String = MSTCOD
        Call oForm.Items("CLTCOD").Specific.Select(CLTCOD, psk_ByValue)
        oForm.Items("DocNum").Specific.VALUE = DocEntry
        
        oForm.Items("DocNum").Update
        oMat1.LoadFromDataSource
        oForm.Update
        oForm.Items("1").CLICK ct_Regular
        
    Else
        oForm.Mode = fm_OK_MODE
        Call Sbo_Application.ActivateMenuItem("1282")
        
        Call oDS_ZPY501H.setValue("U_JSNYER", 0, JSNYER)
        Call oDS_ZPY501H.setValue("U_MSTCOD", 0, MSTCOD)
        Call oDS_ZPY501H.setValue("U_CLTCOD", 0, CLTCOD)
        oMast = MDC_SetMod.Get_EmpID_InFo(MSTCOD)
        Call oDS_ZPY501H.setValue("U_MSTNAM", 0, oMast.MSTNAM)
        Call oDS_ZPY501H.setValue("U_EmpID", 0, oMast.EmpID)
        
        oForm.Update
        
        Sbo_Application.SendKeys "{TAB}"
    End If
    
    Set oRecordSet = Nothing
    
End Sub

Private Sub SavingAmount_Data_Display()
On Error GoTo Error_Message
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim ErrNum      As Integer
    Dim JSNYER      As String:  Dim MSTCOD      As String:  Dim CLTCOD      As String
    
    If oForm.Mode <> fm_ADD_MODE And oForm.Mode <> fm_OK_MODE And oForm.Mode <> fm_UPDATE_MODE Then
        Exit Sub
    End If
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    JSNYER = Trim$(oDS_ZPY501H.GetValue("U_JSNYER", 0))
    MSTCOD = Trim$(oDS_ZPY501H.GetValue("U_MSTCOD", 0))
    CLTCOD = Trim$(oDS_ZPY501H.GetValue("U_CLTCOD", 0))
    
    If JSNYER = "" Then
        ErrNum = 2
        GoTo Error_Message
    ElseIf MSTCOD = "" Then
        ErrNum = 3
        GoTo Error_Message
    ElseIf CLTCOD = "" Then
        ErrNum = 4
        GoTo Error_Message
    End If
    
    sQry = "EXEC ZPY501_2 '" & JSNYER & "', '" & MSTCOD & "', '" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
    
        ErrNum = 1
        
        '2012.02.16 기존에 연금저축명세가 존재 해서 값이 들어온 이후에 연금저축명세를
        '삭제했을때 값을 초기화 되도록 수정(V12.01.08 버젼 이후로 적용되도록 소스만 수정 해 놓음)
        
        oDS_ZPY501H.setValue "U_RETSAV", 0, 0
        oDS_ZPY501H.setValue "U_RETSA1", 0, 0
        oDS_ZPY501H.setValue "U_GYNSAV", 0, 0
        oDS_ZPY501H.setValue "U_YUNSAV", 0, 0
        oDS_ZPY501H.setValue "U_HUSAMT", 0, 0
        oDS_ZPY501H.setValue "U_HU1AMT", 0, 0
        oDS_ZPY501H.setValue "U_HU2AMT", 0, 0
        oDS_ZPY501H.setValue "U_HU3AMT", 0, 0
        oDS_ZPY501H.setValue "U_JFDAM1", 0, 0
        oDS_ZPY501H.setValue "U_JFDAM2", 0, 0
        oDS_ZPY501H.setValue "U_JFDAM3", 0, 0
        
        GoTo Error_Message
        
    End If
    
    oDS_ZPY501H.setValue "U_RETSAV", 0, oRecordSet.Fields("RETSAV").VALUE
    oDS_ZPY501H.setValue "U_RETSA1", 0, oRecordSet.Fields("RETSA1").VALUE
    oDS_ZPY501H.setValue "U_GYNSAV", 0, oRecordSet.Fields("GYNSAV").VALUE
    oDS_ZPY501H.setValue "U_YUNSAV", 0, oRecordSet.Fields("YUNSAV").VALUE
    oDS_ZPY501H.setValue "U_HUSAMT", 0, oRecordSet.Fields("HUSAMT").VALUE
    oDS_ZPY501H.setValue "U_HU1AMT", 0, oRecordSet.Fields("HU1AMT").VALUE
    oDS_ZPY501H.setValue "U_HU2AMT", 0, oRecordSet.Fields("HU2AMT").VALUE
    oDS_ZPY501H.setValue "U_HU3AMT", 0, oRecordSet.Fields("HU3AMT").VALUE
    oDS_ZPY501H.setValue "U_JFDAM1", 0, oRecordSet.Fields("JFDAM1").VALUE
    oDS_ZPY501H.setValue "U_JFDAM2", 0, oRecordSet.Fields("JFDAM2").VALUE
    oDS_ZPY501H.setValue "U_JFDAM3", 0, oRecordSet.Fields("JFDAM3").VALUE
    
    oForm.Update
    
    If oForm.Mode = fm_OK_MODE Then
        oForm.Mode = fm_UPDATE_MODE
    End If
    
    Set oRecordSet = Nothing
    
    Exit Sub
Error_Message:
    Set oRecordSet = Nothing
    Select Case ErrNum
    Case 1:     Sbo_Application.StatusBar.SetText "가져올 연금,저축 명세내역이 없습니다.", bmt_Short, smt_Error
    Case 2:     Sbo_Application.StatusBar.SetText "귀속 년도는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Case 3:     Sbo_Application.StatusBar.SetText "사원 번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Case 4:     Sbo_Application.StatusBar.SetText "자사 코드는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Case Else:  Sbo_Application.StatusBar.SetText "SavingAmount_Data_Display Error : " & Err.Description, bmt_Short, smt_Error
    End Select

End Sub

Private Sub Family_Data_Display()
On Error GoTo Error_Message
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim i           As Integer
    Dim oRow        As Integer
    Dim cnt         As Integer
    Dim JSNYER      As Integer
    Dim GBUAMT      As Double
    Dim MSTCOD      As String:  Dim CLTCOD  As String

'/ Question
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    i = 0
    oRow = 0
'/
    cnt = oDS_ZPY501L.Size
    If cnt > 0 Then
        For i = 0 To cnt - 1
            oDS_ZPY501L.RemoveRecord oDS_ZPY501L.Size - 1
        Next i
        Matrix_AddRow 0, True
    Else
        oMat1.LoadFromDataSource
    End If
'/
    JSNYER = Val(oForm.Items("JSNYER").Specific.VALUE)
    MSTCOD = Trim$(oForm.Items("MSTCOD").Specific.String)
    CLTCOD = Trim(oDS_ZPY501H.GetValue("U_CLTCOD", 0))
    If JSNYER <= 0 Then
        Exit Sub
    End If
    
    i = 0
    sQry = "EXEC ZPY501 " & "'" & JSNYER & "', N'" & MSTCOD & "', '" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        If i + 1 > oDS_ZPY501L.Size Then
            oDS_ZPY501L.InsertRecord (i)
        End If

        BOHAL1 = oRecordSet.Fields("U_JONBOH").VALUE
        BOHAL2 = oRecordSet.Fields("U_BOHAMT3").VALUE
        GBUAMT = oRecordSet.Fields("U_GBUAMT3").VALUE
        If Trim$(oRecordSet.Fields("U_ChkCod").VALUE) = "0" Then  '/ 본인일 경우 전액공제보험료표시
            oDS_ZPY501H.setValue "U_BOHAL1", 0, BOHAL1
            oDS_ZPY501H.setValue "U_BOHAL2", 0, BOHAL2
            oDS_ZPY501H.setValue "U_JI2GBU", 0, GBUAMT
        End If
        oDS_ZPY501L.Offset = i
        oDS_ZPY501L.setValue "U_LineNum", i, i + 1
        oDS_ZPY501L.setValue "U_FamNam", i, oRecordSet.Fields("U_FamNam").VALUE
        oDS_ZPY501L.setValue "U_FamPer", i, oRecordSet.Fields("U_FamPer").VALUE
        oDS_ZPY501L.setValue "U_ChkCod", i, oRecordSet.Fields("U_ChkCod").VALUE
        oDS_ZPY501L.setValue "U_ChkInt", i, oRecordSet.Fields("U_ChkInt").VALUE
        oDS_ZPY501L.setValue "U_ChkBas", i, oRecordSet.Fields("U_ChkBas").VALUE
        oDS_ZPY501L.setValue "U_ChkbuY", i, oRecordSet.Fields("U_ChkbuY").VALUE
        oDS_ZPY501L.setValue "U_ChkJan", i, oRecordSet.Fields("U_ChkJan").VALUE
        oDS_ZPY501L.setValue "U_ChkJeL", i, oRecordSet.Fields("U_ChkJeL").VALUE
        oDS_ZPY501L.setValue "U_ChkChl", i, oRecordSet.Fields("U_ChkChl").VALUE
        oDS_ZPY501L.setValue "U_ChkDaJ", i, oRecordSet.Fields("U_ChkDaJ").VALUE
        oDS_ZPY501L.setValue "U_ChkCHS", i, "0" '출산/입양: 해당년도의 입양자,출산자이니 전년도자료 클리어
        oDS_ZPY501L.setValue "U_BOHGBN", i, oRecordSet.Fields("U_BOHGBN").VALUE
        oDS_ZPY501L.setValue "U_BOHAMT1", i, oRecordSet.Fields("U_BOHAMT1").VALUE
        oDS_ZPY501L.setValue "U_BOHAMT2", i, oRecordSet.Fields("U_BOHAMT2").VALUE
        oDS_ZPY501L.setValue "U_BOHAMT3", i, BOHAL1 + BOHAL2
        oDS_ZPY501L.setValue "U_MEDAMT1", i, oRecordSet.Fields("U_MEDAMT1").VALUE
        oDS_ZPY501L.setValue "U_MEDAMT2", i, oRecordSet.Fields("U_MEDAMT2").VALUE
        oDS_ZPY501L.setValue "U_SCHGBN", i, oRecordSet.Fields("U_SCHGBN").VALUE
        oDS_ZPY501L.setValue "U_EDCAMT1", i, oRecordSet.Fields("U_EDCAMT1").VALUE
        oDS_ZPY501L.setValue "U_EDCAMT2", i, oRecordSet.Fields("U_EDCAMT2").VALUE
        oDS_ZPY501L.setValue "U_CADAMT1", i, oRecordSet.Fields("U_CADAMT1").VALUE
        oDS_ZPY501L.setValue "U_CADAMT2", i, oRecordSet.Fields("U_CADAMT2").VALUE
        oDS_ZPY501L.setValue "U_CSHCAD1", i, oRecordSet.Fields("U_CSHCAD1").VALUE
        oDS_ZPY501L.setValue "U_CSHCAD2", i, oRecordSet.Fields("U_CSHCAD2").VALUE
        oDS_ZPY501L.setValue "U_CSHAMT1", i, oRecordSet.Fields("U_CSHAMT1").VALUE
        oDS_ZPY501L.setValue "U_GBUAMT1", i, oRecordSet.Fields("U_GBUAMT1").VALUE
        oDS_ZPY501L.setValue "U_GBUAMT2", i, oRecordSet.Fields("U_GBUAMT2").VALUE
        oDS_ZPY501L.setValue "U_GBUAMT3", i, oRecordSet.Fields("U_GBUAMT3").VALUE
        oRecordSet.MoveNext
        i = i + 1
    Loop
    
    '/ 메세지
    oMat1.LoadFromDataSource
    i = oMat1.VisualRowCount - 1
    If Trim$(oDS_ZPY501L.GetValue("U_FamNam", i)) <> "" Then
        Matrix_AddRow i + 1
    End If
    
    If Trim$(oDS_ZPY501H.GetValue("U_HUSMAN", 0)) <> "1" And Trim$(oDS_ZPY501H.GetValue("U_HUSMAN", 0)) <> "2" Then
        oDS_ZPY501H.setValue "U_HUSMAN", 0, MDC_SetMod.Get_ReData("U_HUSMAN", "Code", "[@PH_PY001A]", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'")
    End If
    
    sQry = "EXEC ZPY501_3 '" & JSNYER & "', '" & MSTCOD & "', '" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount > 0 Then
        oDS_ZPY501H.setValue "U_LAWGBU", 0, oRecordSet.Fields("LAWGBU").VALUE
        oDS_ZPY501H.setValue "U_POCGBU", 0, oRecordSet.Fields("POCGBU").VALUE
        oDS_ZPY501H.setValue "U_SP1GBU", 0, oRecordSet.Fields("SP1GBU").VALUE
        oDS_ZPY501H.setValue "U_SP2GBU", 0, oRecordSet.Fields("SP2GBU").VALUE
        oDS_ZPY501H.setValue "U_USJGBU", 0, oRecordSet.Fields("USJGBU").VALUE
        oDS_ZPY501H.setValue "U_JIJGBU", 0, oRecordSet.Fields("JIJGBU").VALUE
        oDS_ZPY501H.setValue "U_JI2GBU", 0, oRecordSet.Fields("JI2GBU").VALUE
        oDS_ZPY501H.setValue "U_JNGGBU", 0, oRecordSet.Fields("JNGGBU").VALUE
        oDS_ZPY501H.setValue "U_BONGBU1", 0, oRecordSet.Fields("BONGBU1").VALUE
        oDS_ZPY501H.setValue "U_BONGBU2", 0, oRecordSet.Fields("BONGBU2").VALUE
    End If
    
    oForm.Update
    '//
  ' SBO_Application.StatusBar.SetText "마스터에 적용하였습니다.", bmt_Short, smt_Success
    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "Family_Data_Display Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
    
End Sub

Private Sub Family_Total(Optional TotalOption As String)
    Dim iCol            As Integer
    Dim BOHAMT(1)       As Double
    Dim MEDAMT(1)       As Double
    Dim EDCAMT(1)       As Double
    Dim CADAMT(1)       As Double
    Dim CSHCAD(1)       As Double
    Dim CSHAMT(1)       As Double
    Dim GBUAMT(1)       As Double
    Dim Per_Info        As PerSonNo_Info
    
    Dim U_BOHAMT As Double, U_JGABOA As Double, U_JGAMED As Double, U_GENMED As Double
    Dim U_BONSCH As Double, U_JGASCH As Double, U_JICSCH As Double
    Dim SCHKUM   As Double, TMPAMT   As Double
    Dim U_BAEWOO As Integer, U_BUYN20 As Integer, U_BUYN60 As Integer, U_GYNGLO As Integer, U_GYNGL2 As Integer
    Dim U_JANGAE As Integer, U_MZBURI As Integer, U_BUYN06 As Integer, U_DAGYSU As Integer, U_INJTOT As Integer
    Dim U_CHLSAN As Integer
    Dim U_JI2GBU As Double
    '/
    For iCol = 0 To 1
        BOHAMT(iCol) = 0: MEDAMT(iCol) = 0: EDCAMT(iCol) = 0: CADAMT(iCol) = 0: CSHCAD(iCol) = 0: CSHAMT(iCol) = 0: GBUAMT(iCol) = 0
    Next iCol
    
    U_BOHAMT = 0: U_JGABOA = 0: U_JGAMED = 0: U_GENMED = 0: U_BONSCH = 0: U_JGASCH = 0: U_JICSCH = 0
    TMPAMT = 0
    '/
    U_BAEWOO = 0: U_BUYN20 = 0: U_BUYN60 = 0: U_GYNGLO = 0: U_GYNGL2 = 0: U_JANGAE = 0: U_MZBURI = 0
    U_BUYN06 = 0: U_DAGYSU = 0: U_INJTOT = 0: U_DAGYSU = 0: U_CHLSAN = 0
    '/
    oMat1.FlushToDataSource
    For iCol = 0 To oDS_ZPY501L.Size - 1
        oDS_ZPY501L.Offset = iCol
        '// 인적공제 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        If Trim$(oDS_ZPY501L.GetValue("U_FamNam", iCol)) <> "" Then
            '// 나이확인
             Per_Info = MDC_Com.Age_Chk(oDS_ZPY501L.GetValue("U_FamPer", iCol), Trim$(oDS_ZPY501H.GetValue("U_JSNYER", 0)) & "1231")
            '/ 기본공제자가 아니더라도 공제 가능.
             U_JANGAE = U_JANGAE + Val(oDS_ZPY501L.GetValue("U_ChkJan", iCol))      '/ 장애인공제수
            '/ 본인/배우자/직계존속은 안됨(자녀양육,출산입양)
              If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) > "3" Then  '
                '/ 출산.입양공제는 기본공제대상자는 다른 근로자가 공제, 본인은 자녀양육비추가공제로 받을수 있슴.
                 U_CHLSAN = U_CHLSAN + Val(oDS_ZPY501L.GetValue("U_ChkCHS", iCol))      '/ 출산입양자공제수
                '/ 6세이하비속의 추가공제는 기본공제대상자는 다른 근로자가 공제, 본인은 자녀양육비추가공제로 받을수 있슴.
                 If Per_Info.ManAge <= 6 Then U_BUYN06 = U_BUYN06 + Val(oDS_ZPY501L.GetValue("U_ChkChl", iCol))
              End If
            '/ 기본공제자일경우만 공제일때만 가능.
            If Trim$(oDS_ZPY501L.GetValue("U_ChkBas", iCol)) = "1" Then
                U_INJTOT = U_INJTOT + 1
                
                If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) = "0" And Per_Info.Sex = "1" Then
                    U_MZBURI = U_MZBURI + Val(oDS_ZPY501L.GetValue("U_ChkbuY", iCol))  '/부녀자공제수(본인이고 여성일경우만가능)
                End If

                If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) <> "0" Then     '/ 본인이아닐경우
                    If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) = "3" Then U_BAEWOO = U_BAEWOO + 1
                    If Per_Info.ManAge <= 20 And Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) <> "3" Then U_BUYN20 = U_BUYN20 + 1
                    If Per_Info.ManAge >= 60 And Per_Info.Sex = 0 And Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) <> "3" Then
                        U_BUYN60 = U_BUYN60 + 1
                    End If
                    '// 2009년부터 여자일경우도 55세이상-> 60세이상으로 상향조정
                    If Per_Info.ManAge >= 60 And Per_Info.Sex = 1 And Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) <> "3" Then
                        U_BUYN60 = U_BUYN60 + 1
                    End If
                    '/ 나이한도에서는 제외되나 장애자여서 기본공제인원에 들어가는 사람(배우자는 배우자인원수에 포함하니 제외)
                    If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) <> "3" Then
                       If Val(oDS_ZPY501L.GetValue("U_ChkJan", iCol)) = 1 Then
                           If Per_Info.ManAge > 20 And ((Per_Info.ManAge < 60 And Per_Info.Sex = 1) Or (Per_Info.ManAge < 60 And Per_Info.Sex <> 1)) Then U_BUYN20 = U_BUYN20 + 1
                       End If
                    End If
                    '// 2009년 65세이상 69세이하 경로우대 폐지됨.
                    'If Per_Info.ManAge >= 65 And Per_Info.ManAge <= 69 Then U_GYNGLO = U_GYNGLO + 1
                    U_GYNGLO = 0
                    If Per_Info.ManAge >= 70 Then U_GYNGL2 = U_GYNGL2 + 1
                    If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) = "4" And Per_Info.ManAge <= 20 Then
                        U_DAGYSU = U_DAGYSU + 1
                    End If
                End If
            End If
        End If
        '// 보험료 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        BOHAMT(0) = BOHAMT(0) + Val(oDS_ZPY501L.GetValue("U_BOHAMT1", iCol))
        BOHAMT(1) = BOHAMT(1) + Val(oDS_ZPY501L.GetValue("U_BOHAMT2", iCol)) + Val(oDS_ZPY501L.GetValue("U_BOHAMT3", iCol))
        TMPAMT = 0
        TMPAMT = Val(oDS_ZPY501L.GetValue("U_BOHAMT1", iCol)) + Val(oDS_ZPY501L.GetValue("U_BOHAMT2", iCol)) + Val(oDS_ZPY501L.GetValue("U_BOHAMT3", iCol))
        If Trim$(oDS_ZPY501L.GetValue("U_BOHGBN", iCol)) = "2" Then   '/ 장애인보험
            U_JGABOA = U_JGABOA + TMPAMT
            If Trim$(oDS_ZPY501L.GetValue("U_CHKCOD", iCol)) = "0" Then '/ 본인일경우
                U_JGABOA = U_JGABOA - Val(oDS_ZPY501H.GetValue("U_BOHAL1", 0)) - Val(oDS_ZPY501H.GetValue("U_BOHAL2", 0))
            End If
        Else                                                      '/ 일반인보험
            U_BOHAMT = U_BOHAMT + TMPAMT
            If Trim$(oDS_ZPY501L.GetValue("U_CHKCOD", iCol)) = "0" Then '/ 본인일경우
                U_BOHAMT = U_BOHAMT - Val(oDS_ZPY501H.GetValue("U_BOHAL1", 0)) - Val(oDS_ZPY501H.GetValue("U_BOHAL2", 0))
            End If
        End If
        '/ 의료비 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        MEDAMT(0) = MEDAMT(0) + Val(oDS_ZPY501L.GetValue("U_MEDAMT1", iCol))
        MEDAMT(1) = MEDAMT(1) + Val(oDS_ZPY501L.GetValue("U_MEDAMT2", iCol))
        TMPAMT = Val(oDS_ZPY501L.GetValue("U_MEDAMT1", iCol)) + Val(oDS_ZPY501L.GetValue("U_MEDAMT2", iCol))
'        If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", icol)) = "0" Or _
'           Trim$(oDS_ZPY501L.GetValue("U_ChkJeL", icol)) = "1" Or _
'           Trim$(oDS_ZPY501L.GetValue("U_ChkJan", icol)) = "1" Then   '/ 본인, 경로, 장애인
        If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) = "0" Or _
           Per_Info.ManAge >= 65 Or _
           Trim$(oDS_ZPY501L.GetValue("U_ChkJan", iCol)) = "1" Then   '/ 본인, 65세이상, 장애인
           U_JGAMED = U_JGAMED + TMPAMT
        Else
           U_GENMED = U_GENMED + TMPAMT
        End If
        '/ 교육비 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        EDCAMT(0) = EDCAMT(0) + Val(oDS_ZPY501L.GetValue("U_EDCAMT1", iCol))
        EDCAMT(1) = EDCAMT(1) + Val(oDS_ZPY501L.GetValue("U_EDCAMT2", iCol))
        TMPAMT = Val(oDS_ZPY501L.GetValue("U_EDCAMT1", iCol)) + Val(oDS_ZPY501L.GetValue("U_EDCAMT2", iCol))
        If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) = "0" Then   '/ 본인전액공제(대학원포함)
            U_BONSCH = Val(oDS_ZPY501L.GetValue("U_EDCAMT1", iCol)) + Val(oDS_ZPY501L.GetValue("U_EDCAMT2", iCol))
        Else
            If Trim$(oDS_ZPY501L.GetValue("U_SCHGBN", iCol)) = "5" Then  '/ 장애인특수교육
                U_JGASCH = U_JGASCH + Val(oDS_ZPY501L.GetValue("U_EDCAMT1", iCol)) + Val(oDS_ZPY501L.GetValue("U_EDCAMT2", iCol))
            Else
                Select Case Trim$(oDS_ZPY501L.GetValue("U_SCHGBN", iCol))
                Case "1", "2": SCHKUM = 3000000  '/ 영유아, 초중고
                Case "3": SCHKUM = 9000000   '/ 대학교
                Case "4": SCHKUM = 0         '/ 대학원생
                End Select
                If TMPAMT <= SCHKUM Then
                    U_JICSCH = U_JICSCH + TMPAMT
                Else
                    U_JICSCH = U_JICSCH + SCHKUM
                End If
            End If
        End If
    '/ 신용카드 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        CADAMT(0) = CADAMT(0) + Val(oDS_ZPY501L.GetValue("U_CADAMT1", iCol))
        CADAMT(1) = CADAMT(1) + Val(oDS_ZPY501L.GetValue("U_CADAMT2", iCol))
    '/ 직불카드
        CSHCAD(0) = CSHCAD(0) + Val(oDS_ZPY501L.GetValue("U_CSHCAD1", iCol))
        CSHCAD(1) = CSHCAD(1) + Val(oDS_ZPY501L.GetValue("U_CSHCAD2", iCol))
    '/ 현금영수증 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        CSHAMT(0) = CSHAMT(0) + Val(oDS_ZPY501L.GetValue("U_CSHAMT1", iCol))
    '/ 기부금 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        GBUAMT(0) = GBUAMT(0) + Val(oDS_ZPY501L.GetValue("U_GBUAMT1", iCol))
        GBUAMT(1) = GBUAMT(1) + Val(oDS_ZPY501L.GetValue("U_GBUAMT2", iCol)) + Val(oDS_ZPY501L.GetValue("U_GBUAMT3", iCol))
        U_JI2GBU = U_JI2GBU + Val(oDS_ZPY501L.GetValue("U_GBUAMT3", iCol))
    Next iCol
    
    If TotalOption = "All" Then
        '/ 다자녀일경우 소득공제명세에도 체크자동으로
        For iCol = 0 To oDS_ZPY501L.Size - 1
            oDS_ZPY501L.Offset = iCol
           If Trim$(oDS_ZPY501L.GetValue("U_ChkBas", iCol)) = "1" And Trim$(oDS_ZPY501L.GetValue("U_FamNam", iCol)) <> "" And U_DAGYSU >= 2 Then
                Per_Info = MDC_Com.Age_Chk(oDS_ZPY501L.GetValue("U_FamPer", iCol), Trim$(oDS_ZPY501H.GetValue("U_JSNYER", 0)) & "1231")
                If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) = "4" And Per_Info.ManAge <= 20 Then
                    oDS_ZPY501L.setValue "U_ChkDaJ", iCol, 1
                Else
                    oDS_ZPY501L.setValue "U_ChkDaJ", iCol, 0
                End If
           Else
                oDS_ZPY501L.setValue "U_ChkDaJ", iCol, 0
           End If
        Next iCol
    End If
    oForm.Freeze True
    oMat1.LoadFromDataSource
    If TotalOption = "All" Then
        '/ 소득공제항목 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        oDS_ZPY501H.setValue "U_JGABOA", 0, U_JGABOA    '/ 장애인보험료
        oDS_ZPY501H.setValue "U_BOHAMT", 0, U_BOHAMT    '/ 일반인보험료
        oDS_ZPY501H.setValue "U_JGAMED", 0, U_JGAMED    '/ 본인,경로,장애인의료비
        oDS_ZPY501H.setValue "U_GENMED", 0, U_GENMED    '/ 그외 의료비
        oDS_ZPY501H.setValue "U_BONSCH", 0, U_BONSCH    '/ 본인교육비
        oDS_ZPY501H.setValue "U_JGASCH", 0, U_JGASCH    '/ 장애특수교육비
        oDS_ZPY501H.setValue "U_JICSCH", 0, U_JICSCH    '/ 가족교육비
        If (Val(oDS_ZPY501H.GetValue("U_CADSAV", 0)) + Val(oDS_ZPY501H.GetValue("U_GIRSAV", 0))) <> (CADAMT(0) + CADAMT(1)) Then
            oDS_ZPY501H.setValue "U_CADSAV", 0, CADAMT(0) + CADAMT(1)   '/ 신용카드
            'oDS_ZPY501H.SetValue "U_GIRSAV", 0, CADAMT(1)   '/ 신용카드(학원,지로)
        End If
        oDS_ZPY501H.setValue "U_CA1SAV", 0, CSHCAD(0) + CSHCAD(1)
        oDS_ZPY501H.setValue "U_CSHSAV", 0, CSHAMT(0)   '/ 현금영수증
    
        '/ 기부금 재계산
        TMPAMT = Val(oDS_ZPY501H.GetValue("U_LAWGBU", 0)) + Val(oDS_ZPY501H.GetValue("U_POCGBU", 0)) + _
                 Val(oDS_ZPY501H.GetValue("U_SP1GBU", 0)) + Val(oDS_ZPY501H.GetValue("U_SP2GBU", 0)) + _
                 Val(oDS_ZPY501H.GetValue("U_USJGBU", 0)) + Val(oDS_ZPY501H.GetValue("U_JNGGBU", 0))
        
        If (Val(TMPAMT) + Val(oDS_ZPY501H.GetValue("U_JIJGBU", 0))) <> (GBUAMT(0) + GBUAMT(1) - U_JI2GBU) Then
            '/ 기부금총액> 총기부금-지정기부금
            If (Val(GBUAMT(0)) + Val(GBUAMT(1))) > (TMPAMT + U_JI2GBU + Val(oDS_ZPY501H.GetValue("U_JIJGBU", 0))) Then '/ 그외기부금
                oDS_ZPY501H.setValue "U_JIJGBU", 0, GBUAMT(0) + GBUAMT(1) - U_JI2GBU - TMPAMT '/ 기부금
            ElseIf Val(GBUAMT(0)) = 0 And Val(GBUAMT(1)) = 0 Then
                oDS_ZPY501H.setValue "U_LAWGBU", 0, 0
                oDS_ZPY501H.setValue "U_POCGBU", 0, 0
                oDS_ZPY501H.setValue "U_SP1GBU", 0, 0
                oDS_ZPY501H.setValue "U_SP2GBU", 0, 0
                oDS_ZPY501H.setValue "U_USJGBU", 0, 0
                oDS_ZPY501H.setValue "U_JNGGBU", 0, 0
                oDS_ZPY501H.setValue "U_JIJGBU", 0, 0
                oDS_ZPY501H.setValue "U_JI2GBU", 0, 0
            End If
        End If
         
         oDS_ZPY501H.setValue "U_JI2GBU", 0, U_JI2GBU '/ 기부금(노조)
    
        If U_DAGYSU < 2 Then U_DAGYSU = 0
        '/ 인적공제
        oDS_ZPY501H.setValue "U_BAEWOO", 0, U_BAEWOO
        oDS_ZPY501H.setValue "U_BUYN20", 0, U_BUYN20
        oDS_ZPY501H.setValue "U_BUYN60", 0, U_BUYN60
        oDS_ZPY501H.setValue "U_GYNGLO", 0, U_GYNGLO
        oDS_ZPY501H.setValue "U_GYNGL2", 0, U_GYNGL2
        oDS_ZPY501H.setValue "U_GYNGLO", 0, U_GYNGLO
        oDS_ZPY501H.setValue "U_JANGAE", 0, U_JANGAE
        oDS_ZPY501H.setValue "U_MZBURI", 0, U_MZBURI
        oDS_ZPY501H.setValue "U_BUYN06", 0, U_BUYN06
        oDS_ZPY501H.setValue "U_DAGYSU", 0, U_DAGYSU
        oDS_ZPY501H.setValue "U_CHLSAN", 0, U_CHLSAN
        oDS_ZPY501H.setValue "U_INJTOT", 0, U_INJTOT
    
    
    End If
     
    '/ 부양가족 명세 집계
    oDS_ZPY501H.setValue "U_BONBOH1", 0, BOHAMT(0)
    oDS_ZPY501H.setValue "U_BONBOH2", 0, BOHAMT(1)
    oDS_ZPY501H.setValue "U_BONMED1", 0, MEDAMT(0)
    oDS_ZPY501H.setValue "U_BONMED2", 0, MEDAMT(1)
    oDS_ZPY501H.setValue "U_BONEDC1", 0, EDCAMT(0)
    oDS_ZPY501H.setValue "U_BONEDC2", 0, EDCAMT(1)
    oDS_ZPY501H.setValue "U_BONCAD1", 0, CADAMT(0)
    oDS_ZPY501H.setValue "U_BONCAD2", 0, CADAMT(1)
    oDS_ZPY501H.setValue "U_BONCA11", 0, CSHCAD(0)
    oDS_ZPY501H.setValue "U_BONCA12", 0, CSHCAD(1)
    oDS_ZPY501H.setValue "U_BONCSH1", 0, CSHAMT(0)
    oDS_ZPY501H.setValue "U_BONGBU1", 0, GBUAMT(0)
    oDS_ZPY501H.setValue "U_BONGBU2", 0, GBUAMT(1)
    oForm.Freeze False
    oForm.Update
    
    If oForm.Mode = fm_OK_MODE Then oForm.Mode = fm_UPDATE_MODE
End Sub

'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oRow As Long, Optional Insert_YN As Boolean)
        If Insert_YN = False Then
            oDS_ZPY501L.InsertRecord (oRow)
        End If
        oDS_ZPY501L.Offset = oRow
        oDS_ZPY501L.setValue "U_LineNum", oRow, oRow + 1
        oDS_ZPY501L.setValue "U_FamNam", oRow, ""
        oDS_ZPY501L.setValue "U_FamPer", oRow, ""
        If oRow = 0 Then
        oDS_ZPY501L.setValue "U_ChkCod", oRow, "0"
        Else
        oDS_ZPY501L.setValue "U_ChkCod", oRow, ""
        End If
        oDS_ZPY501L.setValue "U_ChkInt", oRow, "1"
        oDS_ZPY501L.setValue "U_ChkBas", oRow, "1"
        oDS_ZPY501L.setValue "U_ChkbuY", oRow, "0"
        oDS_ZPY501L.setValue "U_ChkJan", oRow, "0"
        oDS_ZPY501L.setValue "U_ChkJeL", oRow, "0"
        oDS_ZPY501L.setValue "U_ChkChl", oRow, "0"
        oDS_ZPY501L.setValue "U_ChkDaJ", oRow, "0"
        oDS_ZPY501L.setValue "U_ChkCHS", oRow, "0"
        oDS_ZPY501L.setValue "U_BOHGBN", oRow, "1"
        oDS_ZPY501L.setValue "U_BOHAMT1", oRow, 0
        oDS_ZPY501L.setValue "U_BOHAMT2", oRow, 0
        oDS_ZPY501L.setValue "U_MEDAMT1", oRow, 0
        oDS_ZPY501L.setValue "U_MEDAMT2", oRow, 0
        oDS_ZPY501L.setValue "U_SCHGBN", oRow, "1"
        oDS_ZPY501L.setValue "U_EDCAMT1", oRow, 0
        oDS_ZPY501L.setValue "U_EDCAMT2", oRow, 0
        oDS_ZPY501L.setValue "U_CADAMT1", oRow, 0
        oDS_ZPY501L.setValue "U_CADAMT2", oRow, 0
        oDS_ZPY501L.setValue "U_CSHCAD1", oRow, 0
        oDS_ZPY501L.setValue "U_CSHCAD2", oRow, 0
        oDS_ZPY501L.setValue "U_CSHAMT1", oRow, 0
        oDS_ZPY501L.setValue "U_GBUAMT2", oRow, 0
        
        oMat1.LoadFromDataSource
End Sub
