VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP001.cls
'//  Module         : PP
'//  Desc           : 공정코드등록
'//  FormType       : PS_PP001
'//  Create Date    : 2010.09.27
'//  Modified Date  :
'//  Creator        : Dong sub Kim
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit         '모듈 수준에서 해당하는 모듈의 모든 변수를 명시적으로 선언하도록 할 때 사용합니다.

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_PP001H       As SAPbouiCOM.DBDataSource    '등록헤더 'Represents a SAP Business One database table to be attached to a form.
Private oDS_PS_PP001L       As SAPbouiCOM.DBDataSource    '등록라인
Private oDS_PS_PP001M       As SAPbouiCOM.DBDataSource

Private oLast_Item_UID   As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID    As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row    As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private oSeq&
Private TmpCode As String
'Private oFormMode As String

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP001.srf")
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
    
    oFormUniqueID01 = "PS_PP001_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    ' 변수나 속성 등에 개체 참조를 지정합니다

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocEntry"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////D////////////////////////////////////////////////
        
    oForm01.Freeze True
    Call CreateItems
    
    oForm01.EnableMenu ("1283"), True      '// 제거
    oForm01.EnableMenu ("1293"), True      '// 행삭제
    oForm01.EnableMenu ("1287"), True      '// 복제
    oForm01.EnableMenu ("1284"), False     '// 취소
    
    Call ComboBox_Setting
    Call CF_ChooseFromList
    Call FormItemEnabled
    Call FormClear '//UDO방식일때
    Call AddMatrixRow(0, oMat01.RowCount, True) '//UDO방식일때
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    
    oForm01.Update      ' Reloads the form from the application.
                        ' Use this method to update an object from the database.

    oForm01.Freeze False
       
    oForm01.Visible = True  ' Indicates whether the form is visible.
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
    Dim sQry As String
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        '// Input Code, Name
                        oDS_PS_PP001H.setValue "Code", 0, Trim(oDS_PS_PP001H.GetValue("U_CpBCode", 0))
                        oDS_PS_PP001H.setValue "Name", 0, Trim(oDS_PS_PP001H.GetValue("U_CpBName", 0))
                        TmpCode = Trim(oDS_PS_PP001H.GetValue("U_CpBCode", 0))
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
            
                ' 공정코드
                If oForm01.Items("CpBCode").Specific.VALUE = "" Then
                    If pval.ItemUID = "CpBCode" And pval.CharPressed = 9 Then
                        oForm01.Items("CpBCode").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' 소속코드
                If pval.ItemUID = "Mat01" And pval.ColUID = "PartCode" And pval.CharPressed = 9 Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                 End If
                
                ' 작업반코드
                If pval.ItemUID = "Mat01" And pval.ColUID = "WkClsCod" And pval.CharPressed = 9 Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                 End If
                
                ' 부서코드
                If pval.ItemUID = "Mat01" And pval.ColUID = "DeptCode" And pval.CharPressed = 9 Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' 단위
                If pval.ItemUID = "Mat01" And pval.ColUID = "Unit" And pval.CharPressed = 9 Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' 대분류코드
                If pval.ItemUID = "Mat01" And pval.ColUID = "ItmBsort" And pval.CharPressed = 9 Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "Mat01" Then
                    If (pval.ColUID = "CCCode") Then
                        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "CCCode") '//사용자값활성
                    ElseIf (pval.ColUID = "ActCode1") Then
                        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ActCode1") '//사용자값활성
                    ElseIf (pval.ColUID = "ActCode2") Then
                        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ActCode2") '//사용자값활성
                    ElseIf (pval.ColUID = "ActCode3") Then
                        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ActCode3") '//사용자값활성
                    ElseIf (pval.ColUID = "ActCode4") Then
                        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ActCode4") '//사용자값활성
                    ElseIf (pval.ColUID = "ActCode5") Then
                        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ActCode5") '//사용자값활성
                    ElseIf (pval.ColUID = "ActCode6") Then
                        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ActCode6") '//사용자값활성
                    ElseIf (pval.ColUID = "ActCode7") Then
                        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ActCode7") '//사용자값활성
                    End If
                End If
                
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            
                If pval.ItemUID = "Mat01" Then

                    If pval.Row > 0 Then
        
                        Call oMat01.SelectRow(pval.Row, True, False)
        
                    End If
                
                End If
            
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then                '
                    If pval.ItemUID = "Mat01" And pval.ColUID = "CpCode" Then
                        FlushToItemValue pval.ItemUID, pval.ColUID, pval.Row, pval.ColUID
                    End If
                End If

            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
                 oLast_Item_UID = pval.ItemUID

            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
            '//메트릭스에 데이터 로드
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
              If pval.ItemUID = "1" Then
                If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                    oForm01.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")
                ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = False Then
                    FormItemEnabled
                    AddMatrixRow 1, oMat01.RowCount, True
                End If
              End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    ' 공정분류
                    If pval.ItemUID = "CpBCode" And pval.ItemChanged = True Then
                        sQry = "Select U_CdName From [@PS_SY001L] Where U_Minor = '" & Trim(oForm01.Items("CpBCode").Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oForm01.Items("CpBName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    '팀명
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "DeptCode")) Then
                        
                        sQry = "           SELECT      T0.U_CodeNm"
                        sQry = sQry & " FROM       [@PS_HR200L] AS T0"
                        sQry = sQry & " WHERE      T0.Code = '1'"
                        sQry = sQry & "                AND T0.U_UseYN = 'Y'"
                        sQry = sQry & "                AND T0.U_Code = '" & Trim(oMat01.Columns("DeptCode").Cells(pval.Row).Specific.VALUE) & "'"
                        
                        'Select  Name From [OUDP] Where Code = '" & Trim(oMat01.Columns("DeptCode").Cells(pval.Row).Specific.VALUE) & "'"
                        
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("DeptName").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        
                    End If
                    
                    '담당명
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "PartCode")) Then
                    
                        sQry = "           SELECT      T0.U_CodeNm"
                        sQry = sQry & " FROM       [@PS_HR200L] AS T0"
                        sQry = sQry & " WHERE      T0.Code = '2'"
                        sQry = sQry & "                AND T0.U_UseYN = 'Y'"
                        sQry = sQry & "                AND T0.U_Code = '" & Trim(oMat01.Columns("PartCode").Cells(pval.Row).Specific.VALUE) & "'"
                        sQry = sQry & "                AND T0.U_Char1 = '" & Trim(oMat01.Columns("DeptCode").Cells(pval.Row).Specific.VALUE) & "'"
                    
'                        sQry = "Select U_CdName From [@PS_SY001L] Where Code = 'P003'"
'                        sQry = sQry & " And U_Minor = '" & Trim(oMat01.Columns("PartCode").Cells(pval.Row).Specific.VALUE) & "'"
                        
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("PartName").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        
                    End If
                    
                    '반명
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "WkClsCod")) Then
                    
                        sQry = "           SELECT      T0.U_CodeNm"
                        sQry = sQry & " FROM       [@PS_HR200L] AS T0"
                        sQry = sQry & " WHERE      T0.Code = '9'"
                        sQry = sQry & "                AND T0.U_UseYN = 'Y'"
                        sQry = sQry & "                AND T0.U_Code = '" & Trim(oMat01.Columns("WkClsCod").Cells(pval.Row).Specific.VALUE) & "'"
                        sQry = sQry & "                AND T0.U_Char1 = '" & Trim(oMat01.Columns("PartCode").Cells(pval.Row).Specific.VALUE) & "'"
                        sQry = sQry & "                AND T0.U_Char2 = '" & Trim(oMat01.Columns("DeptCode").Cells(pval.Row).Specific.VALUE) & "'"
                        
'                        sQry = "Select U_CdName From [@PS_SY001L] Where Code = 'P002'"
'                        sQry = sQry & " And U_Minor = '" & Trim(oMat01.Columns("WkClsCod").Cells(pval.Row).Specific.VALUE) & "'"
                        
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("WkClsNam").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        
                    End If
                    
                    ' 대분류명
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "ItmBsort")) Then
                        sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" & Trim(oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("ItmBname").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    ' 코스트센터명
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "CCCode")) Then
                        sQry = "Select PrcName From [OPRC] Where PrcCode = '" & Trim(oMat01.Columns("CCCode").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("CCName").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    ' 활동유형이름1
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "ActCode1")) Then
                        sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" & Trim(oMat01.Columns("ActCode1").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("ActName1").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    ' 활동유형이름2
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "ActCode2")) Then
                        sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" & Trim(oMat01.Columns("ActCode2").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("ActName2").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    ' 활동유형이름3
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "ActCode3")) Then
                        sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" & Trim(oMat01.Columns("ActCode3").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("ActName3").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    ' 활동유형이름4
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "ActCode4")) Then
                        sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" & Trim(oMat01.Columns("ActCode4").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("ActName4").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    ' 활동유형이름5
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "ActCode5")) Then
                        sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" & Trim(oMat01.Columns("ActCode5").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("ActName5").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    ' 활동유형이름6
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "ActCode6")) Then
                        sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" & Trim(oMat01.Columns("ActCode6").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("ActName6").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    ' 활동유형이름7
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "ActCode7")) Then
                        sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" & Trim(oMat01.Columns("ActCode7").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("ActName7").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                End If

            Case et_MATRIX_LOAD: '//11
                AddMatrixRow 1, oMat01.VisualRowCount, True
                Call oMat01.AutoResizeColumns
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
                oLast_Item_UID = pval.ItemUID
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    If (pval.BeforeAction = True) Then      '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            Case "1293"     '행삭제
        End Select
        
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
                
            Case "1281": '찾기
                FormItemEnabled
                
            Case "1282": '추가
                FormItemEnabled
                FormClear
                AddMatrixRow 0, oMat01.RowCount, True
                oForm01.Items("CpBCode").Click ct_Regular
                
            Case "1287": '복제
                oForm01.Freeze True
                oDS_PS_PP001H.setValue "Code", 0, ""
                oDS_PS_PP001H.setValue "Name", 0, ""
                oDS_PS_PP001H.setValue "U_CpBCode", 0, ""
                oDS_PS_PP001H.setValue "U_CpBName", 0, ""
         
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_PP001L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("LineNum").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        AddMatrixRow 1, oMat01.RowCount, True
'                    End If
'                End If
                
            Case "1293"     '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                    Next i
                    oMat01.FlushToDataSource        ' DBDataSource에 레코드가 한줄 더 생긴다.
                    Call oDS_PS_PP001L.RemoveRecord(oDS_PS_PP001L.Size - 1)     ' 레코드 한 줄을 지운다.
                    oMat01.LoadFromDataSource       ' DBDataSource를 매트릭스에 올리고
                    If oMat01.RowCount = 0 Then
                        Call AddMatrixRow(1, 0, True)
                    Else
                        If Trim(oDS_PS_PP001L.GetValue("U_CpCode", oMat01.RowCount - 1)) <> "" Then
                            Call AddMatrixRow(1, oMat01.RowCount, True)

                        End If
                    End If
                End If
        End Select
    End If
    Exit Sub
Raise_MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
'                FormItemEnabled
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
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

Function CreateItems() As Boolean
On Error GoTo CreateItems_Error
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
        
    Set oDS_PS_PP001H = oForm01.DataSources.DBDataSources("@PS_PP001H")
    Set oDS_PS_PP001L = oForm01.DataSources.DBDataSources("@PS_PP001L")
    
    Set oRecordSet01 = Nothing
    Exit Function
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Exit Sub
ComboBox_Setting_Error:
    Sbo_Application.SetStatusBarMessage "ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Exit Sub
CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
'        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("CpBCode").Enabled = True
'        oForm01.Items("ItmMcode").Enabled = True
'        oForm01.Items("CntcCode").Enabled = True
'        oMat01.Columns("Price1").Editable = True
'        oMat01.Columns("Price2").Editable = True
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
'        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("CpBCode").Enabled = True
'        oForm01.Items("ItmMcode").Enabled = True
'        oForm01.Items("CntcCode").Enabled = True
'        oMat01.Columns("Price1").Editable = False
'        oMat01.Columns("Price2").Editable = False
        
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
'            oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("CpBCode").Enabled = False
'        oForm01.Items("ItmMcode").Enabled = False
'        oForm01.Items("CntcCode").Enabled = False
'        oMat01.Columns("Price1").Editable = True
'        oMat01.Columns("Price2").Editable = True
    End If
    Exit Sub
FormItemEnabled_Error:
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oSeq As Integer, ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo AddMatrixRow_Error
   Select Case oSeq
       Case 0:
            'oMat01.AddRow
            'oMat01.Columns("LineNum").Cells(oRow + 1).Specific.Value = oRow + 1
            oDS_PS_PP001L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
       Case 1:
            oDS_PS_PP001L.InsertRecord oRow
            oDS_PS_PP001L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
   End Select
'AddMatrixRow_Error:
'    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP001'", "")
    If DocNum = 0 Then
'        oForm01.Items("DocEntry").Specific.String = 1
        oDS_PS_PP001H.setValue "DocEntry", 0, "1"
    Else
'        oForm01.Items("DocEntry").Specific.String = DocNum
        oDS_PS_PP001H.setValue "DocEntry", 0, DocNum    ' 화면에 적용이 안되기 때문
    End If
    Exit Sub
FormClear_Error:
    Sbo_Application.SetStatusBarMessage "FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function DataValidCheck() As Boolean
On Error GoTo DataValidCheck_Error
    '//유효성검사
    Exit Function
DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_PP001H.GetValue("U_CpBCode", 0) = "" Or oDS_PS_PP001H.GetValue("U_CpBName", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    HeaderSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "대분류 또는 대분류명은 필수입력 사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

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
    If oMat01.VisualRowCount <= 1 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '// 맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '// 이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 2
            oDS_PS_PP001L.Offset = i
            Select Case True
                Case oDS_PS_PP001L.GetValue("U_CpCode", i) = ""
                    ErrNum = 2
                    GoTo MatrixSpaceLineDel_Error
    
                Case oDS_PS_PP001L.GetValue("U_CpName", i) = ""
                    ErrNum = 3
                    GoTo MatrixSpaceLineDel_Error
                    
                Case oDS_PS_PP001L.GetValue("U_PartCode", i) = ""
                    ErrNum = 4
                    GoTo MatrixSpaceLineDel_Error
            End Select
        Next i
        
        If oDS_PS_PP001L.GetValue("U_CpCode", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_PP001L.RemoveRecord oMat01.VisualRowCount - 1
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
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "공정코드 데이터는 필수입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "공정명 데이터는 필수입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "소속코드는 필수입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Sub FlushToItemValue(ByVal oUID As String, ByVal oCID As String, Optional oRow As Long, Optional oCol As String)
    Dim i$
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "Mat01"
            Select Case oCID
                Case "CpCode"
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 2) And _
                    Trim(oMat01.Columns("CpCode").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        AddMatrixRow 1, oMat01.RowCount, True
                        oMat01.Columns("CpCode").Cells(oRow).Click ct_Regular
                    End If
'                Case "PartCode"
'                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 2) And _
'                    Trim(oMat01.Columns("PartCode").Cells(oRow).Specific.Value) <> "" Then
'                        oMat01.FlushToDataSource
'                        AddMatrixRow 1, oMat01.RowCount, True
'                        oMat01.Columns("PartCode").Cells(oRow).Click ct_Regular
'                    End If
'                Case "WkClsCod"
'                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 2) And _
'                    Trim(oMat01.Columns("WkClsCod").Cells(oRow).Specific.Value) <> "" Then
'                        oMat01.FlushToDataSource
'                        AddMatrixRow 1, oMat01.RowCount, True
'                        oMat01.Columns("WkClsCod").Cells(oRow).Click ct_Regular
'                    End If
'                Case "DeptCode"
'                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 2) And _
'                    Trim(oMat01.Columns("DeptCode").Cells(oRow).Specific.Value) <> "" Then
'                        oMat01.FlushToDataSource
'                        AddMatrixRow 1, oMat01.RowCount, True
'                        oMat01.Columns("DeptCode").Cells(oRow).Click ct_Regular
'                    End If
'                Case "Unit"
'                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 2) And _
'                    Trim(oMat01.Columns("Unit").Cells(oRow).Specific.Value) <> "" Then
'                        oMat01.FlushToDataSource
'                        AddMatrixRow 1, oMat01.RowCount, True
'                        oMat01.Columns("Unit").Cells(oRow).Click ct_Regular
'                    End If
           End Select
    End Select
    Set oRecordSet01 = Nothing
End Sub




