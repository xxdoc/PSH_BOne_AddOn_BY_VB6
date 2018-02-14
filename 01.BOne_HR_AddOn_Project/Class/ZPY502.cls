VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY502"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY502.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 종(전)근무지 등록
'//  FormType       : 2010110502
'//  Create Date    : 2006.01.15
'//  Modified Date  :
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oDS_ZPY502H     As SAPbouiCOM.DBDataSource    '시스템코드 헤더
Private oDS_ZPY502L     As SAPbouiCOM.DBDataSource    '시스템코드 라인
Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private Col_Last_Uid    As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private Col_Last_Row    As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private oOLDCHK         As String


Private Sub FormItemEnabled()
    If oForm.Mode = fm_FIND_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = True
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("DocNum").Enabled = True
        oForm.Items("ENDCHK").Enabled = True
        Call CLTCOD_Select(oForm, "CLTCOD")
        
    ElseIf oForm.Mode = fm_ADD_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = True
        oForm.Items("ENDCHK").Enabled = True
        Call CLTCOD_Select(oForm, "CLTCOD")

    ElseIf oForm.Mode = fm_OK_MODE Then
        oForm.Items("JSNYER").Enabled = False
        oForm.Items("MSTCOD").Enabled = False
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("DocNum").Enabled = False
        '// 급여월마감된것은 비활성화
        oOLDCHK = oDS_ZPY502H.GetValue("U_ENDCHK", 0)
        If MDC_SetMod.Get_ReData("U_ENDCHK", "U_JOBYER", "[@ZPY509L]", "'" & oDS_ZPY502H.GetValue("U_JSNYER", 0) & "'", " AND Code = '" & oDS_ZPY502H.GetValue("U_CLTCOD", 0) & "'") = "Y" Then
            oForm.Items("ENDCHK").Enabled = False
        Else
            oForm.Items("ENDCHK").Enabled = True
        End If

    End If
End Sub

Private Sub FormClear()
    Dim DocNum&
    
    DocNum = MDC_SetMod.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'ZPY502'", "")
    
    If DocNum = 0 Then
        oForm.Items("DocNum").Specific.String = 1
    Else
        oForm.Items("DocNum").Specific.String = DocNum
    End If
    Call FlushToItemValue("JSNYER")
    
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim oMast   As ZPAY_g_EmpID
    Dim TOTAMT  As Double
    Dim JSNYER  As String
    Select Case oUID
    Case "JSNYER"
          If Trim$(oForm.Items(oUID).Specific.String) <> "" Then
              ZPAY_GBL_JSNYER = oForm.Items(oUID).Specific.String
          Else
              oDS_ZPY502H.setValue "U_JSNYER", 0, ZPAY_GBL_JSNYER
          End If
          oForm.Items(oUID).Update
          Call Matrix_TitleSetting
    Case "MSTCOD"
          If oForm.Items(oUID).Specific.String = "" Then
              oDS_ZPY502H.setValue "U_MSTCOD", 0, ""
              oDS_ZPY502H.setValue "U_MSTNAM", 0, ""
              oDS_ZPY502H.setValue "U_EmpID", 0, ""
              oDS_ZPY502H.setValue "U_CLTCOD", 0, ""
          Else
              oDS_ZPY502H.setValue "U_MSTCOD", 0, UCase$(oForm.Items(oUID).Specific.String)
              oMast = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
              oDS_ZPY502H.setValue "U_MSTNAM", 0, oMast.MSTNAM
              oDS_ZPY502H.setValue "U_EmpID", 0, oMast.EmpID
              oDS_ZPY502H.setValue "U_CLTCOD", 0, oMast.CLTCOD
          End If
          oForm.Items("MSTNAM").Update
          oForm.Items("EmpID").Update
          oForm.Items("CLTCOD").Update
          oForm.Items(oUID).Update
    End Select
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If Left$(oUID, 3) = "Col" Then
        oMat1.FlushToDataSource
    
        oDS_ZPY502L.Offset = oRow - 1
        
        Select Case oUID
        Case "Col1", "Col2", "Col3"
'            oMat1.SetLineData oRow
'
            If oRow = oMat1.RowCount And Trim$(oDS_ZPY502L.GetValue("U_JONNAM", oRow - 1)) <> "" Then
                Matrix_AddRow oRow
                oMat1.Columns("Col1").Cells(oRow).CLICK ct_Regular
            End If
        Case "Col5", "Col6", "Col7", "Col13", "Col17", "Col20", "Col23", "Col24", "Col25", "Col26", _
             "Col27", "Col28", "Col29", "Col30", "Col31", "Col32", "Col33", "Col34", "Col35", "Col36", "Col37", "Col38", _
             "Col39", "Col40", "Col41", "Col42", "Col43", "Col44", "Col45", "Col46", "Col47", "Col48", "Col50"
             
            JSNYER = Trim$(oDS_ZPY502H.GetValue("U_JSNYER", 0))
            TOTAMT = 0
            If JSNYER <= "2008" Then
                TOTAMT = Val(oDS_ZPY502L.GetValue("U_JONBT1", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT2", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT3", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JONBU3", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT4", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT5", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JONBT6", oRow - 1))
             
            ElseIf JSNYER = "2009" Then
                TOTAMT = Val(oDS_ZPY502L.GetValue("U_JONBT1", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT2", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT3", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JONBU3", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT4", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT5", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JONBT6", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTG01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH01", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTH05", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH06", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH07", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTH08", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH09", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH10", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTH11", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH12", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH13", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTI01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTK01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTM01", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTM02", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTM03", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTO01", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTQ01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTS01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTT01", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTX01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTY01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTY02", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTY03", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTY20", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTZ01", oRow - 1))
            ElseIf JSNYER >= "2010" Then
                TOTAMT = Val(oDS_ZPY502L.GetValue("U_JONBT1", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT2", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT3", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JONBU3", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT4", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JONBT5", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JONBT6", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTG01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH01", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTH05", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH06", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH07", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTH08", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH09", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH10", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTH11", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH12", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTH13", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTI01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTK01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTM01", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTM02", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTM03", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTO01", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTQ01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTS01", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTT01", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTY02", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTY03", oRow - 1)) + Val(oDS_ZPY502L.GetValue("U_JBTY21", oRow - 1)) + _
                       Val(oDS_ZPY502L.GetValue("U_JBTZ01", oRow - 1))
            End If

            oDS_ZPY502L.setValue "U_JBTTOT", oRow - 1, TOTAMT
            oMat1.SetLineData oRow
        End Select
    End If
End Sub

Private Function MatrixSpaceLineDel() As Boolean
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'저장할 데이터의 유효성을 점검한다
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
On Error GoTo Error_Message
    Dim iRow        As Long
    Dim kRow        As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    Dim Chk_Data1   As String
    Dim JSNYER      As String
    ErrNum = 0
    '/ 헤더부분 체크
    Select Case True
    Case oDS_ZPY502H.GetValue("U_JSNYER", 0) = ""
         ErrNum = 4
         GoTo Error_Message
    Case oDS_ZPY502H.GetValue("U_MSTCOD", 0) = ""
         ErrNum = 5
         GoTo Error_Message
    End Select
    JSNYER = oDS_ZPY502H.GetValue("U_JSNYER", 0)
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat1.FlushToDataSource
        
    If oMat1.RowCount = 1 Then          '// Mat1에 값이 있는지 확인 (ErrorNumber : 1)
        ErrNum = 1
        GoTo Error_Message
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat1.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    For iRow = 0 To oMat1.VisualRowCount - 2      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 3)
        oDS_ZPY502L.Offset = iRow
        If Trim$(oDS_ZPY502L.GetValue("U_JONNAM", iRow)) = "" Then
            ErrNum = 2
            oMat1.Columns("Col1").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY502L.GetValue("U_JONNBR", iRow)) = "" Then
            ErrNum = 6
            oMat1.Columns("Col2").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY502L.GetValue("U_JONSTR", iRow)) = "" Then
            ErrNum = 7
            oMat1.Columns("Col18").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY502L.GetValue("U_JONEND", iRow)) = "" Then
            ErrNum = 8
            oMat1.Columns("Col19").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Left$(oDS_ZPY502L.GetValue("U_JONSTR", iRow), 4) <> Trim$(JSNYER) Or Left$(oDS_ZPY502L.GetValue("U_JONEND", iRow), 4) <> Trim$(JSNYER) Then
            ErrNum = 9
            oMat1.Columns("Col18").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf JSNYER >= "2010" Then    '// 2010년 폐지 비과세 항목에 금액이 입력된 경우 금액을 초기화 시킴
            If Val(oDS_ZPY502L.GetValue("U_JBTX01", iRow)) <> 0 Then
                oDS_ZPY502L.Offset = iRow
                oDS_ZPY502L.setValue "U_JBTX01", iRow, 0
                oMat1.SetLineData (iRow + 1)
            End If
            If Val(oDS_ZPY502L.GetValue("U_JBTY01", iRow)) <> 0 Then
                oDS_ZPY502L.Offset = iRow
                oDS_ZPY502L.setValue "U_JBTY01", iRow, 0
                oMat1.SetLineData (iRow + 1)
            End If
            If Val(oDS_ZPY502L.GetValue("U_JBTY20", iRow)) <> 0 Then
                oDS_ZPY502L.Offset = iRow
                oDS_ZPY502L.setValue "U_JBTY20", iRow, 0
                oMat1.SetLineData (iRow + 1)
            End If
        Else
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '중복체크작업
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            Chk_Data = Trim$(oDS_ZPY502L.GetValue("U_JONNBR", iRow))
            Chk_Data1 = Trim$(oDS_ZPY502L.GetValue("U_JONSTR", iRow))
            For kRow = iRow + 1 To oMat1.VisualRowCount - 2
                oDS_ZPY502L.Offset = kRow
                If Trim$(Chk_Data) = Trim$(oDS_ZPY502L.GetValue("U_JONNBR", kRow)) And _
                   Trim$(Chk_Data1) = Trim$(oDS_ZPY502L.GetValue("U_JONSTR", kRow)) Then
                    ErrNum = 3
                    oMat1.Columns("Col1").Cells(iRow + 1).CLICK ct_Regular
                    GoTo Error_Message
                End If
            Next kRow
        End If
    Next iRow
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oDS_ZPY502L.RemoveRecord oDS_ZPY502L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat1.LoadFromDataSource
    
    MatrixSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "입력할 데이터가 없습니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "근무처명은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "근무처 사업자번호/귀속시작일이 중복입력되었습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "귀속년도는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "사원번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "사업자번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "귀속시작일은 필수입니다. (귀속연도+0101 또는 종전근무지의 당해입사일)을 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 8 Then
        Sbo_Application.StatusBar.SetText "귀속종료일은 필수입니다. (종전근무지의 퇴사일)을 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 9 Then
        Sbo_Application.StatusBar.SetText "귀속연도와 귀속시작일 또는 귀속종료일의 연도가 일치하지 않습니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 10 Then
        Sbo_Application.StatusBar.SetText "2010년부터 폐지된 비과세 항목에 금액이 입력되어 있습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    End If
    
    MatrixSpaceLineDel = False
End Function


Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim DocNum          As String
    
    ErrNum = 0
 '/ Check
    Select Case True
    Case oDS_ZPY502H.GetValue("U_JSNYER", 0) = ""
         ErrNum = 1
         GoTo Error_Message
    Case oDS_ZPY502H.GetValue("U_MSTCOD", 0) = ""
         ErrNum = 2
         GoTo Error_Message
    Case oDS_ZPY502H.GetValue("U_CLTCOD", 0) = ""
         ErrNum = 3
         GoTo Error_Message
    End Select
    
    If Trim$(oDS_ZPY502H.GetValue("U_ENDCHK", 0)) = "Y" And Trim$(oOLDCHK) = "Y" Then
        ErrNum = 5
        GoTo Error_Message
    End If
    
    DocNum = Exist_YN(oDS_ZPY502H.GetValue("U_JSNYER", 0), oDS_ZPY502H.GetValue("U_MSTCOD", 0), oDS_ZPY502H.GetValue("U_CLTCOD", 0))
    If Trim$(DocNum) <> "" And Trim$(oDS_ZPY502H.GetValue("DocNum", 0)) <> Trim$(DocNum) Then
        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        '같은데이터가 존재하는데 자기 자신이 현재 자기자신이 아니라면(같은월에는 취소한거 아니면 하나만 존재해야함)
        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        ErrNum = 4
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
        Sbo_Application.StatusBar.SetText "자사코드는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "문서번호" & DocNum & " 와(과) 데이터가 일치합니다. 저장되지 않습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "잠금 자료입니다. 확인하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "HeaderSpaceLineDel 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    HeaderSpaceLineDel = False
End Function

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
                    '//추가및 업데이트시에
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                        '//추가및 업데이트시에
                        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                            If MatrixSpaceLineDel = False Then
                                BubbleEvent = False
                            End If
                        End If
                    End If
            ElseIf pval.ItemUID = "CBtn1" And oForm.Items("MSTCOD").Enabled = True Then   '/ ChooseBtn사원리스트
                oForm.Items("MSTCOD").CLICK ct_Regular
                Sbo_Application.ActivateMenuItem ("7425")
                BubbleEvent = False
            End If
        Else
            If pval.ItemUID = "1" And pval.ActionSuccess = True And oForm.Mode = fm_ADD_MODE Then
                oForm.Mode = fm_OK_MODE
                Call Sbo_Application.ActivateMenuItem("1282")
            End If
        End If
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_CLICK
        If pval.BeforeAction = True And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" And oForm.Mode <> fm_FIND_MODE Then
            If Last_Item = "MSTCOD" Then
                If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(Last_Item).Specific.String & "'", "") = True _
                And oForm.Items(Last_Item).Specific.String <> "" And Last_Item <> pval.ItemUID Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
        If pval.FormUID = oForm.uniqueID And pval.BeforeAction = True And Last_Item = "Mat1" And Col_Last_Uid = "Col1" And Col_Last_Row > 0 _
          And (Col_Last_Uid <> pval.ColUID Or Col_Last_Row <> pval.Row) And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" Then
            If Col_Last_Row > oMat1.VisualRowCount Then
                Exit Sub
            End If
        End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_VALIDATE
        If pval.BeforeAction = False And pval.ItemChanged = True And _
        (pval.ItemUID = "MSTCOD" Or pval.ItemUID = "JSNYER") Then
            FlushToItemValue pval.ItemUID
        ElseIf pval.BeforeAction = False And pval.ItemChanged = True And pval.ItemUID = "Mat1" _
        And Left$(pval.ColUID, 3) = "Col" Then
            FlushToItemValue pval.ColUID, pval.Row
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
                If oMat1.RowCount > 0 Then
                    oMat1.Columns("Col1").Cells(oMat1.VisualRowCount).CLICK ct_Regular
                    BubbleEvent = False
                End If
            End If
         ElseIf pval.BeforeAction = True And pval.ColUID = "Col1" And pval.CharPressed = 9 Then
            If Trim$(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = "" Then
                Call Sbo_Application.StatusBar.SetText("근무처명은 필수입니다. 입력하세요.", bmt_Short, smt_Error)
                BubbleEvent = False
            End If
         ElseIf pval.BeforeAction = True And pval.ColUID = "Col2" And pval.CharPressed = 9 Then
            If Trim$(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = "" Then
                Call Sbo_Application.StatusBar.SetText("근무처 사업자번호는 필수입니다. 입력하세요.", bmt_Short, smt_Error)
                BubbleEvent = False
            Else
                '/ 사업자번호 체크
                If Len(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) <= 12 Then
                    If MDC_SetMod.TaxNoCheck(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = False Then
                        Call Sbo_Application.StatusBar.SetText("사업자번호가 틀립니다. 확인하세요.", bmt_Short, smt_Error)
                        BubbleEvent = False
                    End If
                End If
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
            Set oDS_ZPY502H = Nothing
            Set oDS_ZPY502L = Nothing
            Set oMat1 = Nothing
        End If
'et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_MATRIX_LOAD:
        If pval.BeforeAction = False Then
            Call Matrix_TitleSetting
            FormItemEnabled
            Call Matrix_AddRow(oMat1.VisualRowCount)
        End If
    End Select
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:", bmt_Short, smt_Error
End Sub


Private Function Exist_YN(JOBYER$, MSTCOD$, CLTCOD$) As String
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'저장할 데이터의 기존데이터가 있는지 확인한다
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "SELECT Top 1 T1.DocNum FROM [@ZPY502H] T1 "
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
'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oRow As Long, Optional Insert_YN As Boolean)
        If Insert_YN = False Then
            oDS_ZPY502L.InsertRecord (oRow)
        End If
        oDS_ZPY502L.Offset = oRow
        oDS_ZPY502L.setValue "U_LineNum", oRow, oRow + 1
        oDS_ZPY502L.setValue "U_JONNAM", oRow, ""
        oDS_ZPY502L.setValue "U_JONNBR", oRow, ""
        oDS_ZPY502L.setValue "U_JONSTR", oRow, ""
        oDS_ZPY502L.setValue "U_JONEND", oRow, ""
        oDS_ZPY502L.setValue "U_JONPAY", oRow, ""
        oDS_ZPY502L.setValue "U_JONBNS", oRow, ""
        oDS_ZPY502L.setValue "U_JONBT1", oRow, ""
        oDS_ZPY502L.setValue "U_JONBT2", oRow, ""
        oDS_ZPY502L.setValue "U_JONBT3", oRow, ""
        oDS_ZPY502L.setValue "U_JONBU3", oRow, ""
        oDS_ZPY502L.setValue "U_JONBT4", oRow, ""
        oDS_ZPY502L.setValue "U_JONBT5", oRow, ""
        oDS_ZPY502L.setValue "U_JONBT6", oRow, ""
        oDS_ZPY502L.setValue "U_JONMED", oRow, ""
        oDS_ZPY502L.setValue "U_JONGBH", oRow, ""
        oDS_ZPY502L.setValue "U_JONKUK", oRow, ""
        oDS_ZPY502L.setValue "U_JONKUE", oRow, ""
        oDS_ZPY502L.setValue "U_JONRET", oRow, ""
        oDS_ZPY502L.setValue "U_JONGAB", oRow, ""
        oDS_ZPY502L.setValue "U_JONJUM", oRow, ""
        '/ 2009년추가
        oDS_ZPY502L.setValue "U_JBTG01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTH01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTH05", oRow, ""
        oDS_ZPY502L.setValue "U_JBTH06", oRow, ""
        oDS_ZPY502L.setValue "U_JBTH07", oRow, ""
        oDS_ZPY502L.setValue "U_JBTH08", oRow, ""
        oDS_ZPY502L.setValue "U_JBTH09", oRow, ""
        oDS_ZPY502L.setValue "U_JBTH10", oRow, ""
        oDS_ZPY502L.setValue "U_JBTH11", oRow, ""
        oDS_ZPY502L.setValue "U_JBTH12", oRow, ""
        oDS_ZPY502L.setValue "U_JBTH13", oRow, ""
        oDS_ZPY502L.setValue "U_JBTI01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTK01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTM01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTM02", oRow, ""
        oDS_ZPY502L.setValue "U_JBTM03", oRow, ""
        oDS_ZPY502L.setValue "U_JBTO01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTQ01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTS01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTT01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTX01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTY01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTY02", oRow, ""
        oDS_ZPY502L.setValue "U_JBTY03", oRow, ""
        oDS_ZPY502L.setValue "U_JBTY20", oRow, ""
        oDS_ZPY502L.setValue "U_JBTY21", oRow, ""
        oDS_ZPY502L.setValue "U_JBTZ01", oRow, ""
        oDS_ZPY502L.setValue "U_JBTTOT", oRow, ""
        
        oMat1.LoadFromDataSource
End Sub
'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i&
    
    If pval.BeforeAction = True Then
        Select Case pval.MenuUID
        Case "1283" '/ 제거
            If Trim$(oDS_ZPY502H.GetValue("U_ENDCHK", 0)) = "Y" Then
                Sbo_Application.StatusBar.SetText "잠금 자료입니다. 확인하여 주십시오.", bmt_Short, smt_Error
                BubbleEvent = False
                Exit Sub
            Else
                If Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
        Case Else
            Exit Sub
        End Select
    Else
    
        Select Case pval.MenuUID
        Case "1287" '/ 복제
           ' oForm.Items("Btn1").Visible = True
        Case "1281", "1282":
            FormItemEnabled
            If pval.MenuUID = "1282" Then
                FormClear
                Call Matrix_AddRow(0, True)
                oForm.Items("JSNYER").CLICK ct_Regular
            End If
        Case "1288" To "1291":
        Case "1293"
            If oMat1.RowCount <> oMat1.VisualRowCount Then
                'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
                '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
                'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                For i = 0 To oMat1.VisualRowCount - 1
                    oMat1.Columns("Col0").Cells(i + 1).Specific.VALUE = i + 1
                Next
                
                oMat1.FlushToDataSource
                oDS_ZPY502L.RemoveRecord oDS_ZPY502L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
                oMat1.Clear
                oMat1.LoadFromDataSource
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

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional JSNYER$, Optional MSTCOD$, Optional CLTCOD$)
'Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY502.srf")
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
    
    oFormUniqueID = "ZPY502_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "ZPY502"
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
'
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim oColumn         As SAPbouiCOM.Column
    Dim oCheck          As SAPbouiCOM.CheckBox
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//디비데이터 소스 개체 할당
    Set oDS_ZPY502H = oForm.DataSources.DBDataSources("@ZPY502H")
    Set oDS_ZPY502L = oForm.DataSources.DBDataSources("@ZPY502L")
    
    Set oMat1 = oForm.Items("Mat1").Specific
    
    '// 사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT Code, Name FROM [@PH_PY005A]"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop
    oForm.Items("CLTCOD").DisplayDesc = True
   '/ Check 버튼
    Set oCheck = oForm.Items("ENDCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"

    Set oCheck = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oCheck = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error:" & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub ShowSource(JSNYER$, MSTCOD$, CLTCOD$)
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim DocEntry$
    Dim oMast           As ZPAY_g_EmpID
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    sQry = "SELECT DocEntry FROM [@ZPY502H]"
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
        
        Call oDS_ZPY502H.setValue("U_JSNYER", 0, JSNYER)
        Call oDS_ZPY502H.setValue("U_MSTCOD", 0, MSTCOD)
        Call oDS_ZPY502H.setValue("U_CLTCOD", 0, CLTCOD)
        oMast = MDC_SetMod.Get_EmpID_InFo(MSTCOD)
        Call oDS_ZPY502H.setValue("U_MSTNAM", 0, oMast.MSTNAM)
        Call oDS_ZPY502H.setValue("U_EmpID", 0, oMast.EmpID)
        
        oForm.Update
        
        Sbo_Application.SendKeys "{TAB}"
    End If
    
    Set oRecordSet = Nothing
    
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
    
    JSNYER = Trim$(oDS_ZPY502H.GetValue("U_JSNYER", 0))
    
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
        oMat1.Columns("Col5").TitleObject.Caption = "비과세-생산"
        oMat1.Columns("Col6").TitleObject.Caption = "비과세-보육"
        oMat1.Columns("Col7").TitleObject.Caption = "비과세-국외"
        oMat1.Columns("Col13").TitleObject.Caption = "비과세-연구"
        oMat1.Columns("Col20").TitleObject.Caption = "비과세-외국인"
        
        oMat1.Columns("Col23").TitleObject.Caption = "비과세(G01)"
        oMat1.Columns("Col24").TitleObject.Caption = "비과세(H01)"
        oMat1.Columns("Col25").TitleObject.Caption = "비과세(H05)"
        oMat1.Columns("Col26").TitleObject.Caption = "비과세(H06)"
        oMat1.Columns("Col27").TitleObject.Caption = "비과세(H07)"
        oMat1.Columns("Col28").TitleObject.Caption = "비과세(H08)"
        oMat1.Columns("Col29").TitleObject.Caption = "비과세(H09)"
        oMat1.Columns("Col30").TitleObject.Caption = "비과세(H10)"
        oMat1.Columns("Col31").TitleObject.Caption = "비과세(H11)"
        oMat1.Columns("Col32").TitleObject.Caption = "비과세(H12)"
        oMat1.Columns("Col33").TitleObject.Caption = "비과세(H13)"
        oMat1.Columns("Col34").TitleObject.Caption = "비과세(I01)"
        oMat1.Columns("Col35").TitleObject.Caption = "비과세(K01)"
        oMat1.Columns("Col36").TitleObject.Caption = "비과세(M01)"
        oMat1.Columns("Col37").TitleObject.Caption = "비과세(M02)"
        oMat1.Columns("Col38").TitleObject.Caption = "비과세(M03)"
        oMat1.Columns("Col39").TitleObject.Caption = "비과세(O01)"
        oMat1.Columns("Col40").TitleObject.Caption = "비과세(Q01)"
        oMat1.Columns("Col41").TitleObject.Caption = "비과세(S01)"
        oMat1.Columns("Col42").TitleObject.Caption = "비과세(T01)"
        oMat1.Columns("Col43").TitleObject.Caption = "비과세(X01)"
        oMat1.Columns("Col44").TitleObject.Caption = "비과세(Y01)"
        oMat1.Columns("Col45").TitleObject.Caption = "비과세(Y02)"
        oMat1.Columns("Col46").TitleObject.Caption = "비과세(Y03)"
        oMat1.Columns("Col47").TitleObject.Caption = "비과세(Y20)"
        oMat1.Columns("Col50").TitleObject.Caption = "비과세(Y21)"
        oMat1.Columns("Col48").TitleObject.Caption = "비과세(Z01)"
        
        '// 비과세 컬럼 전체 표시
        For iCol = 5 To 50
            If (iCol >= 5 And iCol <= 7) Or iCol = 13 Or iCol = 20 Or _
               (iCol >= 23 And iCol <= 50) Then
                oMat1.Columns("Col" & CStr(iCol)).Visible = True
            End If
        Next iCol
    
    ElseIf Val(JSNYER) <= 2008 Then
        Do Until oRecordSet.EOF
            '// 비과세 코드에 따라 컬럼UID 확인
            Select Case oRecordSet.Fields("U_BTXCOD").VALUE
            Case "01":      COLNAM = "Col5"     '/ 비과세-생산
            Case "07":      COLNAM = "Col6"     '/ 비과세-보육
            Case "03":      COLNAM = "Col7"     '/ 비과세-국외
            Case "06":      COLNAM = "Col13"    '/ 비과세-연구
            Case "04":      COLNAM = "Col20"    '/ 비과세-외국인
            Case Else:      COLNAM = ""
            End Select
            
            '// 컬럼명, 화면표시 여부 적용
            If COLNAM <> "" Then
                With oMat1.Columns(COLNAM)
                    .TitleObject.Caption = oRecordSet.Fields("U_BTXNAM").VALUE
                    .Editable = False
                End With
            End If
            
            oRecordSet.MoveNext
        Loop
        
    Else
        Do Until oRecordSet.EOF
            '// 비과세 코드에 따라 컬럼UID 확인
            Select Case oRecordSet.Fields("U_BTXCOD").VALUE
            Case "G01":     COLNAM = "Col23"    '// 비과세(G01)
            Case "H01":     COLNAM = "Col24"    '// 비과세(H01)
            Case "H05":     COLNAM = "Col25"    '// 비과세(H05)
            Case "H06":     COLNAM = "Col26"    '// 비과세(H06)
            Case "H07":     COLNAM = "Col27"    '// 비과세(H07)
            Case "H08":     COLNAM = "Col28"    '// 비과세(H08)
            Case "H09":     COLNAM = "Col29"    '// 비과세(H09)
            Case "H10":     COLNAM = "Col30"    '// 비과세(H10)
            Case "H11":     COLNAM = "Col31"    '// 비과세(H11)
            Case "H12":     COLNAM = "Col32"    '// 비과세(H12)
            Case "H13":     COLNAM = "Col33"    '// 비과세(H13)
            Case "I01":     COLNAM = "Col34"    '// 비과세(I01)
            Case "K01":     COLNAM = "Col35"    '// 비과세(K01)
            Case "M01":     COLNAM = "Col36"    '// 비과세(M01)
            Case "M02":     COLNAM = "Col37"    '// 비과세(M02)
            Case "M03":     COLNAM = "Col38"    '// 비과세(M03)
            Case "O01":     COLNAM = "Col39"    '// 비과세(O01)
            Case "Q01":     COLNAM = "Col40"    '// 비과세(Q01)
            Case "S01":     COLNAM = "Col41"    '// 비과세(S01)
            Case "T01":     COLNAM = "Col42"    '// 비과세(T01)
            Case "X01":     COLNAM = "Col43"    '// 비과세(X01)
            Case "Y01", "R10":
                            COLNAM = "Col44"    '// 비과세(Y01)
            Case "Y02":     COLNAM = "Col45"    '// 비과세(Y02)
            Case "Y03":     COLNAM = "Col46"    '// 비과세(Y03)
            Case "Y20", "Y22":
                            COLNAM = "Col47"    '// 비과세(Y20)
            Case "Y21":     COLNAM = "Col50"    '// 비과세(Y21)
            Case "Z01":     COLNAM = "Col48"    '// 비과세(Z01)
            End Select
            
            '// 컬럼명, 화면표시 여부 적용
            With oMat1.Columns(COLNAM)
                .TitleObject.Caption = oRecordSet.Fields("U_BTXNAM").VALUE
            End With
            
            oRecordSet.MoveNext
        Loop
        If JSNYER >= "2010" Then
            oMat1.Columns("Col43").Visible = False
        End If
        
    End If
    
    oForm.Freeze False
    Set oRecordSet = Nothing
    Exit Sub
Error_Message:
    oForm.Freeze False
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "Matrix_TitleSetting 실행 중 오류가 발생하였습니다.", bmt_Short, smt_Error
End Sub


