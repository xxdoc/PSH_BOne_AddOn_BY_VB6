Attribute VB_Name = "SubMain"
Public Sbo_Application  As SAPbouiCOM.Application
Public Sbo_Company      As SAPbobsCOM.Company
Public FormCurrentCount As Long '//���� �����Ѱ���
Public FormTotalCount   As Long '//������ �����Ѱ���
Public ClassList        As Collection '//�÷��� ��ü
Public ShareFolderPath  As String '//���������ּ�
Public ServerPath       As String '//�����ּ�
Public oZSBO            As ZZMDC

Public Sub Main()
'******************************************************************************
'Function ID    : Main
'�� �� �� ��    : SubMain
'��       ��    : ZZMDC Ŭ������ �ν��Ͻ��� ȣ��, �ý��ۿ��� ���ʷ� ����
'��       ��    : ����
'��   ȯ  ��    : ����
'Ư�̻���       : ����
'******************************************************************************
    If App.PrevInstance = True Then
        Instance_Flg = False
        MsgBox App.EXEName + "�� ����������Դϴ�.", vbExclamation, "oTemp.exe"
        End
    End If
    
    Set oZSBO = New ZZMDC

    Do While GetMessage(Msg, 0&, 0&, 0&)    'Message Loop
        TranslateMessage Msg
        DispatchMessage Msg
        DoEvents
    Loop
    
End Sub

'*******************************************************************
'//����ü�߰�
'*******************************************************************
Public Sub AddForms(ByVal cObject As Variant, ByVal oFormUid As String)
    ClassList.Add cObject, oFormUid
    FormTotalCount = FormTotalCount + 1
    FormCurrentCount = FormCurrentCount + 1
End Sub

'*******************************************************************
'//����ü����
'*******************************************************************
Public Sub RemoveForms(ByVal oFormUniqueID As String)
    Dim oTempClass As Variant
    Set oTempClass = ClassList.Item(oFormUniqueID)
    ClassList.Remove oFormUniqueID
    Set oTempClass = Nothing
    FormCurrentCount = FormCurrentCount - 1
End Sub

'*******************************************************************
'//�����簴ü��
'*******************************************************************
Public Function GetCurrentFormsCount() As Long
    GetCurrentFormsCount = FormCurrentCount
End Function

'*******************************************************************
'//���Ѱ�ü��
'*******************************************************************
Public Function GetTotalFormsCount() As Long
    GetTotalFormsCount = FormTotalCount
End Function
