Attribute VB_Name = "Module5"
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    
    If Success Then
        Dim VBComp As VBIDE.VBComponent
        Dim SaveFolder As String
        SaveFolder = "C:\Users\a0976\OneDrive\AutoLisp �ǲ߻P�����[�c\�M�����O - ø�s\My_Anysis_Support_for_iec"

        For Each VBComp In ThisWorkbook.VBProject.VBComponents
            If VBComp.Type = vbext_ct_StdModule Then
                ' �u�ɥX�зǼҶ�
                VBComp.Export SaveFolder & "\" & VBComp.Name & ".bas"
            End If
        Next VBComp
    End If
End Sub

Sub ExportAllStandardModules()
    Dim VBComp As VBIDE.VBComponent
    Dim SaveFolder As String
    Dim sContent As String
    Dim sFilePath As String
    Dim nFileNum As Integer

    SaveFolder = "C:\Users\a0976\OneDrive\AutoLisp �ǲ߻P�����[�c\�M�����O - ø�s\My_Anysis_Support_for_iec\���B�z" ' ��אּ�z���ɥX���

    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        If VBComp.Type = vbext_ct_StdModule Then
            ' �ͦ������|
            sFilePath = SaveFolder & "\" & VBComp.Name & ".bas"

            ' �ɥX�Ҷ�
            VBComp.Export sFilePath

            ' ���s���}���åH�G�i��Ҧ�Ū�����e
            nFileNum = FreeFile
            Open sFilePath For Binary Access Read As #nFileNum
            sContent = StrConv(InputB(LOF(nFileNum), #nFileNum), vbUnicode)
            Close #nFileNum

            ' �H UTF-8 �榡�O�s
            SaveAsUTF8 sContent, sFilePath
        End If
    Next VBComp

    MsgBox "�Ҧ��зǼҶ��w�ɥX�� " & SaveFolder
End Sub



Function SaveAsUTF8(sContent As String, sFilePath As String)
    Dim nFileNum As Integer
    Dim baBuffer() As Byte

    ' �ഫ�r�Ŧ�� UTF-8
    baBuffer = StrConv(sContent, vbFromUnicode)

    ' �g�J���
    nFileNum = FreeFile
    Open sFilePath For Binary Access Write As #nFileNum
    Put #nFileNum, , baBuffer
    Close #nFileNum
End Function


