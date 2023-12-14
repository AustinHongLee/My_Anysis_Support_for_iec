Attribute VB_Name = "Module5"
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    
    If Success Then
        Dim VBComp As VBIDE.VBComponent
        Dim SaveFolder As String
        SaveFolder = "C:\Users\a0976\OneDrive\AutoLisp 學習與公式架構\專案類別 - 繪製\My_Anysis_Support_for_iec"

        For Each VBComp In ThisWorkbook.VBProject.VBComponents
            If VBComp.Type = vbext_ct_StdModule Then
                ' 只導出標準模塊
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

    SaveFolder = "C:\Users\a0976\OneDrive\AutoLisp 學習與公式架構\專案類別 - 繪製\My_Anysis_Support_for_iec\未處理" ' 更改為您的導出文件夾

    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        If VBComp.Type = vbext_ct_StdModule Then
            ' 生成文件路徑
            sFilePath = SaveFolder & "\" & VBComp.Name & ".bas"

            ' 導出模塊
            VBComp.Export sFilePath

            ' 重新打開文件並以二進位模式讀取內容
            nFileNum = FreeFile
            Open sFilePath For Binary Access Read As #nFileNum
            sContent = StrConv(InputB(LOF(nFileNum), #nFileNum), vbUnicode)
            Close #nFileNum

            ' 以 UTF-8 格式保存
            SaveAsUTF8 sContent, sFilePath
        End If
    Next VBComp

    MsgBox "所有標準模塊已導出到 " & SaveFolder
End Sub



Function SaveAsUTF8(sContent As String, sFilePath As String)
    Dim nFileNum As Integer
    Dim baBuffer() As Byte

    ' 轉換字符串到 UTF-8
    baBuffer = StrConv(sContent, vbFromUnicode)

    ' 寫入文件
    nFileNum = FreeFile
    Open sFilePath For Binary Access Write As #nFileNum
    Put #nFileNum, , baBuffer
    Close #nFileNum
End Function


