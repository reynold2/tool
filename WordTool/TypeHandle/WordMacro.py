import os
import comtypes.client


class WordMacro(object):
    """
    Sub SelectAllTable()

    Dim tempTable As Table

        Application.ScreenUpdating = False

        '判断文档是否被保护
        If ActiveDocument.ProtectionType = wdAllowOnlyFormFields Then
            MsgBox "文档已保护，此时不能选中多个表格！"
            Exit Sub
        End If
        '删除所有可编辑的区域
        ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
        '添加可编辑区域
        For Each tempTable In ActiveDocument.Tables
            tempTable.Range.Editors.Add wdEditorEveryone

        Next
        '选中所有可编辑区域
        ActiveDocument.SelectAllEditableRanges wdEditorEveryone
        '删除所有可编辑的区域
        ActiveDocument.DeleteAllEditableRanges wdEditorEveryone

        Application.ScreenUpdating = True

    End Sub


    Sub table_header()
    '
    '
    If ActiveDocument.Tables.Count >= 1 Then
    Set act_Doc = ActiveDocument
    For Each otable In act_Doc.Tables
    CaptionLabels.Add Name:="表 "
    With otable.Range.InsertCaption(Label:="表 ", Position:=wdCaptionPositionAbove)
    'Position:=wdCaptionPositionBelow
    End With
    Next
    End If

    End Sub
    """

    @staticmethod
    def macro(index, word_path):
        word = comtypes.client.CreateObject("Word.Application")
        word.Documents.Open(word_path, ReadOnly=1)
        if index == 1:
            word.Application.Run("table_header")
        elif index == 2:
            word.Application.Run("SelectAllTable")
        word.Documents(1).Close(SaveChanges=0)
        word.Application.Quit()
