Option Strict Off
Imports System
Imports System.IO
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UI
Imports NXOpen.UF

Module ExportToSpreadsheet
    Dim theSession As Session = Session.GetSession()
    Dim theUI As UI = UI.GetUI()
    Dim ufSession As UFSession = UFSession.GetUFSession()

    Sub Main()
        ' アクティブなワークパートの取得
        Dim workPart As Part = theSession.Parts.Work

        ' アセンブリ内の全てのコンポーネントを取得
        Dim components As Component() = workPart.ComponentAssembly.RootComponent.GetChildren()

        ' スプレッドシートの作成
        Dim sheet As Spreadsheet = theSession.SpreadsheetManager.NewSheet("Part Info")

        ' スプレッドシートに書き出すヘッダを設定
        Dim header As String = "Item ID,Item Revision,Part Name"
        Dim headerRow As Integer = 0
        Dim headerColumn As Integer = 0
        sheet.SetString(headerRow, headerColumn, header)

        ' アセンブリ内の全てのパートに対して、アイテムID、アイテムリビジョン、パート名を取得してスプレッドシートに書き出す
        Dim row As Integer = 1
        For Each component As Component In components
            If component.Prototype Is Nothing Then ' パートファイルの場合のみ処理する
                Dim part As Part = CType(component.Prototype, Part)
                Dim itemId As String = part.GetStringAttribute("ITEM_ID")
                Dim itemRev As String = part.GetStringAttribute("ITEM_REV")
                Dim partName As String = part.Name
                sheet.SetString(row, headerColumn, itemId)
                sheet.SetString(row, headerColumn + 1, itemRev)
                sheet.SetString(row, headerColumn + 2, partName)
                row += 1
            End If
        Next

        ' スプレッドシートを保存する
        sheet.SaveAs("Output", "xlsx")

        ' 完了メッセージを表示
        theUI.NXMessageBox.Show("Export to Spreadsheet", NXMessageBox.DialogType.Information, "Data exported to spreadsheet successfully.")
    End Sub
End Module
