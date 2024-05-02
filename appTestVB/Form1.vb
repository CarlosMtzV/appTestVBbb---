Imports System.Drawing.Text
Imports System.Security.Cryptography
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System

Public Class Form1

    Dim id As Integer

    Public Sub New()
        InitializeComponent()
        id = 0
        lstvData.FullRowSelect = True
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim x As MyItem
        'i++   -->   i = i + 1
        Dim description As String = txtDescription.Text
        id = id + 1
        Dim Price As Random = New Random()

        x = New MyItem(id, description, Math.Round(Price.NextDouble() * 1000, 2))
        'x = New MyItem(id, description, 56)

        lstItems.Items.Add(x.ToString())

        'ListView -- ListViewItems -- SubItems

        For i = 1 To 100
            Dim row As ListViewItem = New ListViewItem(x.Id)
            row.SubItems.Add(x.Description)
            row.SubItems.Add(x.Price)
            lstvData.Items.Add(row)
            x.Id = x.Id + 1
            x.Price = Math.Round(Price.NextDouble() * 1000, 2)
            'x.Price = 56

        Next
        UpdateLabel()
        UpdateTotal()
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        If lstvData.SelectedItems.Count = 0 Then
            Return
        End If
        For Each item As ListViewItem In lstvData.SelectedItems
            lstvData.Items.Remove(item)
        Next
        UpdateLabel()
        UpdateTotal()
    End Sub

    Sub UpdateLabel()
        lblCount.Text = lstvData.Items.Count
    End Sub
    Sub UpdateTotal()
        Dim Total As Decimal = 0
        For Each item As ListViewItem In lstvData.Items
            Total = Total + Decimal.Parse(item.SubItems(2).Text)
        Next
        LblTotal.Text = Total
    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
        saveFileDialog.FilterIndex = 1
        saveFileDialog.RestoreDirectory = True

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = saveFileDialog.FileName

            Dim dataTable As New DataTable()

            For Each column As ColumnHeader In lstvData.Columns
                dataTable.Columns.Add(column.Text)
            Next

            For Each item As ListViewItem In lstvData.Items
                Dim row As DataRow = dataTable.NewRow()
                For i As Integer = 0 To item.SubItems.Count - 1
                    row(i) = item.SubItems(i).Text
                Next
                dataTable.Rows.Add(row)
            Next

            Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook)
                Dim workbookPart As WorkbookPart = spreadsheetDocument.AddWorkbookPart()
                workbookPart.Workbook = New Workbook()
                Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
                worksheetPart.Worksheet = New Worksheet(New SheetData())

                Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(New Sheets())
                Dim sheet As Sheet = New Sheet() With {
                    .Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    .SheetId = 1,
                    .Name = "Sheet1"
                }
                sheets.Append(sheet)

                Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()

                Dim headerRow As New Row()
                For Each column As DataColumn In dataTable.Columns
                    Dim cell As New Cell() With {
                        .DataType = CellValues.String,
                        .CellValue = New CellValue(column.ColumnName)
                    }
                    headerRow.AppendChild(cell)
                Next
                sheetData.AppendChild(headerRow)

                For Each rowItem As DataRow In dataTable.Rows
                    Dim newRow As New Row()
                    For Each columnItem As Object In rowItem.ItemArray
                        Dim cell As New Cell() With {
                            .DataType = CellValues.String,
                            .CellValue = New CellValue(columnItem.ToString())
                        }
                        newRow.AppendChild(cell)
                    Next
                    sheetData.AppendChild(newRow)
                Next

                workbookPart.Workbook.Save()
            End Using

            MessageBox.Show("FileExcel Createddd")
        End If
    End Sub
End Class
