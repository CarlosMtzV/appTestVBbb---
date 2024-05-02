<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()>
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Label1 = New Label()
        txtDescription = New TextBox()
        btnSave = New Button()
        lstItems = New ListBox()
        lstvData = New ListView()
        ColumnHeader1 = New ColumnHeader()
        ColumnHeader2 = New ColumnHeader()
        ColumnHeader3 = New ColumnHeader()
        lblCount = New Label()
        btnRemove = New Button()
        LblTotal = New Label()
        btnInsert = New Button()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 15F)
        Label1.Location = New Point(40, 62)
        Label1.Name = "Label1"
        Label1.Size = New Size(112, 28)
        Label1.TabIndex = 0
        Label1.Text = "Description"
        ' 
        ' txtDescription
        ' 
        txtDescription.Font = New Font("Segoe UI", 15F)
        txtDescription.Location = New Point(158, 62)
        txtDescription.Name = "txtDescription"
        txtDescription.Size = New Size(164, 34)
        txtDescription.TabIndex = 1
        ' 
        ' btnSave
        ' 
        btnSave.Font = New Font("Segoe UI", 15F)
        btnSave.Location = New Point(348, 65)
        btnSave.Name = "btnSave"
        btnSave.Size = New Size(139, 38)
        btnSave.TabIndex = 2
        btnSave.Text = "Save"
        btnSave.UseVisualStyleBackColor = True
        ' 
        ' lstItems
        ' 
        lstItems.Font = New Font("Segoe UI", 15F)
        lstItems.FormattingEnabled = True
        lstItems.ItemHeight = 28
        lstItems.Location = New Point(40, 158)
        lstItems.Name = "lstItems"
        lstItems.Size = New Size(248, 200)
        lstItems.TabIndex = 3
        ' 
        ' lstvData
        ' 
        lstvData.Columns.AddRange(New ColumnHeader() {ColumnHeader1, ColumnHeader2, ColumnHeader3})
        lstvData.Font = New Font("Segoe UI", 15F)
        lstvData.Location = New Point(330, 158)
        lstvData.Name = "lstvData"
        lstvData.Size = New Size(415, 200)
        lstvData.TabIndex = 4
        lstvData.UseCompatibleStateImageBehavior = False
        lstvData.View = View.Details
        ' 
        ' ColumnHeader1
        ' 
        ColumnHeader1.Text = "ID"
        ' 
        ' ColumnHeader2
        ' 
        ColumnHeader2.Text = "DESCRIPTION"
        ColumnHeader2.Width = 200
        ' 
        ' ColumnHeader3
        ' 
        ColumnHeader3.Text = "PRICE"
        ColumnHeader3.Width = 120
        ' 
        ' lblCount
        ' 
        lblCount.AutoSize = True
        lblCount.Font = New Font("Segoe UI", 15F)
        lblCount.Location = New Point(330, 372)
        lblCount.Name = "lblCount"
        lblCount.Size = New Size(0, 28)
        lblCount.TabIndex = 5
        ' 
        ' btnRemove
        ' 
        btnRemove.Font = New Font("Segoe UI", 15F)
        btnRemove.Location = New Point(583, 68)
        btnRemove.Name = "btnRemove"
        btnRemove.Size = New Size(139, 38)
        btnRemove.TabIndex = 6
        btnRemove.Text = "Remove"
        btnRemove.UseVisualStyleBackColor = True
        ' 
        ' LblTotal
        ' 
        LblTotal.AutoSize = True
        LblTotal.Font = New Font("Segoe UI", 15F)
        LblTotal.Location = New Point(583, 391)
        LblTotal.Name = "LblTotal"
        LblTotal.Size = New Size(0, 28)
        LblTotal.TabIndex = 7
        ' 
        ' btnInsert
        ' 
        btnInsert.Font = New Font("Segoe UI", 15F)
        btnInsert.Location = New Point(66, 372)
        btnInsert.Name = "btnInsert"
        btnInsert.Size = New Size(139, 38)
        btnInsert.TabIndex = 8
        btnInsert.Text = "Insert"
        btnInsert.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(btnInsert)
        Controls.Add(LblTotal)
        Controls.Add(btnRemove)
        Controls.Add(lblCount)
        Controls.Add(lstvData)
        Controls.Add(lstItems)
        Controls.Add(btnSave)
        Controls.Add(txtDescription)
        Controls.Add(Label1)
        Name = "Form1"
        Text = "Form1"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Label1 As Label
  Friend WithEvents txtDescription As TextBox
  Friend WithEvents btnSave As Button
  Friend WithEvents lstItems As ListBox
  Friend WithEvents lstvData As ListView
  Friend WithEvents ColumnHeader1 As ColumnHeader
  Friend WithEvents ColumnHeader2 As ColumnHeader
  Friend WithEvents lblCount As Label
  Friend WithEvents btnRemove As Button
  Friend WithEvents ColumnHeader3 As ColumnHeader
  Friend WithEvents LblTotal As Label
    Friend WithEvents btnInsert As Button

End Class
