Namespace XLExportExampleSalesReport
    Partial Public Class Form1
        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        #Region "Windows Form Designer generated code"

        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(Form1))
            Me.btnExportToXLSX = New System.Windows.Forms.Button()
            Me.panel1 = New System.Windows.Forms.Panel()
            Me.label1 = New System.Windows.Forms.Label()
            Me.btnExportToXLS = New System.Windows.Forms.Button()
            Me.btnExportToCSV = New System.Windows.Forms.Button()
            Me.saveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
            Me.panel1.SuspendLayout()
            Me.SuspendLayout()
            ' 
            ' btnExportToXLSX
            ' 
            Me.btnExportToXLSX.Image = (DirectCast(resources.GetObject("btnExportToXLSX.Image"), System.Drawing.Image))
            Me.btnExportToXLSX.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnExportToXLSX.Location = New System.Drawing.Point(11, 80)
            Me.btnExportToXLSX.Name = "btnExportToXLSX"
            Me.btnExportToXLSX.Size = New System.Drawing.Size(109, 28)
            Me.btnExportToXLSX.TabIndex = 0
            Me.btnExportToXLSX.Text = "Export to XLSX"
            Me.btnExportToXLSX.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            Me.btnExportToXLSX.UseVisualStyleBackColor = True
            ' 
            ' panel1
            ' 
            Me.panel1.BackColor = System.Drawing.SystemColors.Window
            Me.panel1.Controls.Add(Me.label1)
            Me.panel1.Dock = System.Windows.Forms.DockStyle.Top
            Me.panel1.Location = New System.Drawing.Point(0, 0)
            Me.panel1.Name = "panel1"
            Me.panel1.Padding = New System.Windows.Forms.Padding(8)
            Me.panel1.Size = New System.Drawing.Size(365, 63)
            Me.panel1.TabIndex = 1
            ' 
            ' label1
            ' 
            Me.label1.AutoSize = True
            Me.label1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(204)))
            Me.label1.Location = New System.Drawing.Point(8, 8)
            Me.label1.Name = "label1"
            Me.label1.Size = New System.Drawing.Size(311, 39)
            Me.label1.TabIndex = 0
            Me.label1.Text = "Generate a sales report using the XL Export API." & ControlChars.CrLf & "Click one of the buttons below " & "to save the document " & ControlChars.CrLf & "in the corresponding format."
            ' 
            ' btnExportToXLS
            ' 
            Me.btnExportToXLS.Image = (DirectCast(resources.GetObject("btnExportToXLS.Image"), System.Drawing.Image))
            Me.btnExportToXLS.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnExportToXLS.Location = New System.Drawing.Point(126, 80)
            Me.btnExportToXLS.Name = "btnExportToXLS"
            Me.btnExportToXLS.Size = New System.Drawing.Size(109, 28)
            Me.btnExportToXLS.TabIndex = 2
            Me.btnExportToXLS.Text = "Export to XLS"
            Me.btnExportToXLS.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            Me.btnExportToXLS.UseVisualStyleBackColor = True
            ' 
            ' btnExportToCSV
            ' 
            Me.btnExportToCSV.Image = (DirectCast(resources.GetObject("btnExportToCSV.Image"), System.Drawing.Image))
            Me.btnExportToCSV.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnExportToCSV.Location = New System.Drawing.Point(241, 80)
            Me.btnExportToCSV.Name = "btnExportToCSV"
            Me.btnExportToCSV.Size = New System.Drawing.Size(109, 28)
            Me.btnExportToCSV.TabIndex = 3
            Me.btnExportToCSV.Text = "Export to CSV"
            Me.btnExportToCSV.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            Me.btnExportToCSV.UseVisualStyleBackColor = True
            ' 
            ' Form1
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(365, 120)
            Me.Controls.Add(Me.btnExportToCSV)
            Me.Controls.Add(Me.btnExportToXLS)
            Me.Controls.Add(Me.panel1)
            Me.Controls.Add(Me.btnExportToXLSX)
            Me.Name = "Form1"
            Me.Text = "XL Export example"
            Me.panel1.ResumeLayout(False)
            Me.panel1.PerformLayout()
            Me.ResumeLayout(False)

        End Sub

        #End Region

        Private WithEvents btnExportToXLSX As System.Windows.Forms.Button
        Private panel1 As System.Windows.Forms.Panel
        Private label1 As System.Windows.Forms.Label
        Private WithEvents btnExportToXLS As System.Windows.Forms.Button
        Private WithEvents btnExportToCSV As System.Windows.Forms.Button
        Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
    End Class
End Namespace

