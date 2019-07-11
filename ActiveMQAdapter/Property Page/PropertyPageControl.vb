Friend Class PropertyPageControl
    Inherits System.Windows.Forms.UserControl

    ' Используйте конструктор с параметрами для инициализации страницы

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNameQueueOut As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtNameQueueIn As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PropertyPageControl))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.txtPort = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNameQueueOut = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtNameQueueIn = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 109)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Сервер:"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(95, 106)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(254, 20)
        Me.txtServer.TabIndex = 2
        '
        'txtPort
        '
        Me.txtPort.Location = New System.Drawing.Point(95, 132)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(254, 20)
        Me.txtPort.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 135)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Порт:"
        '
        'txtNameQueueOut
        '
        Me.txtNameQueueOut.Location = New System.Drawing.Point(165, 184)
        Me.txtNameQueueOut.Name = "txtNameQueueOut"
        Me.txtNameQueueOut.Size = New System.Drawing.Size(184, 20)
        Me.txtNameQueueOut.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 187)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(159, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Имя очереди исх. сообщений:"
        '
        'txtNameQueueIn
        '
        Me.txtNameQueueIn.Location = New System.Drawing.Point(165, 158)
        Me.txtNameQueueIn.Name = "txtNameQueueIn"
        Me.txtNameQueueIn.Size = New System.Drawing.Size(184, 20)
        Me.txtNameQueueIn.TabIndex = 10
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 161)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(153, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Имя очереди вх. сообщений:"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(3, 28)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(99, 60)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 11
        Me.PictureBox1.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.25!)
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(108, 28)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(227, 65)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "ООО «ППВТИ» " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "614002 ,г. Пермь, ул. Чернышевского, д.3" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Тел./Факс: (342) 217-00-0" & _
            "0 (многоканальный)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "E-mail: ppvti@ppvti.ru" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "http://www.ppvti.ru"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(54, 9)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(246, 16)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Адаптер для 1С и Active MQ v 1.0"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(216, 207)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(133, 12)
        Me.Label8.TabIndex = 15
        Me.Label8.Text = "Copyright © ООО ППВТИ 2011"
        '
        'PropertyPageControl
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.txtNameQueueIn)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtNameQueueOut)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtPort)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.Label1)
        Me.Name = "PropertyPageControl"
        Me.Size = New System.Drawing.Size(352, 296)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Event DirtyChanged()

    Private Sub txtNameQueueIn_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNameQueueIn.TextChanged
        RaiseEvent DirtyChanged()
    End Sub
    Private Sub txtNameQueueOut_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNameQueueOut.TextChanged
        RaiseEvent DirtyChanged()
    End Sub
    Private Sub txtPort_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPort.TextChanged
        RaiseEvent DirtyChanged()
    End Sub
    Private Sub txtServer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtServer.TextChanged
        RaiseEvent DirtyChanged()
    End Sub
End Class
