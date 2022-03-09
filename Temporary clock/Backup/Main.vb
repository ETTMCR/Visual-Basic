Imports System.Windows.Forms

REM Note: All controls on screen are created starting at 1. The arrays are all 0 based.
REM       So in these samples, the control array will be 1 size larger than needed, and the 0 index will be set to Nothing.
REM       If you name your controls starting at 0, there will not be this extra un-needed index. 
REM       Example:  TextBox0, Textbox1, etc...


Public Class Main
    Inherits System.Windows.Forms.Form

    'Can create a class variable to be used throughout the form
    'Just be sure to initialize this member on form_Load
    Dim mTextBoxes() As TextBox
    Dim mButtons() As Button
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Mixed2 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents Mixed5 As System.Windows.Forms.TextBox
    Friend WithEvents Mixed3 As System.Windows.Forms.TextBox
    Friend WithEvents Mixed1 As System.Windows.Forms.TextBox
    Friend WithEvents Mixed4 As System.Windows.Forms.Label
    Friend WithEvents lblLikeControls As System.Windows.Forms.Label
    Friend WithEvents lblMixedControls As System.Windows.Forms.Label
    Friend WithEvents btnLike As System.Windows.Forms.Button
    Friend WithEvents btnMixed As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnArray1 As System.Windows.Forms.Button
    Friend WithEvents btnArray2 As System.Windows.Forms.Button
    Friend WithEvents btnArray3 As System.Windows.Forms.Button
    Friend WithEvents btnArray5 As System.Windows.Forms.Button
    Friend WithEvents btnArray4 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
  Friend WithEvents Panel3 As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.TextBox1 = New System.Windows.Forms.TextBox
Me.TextBox2 = New System.Windows.Forms.TextBox
Me.TextBox3 = New System.Windows.Forms.TextBox
Me.TextBox4 = New System.Windows.Forms.TextBox
Me.Mixed2 = New System.Windows.Forms.Label
Me.TextBox5 = New System.Windows.Forms.TextBox
Me.Mixed5 = New System.Windows.Forms.TextBox
Me.Mixed3 = New System.Windows.Forms.TextBox
Me.Mixed1 = New System.Windows.Forms.TextBox
Me.Mixed4 = New System.Windows.Forms.Label
Me.lblLikeControls = New System.Windows.Forms.Label
Me.lblMixedControls = New System.Windows.Forms.Label
Me.btnLike = New System.Windows.Forms.Button
Me.btnMixed = New System.Windows.Forms.Button
Me.Label3 = New System.Windows.Forms.Label
Me.btnArray1 = New System.Windows.Forms.Button
Me.btnArray2 = New System.Windows.Forms.Button
Me.btnArray3 = New System.Windows.Forms.Button
Me.btnArray4 = New System.Windows.Forms.Button
Me.btnArray5 = New System.Windows.Forms.Button
Me.GroupBox1 = New System.Windows.Forms.GroupBox
Me.Panel1 = New System.Windows.Forms.Panel
Me.Panel2 = New System.Windows.Forms.Panel
Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
Me.Panel3 = New System.Windows.Forms.Panel
Me.GroupBox1.SuspendLayout()
Me.Panel1.SuspendLayout()
Me.Panel2.SuspendLayout()
Me.Panel3.SuspendLayout()
Me.SuspendLayout()
'
'TextBox1
'
Me.TextBox1.Location = New System.Drawing.Point(40, 56)
Me.TextBox1.Name = "TextBox1"
Me.TextBox1.Size = New System.Drawing.Size(104, 20)
Me.TextBox1.TabIndex = 0
Me.TextBox1.Text = "TextBox1"
'
'TextBox2
'
Me.TextBox2.Location = New System.Drawing.Point(40, 88)
Me.TextBox2.Name = "TextBox2"
Me.TextBox2.Size = New System.Drawing.Size(104, 20)
Me.TextBox2.TabIndex = 1
Me.TextBox2.Text = "TextBox2"
'
'TextBox3
'
Me.TextBox3.Location = New System.Drawing.Point(40, 120)
Me.TextBox3.Name = "TextBox3"
Me.TextBox3.Size = New System.Drawing.Size(104, 20)
Me.TextBox3.TabIndex = 2
Me.TextBox3.Text = "TextBox3"
'
'TextBox4
'
Me.TextBox4.Location = New System.Drawing.Point(24, 8)
Me.TextBox4.Name = "TextBox4"
Me.TextBox4.Size = New System.Drawing.Size(104, 20)
Me.TextBox4.TabIndex = 3
Me.TextBox4.Text = "TextBox4"
'
'Mixed2
'
Me.Mixed2.Location = New System.Drawing.Point(256, 80)
Me.Mixed2.Name = "Mixed2"
Me.Mixed2.Size = New System.Drawing.Size(104, 24)
Me.Mixed2.TabIndex = 6
Me.Mixed2.Text = "Mixed2"
'
'TextBox5
'
Me.TextBox5.Location = New System.Drawing.Point(16, 8)
Me.TextBox5.Name = "TextBox5"
Me.TextBox5.Size = New System.Drawing.Size(104, 20)
Me.TextBox5.TabIndex = 7
Me.TextBox5.Text = "TextBox5"
'
'Mixed5
'
Me.Mixed5.Location = New System.Drawing.Point(40, 88)
Me.Mixed5.Name = "Mixed5"
Me.Mixed5.Size = New System.Drawing.Size(104, 20)
Me.Mixed5.TabIndex = 12
Me.Mixed5.Text = "Mixed5"
'
'Mixed3
'
Me.Mixed3.Location = New System.Drawing.Point(40, 24)
Me.Mixed3.Name = "Mixed3"
Me.Mixed3.Size = New System.Drawing.Size(104, 20)
Me.Mixed3.TabIndex = 10
Me.Mixed3.Text = "Mixed3"
'
'Mixed1
'
Me.Mixed1.Location = New System.Drawing.Point(256, 48)
Me.Mixed1.Name = "Mixed1"
Me.Mixed1.Size = New System.Drawing.Size(104, 20)
Me.Mixed1.TabIndex = 8
Me.Mixed1.Text = "Mixed1"
'
'Mixed4
'
Me.Mixed4.Location = New System.Drawing.Point(40, 56)
Me.Mixed4.Name = "Mixed4"
Me.Mixed4.Size = New System.Drawing.Size(104, 24)
Me.Mixed4.TabIndex = 13
Me.Mixed4.Text = "Mixed4"
'
'lblLikeControls
'
Me.lblLikeControls.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblLikeControls.Location = New System.Drawing.Point(40, 24)
Me.lblLikeControls.Name = "lblLikeControls"
Me.lblLikeControls.Size = New System.Drawing.Size(104, 16)
Me.lblLikeControls.TabIndex = 14
Me.lblLikeControls.Text = "Like Controls"
Me.lblLikeControls.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'lblMixedControls
'
Me.lblMixedControls.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblMixedControls.Location = New System.Drawing.Point(248, 24)
Me.lblMixedControls.Name = "lblMixedControls"
Me.lblMixedControls.Size = New System.Drawing.Size(104, 16)
Me.lblMixedControls.TabIndex = 15
Me.lblMixedControls.Text = "Mixed Controls"
Me.lblMixedControls.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'btnLike
'
Me.btnLike.Location = New System.Drawing.Point(56, 288)
Me.btnLike.Name = "btnLike"
Me.btnLike.Size = New System.Drawing.Size(88, 32)
Me.btnLike.TabIndex = 16
Me.btnLike.Text = "Update Controls"
'
'btnMixed
'
Me.btnMixed.Location = New System.Drawing.Point(264, 296)
Me.btnMixed.Name = "btnMixed"
Me.btnMixed.Size = New System.Drawing.Size(88, 32)
Me.btnMixed.TabIndex = 17
Me.btnMixed.Text = "Update Controls"
'
'Label3
'
Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label3.Location = New System.Drawing.Point(424, 24)
Me.Label3.Name = "Label3"
Me.Label3.Size = New System.Drawing.Size(104, 16)
Me.Label3.TabIndex = 24
Me.Label3.Text = "Call Back Sample"
Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'btnArray1
'
Me.btnArray1.Location = New System.Drawing.Point(440, 48)
Me.btnArray1.Name = "btnArray1"
Me.btnArray1.Size = New System.Drawing.Size(80, 24)
Me.btnArray1.TabIndex = 25
Me.btnArray1.Text = "btnArray1"
'
'btnArray2
'
Me.btnArray2.Location = New System.Drawing.Point(440, 80)
Me.btnArray2.Name = "btnArray2"
Me.btnArray2.Size = New System.Drawing.Size(80, 24)
Me.btnArray2.TabIndex = 26
Me.btnArray2.Text = "btnArray2"
'
'btnArray3
'
Me.btnArray3.Location = New System.Drawing.Point(440, 112)
Me.btnArray3.Name = "btnArray3"
Me.btnArray3.Size = New System.Drawing.Size(80, 24)
Me.btnArray3.TabIndex = 27
Me.btnArray3.Text = "btnArray3"
'
'btnArray4
'
Me.btnArray4.Location = New System.Drawing.Point(440, 144)
Me.btnArray4.Name = "btnArray4"
Me.btnArray4.Size = New System.Drawing.Size(80, 24)
Me.btnArray4.TabIndex = 28
Me.btnArray4.Text = "btnArray4"
'
'btnArray5
'
Me.btnArray5.Location = New System.Drawing.Point(440, 176)
Me.btnArray5.Name = "btnArray5"
Me.btnArray5.Size = New System.Drawing.Size(80, 24)
Me.btnArray5.TabIndex = 29
Me.btnArray5.Text = "btnArray5"
'
'GroupBox1
'
Me.GroupBox1.Controls.Add(Me.Panel3)
Me.GroupBox1.Controls.Add(Me.Panel1)
Me.GroupBox1.Controls.Add(Me.TextBox2)
Me.GroupBox1.Controls.Add(Me.TextBox1)
Me.GroupBox1.Controls.Add(Me.lblLikeControls)
Me.GroupBox1.Controls.Add(Me.TextBox3)
Me.GroupBox1.Location = New System.Drawing.Point(16, 16)
Me.GroupBox1.Name = "GroupBox1"
Me.GroupBox1.Size = New System.Drawing.Size(184, 248)
Me.GroupBox1.TabIndex = 31
Me.GroupBox1.TabStop = False
Me.GroupBox1.Text = "GroupBox1"
'
'Panel1
'
Me.Panel1.BackColor = System.Drawing.Color.CornflowerBlue
Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.Panel1.Controls.Add(Me.TextBox5)
Me.Panel1.Location = New System.Drawing.Point(24, 184)
Me.Panel1.Name = "Panel1"
Me.Panel1.Size = New System.Drawing.Size(136, 40)
Me.Panel1.TabIndex = 15
'
'Panel2
'
Me.Panel2.BackColor = System.Drawing.SystemColors.Control
Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.Panel2.Controls.Add(Me.Mixed5)
Me.Panel2.Controls.Add(Me.Mixed4)
Me.Panel2.Controls.Add(Me.Mixed3)
Me.Panel2.Location = New System.Drawing.Point(216, 112)
Me.Panel2.Name = "Panel2"
Me.Panel2.Size = New System.Drawing.Size(176, 128)
Me.Panel2.TabIndex = 32
'
'LinkLabel1
'
Me.LinkLabel1.Location = New System.Drawing.Point(256, 256)
Me.LinkLabel1.Name = "LinkLabel1"
Me.LinkLabel1.Size = New System.Drawing.Size(112, 16)
Me.LinkLabel1.TabIndex = 33
Me.LinkLabel1.TabStop = True
Me.LinkLabel1.Text = "LinkLabel1"
'
'Panel3
'
Me.Panel3.BackColor = System.Drawing.Color.IndianRed
Me.Panel3.Controls.Add(Me.TextBox4)
Me.Panel3.Location = New System.Drawing.Point(16, 144)
Me.Panel3.Name = "Panel3"
Me.Panel3.Size = New System.Drawing.Size(144, 32)
Me.Panel3.TabIndex = 16
'
'Main
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(592, 334)
Me.Controls.Add(Me.LinkLabel1)
Me.Controls.Add(Me.Panel2)
Me.Controls.Add(Me.GroupBox1)
Me.Controls.Add(Me.btnArray5)
Me.Controls.Add(Me.btnArray4)
Me.Controls.Add(Me.btnArray3)
Me.Controls.Add(Me.btnArray2)
Me.Controls.Add(Me.btnArray1)
Me.Controls.Add(Me.Label3)
Me.Controls.Add(Me.btnMixed)
Me.Controls.Add(Me.btnLike)
Me.Controls.Add(Me.lblMixedControls)
Me.Controls.Add(Me.Mixed1)
Me.Controls.Add(Me.Mixed2)
Me.Name = "Main"
Me.Text = "Form1"
Me.GroupBox1.ResumeLayout(False)
Me.Panel1.ResumeLayout(False)
Me.Panel2.ResumeLayout(False)
Me.Panel3.ResumeLayout(False)
Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        mTextBoxes = ControlArrayUtils.getControlArray(Me, "TextBox")
        mTextBoxes(1).Text = "Form Load Test 1"
        mTextBoxes(2).Text = "Form Load Test 2"
        mTextBoxes(3).Text = "Form Load Test 3"
        mTextBoxes(4).Text = "Form Load Test 4"
        mTextBoxes(5).Text = "Form Load Test 5"

        'Setup the buttons
        'All buttons share the same Event Handler
        Dim i As Integer
        mButtons = ControlArrayUtils.getControlArray(Me, "btnArray")
        For i = 1 To 5
            AddHandler mButtons(i).Click, AddressOf Button_Click
        Next i


    End Sub
    Private Sub Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox(sender.name & " was clicked.")
    End Sub


    Private Sub btnLike_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLike.Click
        'Can create local control arrays, but need to call getControlArray each time.


        'This shows how a specific control array can be created.
        'The advantage of the specific control array is that you have full intellisense support of that control
        Dim textBoxes As TextBox() = ControlArrayUtils.getControlArray(Me, "TextBox")

        textBoxes(1).Text = "Update TextBox 1"
        textBoxes(2).Text = "Update TextBox 2"
        textBoxes(3).Text = "Update TextBox 3"
        textBoxes(4).Text = "Update TextBox 4"
        textBoxes(5).Text = "Update TextBox 5"


        'Could also create a member variable for the class, and just call it once on Form_Load
        'This is closest to the VB6 control array style.
        'Then use it here like  this:
        ' mTextBoxes(1).Text = "Form Load Test 1"


    End Sub

    Private Sub btnMixed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMixed.Click
        'This shows how a generic control array can be created.
        'The advantage is that you can have different types of controls within the array.
        'The disadvantage is that you have no intellisense for the specific control properties
        Dim controls As Control() = ControlArrayUtils.getMixedControlArray(Me, "Mixed")

        controls(1).Text = "Update Control 1" 'This is a textbox
        controls(2).Text = "Update Control 2" 'This is a label
        controls(3).Text = "Update Control 3" 'This is a textbox
        controls(4).Text = "Update Control 4" 'This is a label
        controls(5).Text = "Update Control 5" 'This is a textbox
        'controls(29).Text = "Yes!"
    End Sub


   
End Class
