
Imports System.Windows.Forms

REM Note: All controls on screen are created starting at 1. The arrays are all 0 based.
REM       So in these samples, the control array will be 1 size larger than needed, and the 0 index will be set to Nothing.
REM       If you name your controls starting at 0, there will not be this extra un-needed index. 
REM       Example:  TextBox0, Textbox1, etc...

'http://stackoverflow.com/questions/4956458/having-trouble-in-changing-the-image-opacity
'http://stackoverflow.com/questions/24817235/change-opacity-in-picturebox


Imports System.Drawing.Imaging



Public Class Main



    Inherits System.Windows.Forms.Form

    'Can create a class variable to be used throughout the form
    'Just be sure to initialize this member on form_Load  ' !!!!!!!!!!!!!!!!!!!
    Dim picturebox_ex() As PictureBox  ' !!!!!!!!!!!!!!!!!!! see form_load
    Dim labels_ex() As Label
    Friend WithEvents Timer_sec As System.Windows.Forms.Timer
    '@@@@@ generated automaticly
    Dim mButtons() As Button
    Friend WithEvents PictureBox27 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox28 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox29 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox30 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox31 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox32 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox33 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox34 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox35 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox36 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox37 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox38 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox41 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox40 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox42 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox43 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox44 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox46 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox47 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox48 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox49 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox51 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox52 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox53 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox54 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox56 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox57 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox58 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox59 As System.Windows.Forms.PictureBox
    Dim before_sec As Integer
    Dim picturebox_ex_HR() As PictureBox
    Friend WithEvents PicBxHR1 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR2 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR3 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR4 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR5 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR6 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR7 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR8 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR9 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR10 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR11 As System.Windows.Forms.PictureBox
    Friend WithEvents PicBxHR0 As System.Windows.Forms.PictureBox
    Dim before_min As Integer
    Friend WithEvents lbl_min1 As System.Windows.Forms.Label
    Friend WithEvents lbl_min2 As System.Windows.Forms.Label
    Friend WithEvents Timer_min As System.Windows.Forms.Timer
    Dim before_HR As Integer
    Friend WithEvents lbl_min3 As System.Windows.Forms.Label
    Friend WithEvents lbl_min6 As System.Windows.Forms.Label
    Friend WithEvents lbl_min5 As System.Windows.Forms.Label
    Friend WithEvents lbl_min14 As System.Windows.Forms.Label
    Friend WithEvents lbl_min13 As System.Windows.Forms.Label
    Friend WithEvents lbl_min12 As System.Windows.Forms.Label
    Friend WithEvents lbl_min11 As System.Windows.Forms.Label
    Friend WithEvents lbl_min10 As System.Windows.Forms.Label
    Friend WithEvents lbl_min9 As System.Windows.Forms.Label
    Friend WithEvents lbl_min8 As System.Windows.Forms.Label
    Friend WithEvents lbl_min23 As System.Windows.Forms.Label
    Friend WithEvents lbl_min20 As System.Windows.Forms.Label
    Friend WithEvents lbl_min19 As System.Windows.Forms.Label
    Friend WithEvents lbl_min17 As System.Windows.Forms.Label
    Friend WithEvents lbl_min25 As System.Windows.Forms.Label
    Friend WithEvents lbl_min18 As System.Windows.Forms.Label
    Friend WithEvents lbl_min16 As System.Windows.Forms.Label
    Friend WithEvents lbl_min26 As System.Windows.Forms.Label
    Friend WithEvents lbl_min27 As System.Windows.Forms.Label
    Friend WithEvents lbl_min28 As System.Windows.Forms.Label
    Friend WithEvents lbl_min21 As System.Windows.Forms.Label
    Friend WithEvents lbl_min15 As System.Windows.Forms.Label
    Friend WithEvents lbl_min29 As System.Windows.Forms.Label
    Friend WithEvents lbl_min30 As System.Windows.Forms.Label
    Friend WithEvents lbl_min31 As System.Windows.Forms.Label
    Friend WithEvents lbl_min32 As System.Windows.Forms.Label
    Friend WithEvents lbl_min35 As System.Windows.Forms.Label
    Friend WithEvents lbl_min34 As System.Windows.Forms.Label
    Friend WithEvents lbl_min33 As System.Windows.Forms.Label
    Friend WithEvents lbl_min40 As System.Windows.Forms.Label
    Friend WithEvents lbl_min39 As System.Windows.Forms.Label
    Friend WithEvents lbl_min38 As System.Windows.Forms.Label
    Friend WithEvents lbl_min37 As System.Windows.Forms.Label
    Friend WithEvents lbl_min36 As System.Windows.Forms.Label
    Friend WithEvents PictureBox39 As System.Windows.Forms.PictureBox
    Friend WithEvents lbl_min50 As System.Windows.Forms.Label
    Friend WithEvents lbl_min49 As System.Windows.Forms.Label
    Friend WithEvents lbl_min47 As System.Windows.Forms.Label
    Friend WithEvents lbl_min46 As System.Windows.Forms.Label
    Friend WithEvents lbl_min45 As System.Windows.Forms.Label
    Friend WithEvents lbl_min48 As System.Windows.Forms.Label
    Friend WithEvents lbl_min44 As System.Windows.Forms.Label
    Friend WithEvents lbl_min43 As System.Windows.Forms.Label
    Friend WithEvents lbl_min42 As System.Windows.Forms.Label
    Friend WithEvents lbl_min41 As System.Windows.Forms.Label
    Friend WithEvents lbl_min0 As System.Windows.Forms.Label
    Friend WithEvents lbl_min59 As System.Windows.Forms.Label
    Friend WithEvents lbl_min57 As System.Windows.Forms.Label
    Friend WithEvents lbl_min56 As System.Windows.Forms.Label
    Friend WithEvents lbl_min55 As System.Windows.Forms.Label
    Friend WithEvents lbl_min58 As System.Windows.Forms.Label
    Friend WithEvents lbl_min54 As System.Windows.Forms.Label
    Friend WithEvents lbl_min53 As System.Windows.Forms.Label
    Friend WithEvents lbl_min52 As System.Windows.Forms.Label
    Friend WithEvents lbl_min51 As System.Windows.Forms.Label
    Friend WithEvents lbl_min4 As System.Windows.Forms.Label
    Friend WithEvents lbl_min7 As System.Windows.Forms.Label
    Friend WithEvents lbl_min22 As System.Windows.Forms.Label
    Friend WithEvents lbl_min24 As System.Windows.Forms.Label
    Friend WithEvents PictureBox_clock As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox26 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox24 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox23 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox22 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox25 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox21 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox20 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox19 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox18 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox17 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox16 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox15 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox13 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox14 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox12 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox11 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox10 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox9 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox8 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox7 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox0 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox45 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox50 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox55 As System.Windows.Forms.PictureBox

    'Dim now_HR_PMAM As Integer


    Dim font_size_MIN As Integer
    'Private Sub PictureBox_clock_MouseHover(sender As Object, e As System.EventArgs) Handles PictureBox_clock.MouseHover
    '    PictureBox_clock.tooltip = Date.Now
    'End Sub

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

    Private Sub Timer_min_Tick(sender As System.Object, e As System.EventArgs) Handles Timer_min.Tick
        labels_ex(DateTime.Now.Minute).Visible = True

        'labels_ex(DateTime.Now.Second).Font = New Font(lbl_min1.Font.Name, DateTime.Now.Second)
        If DateTime.Now.Second < 30 Then
            If DateTime.Now.Second = 0 Then
                'lbl_min1.Font() = New Font(lbl_min1.Font.Name, 1)
                labels_ex(DateTime.Now.Minute).Font = New Font(lbl_min1.Font.Name, 1)
                labels_ex(DateTime.Now.Minute).BringToFront()
            Else
                'lbl_min1.Font() = New Font(lbl_min1.Font.Name, DateTime.Now.Second)
                labels_ex(DateTime.Now.Minute).Font = New Font(lbl_min1.Font.Name, DateTime.Now.Second)
                labels_ex(DateTime.Now.Minute).BringToFront()
            End If
        Else
            'lbl_min1.Font() = New Font(lbl_min1.Font.Name, 61 - DateTime.Now.Second)
            labels_ex(DateTime.Now.Minute).Font = New Font(lbl_min1.Font.Name, 61 - DateTime.Now.Second)
            labels_ex(DateTime.Now.Minute).BringToFront()
        End If
        If DateTime.Now.Second = 0 Then
            labels_ex(DateTime.Now.Minute - 1).Visible = False
        End If
    End Sub

    Private Sub Timer_sec_Tick(sender As System.Object, e As System.EventArgs) Handles Timer_sec.Tick

        '' @@@sec ' making the seconds appear only in is second
        'picturebox_ex(before_sec).Visible = True
        'before_sec = DateTime.Now.Second
        'picturebox_ex(DateTime.Now.Second()).Visible = False
        '' @@@sec

        '' @@@ MIN
        'picturebox_ex(before_min).Visible = True
        'before_min = DateTime.Now.Minute
        'picturebox_ex(DateTime.Now.Minute()).Visible = False
        '' @@@ MIN

        ' @@@ HOUR '~~~~~~~~~~~~~~~
        picturebox_ex_HR(before_HR).Visible = True
        before_HR = DateTime.Now.Hour

        If DateTime.Now.Hour > 12 Then ' a problem with 00 hour
            picturebox_ex_HR(before_HR - 12).Visible = True
            before_HR = DateTime.Now.Hour - 12
        Else
            picturebox_ex_HR(before_HR).Visible = True
            before_HR = DateTime.Now.Hour
        End If

        If DateTime.Now.Hour > 12 Then ' a problem with 00 hour
            picturebox_ex_HR(DateTime.Now.Hour() - 13).Visible = True
        Else
            picturebox_ex_HR(DateTime.Now.Hour()).Visible = True
        End If

        If DateTime.Now.Hour > 12 Then
            picturebox_ex_HR(DateTime.Now.Hour() - 12).Visible = False
        Else
            picturebox_ex_HR(DateTime.Now.Hour()).Visible = False
        End If
        ' @@@ HOUR


    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        before_sec = Date.Now.Second '- 1 will make problem if you start the app excactly at 00 seconds
        before_min = Date.Now.Minute

        'MsgBox(Format(Now, "hh"))

        If DateTime.Now.Hour > 12 Then ' ~~~~~~~~~~
            'before_HR = Int(Format(DateTime.Now.Hour, "hh")) '- 13
            before_HR = DateTime.Now.Hour - 13
        Else
            before_HR = DateTime.Now.Hour
        End If


        picturebox_ex = ControlArrayUtils.getControlArray(Me, "picturebox") ' !!!!!!!!!!!!!!!!!!! making an array for controls with the name perfix "picturebox"

        picturebox_ex_HR = ControlArrayUtils.getControlArray(Me, "PicBxHR") ' !!!!!!!!!!!!!!!!!!!   

        labels_ex = ControlArrayUtils.getControlArray(Me, "lbl_min")
        Dim i As Integer
        For i = 0 To 59
            labels_ex(i).Visible = False
            'labels_ex(i).Font = New Font(lbl_min1.Font.Name, 1)
            labels_ex(i).Parent = PictureBox_clock
            'labels_ex(i).BackColor .

            picturebox_ex(i).Visible = False
        Next

        'Setup the buttons
        'All buttons share the same Event Handler
        'Dim i As Integer
        'mButtons = ControlArrayUtils.getControlArray(Me, "btnArray")
        'For i = 1 To 5
        '    AddHandler mButtons(i).Click, AddressOf Button_Click
        'Next i


    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Timer_sec = New System.Windows.Forms.Timer(Me.components)
        Me.PictureBox27 = New System.Windows.Forms.PictureBox()
        Me.PictureBox28 = New System.Windows.Forms.PictureBox()
        Me.PictureBox29 = New System.Windows.Forms.PictureBox()
        Me.PictureBox30 = New System.Windows.Forms.PictureBox()
        Me.PictureBox31 = New System.Windows.Forms.PictureBox()
        Me.PictureBox32 = New System.Windows.Forms.PictureBox()
        Me.PictureBox33 = New System.Windows.Forms.PictureBox()
        Me.PictureBox34 = New System.Windows.Forms.PictureBox()
        Me.PictureBox35 = New System.Windows.Forms.PictureBox()
        Me.PictureBox36 = New System.Windows.Forms.PictureBox()
        Me.PictureBox37 = New System.Windows.Forms.PictureBox()
        Me.PictureBox38 = New System.Windows.Forms.PictureBox()
        Me.PictureBox41 = New System.Windows.Forms.PictureBox()
        Me.PictureBox40 = New System.Windows.Forms.PictureBox()
        Me.PictureBox42 = New System.Windows.Forms.PictureBox()
        Me.PictureBox43 = New System.Windows.Forms.PictureBox()
        Me.PictureBox44 = New System.Windows.Forms.PictureBox()
        Me.PictureBox46 = New System.Windows.Forms.PictureBox()
        Me.PictureBox47 = New System.Windows.Forms.PictureBox()
        Me.PictureBox48 = New System.Windows.Forms.PictureBox()
        Me.PictureBox49 = New System.Windows.Forms.PictureBox()
        Me.PictureBox51 = New System.Windows.Forms.PictureBox()
        Me.PictureBox52 = New System.Windows.Forms.PictureBox()
        Me.PictureBox53 = New System.Windows.Forms.PictureBox()
        Me.PictureBox54 = New System.Windows.Forms.PictureBox()
        Me.PictureBox56 = New System.Windows.Forms.PictureBox()
        Me.PictureBox57 = New System.Windows.Forms.PictureBox()
        Me.PictureBox58 = New System.Windows.Forms.PictureBox()
        Me.PictureBox59 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR1 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR2 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR3 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR4 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR5 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR6 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR7 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR8 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR9 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR10 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR11 = New System.Windows.Forms.PictureBox()
        Me.PicBxHR0 = New System.Windows.Forms.PictureBox()
        Me.lbl_min1 = New System.Windows.Forms.Label()
        Me.Timer_min = New System.Windows.Forms.Timer(Me.components)
        Me.lbl_min2 = New System.Windows.Forms.Label()
        Me.lbl_min3 = New System.Windows.Forms.Label()
        Me.lbl_min6 = New System.Windows.Forms.Label()
        Me.lbl_min5 = New System.Windows.Forms.Label()
        Me.lbl_min14 = New System.Windows.Forms.Label()
        Me.lbl_min13 = New System.Windows.Forms.Label()
        Me.lbl_min12 = New System.Windows.Forms.Label()
        Me.lbl_min11 = New System.Windows.Forms.Label()
        Me.lbl_min10 = New System.Windows.Forms.Label()
        Me.lbl_min9 = New System.Windows.Forms.Label()
        Me.lbl_min8 = New System.Windows.Forms.Label()
        Me.lbl_min23 = New System.Windows.Forms.Label()
        Me.lbl_min20 = New System.Windows.Forms.Label()
        Me.lbl_min19 = New System.Windows.Forms.Label()
        Me.lbl_min17 = New System.Windows.Forms.Label()
        Me.lbl_min25 = New System.Windows.Forms.Label()
        Me.lbl_min18 = New System.Windows.Forms.Label()
        Me.lbl_min16 = New System.Windows.Forms.Label()
        Me.lbl_min26 = New System.Windows.Forms.Label()
        Me.lbl_min27 = New System.Windows.Forms.Label()
        Me.lbl_min28 = New System.Windows.Forms.Label()
        Me.lbl_min21 = New System.Windows.Forms.Label()
        Me.lbl_min15 = New System.Windows.Forms.Label()
        Me.lbl_min29 = New System.Windows.Forms.Label()
        Me.lbl_min30 = New System.Windows.Forms.Label()
        Me.lbl_min31 = New System.Windows.Forms.Label()
        Me.lbl_min32 = New System.Windows.Forms.Label()
        Me.lbl_min35 = New System.Windows.Forms.Label()
        Me.lbl_min34 = New System.Windows.Forms.Label()
        Me.lbl_min33 = New System.Windows.Forms.Label()
        Me.lbl_min40 = New System.Windows.Forms.Label()
        Me.lbl_min39 = New System.Windows.Forms.Label()
        Me.lbl_min38 = New System.Windows.Forms.Label()
        Me.lbl_min37 = New System.Windows.Forms.Label()
        Me.lbl_min36 = New System.Windows.Forms.Label()
        Me.PictureBox39 = New System.Windows.Forms.PictureBox()
        Me.lbl_min50 = New System.Windows.Forms.Label()
        Me.lbl_min49 = New System.Windows.Forms.Label()
        Me.lbl_min47 = New System.Windows.Forms.Label()
        Me.lbl_min46 = New System.Windows.Forms.Label()
        Me.lbl_min45 = New System.Windows.Forms.Label()
        Me.lbl_min48 = New System.Windows.Forms.Label()
        Me.lbl_min44 = New System.Windows.Forms.Label()
        Me.lbl_min43 = New System.Windows.Forms.Label()
        Me.lbl_min42 = New System.Windows.Forms.Label()
        Me.lbl_min41 = New System.Windows.Forms.Label()
        Me.lbl_min0 = New System.Windows.Forms.Label()
        Me.lbl_min59 = New System.Windows.Forms.Label()
        Me.lbl_min57 = New System.Windows.Forms.Label()
        Me.lbl_min56 = New System.Windows.Forms.Label()
        Me.lbl_min55 = New System.Windows.Forms.Label()
        Me.lbl_min58 = New System.Windows.Forms.Label()
        Me.lbl_min54 = New System.Windows.Forms.Label()
        Me.lbl_min53 = New System.Windows.Forms.Label()
        Me.lbl_min52 = New System.Windows.Forms.Label()
        Me.lbl_min51 = New System.Windows.Forms.Label()
        Me.lbl_min4 = New System.Windows.Forms.Label()
        Me.lbl_min7 = New System.Windows.Forms.Label()
        Me.lbl_min22 = New System.Windows.Forms.Label()
        Me.lbl_min24 = New System.Windows.Forms.Label()
        Me.PictureBox_clock = New System.Windows.Forms.PictureBox()
        Me.PictureBox26 = New System.Windows.Forms.PictureBox()
        Me.PictureBox24 = New System.Windows.Forms.PictureBox()
        Me.PictureBox23 = New System.Windows.Forms.PictureBox()
        Me.PictureBox22 = New System.Windows.Forms.PictureBox()
        Me.PictureBox25 = New System.Windows.Forms.PictureBox()
        Me.PictureBox21 = New System.Windows.Forms.PictureBox()
        Me.PictureBox20 = New System.Windows.Forms.PictureBox()
        Me.PictureBox19 = New System.Windows.Forms.PictureBox()
        Me.PictureBox18 = New System.Windows.Forms.PictureBox()
        Me.PictureBox17 = New System.Windows.Forms.PictureBox()
        Me.PictureBox16 = New System.Windows.Forms.PictureBox()
        Me.PictureBox15 = New System.Windows.Forms.PictureBox()
        Me.PictureBox13 = New System.Windows.Forms.PictureBox()
        Me.PictureBox14 = New System.Windows.Forms.PictureBox()
        Me.PictureBox12 = New System.Windows.Forms.PictureBox()
        Me.PictureBox11 = New System.Windows.Forms.PictureBox()
        Me.PictureBox10 = New System.Windows.Forms.PictureBox()
        Me.PictureBox9 = New System.Windows.Forms.PictureBox()
        Me.PictureBox8 = New System.Windows.Forms.PictureBox()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.PictureBox6 = New System.Windows.Forms.PictureBox()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox0 = New System.Windows.Forms.PictureBox()
        Me.PictureBox45 = New System.Windows.Forms.PictureBox()
        Me.PictureBox50 = New System.Windows.Forms.PictureBox()
        Me.PictureBox55 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox27, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox28, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox29, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox30, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox31, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox32, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox33, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox34, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox35, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox36, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox37, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox38, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox41, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox40, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox42, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox43, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox44, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox46, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox47, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox48, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox49, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox51, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox52, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox53, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox54, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox56, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox57, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox58, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox59, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBxHR0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox39, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox_clock, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox26, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox24, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox23, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox22, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox25, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox21, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox20, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox19, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox18, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox16, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox14, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox45, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox50, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox55, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Timer_sec
        '
        Me.Timer_sec.Enabled = True
        Me.Timer_sec.Interval = 1000
        '
        'PictureBox27
        '
        Me.PictureBox27.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox27.Location = New System.Drawing.Point(486, 570)
        Me.PictureBox27.Name = "PictureBox27"
        Me.PictureBox27.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox27.TabIndex = 62
        Me.PictureBox27.TabStop = False
        '
        'PictureBox28
        '
        Me.PictureBox28.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox28.Location = New System.Drawing.Point(456, 578)
        Me.PictureBox28.Name = "PictureBox28"
        Me.PictureBox28.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox28.TabIndex = 63
        Me.PictureBox28.TabStop = False
        '
        'PictureBox29
        '
        Me.PictureBox29.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox29.Location = New System.Drawing.Point(427, 582)
        Me.PictureBox29.Name = "PictureBox29"
        Me.PictureBox29.Size = New System.Drawing.Size(23, 20)
        Me.PictureBox29.TabIndex = 64
        Me.PictureBox29.TabStop = False
        '
        'PictureBox30
        '
        Me.PictureBox30.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox30.Location = New System.Drawing.Point(396, 546)
        Me.PictureBox30.Name = "PictureBox30"
        Me.PictureBox30.Size = New System.Drawing.Size(50, 29)
        Me.PictureBox30.TabIndex = 65
        Me.PictureBox30.TabStop = False
        '
        'PictureBox31
        '
        Me.PictureBox31.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox31.Location = New System.Drawing.Point(378, 584)
        Me.PictureBox31.Name = "PictureBox31"
        Me.PictureBox31.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox31.TabIndex = 66
        Me.PictureBox31.TabStop = False
        '
        'PictureBox32
        '
        Me.PictureBox32.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox32.Location = New System.Drawing.Point(352, 579)
        Me.PictureBox32.Name = "PictureBox32"
        Me.PictureBox32.Size = New System.Drawing.Size(22, 21)
        Me.PictureBox32.TabIndex = 67
        Me.PictureBox32.TabStop = False
        '
        'PictureBox33
        '
        Me.PictureBox33.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox33.Location = New System.Drawing.Point(326, 572)
        Me.PictureBox33.Name = "PictureBox33"
        Me.PictureBox33.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox33.TabIndex = 68
        Me.PictureBox33.TabStop = False
        '
        'PictureBox34
        '
        Me.PictureBox34.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox34.Location = New System.Drawing.Point(301, 563)
        Me.PictureBox34.Name = "PictureBox34"
        Me.PictureBox34.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox34.TabIndex = 69
        Me.PictureBox34.TabStop = False
        '
        'PictureBox35
        '
        Me.PictureBox35.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox35.Location = New System.Drawing.Point(288, 518)
        Me.PictureBox35.Name = "PictureBox35"
        Me.PictureBox35.Size = New System.Drawing.Size(50, 29)
        Me.PictureBox35.TabIndex = 70
        Me.PictureBox35.TabStop = False
        '
        'PictureBox36
        '
        Me.PictureBox36.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox36.Location = New System.Drawing.Point(265, 543)
        Me.PictureBox36.Name = "PictureBox36"
        Me.PictureBox36.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox36.TabIndex = 71
        Me.PictureBox36.TabStop = False
        '
        'PictureBox37
        '
        Me.PictureBox37.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox37.Location = New System.Drawing.Point(245, 530)
        Me.PictureBox37.Name = "PictureBox37"
        Me.PictureBox37.Size = New System.Drawing.Size(23, 20)
        Me.PictureBox37.TabIndex = 72
        Me.PictureBox37.TabStop = False
        '
        'PictureBox38
        '
        Me.PictureBox38.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox38.Location = New System.Drawing.Point(224, 510)
        Me.PictureBox38.Name = "PictureBox38"
        Me.PictureBox38.Size = New System.Drawing.Size(24, 21)
        Me.PictureBox38.TabIndex = 73
        Me.PictureBox38.TabStop = False
        '
        'PictureBox41
        '
        Me.PictureBox41.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox41.Location = New System.Drawing.Point(176, 453)
        Me.PictureBox41.Name = "PictureBox41"
        Me.PictureBox41.Size = New System.Drawing.Size(23, 22)
        Me.PictureBox41.TabIndex = 75
        Me.PictureBox41.TabStop = False
        '
        'PictureBox40
        '
        Me.PictureBox40.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox40.Location = New System.Drawing.Point(227, 453)
        Me.PictureBox40.Name = "PictureBox40"
        Me.PictureBox40.Size = New System.Drawing.Size(50, 29)
        Me.PictureBox40.TabIndex = 76
        Me.PictureBox40.TabStop = False
        '
        'PictureBox42
        '
        Me.PictureBox42.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox42.Location = New System.Drawing.Point(169, 430)
        Me.PictureBox42.Name = "PictureBox42"
        Me.PictureBox42.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox42.TabIndex = 81
        Me.PictureBox42.TabStop = False
        '
        'PictureBox43
        '
        Me.PictureBox43.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox43.Location = New System.Drawing.Point(164, 410)
        Me.PictureBox43.Name = "PictureBox43"
        Me.PictureBox43.Size = New System.Drawing.Size(23, 20)
        Me.PictureBox43.TabIndex = 82
        Me.PictureBox43.TabStop = False
        '
        'PictureBox44
        '
        Me.PictureBox44.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox44.Location = New System.Drawing.Point(158, 385)
        Me.PictureBox44.Name = "PictureBox44"
        Me.PictureBox44.Size = New System.Drawing.Size(33, 21)
        Me.PictureBox44.TabIndex = 83
        Me.PictureBox44.TabStop = False
        '
        'PictureBox46
        '
        Me.PictureBox46.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox46.Location = New System.Drawing.Point(158, 338)
        Me.PictureBox46.Name = "PictureBox46"
        Me.PictureBox46.Size = New System.Drawing.Size(27, 21)
        Me.PictureBox46.TabIndex = 84
        Me.PictureBox46.TabStop = False
        '
        'PictureBox47
        '
        Me.PictureBox47.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox47.Location = New System.Drawing.Point(163, 315)
        Me.PictureBox47.Name = "PictureBox47"
        Me.PictureBox47.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox47.TabIndex = 85
        Me.PictureBox47.TabStop = False
        '
        'PictureBox48
        '
        Me.PictureBox48.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox48.Location = New System.Drawing.Point(169, 291)
        Me.PictureBox48.Name = "PictureBox48"
        Me.PictureBox48.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox48.TabIndex = 86
        Me.PictureBox48.TabStop = False
        '
        'PictureBox49
        '
        Me.PictureBox49.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox49.Location = New System.Drawing.Point(179, 269)
        Me.PictureBox49.Name = "PictureBox49"
        Me.PictureBox49.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox49.TabIndex = 87
        Me.PictureBox49.TabStop = False
        '
        'PictureBox51
        '
        Me.PictureBox51.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox51.Location = New System.Drawing.Point(208, 230)
        Me.PictureBox51.Name = "PictureBox51"
        Me.PictureBox51.Size = New System.Drawing.Size(22, 20)
        Me.PictureBox51.TabIndex = 88
        Me.PictureBox51.TabStop = False
        '
        'PictureBox52
        '
        Me.PictureBox52.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox52.Location = New System.Drawing.Point(223, 210)
        Me.PictureBox52.Name = "PictureBox52"
        Me.PictureBox52.Size = New System.Drawing.Size(23, 22)
        Me.PictureBox52.TabIndex = 89
        Me.PictureBox52.TabStop = False
        '
        'PictureBox53
        '
        Me.PictureBox53.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox53.Location = New System.Drawing.Point(244, 195)
        Me.PictureBox53.Name = "PictureBox53"
        Me.PictureBox53.Size = New System.Drawing.Size(22, 21)
        Me.PictureBox53.TabIndex = 90
        Me.PictureBox53.TabStop = False
        '
        'PictureBox54
        '
        Me.PictureBox54.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox54.Location = New System.Drawing.Point(265, 180)
        Me.PictureBox54.Name = "PictureBox54"
        Me.PictureBox54.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox54.TabIndex = 91
        Me.PictureBox54.TabStop = False
        '
        'PictureBox56
        '
        Me.PictureBox56.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox56.Location = New System.Drawing.Point(300, 159)
        Me.PictureBox56.Name = "PictureBox56"
        Me.PictureBox56.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox56.TabIndex = 92
        Me.PictureBox56.TabStop = False
        '
        'PictureBox57
        '
        Me.PictureBox57.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox57.Location = New System.Drawing.Point(324, 149)
        Me.PictureBox57.Name = "PictureBox57"
        Me.PictureBox57.Size = New System.Drawing.Size(26, 19)
        Me.PictureBox57.TabIndex = 93
        Me.PictureBox57.TabStop = False
        '
        'PictureBox58
        '
        Me.PictureBox58.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox58.Location = New System.Drawing.Point(346, 142)
        Me.PictureBox58.Name = "PictureBox58"
        Me.PictureBox58.Size = New System.Drawing.Size(32, 20)
        Me.PictureBox58.TabIndex = 94
        Me.PictureBox58.TabStop = False
        '
        'PictureBox59
        '
        Me.PictureBox59.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox59.Location = New System.Drawing.Point(378, 138)
        Me.PictureBox59.Name = "PictureBox59"
        Me.PictureBox59.Size = New System.Drawing.Size(29, 21)
        Me.PictureBox59.TabIndex = 95
        Me.PictureBox59.TabStop = False
        '
        'PicBxHR1
        '
        Me.PicBxHR1.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR1.Location = New System.Drawing.Point(574, 80)
        Me.PicBxHR1.Name = "PicBxHR1"
        Me.PicBxHR1.Size = New System.Drawing.Size(33, 54)
        Me.PicBxHR1.TabIndex = 96
        Me.PicBxHR1.TabStop = False
        '
        'PicBxHR2
        '
        Me.PicBxHR2.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR2.Location = New System.Drawing.Point(703, 190)
        Me.PicBxHR2.Name = "PicBxHR2"
        Me.PicBxHR2.Size = New System.Drawing.Size(37, 53)
        Me.PicBxHR2.TabIndex = 97
        Me.PicBxHR2.TabStop = False
        '
        'PicBxHR3
        '
        Me.PicBxHR3.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR3.Location = New System.Drawing.Point(748, 342)
        Me.PicBxHR3.Name = "PicBxHR3"
        Me.PicBxHR3.Size = New System.Drawing.Size(37, 54)
        Me.PicBxHR3.TabIndex = 98
        Me.PicBxHR3.TabStop = False
        '
        'PicBxHR4
        '
        Me.PicBxHR4.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR4.Location = New System.Drawing.Point(686, 495)
        Me.PicBxHR4.Name = "PicBxHR4"
        Me.PicBxHR4.Size = New System.Drawing.Size(46, 53)
        Me.PicBxHR4.TabIndex = 99
        Me.PicBxHR4.TabStop = False
        '
        'PicBxHR5
        '
        Me.PicBxHR5.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR5.Location = New System.Drawing.Point(566, 612)
        Me.PicBxHR5.Name = "PicBxHR5"
        Me.PicBxHR5.Size = New System.Drawing.Size(47, 54)
        Me.PicBxHR5.TabIndex = 100
        Me.PicBxHR5.TabStop = False
        '
        'PicBxHR6
        '
        Me.PicBxHR6.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR6.Location = New System.Drawing.Point(396, 651)
        Me.PicBxHR6.Name = "PicBxHR6"
        Me.PicBxHR6.Size = New System.Drawing.Size(46, 54)
        Me.PicBxHR6.TabIndex = 101
        Me.PicBxHR6.TabStop = False
        '
        'PicBxHR7
        '
        Me.PicBxHR7.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR7.Location = New System.Drawing.Point(227, 615)
        Me.PicBxHR7.Name = "PicBxHR7"
        Me.PicBxHR7.Size = New System.Drawing.Size(45, 54)
        Me.PicBxHR7.TabIndex = 102
        Me.PicBxHR7.TabStop = False
        '
        'PicBxHR8
        '
        Me.PicBxHR8.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR8.Location = New System.Drawing.Point(94, 496)
        Me.PicBxHR8.Name = "PicBxHR8"
        Me.PicBxHR8.Size = New System.Drawing.Size(46, 54)
        Me.PicBxHR8.TabIndex = 103
        Me.PicBxHR8.TabStop = False
        '
        'PicBxHR9
        '
        Me.PicBxHR9.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR9.Location = New System.Drawing.Point(50, 342)
        Me.PicBxHR9.Name = "PicBxHR9"
        Me.PicBxHR9.Size = New System.Drawing.Size(47, 64)
        Me.PicBxHR9.TabIndex = 104
        Me.PicBxHR9.TabStop = False
        '
        'PicBxHR10
        '
        Me.PicBxHR10.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR10.Location = New System.Drawing.Point(82, 197)
        Me.PicBxHR10.Name = "PicBxHR10"
        Me.PicBxHR10.Size = New System.Drawing.Size(69, 49)
        Me.PicBxHR10.TabIndex = 105
        Me.PicBxHR10.TabStop = False
        '
        'PicBxHR11
        '
        Me.PicBxHR11.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR11.Location = New System.Drawing.Point(220, 78)
        Me.PicBxHR11.Name = "PicBxHR11"
        Me.PicBxHR11.Size = New System.Drawing.Size(55, 52)
        Me.PicBxHR11.TabIndex = 106
        Me.PicBxHR11.TabStop = False
        '
        'PicBxHR0
        '
        Me.PicBxHR0.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PicBxHR0.Location = New System.Drawing.Point(355, 38)
        Me.PicBxHR0.Name = "PicBxHR0"
        Me.PicBxHR0.Size = New System.Drawing.Size(112, 54)
        Me.PicBxHR0.TabIndex = 107
        Me.PicBxHR0.TabStop = False
        Me.PicBxHR0.Tag = ""
        '
        'lbl_min1
        '
        Me.lbl_min1.AutoSize = True
        Me.lbl_min1.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min1.Location = New System.Drawing.Point(428, 144)
        Me.lbl_min1.Name = "lbl_min1"
        Me.lbl_min1.Size = New System.Drawing.Size(18, 20)
        Me.lbl_min1.TabIndex = 108
        Me.lbl_min1.Text = "1"
        '
        'Timer_min
        '
        Me.Timer_min.Enabled = True
        Me.Timer_min.Interval = 1000
        '
        'lbl_min2
        '
        Me.lbl_min2.AutoSize = True
        Me.lbl_min2.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min2.Location = New System.Drawing.Point(464, 147)
        Me.lbl_min2.Name = "lbl_min2"
        Me.lbl_min2.Size = New System.Drawing.Size(18, 20)
        Me.lbl_min2.TabIndex = 109
        Me.lbl_min2.Text = "2"
        '
        'lbl_min3
        '
        Me.lbl_min3.AutoSize = True
        Me.lbl_min3.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min3.Location = New System.Drawing.Point(488, 154)
        Me.lbl_min3.Name = "lbl_min3"
        Me.lbl_min3.Size = New System.Drawing.Size(18, 20)
        Me.lbl_min3.TabIndex = 110
        Me.lbl_min3.Text = "3"
        '
        'lbl_min6
        '
        Me.lbl_min6.AutoSize = True
        Me.lbl_min6.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min6.Location = New System.Drawing.Point(554, 183)
        Me.lbl_min6.Name = "lbl_min6"
        Me.lbl_min6.Size = New System.Drawing.Size(18, 20)
        Me.lbl_min6.TabIndex = 113
        Me.lbl_min6.Text = "6"
        '
        'lbl_min5
        '
        Me.lbl_min5.AutoSize = True
        Me.lbl_min5.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min5.Location = New System.Drawing.Point(512, 202)
        Me.lbl_min5.Name = "lbl_min5"
        Me.lbl_min5.Size = New System.Drawing.Size(18, 20)
        Me.lbl_min5.TabIndex = 112
        Me.lbl_min5.Text = "5"
        '
        'lbl_min14
        '
        Me.lbl_min14.AutoSize = True
        Me.lbl_min14.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min14.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min14.Location = New System.Drawing.Point(649, 342)
        Me.lbl_min14.Name = "lbl_min14"
        Me.lbl_min14.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min14.TabIndex = 121
        Me.lbl_min14.Text = "14"
        '
        'lbl_min13
        '
        Me.lbl_min13.AutoSize = True
        Me.lbl_min13.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min13.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min13.Location = New System.Drawing.Point(649, 315)
        Me.lbl_min13.Name = "lbl_min13"
        Me.lbl_min13.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min13.TabIndex = 120
        Me.lbl_min13.Text = "13"
        '
        'lbl_min12
        '
        Me.lbl_min12.AutoSize = True
        Me.lbl_min12.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min12.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min12.Location = New System.Drawing.Point(638, 291)
        Me.lbl_min12.Name = "lbl_min12"
        Me.lbl_min12.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min12.TabIndex = 119
        Me.lbl_min12.Text = "12"
        '
        'lbl_min11
        '
        Me.lbl_min11.AutoSize = True
        Me.lbl_min11.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min11.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min11.Location = New System.Drawing.Point(629, 267)
        Me.lbl_min11.Name = "lbl_min11"
        Me.lbl_min11.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min11.TabIndex = 118
        Me.lbl_min11.Text = "11"
        '
        'lbl_min10
        '
        Me.lbl_min10.AutoSize = True
        Me.lbl_min10.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min10.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min10.Location = New System.Drawing.Point(576, 267)
        Me.lbl_min10.Name = "lbl_min10"
        Me.lbl_min10.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min10.TabIndex = 117
        Me.lbl_min10.Text = "10"
        '
        'lbl_min9
        '
        Me.lbl_min9.AutoSize = True
        Me.lbl_min9.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min9.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min9.Location = New System.Drawing.Point(610, 230)
        Me.lbl_min9.Name = "lbl_min9"
        Me.lbl_min9.Size = New System.Drawing.Size(18, 20)
        Me.lbl_min9.TabIndex = 116
        Me.lbl_min9.Text = "9"
        '
        'lbl_min8
        '
        Me.lbl_min8.AutoSize = True
        Me.lbl_min8.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min8.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min8.Location = New System.Drawing.Point(595, 211)
        Me.lbl_min8.Name = "lbl_min8"
        Me.lbl_min8.Size = New System.Drawing.Size(18, 20)
        Me.lbl_min8.TabIndex = 115
        Me.lbl_min8.Text = "8"
        '
        'lbl_min23
        '
        Me.lbl_min23.AutoSize = True
        Me.lbl_min23.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min23.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min23.Location = New System.Drawing.Point(572, 526)
        Me.lbl_min23.Name = "lbl_min23"
        Me.lbl_min23.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min23.TabIndex = 135
        Me.lbl_min23.Text = "23"
        '
        'lbl_min20
        '
        Me.lbl_min20.AutoSize = True
        Me.lbl_min20.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min20.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min20.Location = New System.Drawing.Point(570, 453)
        Me.lbl_min20.Name = "lbl_min20"
        Me.lbl_min20.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min20.TabIndex = 133
        Me.lbl_min20.Text = "20"
        '
        'lbl_min19
        '
        Me.lbl_min19.AutoSize = True
        Me.lbl_min19.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min19.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min19.Location = New System.Drawing.Point(628, 453)
        Me.lbl_min19.Name = "lbl_min19"
        Me.lbl_min19.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min19.TabIndex = 132
        Me.lbl_min19.Text = "19"
        '
        'lbl_min17
        '
        Me.lbl_min17.AutoSize = True
        Me.lbl_min17.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min17.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min17.Location = New System.Drawing.Point(649, 406)
        Me.lbl_min17.Name = "lbl_min17"
        Me.lbl_min17.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min17.TabIndex = 131
        Me.lbl_min17.Text = "17"
        '
        'lbl_min25
        '
        Me.lbl_min25.AutoSize = True
        Me.lbl_min25.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min25.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min25.Location = New System.Drawing.Point(520, 518)
        Me.lbl_min25.Name = "lbl_min25"
        Me.lbl_min25.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min25.TabIndex = 129
        Me.lbl_min25.Text = "25"
        '
        'lbl_min18
        '
        Me.lbl_min18.AutoSize = True
        Me.lbl_min18.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min18.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min18.Location = New System.Drawing.Point(645, 426)
        Me.lbl_min18.Name = "lbl_min18"
        Me.lbl_min18.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min18.TabIndex = 128
        Me.lbl_min18.Text = "18"
        '
        'lbl_min16
        '
        Me.lbl_min16.AutoSize = True
        Me.lbl_min16.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min16.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min16.Location = New System.Drawing.Point(649, 380)
        Me.lbl_min16.Name = "lbl_min16"
        Me.lbl_min16.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min16.TabIndex = 127
        Me.lbl_min16.Text = "16"
        '
        'lbl_min26
        '
        Me.lbl_min26.AutoSize = True
        Me.lbl_min26.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min26.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min26.Location = New System.Drawing.Point(515, 563)
        Me.lbl_min26.Name = "lbl_min26"
        Me.lbl_min26.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min26.TabIndex = 126
        Me.lbl_min26.Text = "26"
        '
        'lbl_min27
        '
        Me.lbl_min27.AutoSize = True
        Me.lbl_min27.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min27.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min27.Location = New System.Drawing.Point(491, 570)
        Me.lbl_min27.Name = "lbl_min27"
        Me.lbl_min27.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min27.TabIndex = 125
        Me.lbl_min27.Text = "27"
        '
        'lbl_min28
        '
        Me.lbl_min28.AutoSize = True
        Me.lbl_min28.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min28.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min28.Location = New System.Drawing.Point(460, 578)
        Me.lbl_min28.Name = "lbl_min28"
        Me.lbl_min28.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min28.TabIndex = 124
        Me.lbl_min28.Text = "28"
        '
        'lbl_min21
        '
        Me.lbl_min21.AutoSize = True
        Me.lbl_min21.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min21.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min21.Location = New System.Drawing.Point(610, 489)
        Me.lbl_min21.Name = "lbl_min21"
        Me.lbl_min21.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min21.TabIndex = 123
        Me.lbl_min21.Text = "21"
        '
        'lbl_min15
        '
        Me.lbl_min15.AutoSize = True
        Me.lbl_min15.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min15.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min15.Location = New System.Drawing.Point(610, 362)
        Me.lbl_min15.Name = "lbl_min15"
        Me.lbl_min15.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min15.TabIndex = 122
        Me.lbl_min15.Text = "15"
        '
        'lbl_min29
        '
        Me.lbl_min29.AutoSize = True
        Me.lbl_min29.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min29.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min29.Location = New System.Drawing.Point(426, 582)
        Me.lbl_min29.Name = "lbl_min29"
        Me.lbl_min29.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min29.TabIndex = 136
        Me.lbl_min29.Text = "29"
        '
        'lbl_min30
        '
        Me.lbl_min30.AutoSize = True
        Me.lbl_min30.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min30.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min30.Location = New System.Drawing.Point(406, 546)
        Me.lbl_min30.Name = "lbl_min30"
        Me.lbl_min30.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min30.TabIndex = 137
        Me.lbl_min30.Text = "30"
        '
        'lbl_min31
        '
        Me.lbl_min31.AutoSize = True
        Me.lbl_min31.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min31.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min31.Location = New System.Drawing.Point(380, 582)
        Me.lbl_min31.Name = "lbl_min31"
        Me.lbl_min31.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min31.TabIndex = 138
        Me.lbl_min31.Text = "31"
        '
        'lbl_min32
        '
        Me.lbl_min32.AutoSize = True
        Me.lbl_min32.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min32.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min32.Location = New System.Drawing.Point(351, 578)
        Me.lbl_min32.Name = "lbl_min32"
        Me.lbl_min32.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min32.TabIndex = 139
        Me.lbl_min32.Text = "32"
        '
        'lbl_min35
        '
        Me.lbl_min35.AutoSize = True
        Me.lbl_min35.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min35.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min35.Location = New System.Drawing.Point(296, 518)
        Me.lbl_min35.Name = "lbl_min35"
        Me.lbl_min35.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min35.TabIndex = 142
        Me.lbl_min35.Text = "35"
        '
        'lbl_min34
        '
        Me.lbl_min34.AutoSize = True
        Me.lbl_min34.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min34.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min34.Location = New System.Drawing.Point(297, 563)
        Me.lbl_min34.Name = "lbl_min34"
        Me.lbl_min34.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min34.TabIndex = 141
        Me.lbl_min34.Text = "34"
        '
        'lbl_min33
        '
        Me.lbl_min33.AutoSize = True
        Me.lbl_min33.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min33.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min33.Location = New System.Drawing.Point(326, 570)
        Me.lbl_min33.Name = "lbl_min33"
        Me.lbl_min33.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min33.TabIndex = 140
        Me.lbl_min33.Text = "33"
        '
        'lbl_min40
        '
        Me.lbl_min40.AutoSize = True
        Me.lbl_min40.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min40.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min40.Location = New System.Drawing.Point(223, 453)
        Me.lbl_min40.Name = "lbl_min40"
        Me.lbl_min40.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min40.TabIndex = 147
        Me.lbl_min40.Text = "40"
        '
        'lbl_min39
        '
        Me.lbl_min39.AutoSize = True
        Me.lbl_min39.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min39.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min39.Location = New System.Drawing.Point(204, 489)
        Me.lbl_min39.Name = "lbl_min39"
        Me.lbl_min39.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min39.TabIndex = 146
        Me.lbl_min39.Text = "39"
        '
        'lbl_min38
        '
        Me.lbl_min38.AutoSize = True
        Me.lbl_min38.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min38.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min38.Location = New System.Drawing.Point(221, 507)
        Me.lbl_min38.Name = "lbl_min38"
        Me.lbl_min38.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min38.TabIndex = 145
        Me.lbl_min38.Text = "38"
        '
        'lbl_min37
        '
        Me.lbl_min37.AutoSize = True
        Me.lbl_min37.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min37.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min37.Location = New System.Drawing.Point(240, 525)
        Me.lbl_min37.Name = "lbl_min37"
        Me.lbl_min37.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min37.TabIndex = 144
        Me.lbl_min37.Text = "37"
        '
        'lbl_min36
        '
        Me.lbl_min36.AutoSize = True
        Me.lbl_min36.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min36.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min36.Location = New System.Drawing.Point(261, 543)
        Me.lbl_min36.Name = "lbl_min36"
        Me.lbl_min36.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min36.TabIndex = 143
        Me.lbl_min36.Text = "36"
        '
        'PictureBox39
        '
        Me.PictureBox39.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox39.Location = New System.Drawing.Point(208, 488)
        Me.PictureBox39.Name = "PictureBox39"
        Me.PictureBox39.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox39.TabIndex = 74
        Me.PictureBox39.TabStop = False
        '
        'lbl_min50
        '
        Me.lbl_min50.AutoSize = True
        Me.lbl_min50.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min50.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min50.Location = New System.Drawing.Point(228, 265)
        Me.lbl_min50.Name = "lbl_min50"
        Me.lbl_min50.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min50.TabIndex = 157
        Me.lbl_min50.Text = "50"
        '
        'lbl_min49
        '
        Me.lbl_min49.AutoSize = True
        Me.lbl_min49.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min49.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min49.Location = New System.Drawing.Point(175, 268)
        Me.lbl_min49.Name = "lbl_min49"
        Me.lbl_min49.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min49.TabIndex = 156
        Me.lbl_min49.Text = "49"
        '
        'lbl_min47
        '
        Me.lbl_min47.AutoSize = True
        Me.lbl_min47.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min47.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min47.Location = New System.Drawing.Point(165, 316)
        Me.lbl_min47.Name = "lbl_min47"
        Me.lbl_min47.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min47.TabIndex = 155
        Me.lbl_min47.Text = "47"
        '
        'lbl_min46
        '
        Me.lbl_min46.AutoSize = True
        Me.lbl_min46.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min46.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min46.Location = New System.Drawing.Point(159, 339)
        Me.lbl_min46.Name = "lbl_min46"
        Me.lbl_min46.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min46.TabIndex = 154
        Me.lbl_min46.Text = "46"
        '
        'lbl_min45
        '
        Me.lbl_min45.AutoSize = True
        Me.lbl_min45.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min45.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min45.Location = New System.Drawing.Point(203, 362)
        Me.lbl_min45.Name = "lbl_min45"
        Me.lbl_min45.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min45.TabIndex = 153
        Me.lbl_min45.Text = "45"
        '
        'lbl_min48
        '
        Me.lbl_min48.AutoSize = True
        Me.lbl_min48.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min48.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min48.Location = New System.Drawing.Point(165, 291)
        Me.lbl_min48.Name = "lbl_min48"
        Me.lbl_min48.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min48.TabIndex = 152
        Me.lbl_min48.Text = "48"
        '
        'lbl_min44
        '
        Me.lbl_min44.AutoSize = True
        Me.lbl_min44.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min44.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min44.Location = New System.Drawing.Point(160, 385)
        Me.lbl_min44.Name = "lbl_min44"
        Me.lbl_min44.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min44.TabIndex = 151
        Me.lbl_min44.Text = "44"
        '
        'lbl_min43
        '
        Me.lbl_min43.AutoSize = True
        Me.lbl_min43.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min43.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min43.Location = New System.Drawing.Point(160, 409)
        Me.lbl_min43.Name = "lbl_min43"
        Me.lbl_min43.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min43.TabIndex = 150
        Me.lbl_min43.Text = "43"
        '
        'lbl_min42
        '
        Me.lbl_min42.AutoSize = True
        Me.lbl_min42.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min42.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min42.Location = New System.Drawing.Point(165, 430)
        Me.lbl_min42.Name = "lbl_min42"
        Me.lbl_min42.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min42.TabIndex = 149
        Me.lbl_min42.Text = "42"
        '
        'lbl_min41
        '
        Me.lbl_min41.AutoSize = True
        Me.lbl_min41.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min41.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min41.Location = New System.Drawing.Point(175, 453)
        Me.lbl_min41.Name = "lbl_min41"
        Me.lbl_min41.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min41.TabIndex = 148
        Me.lbl_min41.Text = "41"
        '
        'lbl_min0
        '
        Me.lbl_min0.AutoSize = True
        Me.lbl_min0.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min0.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min0.Location = New System.Drawing.Point(406, 181)
        Me.lbl_min0.Name = "lbl_min0"
        Me.lbl_min0.Size = New System.Drawing.Size(18, 20)
        Me.lbl_min0.TabIndex = 167
        Me.lbl_min0.Text = "0"
        '
        'lbl_min59
        '
        Me.lbl_min59.AutoSize = True
        Me.lbl_min59.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min59.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min59.Location = New System.Drawing.Point(381, 138)
        Me.lbl_min59.Name = "lbl_min59"
        Me.lbl_min59.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min59.TabIndex = 166
        Me.lbl_min59.Text = "59"
        '
        'lbl_min57
        '
        Me.lbl_min57.AutoSize = True
        Me.lbl_min57.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min57.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min57.Location = New System.Drawing.Point(326, 151)
        Me.lbl_min57.Name = "lbl_min57"
        Me.lbl_min57.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min57.TabIndex = 165
        Me.lbl_min57.Text = "57"
        '
        'lbl_min56
        '
        Me.lbl_min56.AutoSize = True
        Me.lbl_min56.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min56.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min56.Location = New System.Drawing.Point(297, 161)
        Me.lbl_min56.Name = "lbl_min56"
        Me.lbl_min56.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min56.TabIndex = 164
        Me.lbl_min56.Text = "56"
        '
        'lbl_min55
        '
        Me.lbl_min55.AutoSize = True
        Me.lbl_min55.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min55.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min55.Location = New System.Drawing.Point(297, 202)
        Me.lbl_min55.Name = "lbl_min55"
        Me.lbl_min55.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min55.TabIndex = 163
        Me.lbl_min55.Text = "55"
        '
        'lbl_min58
        '
        Me.lbl_min58.AutoSize = True
        Me.lbl_min58.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min58.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min58.Location = New System.Drawing.Point(348, 142)
        Me.lbl_min58.Name = "lbl_min58"
        Me.lbl_min58.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min58.TabIndex = 162
        Me.lbl_min58.Text = "58"
        '
        'lbl_min54
        '
        Me.lbl_min54.AutoSize = True
        Me.lbl_min54.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min54.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min54.Location = New System.Drawing.Point(261, 183)
        Me.lbl_min54.Name = "lbl_min54"
        Me.lbl_min54.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min54.TabIndex = 161
        Me.lbl_min54.Text = "54"
        '
        'lbl_min53
        '
        Me.lbl_min53.AutoSize = True
        Me.lbl_min53.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min53.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min53.Location = New System.Drawing.Point(241, 195)
        Me.lbl_min53.Name = "lbl_min53"
        Me.lbl_min53.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min53.TabIndex = 160
        Me.lbl_min53.Text = "53"
        '
        'lbl_min52
        '
        Me.lbl_min52.AutoSize = True
        Me.lbl_min52.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min52.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min52.Location = New System.Drawing.Point(220, 211)
        Me.lbl_min52.Name = "lbl_min52"
        Me.lbl_min52.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min52.TabIndex = 159
        Me.lbl_min52.Text = "52"
        '
        'lbl_min51
        '
        Me.lbl_min51.AutoSize = True
        Me.lbl_min51.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min51.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min51.Location = New System.Drawing.Point(204, 230)
        Me.lbl_min51.Name = "lbl_min51"
        Me.lbl_min51.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min51.TabIndex = 158
        Me.lbl_min51.Text = "51"
        '
        'lbl_min4
        '
        Me.lbl_min4.AutoSize = True
        Me.lbl_min4.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min4.Location = New System.Drawing.Point(512, 161)
        Me.lbl_min4.Name = "lbl_min4"
        Me.lbl_min4.Size = New System.Drawing.Size(18, 20)
        Me.lbl_min4.TabIndex = 111
        Me.lbl_min4.Text = "4"
        '
        'lbl_min7
        '
        Me.lbl_min7.AutoSize = True
        Me.lbl_min7.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min7.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min7.Location = New System.Drawing.Point(572, 197)
        Me.lbl_min7.Name = "lbl_min7"
        Me.lbl_min7.Size = New System.Drawing.Size(18, 20)
        Me.lbl_min7.TabIndex = 114
        Me.lbl_min7.Text = "7"
        '
        'lbl_min22
        '
        Me.lbl_min22.AutoSize = True
        Me.lbl_min22.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min22.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min22.Location = New System.Drawing.Point(595, 507)
        Me.lbl_min22.Name = "lbl_min22"
        Me.lbl_min22.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min22.TabIndex = 134
        Me.lbl_min22.Text = "22"
        '
        'lbl_min24
        '
        Me.lbl_min24.AutoSize = True
        Me.lbl_min24.BackColor = System.Drawing.Color.Transparent
        Me.lbl_min24.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lbl_min24.Location = New System.Drawing.Point(562, 543)
        Me.lbl_min24.Name = "lbl_min24"
        Me.lbl_min24.Size = New System.Drawing.Size(27, 20)
        Me.lbl_min24.TabIndex = 130
        Me.lbl_min24.Text = "24"
        '
        'PictureBox_clock
        '
        Me.PictureBox_clock.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox_clock.Image = Global.junkVB.My.Resources.Resources.Copy_of_clock___Copy
        Me.PictureBox_clock.Location = New System.Drawing.Point(14, 13)
        Me.PictureBox_clock.Name = "PictureBox_clock"
        Me.PictureBox_clock.Size = New System.Drawing.Size(814, 717)
        Me.PictureBox_clock.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox_clock.TabIndex = 35
        Me.PictureBox_clock.TabStop = False
        '
        'PictureBox26
        '
        Me.PictureBox26.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox26.Location = New System.Drawing.Point(515, 564)
        Me.PictureBox26.Name = "PictureBox26"
        Me.PictureBox26.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox26.TabIndex = 61
        Me.PictureBox26.TabStop = False
        '
        'PictureBox24
        '
        Me.PictureBox24.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox24.Location = New System.Drawing.Point(558, 543)
        Me.PictureBox24.Name = "PictureBox24"
        Me.PictureBox24.Size = New System.Drawing.Size(22, 22)
        Me.PictureBox24.TabIndex = 59
        Me.PictureBox24.TabStop = False
        '
        'PictureBox23
        '
        Me.PictureBox23.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox23.Location = New System.Drawing.Point(575, 525)
        Me.PictureBox23.Name = "PictureBox23"
        Me.PictureBox23.Size = New System.Drawing.Size(23, 21)
        Me.PictureBox23.TabIndex = 58
        Me.PictureBox23.TabStop = False
        '
        'PictureBox22
        '
        Me.PictureBox22.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox22.Location = New System.Drawing.Point(593, 509)
        Me.PictureBox22.Name = "PictureBox22"
        Me.PictureBox22.Size = New System.Drawing.Size(23, 18)
        Me.PictureBox22.TabIndex = 57
        Me.PictureBox22.TabStop = False
        '
        'PictureBox25
        '
        Me.PictureBox25.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox25.Location = New System.Drawing.Point(508, 512)
        Me.PictureBox25.Name = "PictureBox25"
        Me.PictureBox25.Size = New System.Drawing.Size(40, 35)
        Me.PictureBox25.TabIndex = 60
        Me.PictureBox25.TabStop = False
        '
        'PictureBox21
        '
        Me.PictureBox21.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox21.Location = New System.Drawing.Point(608, 488)
        Me.PictureBox21.Name = "PictureBox21"
        Me.PictureBox21.Size = New System.Drawing.Size(24, 21)
        Me.PictureBox21.TabIndex = 56
        Me.PictureBox21.TabStop = False
        '
        'PictureBox20
        '
        Me.PictureBox20.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox20.Location = New System.Drawing.Point(556, 444)
        Me.PictureBox20.Name = "PictureBox20"
        Me.PictureBox20.Size = New System.Drawing.Size(51, 33)
        Me.PictureBox20.TabIndex = 55
        Me.PictureBox20.TabStop = False
        '
        'PictureBox19
        '
        Me.PictureBox19.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox19.Location = New System.Drawing.Point(632, 449)
        Me.PictureBox19.Name = "PictureBox19"
        Me.PictureBox19.Size = New System.Drawing.Size(24, 28)
        Me.PictureBox19.TabIndex = 54
        Me.PictureBox19.TabStop = False
        '
        'PictureBox18
        '
        Me.PictureBox18.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox18.Location = New System.Drawing.Point(642, 426)
        Me.PictureBox18.Name = "PictureBox18"
        Me.PictureBox18.Size = New System.Drawing.Size(23, 24)
        Me.PictureBox18.TabIndex = 53
        Me.PictureBox18.TabStop = False
        '
        'PictureBox17
        '
        Me.PictureBox17.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox17.Location = New System.Drawing.Point(649, 403)
        Me.PictureBox17.Name = "PictureBox17"
        Me.PictureBox17.Size = New System.Drawing.Size(23, 27)
        Me.PictureBox17.TabIndex = 52
        Me.PictureBox17.TabStop = False
        '
        'PictureBox16
        '
        Me.PictureBox16.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox16.Location = New System.Drawing.Point(652, 380)
        Me.PictureBox16.Name = "PictureBox16"
        Me.PictureBox16.Size = New System.Drawing.Size(22, 27)
        Me.PictureBox16.TabIndex = 51
        Me.PictureBox16.TabStop = False
        '
        'PictureBox15
        '
        Me.PictureBox15.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox15.Location = New System.Drawing.Point(580, 352)
        Me.PictureBox15.Name = "PictureBox15"
        Me.PictureBox15.Size = New System.Drawing.Size(61, 44)
        Me.PictureBox15.TabIndex = 50
        Me.PictureBox15.TabStop = False
        '
        'PictureBox13
        '
        Me.PictureBox13.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox13.Location = New System.Drawing.Point(649, 313)
        Me.PictureBox13.Name = "PictureBox13"
        Me.PictureBox13.Size = New System.Drawing.Size(23, 27)
        Me.PictureBox13.TabIndex = 45
        Me.PictureBox13.TabStop = False
        '
        'PictureBox14
        '
        Me.PictureBox14.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox14.Location = New System.Drawing.Point(653, 339)
        Me.PictureBox14.Name = "PictureBox14"
        Me.PictureBox14.Size = New System.Drawing.Size(23, 28)
        Me.PictureBox14.TabIndex = 49
        Me.PictureBox14.TabStop = False
        '
        'PictureBox12
        '
        Me.PictureBox12.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox12.Location = New System.Drawing.Point(626, 291)
        Me.PictureBox12.Name = "PictureBox12"
        Me.PictureBox12.Size = New System.Drawing.Size(39, 23)
        Me.PictureBox12.TabIndex = 46
        Me.PictureBox12.TabStop = False
        '
        'PictureBox11
        '
        Me.PictureBox11.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox11.Location = New System.Drawing.Point(628, 267)
        Me.PictureBox11.Name = "PictureBox11"
        Me.PictureBox11.Size = New System.Drawing.Size(24, 23)
        Me.PictureBox11.TabIndex = 47
        Me.PictureBox11.TabStop = False
        '
        'PictureBox10
        '
        Me.PictureBox10.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox10.Location = New System.Drawing.Point(565, 265)
        Me.PictureBox10.Name = "PictureBox10"
        Me.PictureBox10.Size = New System.Drawing.Size(42, 32)
        Me.PictureBox10.TabIndex = 48
        Me.PictureBox10.TabStop = False
        '
        'PictureBox9
        '
        Me.PictureBox9.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox9.Location = New System.Drawing.Point(611, 230)
        Me.PictureBox9.Name = "PictureBox9"
        Me.PictureBox9.Size = New System.Drawing.Size(17, 27)
        Me.PictureBox9.TabIndex = 36
        Me.PictureBox9.TabStop = False
        '
        'PictureBox8
        '
        Me.PictureBox8.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox8.Location = New System.Drawing.Point(594, 207)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(19, 31)
        Me.PictureBox8.TabIndex = 44
        Me.PictureBox8.TabStop = False
        '
        'PictureBox7
        '
        Me.PictureBox7.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox7.Location = New System.Drawing.Point(574, 196)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(16, 23)
        Me.PictureBox7.TabIndex = 43
        Me.PictureBox7.TabStop = False
        '
        'PictureBox6
        '
        Me.PictureBox6.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox6.Location = New System.Drawing.Point(556, 180)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(16, 23)
        Me.PictureBox6.TabIndex = 42
        Me.PictureBox6.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox4.Location = New System.Drawing.Point(514, 158)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(16, 23)
        Me.PictureBox4.TabIndex = 40
        Me.PictureBox4.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox5.Location = New System.Drawing.Point(514, 201)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(16, 33)
        Me.PictureBox5.TabIndex = 41
        Me.PictureBox5.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox3.Location = New System.Drawing.Point(488, 151)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 23)
        Me.PictureBox3.TabIndex = 39
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox2.Location = New System.Drawing.Point(464, 144)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(17, 27)
        Me.PictureBox2.TabIndex = 38
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox1.Location = New System.Drawing.Point(430, 140)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 27)
        Me.PictureBox1.TabIndex = 37
        Me.PictureBox1.TabStop = False
        '
        'PictureBox0
        '
        Me.PictureBox0.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox0.Location = New System.Drawing.Point(396, 180)
        Me.PictureBox0.Name = "PictureBox0"
        Me.PictureBox0.Size = New System.Drawing.Size(50, 29)
        Me.PictureBox0.TabIndex = 80
        Me.PictureBox0.TabStop = False
        '
        'PictureBox45
        '
        Me.PictureBox45.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox45.Location = New System.Drawing.Point(200, 362)
        Me.PictureBox45.Name = "PictureBox45"
        Me.PictureBox45.Size = New System.Drawing.Size(50, 29)
        Me.PictureBox45.TabIndex = 168
        Me.PictureBox45.TabStop = False
        '
        'PictureBox50
        '
        Me.PictureBox50.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox50.Location = New System.Drawing.Point(227, 261)
        Me.PictureBox50.Name = "PictureBox50"
        Me.PictureBox50.Size = New System.Drawing.Size(28, 29)
        Me.PictureBox50.TabIndex = 169
        Me.PictureBox50.TabStop = False
        '
        'PictureBox55
        '
        Me.PictureBox55.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PictureBox55.Location = New System.Drawing.Point(299, 203)
        Me.PictureBox55.Name = "PictureBox55"
        Me.PictureBox55.Size = New System.Drawing.Size(50, 29)
        Me.PictureBox55.TabIndex = 170
        Me.PictureBox55.TabStop = False
        '
        'Main
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(911, 743)
        Me.Controls.Add(Me.lbl_min50)
        Me.Controls.Add(Me.PictureBox50)
        Me.Controls.Add(Me.lbl_min45)
        Me.Controls.Add(Me.PictureBox55)
        Me.Controls.Add(Me.PictureBox45)
        Me.Controls.Add(Me.lbl_min0)
        Me.Controls.Add(Me.lbl_min59)
        Me.Controls.Add(Me.lbl_min57)
        Me.Controls.Add(Me.lbl_min56)
        Me.Controls.Add(Me.lbl_min55)
        Me.Controls.Add(Me.lbl_min58)
        Me.Controls.Add(Me.lbl_min54)
        Me.Controls.Add(Me.lbl_min53)
        Me.Controls.Add(Me.lbl_min52)
        Me.Controls.Add(Me.lbl_min51)
        Me.Controls.Add(Me.lbl_min49)
        Me.Controls.Add(Me.lbl_min47)
        Me.Controls.Add(Me.lbl_min46)
        Me.Controls.Add(Me.lbl_min48)
        Me.Controls.Add(Me.lbl_min44)
        Me.Controls.Add(Me.lbl_min43)
        Me.Controls.Add(Me.lbl_min42)
        Me.Controls.Add(Me.lbl_min41)
        Me.Controls.Add(Me.lbl_min40)
        Me.Controls.Add(Me.lbl_min39)
        Me.Controls.Add(Me.lbl_min38)
        Me.Controls.Add(Me.lbl_min37)
        Me.Controls.Add(Me.lbl_min36)
        Me.Controls.Add(Me.lbl_min35)
        Me.Controls.Add(Me.lbl_min34)
        Me.Controls.Add(Me.lbl_min33)
        Me.Controls.Add(Me.lbl_min32)
        Me.Controls.Add(Me.lbl_min31)
        Me.Controls.Add(Me.lbl_min30)
        Me.Controls.Add(Me.lbl_min29)
        Me.Controls.Add(Me.lbl_min23)
        Me.Controls.Add(Me.lbl_min22)
        Me.Controls.Add(Me.lbl_min20)
        Me.Controls.Add(Me.lbl_min19)
        Me.Controls.Add(Me.lbl_min17)
        Me.Controls.Add(Me.lbl_min24)
        Me.Controls.Add(Me.lbl_min25)
        Me.Controls.Add(Me.lbl_min18)
        Me.Controls.Add(Me.lbl_min16)
        Me.Controls.Add(Me.lbl_min26)
        Me.Controls.Add(Me.lbl_min27)
        Me.Controls.Add(Me.lbl_min28)
        Me.Controls.Add(Me.lbl_min21)
        Me.Controls.Add(Me.lbl_min15)
        Me.Controls.Add(Me.lbl_min14)
        Me.Controls.Add(Me.lbl_min13)
        Me.Controls.Add(Me.lbl_min12)
        Me.Controls.Add(Me.lbl_min11)
        Me.Controls.Add(Me.lbl_min10)
        Me.Controls.Add(Me.lbl_min9)
        Me.Controls.Add(Me.lbl_min8)
        Me.Controls.Add(Me.lbl_min7)
        Me.Controls.Add(Me.lbl_min6)
        Me.Controls.Add(Me.lbl_min5)
        Me.Controls.Add(Me.lbl_min4)
        Me.Controls.Add(Me.lbl_min3)
        Me.Controls.Add(Me.lbl_min2)
        Me.Controls.Add(Me.lbl_min1)
        Me.Controls.Add(Me.PicBxHR0)
        Me.Controls.Add(Me.PicBxHR11)
        Me.Controls.Add(Me.PicBxHR10)
        Me.Controls.Add(Me.PicBxHR9)
        Me.Controls.Add(Me.PicBxHR8)
        Me.Controls.Add(Me.PicBxHR7)
        Me.Controls.Add(Me.PicBxHR6)
        Me.Controls.Add(Me.PicBxHR5)
        Me.Controls.Add(Me.PicBxHR4)
        Me.Controls.Add(Me.PicBxHR3)
        Me.Controls.Add(Me.PicBxHR2)
        Me.Controls.Add(Me.PicBxHR1)
        Me.Controls.Add(Me.PictureBox59)
        Me.Controls.Add(Me.PictureBox58)
        Me.Controls.Add(Me.PictureBox57)
        Me.Controls.Add(Me.PictureBox56)
        Me.Controls.Add(Me.PictureBox54)
        Me.Controls.Add(Me.PictureBox53)
        Me.Controls.Add(Me.PictureBox52)
        Me.Controls.Add(Me.PictureBox51)
        Me.Controls.Add(Me.PictureBox49)
        Me.Controls.Add(Me.PictureBox48)
        Me.Controls.Add(Me.PictureBox47)
        Me.Controls.Add(Me.PictureBox46)
        Me.Controls.Add(Me.PictureBox44)
        Me.Controls.Add(Me.PictureBox43)
        Me.Controls.Add(Me.PictureBox42)
        Me.Controls.Add(Me.PictureBox0)
        Me.Controls.Add(Me.PictureBox40)
        Me.Controls.Add(Me.PictureBox41)
        Me.Controls.Add(Me.PictureBox39)
        Me.Controls.Add(Me.PictureBox38)
        Me.Controls.Add(Me.PictureBox37)
        Me.Controls.Add(Me.PictureBox36)
        Me.Controls.Add(Me.PictureBox35)
        Me.Controls.Add(Me.PictureBox34)
        Me.Controls.Add(Me.PictureBox33)
        Me.Controls.Add(Me.PictureBox32)
        Me.Controls.Add(Me.PictureBox31)
        Me.Controls.Add(Me.PictureBox30)
        Me.Controls.Add(Me.PictureBox29)
        Me.Controls.Add(Me.PictureBox28)
        Me.Controls.Add(Me.PictureBox27)
        Me.Controls.Add(Me.PictureBox26)
        Me.Controls.Add(Me.PictureBox25)
        Me.Controls.Add(Me.PictureBox24)
        Me.Controls.Add(Me.PictureBox23)
        Me.Controls.Add(Me.PictureBox22)
        Me.Controls.Add(Me.PictureBox21)
        Me.Controls.Add(Me.PictureBox20)
        Me.Controls.Add(Me.PictureBox19)
        Me.Controls.Add(Me.PictureBox18)
        Me.Controls.Add(Me.PictureBox17)
        Me.Controls.Add(Me.PictureBox16)
        Me.Controls.Add(Me.PictureBox15)
        Me.Controls.Add(Me.PictureBox14)
        Me.Controls.Add(Me.PictureBox10)
        Me.Controls.Add(Me.PictureBox11)
        Me.Controls.Add(Me.PictureBox12)
        Me.Controls.Add(Me.PictureBox13)
        Me.Controls.Add(Me.PictureBox8)
        Me.Controls.Add(Me.PictureBox7)
        Me.Controls.Add(Me.PictureBox6)
        Me.Controls.Add(Me.PictureBox5)
        Me.Controls.Add(Me.PictureBox4)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.PictureBox9)
        Me.Controls.Add(Me.PictureBox_clock)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Main"
        Me.Text = "MY CLOCK"
        CType(Me.PictureBox27, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox28, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox29, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox30, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox31, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox32, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox33, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox34, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox35, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox36, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox37, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox38, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox41, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox40, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox42, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox43, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox44, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox46, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox47, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox48, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox49, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox51, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox52, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox53, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox54, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox56, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox57, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox58, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox59, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBxHR0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox39, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox_clock, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox26, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox24, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox23, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox22, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox25, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox21, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox20, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox19, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox18, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox16, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox14, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox45, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox50, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox55, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region






  

   
   
End Class