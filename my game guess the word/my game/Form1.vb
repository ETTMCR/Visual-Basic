Public Class Form1
    Dim num_up As Integer
    Dim num_down As Integer
    Dim num_right As Integer
    Dim num_left As Integer
    Dim num_dell As Integer
    Dim num_back As Integer

    Dim II As Integer
    Dim num_now, last_num_now As Integer
    Dim new_here_y, new_here_x As Integer
    Dim int_lbl_gen As Integer


    ' Dim GFX As Graphics


    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        HScrollBar1.Minimum = 1

        HScrollBar1.Maximum = 1000
        'HScrollBar1.SmallChange = 1000


        Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

        Me.Width = screenWidth
        Me.Height = screenHeight
        Me.Left = 0
        Me.Top = 0
        'GFX = Me.CreateGraphics

        'num_up = 1
        'num_down = 2
        'num_right = 3
        'num_left = 4
        'num_dell = 5
        'num_back = 6

        new_here_x = (Me.Width \ 2)
        new_here_y = (Me.Height \ 2)
        int_lbl_gen = 0

        'For i = 1 To 100
        '    Me.SetPixel((Me.Width \ 2) + II, (Me.Height \ 2) + II, Color.Red)
        '    II = II + 1
        'Next



        Randomize()



        'randoming num_up
        num_up = Int(Rnd() * 6 + 1)
        Label1.Text = Label1.Text & (Str(num_up))

        'randoming num_down
        num_down = Int(Rnd() * 6 + 1)
        Do While num_down = num_up
            num_down = Int(Rnd() * 6 + 1)
        Loop
        Label2.Text = Label2.Text & (Str(num_down))

        'randoming num_left
        num_left = Int(Rnd() * 6 + 1)
        Do While num_left = num_up Or num_left = num_down
            num_left = Int(Rnd() * 6 + 1)
        Loop
        Label3.Text = Label3.Text & (Str(num_left))

        'randoming num_right
        num_right = Int(Rnd() * 6 + 1)
        Do While num_right = num_up Or num_right = num_down Or num_right = num_left
            num_right = Int(Rnd() * 6 + 1)
        Loop
        Label4.Text = Label4.Text & (Str(num_right))

        'randoming num_dell
        num_dell = Int(Rnd() * 6 + 1)
        Do While num_dell = num_up Or num_dell = num_down Or num_dell = num_left Or num_dell = num_right
            num_dell = Int(Rnd() * 6 + 1)
        Loop
        Label5.Text = Label5.Text & (Str(num_dell))

        'randoming num_back
        num_back = Int(Rnd() * 6 + 1)
        Do While num_back = num_up Or num_back = num_down Or num_back = num_left Or num_back = num_right Or num_back = num_dell
            num_back = Int(Rnd() * 6 + 1)
        Loop
        Label6.Text = Label6.Text & (Str(num_back))


    End Sub

    Private Sub HScrollBar1_Scroll(sender As System.Object, e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar1.Scroll
        'HScrollBar1.Value = HScrollBar1.Value * 10
        ' lbl_timer1.Text = 1000 / HScrollBar1.Value
        'Label8.Visible = False
        lbl_timer1.Text = Format((1000 / HScrollBar1.Value), "0.00")
        'lbl_timer1.Text = HScrollBar1.Value

        '        Format(Per1, "0.00")
        'Documentation can be found here http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.strings.format.aspx
        'If you want to round the numbers you can use
        'Math.Round(Per1, 2)

        Timer1.Interval = HScrollBar1.Value '* 20 if max = 50
    End Sub


    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        'For i = 1 To 1000
        '    Me.SetPixel((Me.Width \ 2) + II, (Me.Height \ 2) + II, Color.Red)
        '    II = II + 1
        'Next

        num_now = Int(Rnd() * 6 + 1)
        lbl_gen.Text = int_lbl_gen

        Select Case num_now
            Case num_up
                Me.CreateGraphics.DrawLine(Pens.Black, new_here_x, new_here_y, new_here_x, new_here_y - 50)
                new_here_y = new_here_y - 50
                Label1.BackColor = Color.Yellow
                Label3.BackColor = SystemColors.Control
                Label2.BackColor = SystemColors.Control
                Label4.BackColor = SystemColors.Control
                Label5.BackColor = SystemColors.Control
                Label6.BackColor = SystemColors.Control

            Case num_down
                Me.CreateGraphics.DrawLine(Pens.Black, new_here_x, new_here_y, new_here_x, 50 + new_here_y)
                new_here_y = new_here_y + 50
                Label2.BackColor = Color.Yellow
                Label1.BackColor = SystemColors.Control
                Label3.BackColor = SystemColors.Control
                Label4.BackColor = SystemColors.Control
                Label5.BackColor = SystemColors.Control
                Label6.BackColor = SystemColors.Control

            Case num_left
                Me.CreateGraphics.DrawLine(Pens.Black, new_here_x, new_here_y, new_here_x - 50, new_here_y)
                new_here_x = new_here_x - 50
                Label3.BackColor = Color.Yellow
                Label1.BackColor = SystemColors.Control
                Label2.BackColor = SystemColors.Control
                Label4.BackColor = SystemColors.Control
                Label5.BackColor = SystemColors.Control
                Label6.BackColor = SystemColors.Control

            Case num_right
                Me.CreateGraphics.DrawLine(Pens.Black, new_here_x, new_here_y, 50 + new_here_x, new_here_y)
                new_here_x = new_here_x + 50
                Label4.BackColor = Color.Yellow
                Label1.BackColor = SystemColors.Control
                Label2.BackColor = SystemColors.Control
                Label3.BackColor = SystemColors.Control
                Label5.BackColor = SystemColors.Control
                Label6.BackColor = SystemColors.Control

            Case num_dell ' not leaving the last position, thats why new_here_x / y will not change
                Label1.BackColor = SystemColors.Control
                Label2.BackColor = SystemColors.Control
                Label3.BackColor = SystemColors.Control
                Label4.BackColor = SystemColors.Control
                Label6.BackColor = SystemColors.Control
                Label5.BackColor = Color.Yellow
                ' not more than one step

                Select Case last_num_now
                    Case num_up
                        Me.CreateGraphics.DrawLine(Pens.White, new_here_x, new_here_y, new_here_x, new_here_y + 50)
                        'new_here_y = new_here_y + 50

                    Case num_down
                        Me.CreateGraphics.DrawLine(Pens.White, new_here_x, new_here_y, new_here_x, new_here_y - 50)
                        'new_here_y = new_here_y - 50

                    Case num_left
                        Me.CreateGraphics.DrawLine(Pens.White, new_here_x, new_here_y, new_here_x + 50, new_here_y)
                        'new_here_x = new_here_x + 50

                    Case num_right
                        Me.CreateGraphics.DrawLine(Pens.White, new_here_x, new_here_y, new_here_x - 50, new_here_y)
                        'new_here_x = new_here_x - 50

                        'Case num_dell ' for dell after dell i need 2d array ! or var like last_dell_x / y 
                        '    Me.CreateGraphics.DrawLine(Pens.White, new_here_x, new_here_y, new_here_x - 50, new_here_y)

                End Select 'of dell ONE step !

            Case num_back
                Label6.BackColor = Color.Yellow
                Label2.BackColor = SystemColors.Control
                Label3.BackColor = SystemColors.Control
                Label4.BackColor = SystemColors.Control
                Label5.BackColor = SystemColors.Control
                Label1.BackColor = SystemColors.Control

                Select Case last_num_now
                    Case num_up
                        Me.CreateGraphics.DrawLine(Pens.Black, new_here_x, new_here_y, new_here_x, new_here_y + 50)
                        new_here_y = new_here_y + 50

                    Case num_down
                        Me.CreateGraphics.DrawLine(Pens.Black, new_here_x, new_here_y, new_here_x, new_here_y - 50)
                        new_here_y = new_here_y - 50

                    Case num_left
                        Me.CreateGraphics.DrawLine(Pens.Black, new_here_x, new_here_y, new_here_x + 50, new_here_y)
                        new_here_x = new_here_x + 50

                    Case num_right
                        Me.CreateGraphics.DrawLine(Pens.Black, new_here_x, new_here_y, new_here_x - 50, new_here_y)
                        new_here_x = new_here_x - 50

                        'Case num_dell
                        '    Me.CreateGraphics.DrawLine(Pens.White, new_here_x, new_here_y, new_here_x - 50, new_here_y)

                End Select 'of back ONE MOVE !

                ' what was the last num_now will be last_num_now than respect to it. to do the back move, BUT ....
                ' this is only good foronly one back 
                ' else we need array
        End Select


        last_num_now = num_now
        int_lbl_gen = int_lbl_gen + 1
        'Me.CreateGraphics.DrawLine(Pens.Black, (Me.Width \ 2), (Me.Height \ 2), (Me.Width \ 2) + II, (Me.Height \ 2))
        'II = II + 1



    End Sub

    Private Sub btn_start_Click(sender As System.Object, e As System.EventArgs) Handles btn_start.Click
        Timer1.Enabled = True
        'Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) ' / 2

        'Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) ' / 2


    End Sub

    Private Sub btn_pause_Click(sender As System.Object, e As System.EventArgs) Handles btn_pause.Click
        Timer1.Enabled = False
    End Sub

    Private Sub btn_clear_Click(sender As System.Object, e As System.EventArgs) Handles btn_clear.Click
        ' clear the form1
        Me.Hide()
        Me.Show()

        new_here_x = (Me.Width \ 2)
        new_here_y = (Me.Height \ 2)
        int_lbl_gen = 0
        lbl_gen.Text = int_lbl_gen


        'Randomize()

        ''randoming num_up
        'num_up = Int(Rnd() * 6 + 1)
        'Label1.Text = Label1.Text & (Str(num_up))

        ''randoming num_down
        'num_down = Int(Rnd() * 6 + 1)
        'Do While num_down = num_up
        '    num_down = Int(Rnd() * 6 + 1)
        'Loop
        'Label2.Text = Label2.Text & (Str(num_down))

        ''randoming num_left
        'num_left = Int(Rnd() * 6 + 1)
        'Do While num_left = num_up Or num_left = num_down
        '    num_left = Int(Rnd() * 6 + 1)
        'Loop
        'Label3.Text = Label3.Text & (Str(num_left))

        ''randoming num_right
        'num_right = Int(Rnd() * 6 + 1)
        'Do While num_right = num_up Or num_right = num_down Or num_right = num_left
        '    num_right = Int(Rnd() * 6 + 1)
        'Loop
        'Label4.Text = Label4.Text & (Str(num_right))

        ''randoming num_dell
        'num_dell = Int(Rnd() * 6 + 1)
        'Do While num_dell = num_up Or num_dell = num_down Or num_dell = num_left Or num_dell = num_right
        '    num_dell = Int(Rnd() * 6 + 1)
        'Loop
        'Label5.Text = Label5.Text & (Str(num_dell))

        ''randoming num_back
        'num_back = Int(Rnd() * 6 + 1)
        'Do While num_back = num_up Or num_back = num_down Or num_back = num_left Or num_back = num_right Or num_back = num_dell
        '    num_back = Int(Rnd() * 6 + 1)
        'Loop
        'Label6.Text = Label6.Text & (Str(num_back))



    End Sub


 
  
End Class
