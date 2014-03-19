Public Class UserControl1

    Dim playername As String

    Dim FGmade As Integer = 0
    Dim FGmissed As Integer = 0
    Dim Tmade As Integer = 0
    Dim Tmissed As Integer = 0
    Dim FTmade As Integer = 0
    Dim FTmissed As Integer = 0

    Dim REB As Integer = 0
    Dim AST As Integer = 0
    Dim STL As Integer = 0
    Dim BLK As Integer = 0
    Dim TOv As Integer = 0
    Dim PF As Integer = 0




    Public Property player()
        Get
            Return Me.TextBox13.Text
        End Get
        Set(ByVal value)
            Me.TextBox13.Text = value
        End Set
    End Property
    Public Property points()
        Get
            Return Me.Label6.Text
        End Get
        Set(ByVal value)
            Me.Label6.Text = value
        End Set
    End Property
    Public Property FGA()
        Get
            Return Me.FGmissed
        End Get
        Set(ByVal value)
            Me.FGmissed = value
        End Set
    End Property
    Public Property FGM()
        Get
            Return Me.FGmade
        End Get
        Set(ByVal value)
            Me.FGmade = value
        End Set
    End Property
    Public Property TA()
        Get
            Return Me.Tmissed
        End Get
        Set(ByVal value)
            Me.Tmissed = value
        End Set
    End Property
    Public Property TM()
        Get
            Return Me.Tmade
        End Get
        Set(ByVal value)
            Me.Tmade = value
        End Set
    End Property
    Public Property FTA()
        Get
            Return Me.FTmissed
        End Get
        Set(ByVal value)
            Me.FTmissed = value
        End Set
    End Property
    Public Property FTM()
        Get
            Return Me.FTmade
        End Get
        Set(ByVal value)
            Me.FTmade = value
        End Set
    End Property
    Public Property Rebounds()
        Get
            Return Me.REB
        End Get
        Set(ByVal value)
            Me.REB = value
        End Set
    End Property
    Public Property Assists()
        Get
            Return Me.AST
        End Get
        Set(ByVal value)
            Me.AST = value
        End Set
    End Property
    Public Property Steals()
        Get
            Return Me.STL
        End Get
        Set(ByVal value)
            Me.STL = value
        End Set
    End Property
    Public Property Blocks()
        Get
            Return Me.BLK
        End Get
        Set(ByVal value)
            Me.BLK = value
        End Set
    End Property
    Public Property Turnovers()
        Get
            Return Me.TOv
        End Get
        Set(ByVal value)
            Me.TOv = value
        End Set
    End Property
    Public Property Fouls()
        Get
            Return Me.PF
        End Get
        Set(ByVal value)
            Me.PF = value
        End Set
    End Property


    Private Sub UserControl1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FGmade = FGmade + 1
        FGmissed = FGmissed + 1
        TextBox1.Text = FGmade
        TextBox2.Text = FGmissed
        ReturnPoints()
        ReturnPct()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FGmade = FGmade - 1
        FGmissed = FGmissed - 1
        TextBox1.Text = FGmade
        TextBox2.Text = FGmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FGmissed = FGmissed + 1
        TextBox2.Text = FGmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        FGmissed = FGmissed - 1
        TextBox2.Text = FGmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Tmade = Tmade + 1
        Tmissed = Tmissed + 1
        TextBox3.Text = Tmade
        TextBox4.Text = Tmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Tmade = Tmade - 1
        TextBox3.Text = Tmade
        Tmissed = Tmissed - 1
        TextBox4.Text = Tmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Tmissed = Tmissed + 1
        TextBox4.Text = Tmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Tmissed = Tmissed - 1
        TextBox4.Text = Tmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        FTmade = FTmade + 1
        TextBox5.Text = FTmade
        FTmissed = FTmissed + 1
        TextBox6.Text = FTmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        FTmade = FTmade - 1
        TextBox5.Text = FTmade
        FTmissed = FTmissed - 1
        TextBox6.Text = FTmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        FTmissed = FTmissed + 1
        TextBox6.Text = FTmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        FTmissed = FTmissed - 1
        TextBox6.Text = FTmissed
        ReturnPoints()
        ReturnPct()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        REB = REB + 1
        TextBox7.Text = REB
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        REB = REB - 1
        TextBox7.Text = REB
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        AST = AST + 1
        TextBox8.Text = AST
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        AST = AST - 1
        TextBox8.Text = AST
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        STL = STL + 1
        TextBox9.Text = STL
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        STL = STL - 1
        TextBox9.Text = STL
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        BLK = BLK + 1
        TextBox10.Text = BLK
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        BLK = BLK - 1
        TextBox10.Text = BLK
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        TOv = TOv + 1
        TextBox11.Text = TOv
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        TOv = TOv - 1
        TextBox11.Text = TOv
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        PF = PF + 1
        TextBox12.Text = PF
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        PF = PF - 1
        TextBox12.Text = PF
    End Sub

    Public Sub ReturnPoints()
        Dim val As Integer = (FGmade * 2) + (Tmade * 3) + (FTmade)
        Label6.Text = val
    End Sub
    Public Sub ReturnPct()
        Try

            Dim val As Decimal = ((FGmade + Tmade) / (FGmissed + Tmissed))
            Label7.Text = FormatNumber((val * 100), 2) & "%"

        Catch ex As Exception

        End Try
    End Sub

End Class
