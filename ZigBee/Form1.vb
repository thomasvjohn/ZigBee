Imports System.Threading.Thread
Imports System.AppDomain
Public Class Form1

#Region " Global Variables "
    Dim ZigBee1Data As String = Nothing
    Dim ZigBee2Data As String = Nothing
    Dim AryList As New ArrayList()
#End Region

#Region " Form Load "
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            RadioButton3.Checked = False
        ElseIf RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            RadioButton3.Checked = False
        ElseIf RadioButton3.Checked = True Then
            RadioButton1.Checked = False
            RadioButton2.Checked = False
        End If

        T1.Text = "ZigBee 1: Disconnected"
        T2.Text = "ZigBee 2: Disconnected"

        ToolTipSub()
    End Sub
#End Region

#Region " Connect / Disconnect "
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ZigBee1.IsOpen = False Then
            ZigBee1.Open()
            T1.Text = "ZigBee 1: Connected"
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Timer1.Enabled = False
        If ZigBee1.IsOpen = True Then
            ZigBee1.Close()
            T1.Text = "ZigBee 1: Disconnected"
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If ZigBee2.IsOpen = False Then
            ZigBee2.Open()
            T2.Text = "ZigBee 2: Connected"
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Timer1.Enabled = False
        If ZigBee2.IsOpen = True Then
            ZigBee2.Close()
            T2.Text = "ZigBee 2: Disconnected"
        End If
    End Sub
#End Region

#Region " Send/Receive Command "

    Private Sub ZigBee1_DataReceived(ByVal sender As System.Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles ZigBee1.DataReceived
        ZigBee1Data = ZigBee1Data & ZigBee1.ReadExisting
    End Sub

    Private Sub ZigBee2_DataReceived(ByVal sender As System.Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles ZigBee2.DataReceived
        ZigBee2Data = ZigBee2Data & ZigBee2.ReadExisting
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If CheckBox1.Checked = False Then
            If ZigBee1Data <> Nothing Then
                TextBox2.Text = TextBox2.Text & vbNewLine & "ZigBee1 Recv: " & ZigBee1Data
                ZigBee1Data = Nothing
            End If
            If ZigBee2Data <> Nothing Then
                TextBox2.Text = TextBox2.Text & vbNewLine & "ZigBee2 Recv: " & ZigBee2Data
                ZigBee2Data = Nothing
            End If
            Application.DoEvents()
        End If
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox1.Text = TextBox1.Text.ToUpper
        If TextBox1.Text = "CLS" Then
            TextBox2.Clear()
        Else
            If RadioButton1.Checked = True Then
                ZigBee1.WriteLine(TextBox1.Text.Trim & vbCrLf)
                TextBox2.Text = TextBox2.Text & "ZigBee1 Send: " & TextBox1.Text
                TextBox1.Text = Nothing
            ElseIf RadioButton2.Checked = True Then
                ZigBee2.WriteLine(TextBox2.Text.Trim & vbCrLf)
                TextBox2.Text = TextBox2.Text & "ZigBee2 Send: " & TextBox1.Text
                TextBox1.Text = Nothing
            ElseIf RadioButton3.Checked = True Then
                ZigBee1.WriteLine(TextBox1.Text.Trim & vbCrLf)
                TextBox2.Text = TextBox2.Text & "ZigBee1 Send: " & TextBox1.Text
                ZigBee2.WriteLine(TextBox1.Text.Trim & vbCrLf)
                TextBox2.Text = TextBox2.Text & vbNewLine & "ZigBee2 Send: " & TextBox1.Text
                TextBox1.Text = Nothing
            End If
            Timer1.Enabled = True
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            RadioButton3.Checked = False
        ElseIf RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            RadioButton3.Checked = False
        ElseIf RadioButton3.Checked = True Then
            RadioButton1.Checked = False
            RadioButton2.Checked = False
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            RadioButton3.Checked = False
        ElseIf RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            RadioButton3.Checked = False
        ElseIf RadioButton3.Checked = True Then
            RadioButton1.Checked = False
            RadioButton2.Checked = False
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            RadioButton3.Checked = False
        ElseIf RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            RadioButton3.Checked = False
        ElseIf RadioButton3.Checked = True Then
            RadioButton1.Checked = False
            RadioButton2.Checked = False
        End If
    End Sub
#End Region

#Region " S-Register "

#Region " Tool Tip "
    Sub ToolTipSub()
        'S0
        ToolTip1.SetToolTip(Me.TextBox3, "Range 0001 - FFFF")
        ToolTip1.SetToolTip(Me.TextBox5, "Range 0000 - FFFF")
        ToolTip1.SetToolTip(Me.TextBox6, "Range 0000000000000000 – FFFFFFFFFFFFFFFF")
        ToolTip1.SetToolTip(Me.TextBox7, "Range 0000000000000000 – FFFFFFFFFFFFFFFF")
        ToolTip1.SetToolTip(Me.TextBox8, "Range 0000 - FFFF")
        ToolTip1.SetToolTip(Me.TextBox9, "Range 0000000000000000 – FFFFFFFFFFFFFFFF")
        ToolTip1.SetToolTip(Me.TextBox10, "Range 0000 - FFFF")
        ToolTip1.SetToolTip(Me.TextBox13, "Range 0000 - FFFF")
        ToolTip1.SetToolTip(Me.TextBox18, "Range 0000 - FFFF")
        'S1
        ToolTip1.SetToolTip(Me.TextBox19, "Range 0000 - FFFF")
        ToolTip1.SetToolTip(Me.TextBox23, "Range 0000 - FFFF")
        ToolTip1.SetToolTip(Me.TextBox29, "Range 00000000 to FFFFFFFF")
        ToolTip1.SetToolTip(Me.TextBox30, "Range 0000 - FFFF")
        ToolTip1.SetToolTip(Me.TextBox32, "Range 0000 - FFFF")
        ToolTip1.SetToolTip(Me.TextBox34, "Range 0000 – 2EE0 (0 – 12000)")
        'S2
        ToolTip1.SetToolTip(Me.TextBox35, "Range 0000 – 2EE0 (0 – 12000)")
        ToolTip1.SetToolTip(Me.TextBox36, "Range 0000 – 2EE0 (0 – 12000)")
        ToolTip1.SetToolTip(Me.TextBox37, "Range 0000 – 2EE0 (0 – 12000)")
        'S3
        ToolTip1.SetToolTip(Me.TextBox60, "Range 0000 – 0003")
        ToolTip1.SetToolTip(Me.TextBox61, "Range 0000 – 0003")
        'S4
    End Sub
#End Region

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ClearTestbox()
        If CheckBox1.Checked = True Then
            If ZigBee1.IsOpen = True AndAlso ComboBox1.SelectedIndex = 0 Then
                ZigBee1.WriteLine("AT+TOKDUMP" & vbCrLf)
                Timer2.Enabled = True
            ElseIf ZigBee2.IsOpen = True AndAlso ComboBox1.SelectedIndex = 1 Then
                ZigBee1.WriteLine("AT+TOKDUMP" & vbCrLf)
            End If
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False

        Dim DT As String() = ZigBee1Data.Split(vbNewLine)
        For i = 0 To DT.Length - 1
            If DT(i).Contains(":") = True Then
                Dim DT1 As String() = DT(i).Split(":")
                AryList.Add(DT1(1))
            Else
                AryList.Add(DT(i))
            End If
        Next
        If AryList.Count > 80 Then LoadSRegister()
    End Sub

    Sub ClearTestbox()
        TextBox3.Text = Nothing
        TextBox4.Text = Nothing
        TextBox5.Text = Nothing
        TextBox6.Text = Nothing
        TextBox7.Text = Nothing
        TextBox8.Text = Nothing
        TextBox9.Text = Nothing
        TextBox10.Text = Nothing
        TextBox11.Text = Nothing
        TextBox12.Text = Nothing
        TextBox13.Text = Nothing
        TextBox14.Text = Nothing
        TextBox15.Text = Nothing
        TextBox16.Text = Nothing
        TextBox17.Text = Nothing
        TextBox18.Text = Nothing
        TextBox19.Text = Nothing
        TextBox20.Text = Nothing
        TextBox21.Text = Nothing
        TextBox22.Text = Nothing
        TextBox23.Text = Nothing
        TextBox24.Text = Nothing
        TextBox25.Text = Nothing
        TextBox26.Text = Nothing
        TextBox27.Text = Nothing
        TextBox28.Text = Nothing
        TextBox29.Text = Nothing
        TextBox30.Text = Nothing
        TextBox31.Text = Nothing
        TextBox32.Text = Nothing
        TextBox33.Text = Nothing
        TextBox34.Text = Nothing
        TextBox35.Text = Nothing
        TextBox36.Text = Nothing
        TextBox37.Text = Nothing
        TextBox38.Text = Nothing
        TextBox39.Text = Nothing
        TextBox40.Text = Nothing
        TextBox41.Text = Nothing
        TextBox42.Text = Nothing
        TextBox43.Text = Nothing
        TextBox44.Text = Nothing
        TextBox45.Text = Nothing
        TextBox46.Text = Nothing
        TextBox47.Text = Nothing
        TextBox48.Text = Nothing
        TextBox49.Text = Nothing
        TextBox50.Text = Nothing
        TextBox51.Text = Nothing
        TextBox52.Text = Nothing
        TextBox53.Text = Nothing
        TextBox54.Text = Nothing
        TextBox55.Text = Nothing
        TextBox56.Text = Nothing
        TextBox57.Text = Nothing
        TextBox58.Text = Nothing
        TextBox59.Text = Nothing
        TextBox60.Text = Nothing
        TextBox61.Text = Nothing
        TextBox62.Text = Nothing
        TextBox63.Text = Nothing
        TextBox64.Text = Nothing
        TextBox65.Text = Nothing
        TextBox66.Text = Nothing
        TextBox67.Text = Nothing
        TextBox68.Text = Nothing
        TextBox69.Text = Nothing
        TextBox70.Text = Nothing
        TextBox71.Text = Nothing
        TextBox72.Text = Nothing
        TextBox73.Text = Nothing
        TextBox74.Text = Nothing
        TextBox75.Text = Nothing
        TextBox76.Text = Nothing
        TextBox77.Text = Nothing
        TextBox78.Text = Nothing
        TextBox79.Text = Nothing
        TextBox80.Text = Nothing
        TextBox81.Text = Nothing
        TextBox82.Text = Nothing
    End Sub

    Sub LoadSRegister()
        TextBox3.Text = AryList.Item(1)
        TextBox4.Text = AryList.Item(2)
        TextBox5.Text = AryList.Item(3)
        TextBox6.Text = AryList.Item(4)
        TextBox7.Text = AryList.Item(5)
        TextBox8.Text = AryList.Item(6)
        TextBox9.Text = AryList.Item(7)
        TextBox10.Text = AryList.Item(8)
        TextBox11.Text = AryList.Item(9)
        TextBox12.Text = AryList.Item(10)
        TextBox13.Text = AryList.Item(11)
        TextBox14.Text = AryList.Item(12)
        TextBox15.Text = AryList.Item(13)
        TextBox16.Text = AryList.Item(14)
        TextBox17.Text = AryList.Item(15)
        TextBox18.Text = AryList.Item(16)
        TextBox19.Text = AryList.Item(17)
        TextBox20.Text = AryList.Item(18)
        TextBox21.Text = AryList.Item(19)
        TextBox22.Text = AryList.Item(20)
        TextBox23.Text = AryList.Item(21)
        TextBox24.Text = AryList.Item(22)
        TextBox25.Text = AryList.Item(23)
        TextBox26.Text = AryList.Item(24)
        TextBox27.Text = AryList.Item(25)
        TextBox28.Text = AryList.Item(26)
        TextBox29.Text = AryList.Item(27)
        TextBox30.Text = AryList.Item(28)
        TextBox31.Text = AryList.Item(29)
        TextBox32.Text = AryList.Item(30)
        TextBox33.Text = AryList.Item(31)
        TextBox34.Text = AryList.Item(32)
        TextBox35.Text = AryList.Item(33)
        TextBox36.Text = AryList.Item(34)
        TextBox37.Text = AryList.Item(35)
        TextBox38.Text = AryList.Item(36)
        TextBox39.Text = AryList.Item(37)
        TextBox40.Text = AryList.Item(38)
        TextBox41.Text = AryList.Item(39)
        TextBox42.Text = AryList.Item(40)
        TextBox43.Text = AryList.Item(41)
        TextBox44.Text = AryList.Item(42)
        TextBox45.Text = AryList.Item(43)
        TextBox46.Text = AryList.Item(44)
        TextBox47.Text = AryList.Item(45)
        TextBox48.Text = AryList.Item(46)
        TextBox49.Text = AryList.Item(47)
        TextBox50.Text = AryList.Item(48)
        TextBox51.Text = AryList.Item(49)
        TextBox52.Text = AryList.Item(50)
        TextBox53.Text = AryList.Item(51)
        TextBox54.Text = AryList.Item(52)
        TextBox55.Text = AryList.Item(53)
        TextBox56.Text = AryList.Item(54)
        TextBox57.Text = AryList.Item(55)
        TextBox58.Text = AryList.Item(56)
        TextBox59.Text = AryList.Item(57)
        TextBox60.Text = AryList.Item(58)
        TextBox61.Text = AryList.Item(59)
        TextBox62.Text = AryList.Item(60)
        TextBox63.Text = AryList.Item(61)
        TextBox64.Text = AryList.Item(62)
        TextBox65.Text = AryList.Item(63)
        TextBox66.Text = AryList.Item(64)
        TextBox67.Text = AryList.Item(65)
        TextBox68.Text = AryList.Item(66)
        TextBox69.Text = AryList.Item(67)
        TextBox70.Text = AryList.Item(68)
        TextBox71.Text = AryList.Item(69)
        TextBox72.Text = AryList.Item(70)
        TextBox73.Text = AryList.Item(71)
        TextBox74.Text = AryList.Item(72)
        TextBox75.Text = AryList.Item(73)
        TextBox76.Text = AryList.Item(74)
        TextBox77.Text = AryList.Item(75)
        TextBox78.Text = AryList.Item(76)
        TextBox79.Text = AryList.Item(77)
        TextBox80.Text = AryList.Item(78)
        TextBox81.Text = AryList.Item(79)
        TextBox82.Text = AryList.Item(80)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        'S00  Channel Mask
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S01  Transmit Power Level
        If TextBox4.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S02  Preferred PAN ID
        If TextBox5.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S03  Preferred Extended PAN ID
        If TextBox6.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S04  Local EUI
        If TextBox7.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S05  Local NodeID
        If TextBox8.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S06  (Parent)'s EUI
        If TextBox9.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S07  (Parent)'s NodeID
        If TextBox10.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S08  Network Key
        If TextBox11.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S09  Link Key
        If TextBox12.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S0A  Main Function
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S0B  User Readable Name
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S0C  (Password)
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S0D  Device Information
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S0E  Prompt Enable 1 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S0F  Prompt Enable 2 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S10  Extended Function 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S11  Device Specific
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S12  UART Setup  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S13  Pull-up enable  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S14  Pull-down enable
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S15  I/O Configuration
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S16  Data Direction of I/O Port  (volatile)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S17  Initial Value of S16  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S18  Output Buffer of I/O Port  (volatile)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S19  Initial Value of S18 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S1A  Input Buffer of I/O Port  (volatile) 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S1B  Special Function pin 1 (volatile)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S1C  Initial Value of S1B  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S1D  Special Function Pin 2 (volatile)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S1E  Initial Value of S1D 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S1F  A/D1 (ETRX3: ADC0)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S20  A/D2 (ETRX3: ADC1)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S21  A/D3 (ETRX3: ADC2)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S22  A/D4 (ETRX3: ADC3) 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S23  Immediate functionality at IRQ0   
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S24  Immediate functionality at IRQ1   
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S25  Immediate functionality at IRQ2   
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S26  Immediate functionality at IRQ3 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S27  Functionality 1 at Boot-up  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S28  Functionality at Network Join  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S29  Timer/Counter 0 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S2A  Functionality for Timer/Counter 0  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S2B  Timer/Counter 1  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S2C  Functionality for Timer/Counter 1  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S2D  Timer/Counter 2  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S2E  Functionality for Timer/Counter 2 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S2F  Timer/Counter 3  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S30  Functionality for Timer/Counter 3 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S31  Timer/Counter 4  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S32  Functionality for Timer/Counter 4  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S33  Timer/Counter 5  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S34  Functionality for Timer/Counter 5 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S35  Timer/Counter 6  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S36  Functionality for Timer/Counter 6 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S37  Timer/Counter 7  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S38  Functionality for Timer/Counter 7  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S39  Power mode (volatile)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S3A  Initial Power Mode  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S3B  Start-up Functionality Plaintext A  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S3C  Start-up Functionality Plaintext B  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S3D  Supply Voltage  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S3E  Multicast Table Entry 00 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S3F  Multicast Table Entry 01  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S40  Source and Destination Endpoints for xCASTs (volatile) 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S41  Initial Value of S40  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S42  Cluster ID for xCASTs (volatile)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S43  Initial Value of S42  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S44  Profile ID for xCASTs (volatile)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S45  Initial Value of S44 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S46  Start-up Functionality 32 bit number (volatile)  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S47  Power Descriptor  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S48  Endpoint 2 Profile ID  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S49  Endpoint 2 Device ID  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S4A  Endpoint 2 Device Version  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S4B  Endpoint 2 Input Cluster List  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S4C  Endpoint 2 Output Cluster List  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S4D  Mobile End Device Poll Timeout 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S4E  End Device Poll Timeout 
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
        'S4F  MAC Timeout  
        If TextBox3.Text <> AryList(1) Then ZigBee1.WriteLine("ATS00=" & TextBox3.Text & vbCrLf)
        Sleep(50)
        Application.DoEvents()
    End Sub

#End Region



End Class
