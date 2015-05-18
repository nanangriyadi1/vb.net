Imports System.Data.OleDb
Imports System.Data


Public Class Ujian
    Dim plhj As String = vbNullString
    Dim jam, menit, detik, n1, ulang, banyakSoal As Integer
    Dim cnn As New OleDbConnection
    Dim sQl As String
    Dim nilai(101) As String
    Private Sub Ujian_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        banyakSoal = 0
        Label11.Text = Format(Today, "MM/dd/yyyy")
        BTNSelesai.Enabled = False
        Call BersihkanJawaban()
        Call Koneksi()
        Call soal()

        DGV.ReadOnly = True
        DGV.Rows.Clear()
        jam = 1
        menit = 0
        detik = 0
        Label29.Text = Format(jam, "00") & ":" & Format(menit, "00") & ":" & Format(detik, "00")
        'ListBox1.SelectedIndex = 0
        ulang = 0
    End Sub

    Sub timerreset()
        jam = 1
        menit = 0
        detik = 0
        Label29.Text = Format(jam, "00") & ":" & Format(menit, "00") & ":" & Format(detik, "00")
        'Timer4.Enabled = True
    End Sub

    Sub soal()
        cnn = New OleDbConnection(kn)
        If cnn.State <> ConnectionState.Closed Then cnn.Close()
        cnn.Open()

    sQl = "SELECT  IdSimulasi, Materi  FROM TBLSimulasi"
    CMD = New OleDbCommand(sQl, cnn)
    DR = CMD.ExecuteReader

    While DR.Read
      'ComboBox1.Text = (DR.Item("IdSimulasi") & Space(2) & DR.Item("Materi"))
      ComboBox1.Items.Add(DR.Item("IdSimulasi") & Space(3) & DR.Item("Materi"))

      'TextBox2.Text = (DR.Item("IdSimulasi") & Space(3) & DR.Item("Materi"))
      'ComboBox1.Items.Add(DR.Item("IdSimulasi") & Space(5) & DR.Item("Materi"))
    End While

  End Sub

  Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
    ' ListBox1.SelectedIndex = 0
    Call BersihkanJawaban()
    'tampilkan pertanyaan soal ujian dalam textbox1 sesuai dengan nomor soal yang dipilih
    CMD = New OleDbCommand("select * from tblsoal where IdSimulasi='" & Microsoft.VisualBasic.Left(ComboBox1.Text, 3) & "' and VAL(nomor)='" & Val(ListBox1.Text) & "'", Conn)
    DR = CMD.ExecuteReader
    DR.Read()
    If DR.HasRows Then
      TextBox1.Text = DR.Item("pertanyaan")
      RadioButton1.Text = DR.Item("A")
      RadioButton2.Text = DR.Item("B")
      RadioButton3.Text = DR.Item("C")
      RadioButton4.Text = DR.Item("D")
      RadioButton5.Text = DR.Item("E")
    End If
    Timer4.Enabled = True

    Dim kk As Integer = ListBox1.SelectedIndex
    Dim mm As String = nilai(kk)
    If (mm = "") Then

    Else
      Call tampilKeRadioButton(kk)
    End If
  End Sub

    Sub tampilKeRadioButton(ByVal pNilaiIndexList As Integer)
        Dim nArray As String
        nArray = nilai(pNilaiIndexList)
        If (nArray = vbNullString) Then
            MsgBox("ada yang null")
        Else
            If (nArray = "A") Then
                RadioButton1.Checked = True
            ElseIf (nArray = "B") Then
                RadioButton2.Checked = True
            ElseIf (nArray = "C") Then
                RadioButton3.Checked = True
            ElseIf (nArray = "D") Then
                RadioButton4.Checked = True
            ElseIf (nArray = "E") Then
                RadioButton5.Checked = True
            End If
        End If
    End Sub

  Sub BersihkanJawaban()
    RadioButton1.Checked = False
    RadioButton2.Checked = False
    RadioButton3.Checked = False
    RadioButton4.Checked = False
    RadioButton5.Checked = False
  End Sub

  Private Sub RadioButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.Click
    'jika jawaban A dipilih maka lakukan proses penyesuaian jawaban apakah benar atau salah
    CMD = New OleDbCommand("select * from tblsoal where IdSimulasi='" & Microsoft.VisualBasic.Left(ComboBox1.Text, 3) & "' and VAL(nomor)='" & ListBox1.Text & "'", Conn)
    DR = CMD.ExecuteReader
    DR.Read()
    If DR.HasRows Then
      Label20.Text = "A"
      Label21.Text = DR.Item("Jawaban")
      If Label20.Text = Label21.Text Then Label22.Text = "BENAR" Else Label22.Text = "SALAH"
    End If
  End Sub

  Private Sub RadioButton2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton2.Click
    'jika jawaban B dipilih maka lakukan proses penyesuaian jawaban apakah benar atau salah
    CMD = New OleDbCommand("select * from tblsoal where IdSimulasi='" & Microsoft.VisualBasic.Left(ComboBox1.Text, 3) & "' and VAL(nomor)='" & ListBox1.Text & "'", Conn)
    DR = CMD.ExecuteReader
    DR.Read()
    If DR.HasRows Then
      Label20.Text = "B"
      Label21.Text = DR.Item("Jawaban")
      If Label20.Text = Label21.Text Then Label22.Text = "BENAR" Else Label22.Text = "SALAH"
    End If
  End Sub

  Private Sub RadioButton3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton3.Click
    'jika jawaban C dipilih maka lakukan proses penyesuaian jawaban apakah benar atau salah
    CMD = New OleDbCommand("select * from tblsoal where IdSimulasi='" & Microsoft.VisualBasic.Left(ComboBox1.Text, 3) & "' and VAL(nomor)='" & ListBox1.Text & "'", Conn)
    DR = CMD.ExecuteReader
    DR.Read()
    If DR.HasRows Then
      Label20.Text = "C"
      Label21.Text = DR.Item("Jawaban")
      If Label20.Text = Label21.Text Then Label22.Text = "BENAR" Else Label22.Text = "SALAH"
    End If
  End Sub

  Private Sub RadioButton4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton4.Click
    'jika jawaban D dipilih maka lakukan proses penyesuaian jawaban apakah benar atau salah
    CMD = New OleDbCommand("select * from tblsoal where IdSimulasi='" & Microsoft.VisualBasic.Left(ComboBox1.Text, 3) & "' and VAL(nomor)='" & ListBox1.Text & "'", Conn)
    DR = CMD.ExecuteReader
    DR.Read()
    If DR.HasRows Then
      Label20.Text = "D"
      Label21.Text = DR.Item("Jawaban")
      If Label20.Text = Label21.Text Then Label22.Text = "BENAR" Else Label22.Text = "SALAH"
    End If
  End Sub

  Private Sub RadioButton5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton5.Click
    'jika jawaban E dipilih maka lakukan proses penyesuaian jawaban apakah benar atau salah
    CMD = New OleDbCommand("select * from tblsoal where IdSimulasi='" & Microsoft.VisualBasic.Left(ComboBox1.Text, 3) & "' and VAL(nomor)='" & ListBox1.Text & "'", Conn)
    DR = CMD.ExecuteReader
    DR.Read()
    If DR.HasRows Then
      Label20.Text = "E"
      Label21.Text = DR.Item("Jawaban")
      If Label20.Text = Label21.Text Then Label22.Text = "BENAR" Else Label22.Text = "SALAH"
    End If
  End Sub


  Private Sub BTNJawab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNJawab.Click

    'jika soal belum dipilih
    If ComboBox1.Text = "" Then
      MsgBox("Anda belum memilih soal")
      Exit Sub
    End If

    'jika belum memeilih nomor soal
    If ListBox1.Text = "" Or TextBox1.Text = "" Then
      MsgBox("Anda belum memilih nomor soal")
      Exit Sub
    End If

    'jika belum memilih jawaban
    If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False Then
      MsgBox("Anda belum memilih jawaban")
      Exit Sub
    End If

        If ulang = 0 Then
            DGV.Rows.Add(ListBox1.Text, Label20.Text, Label21.Text, Label22.Text)
            BTNSelesai.Enabled = True
            ListBox1.Focus()
            TabControl1.SelectTab(TabPage2)
        Else
            For BARIS As Integer = 0 To DGV.RowCount - 1
                If ListBox1.Text = DGV.Rows(BARIS).Cells(0).Value Then
                    Dim x As String = MessageBox.Show("Nomer " + DGV.Rows(BARIS).Cells(0).Value + " Sudah Terjawab " + Chr(13) + "Apakah Ingin Dirubah?", "Konfirmasi Perubahan Jawaban", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (x = vbYes) Then
                        If (RadioButton1.Checked = True) Then
                            DGV.Rows(BARIS).Cells(1).Value = "A"
                            DGV.Rows(BARIS).Cells(3).Value = Label22.Text
                        ElseIf (RadioButton2.Checked = True) Then
                            DGV.Rows(BARIS).Cells(1).Value = "B"
                            DGV.Rows(BARIS).Cells(3).Value = Label22.Text
                        ElseIf (RadioButton3.Checked = True) Then
                            DGV.Rows(BARIS).Cells(1).Value = "C"
                            DGV.Rows(BARIS).Cells(3).Value = Label22.Text
                        ElseIf (RadioButton4.Checked = True) Then
                            DGV.Rows(BARIS).Cells(1).Value = "D"
                            DGV.Rows(BARIS).Cells(3).Value = Label22.Text
                        ElseIf (RadioButton5.Checked = True) Then
                            DGV.Rows(BARIS).Cells(1).Value = "E"
                            DGV.Rows(BARIS).Cells(3).Value = Label22.Text
                        End If
                        n1 = ListBox1.SelectedIndex
                        Call cekRadioButton(n1)
                        If ListBox1.Items.Count - 1 = ListBox1.SelectedIndex Then
                            MsgBox("soal sudah habis2")
                            'panggil method yang menandakan bahwa soal sudah habis (hasil evaluasi nilai keseluruhan)
                        Else
                            ListBox1.SelectedIndex = ListBox1.SelectedIndex + 1
                            Call cekRadioButton(ListBox1.SelectedIndex)
                        End If
                    Else
                        MessageBox.Show("Batal Merubah Jawaban, Maka Soal Selanjutnya...", "Informasi Perubahan Jawaban", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ListBox1.SelectedIndex = ListBox1.SelectedIndex + 1
                    End If
                    Exit Sub
                End If
            Next
            DGV.Rows.Add(ListBox1.Text, Label20.Text, Label21.Text, Label22.Text)

        End If
        ListBox1.Focus()
        TabControl1.SelectTab(TabPage2)
        n1 = ListBox1.SelectedIndex
        Call cekRadioButton(n1)
        If ListBox1.Items.Count - 1 = ListBox1.SelectedIndex Then
            MsgBox("soal sudah habis2")
            'panggil method yang menandakan bahwa soal sudah habis (hasil evaluasi nilai keseluruhan)
        Else
            ListBox1.SelectedIndex = ListBox1.SelectedIndex + 1
            Call cekRadioButton(ListBox1.SelectedIndex)
        End If
        'ListBox1.SelectedIndex = ListBox1.SelectedIndex + 1
        ulang = 1
  End Sub


  Sub validasikategori1()
    If ListBox1.SelectedIndex = "99" Then
      MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Sub validasikategori2()
    If ListBox1.SelectedIndex = "3" Then
      MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Sub validasikategori3()
    If ListBox1.SelectedIndex = "0" Then
      MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Sub validasikategori4()
    If ListBox1.SelectedIndex = "5" Then
      MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Sub validasikategori5()
    If ListBox1.SelectedIndex = "6" Then
      MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Sub validasikategori6()
    If ListBox1.SelectedIndex = "7" Then
      MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Sub validasikategori7()
    If ListBox1.SelectedIndex = "8" Then
            MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Sub validasikategori8()
    If ListBox1.SelectedIndex = "9" Then
      MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Sub validasikategori9()
    If ListBox1.SelectedIndex = "9" Then
      MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Sub validasikategori10()
    If ListBox1.SelectedIndex = "9" Then
      MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Sub validasikategori11()
    If ListBox1.SelectedIndex = "9" Then
      MsgBox("Silahkan klik 'SELESAI' untuk mengakhiri ujian ATAU klik 'RESET' untuk mengulanginya dari awal")
    End If
  End Sub

  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Label12.Text = TimeOfDay
    Timer1.Enabled = False
  End Sub

  Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
    Label13.Text = TimeOfDay
  End Sub

  'membuat fungsi untuk menghitung jumlah jawaban yang benar
  Sub JumlahBenar()
    Dim hitung As Integer = 0
    For baris As Integer = 0 To DGV.RowCount - 2
      If DGV.Rows(baris).Cells(3).Value = "BENAR" Then
        hitung = hitung + 1
        Label17.Text = hitung
      End If
    Next
  End Sub

  'membuat fungsi untuk menghitung jumlah jawaban yang benar
  Sub JumlahSalah()
    Dim hitung As Integer = 0
    For baris As Integer = 0 To DGV.RowCount - 2
      If DGV.Rows(baris).Cells(3).Value = "SALAH" Then
        hitung = hitung + 1
        Label18.Text = hitung
      End If
    Next
  End Sub

  Private Sub BTNTutup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNTutup.Click
    ComboBox1.Text = ""
    ListBox1.Items.Clear()
    TextBox1.Clear()
    DGV.Rows.Clear()
    Me.Close()
    Me.Dispose()
    splash.Dispose()
  End Sub


  Private Sub BTNPetunjuk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPetunjuk.Click
    MsgBox("1. Pilih Soal yang tersedia pada combo box" & Chr(13) & _
           "2. Pilih nomor soal dalam list di sebelah kiri" & Chr(13) & _
           "3. Pilih jawaban pada option button" & vbCrLf & _
           "4. Klik jawab untuk melanjutkan ke soal berikutnya'" & vbCrLf & _
           "5. klik tombol 'Selesai Ujian'" & vbCrLf & _
           "5. Jika Ingin mengulangi ujian klik tombol 'Reset'")
  End Sub

  Private Sub BTNBatal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNBatal.Click
    Call BersihkanJawaban()
    Call timerreset()
    For x As Integer = 0 To nilai.Length - 1
      nilai(x) = ""
    Next

    Label15.Text = "0"
    Label16.Text = "0"
    Label17.Text = "0"
    Label18.Text = "0"
    Label19.Text = "-"

    TextBox1.Text = ""
    'ListBox1.SelectedIndex = "0"
    ListBox1.Enabled = True

    DGV.Rows.Clear()
    Me.Timer4.Enabled = False

    BTNSelesai.Enabled = True
    BTNJawab.Enabled = True
    'ListBox1.SelectedIndex = 0
    TabControl1.SelectTab(TabPage1)

    ComboBox1.Enabled = True
    ListBox1.Items.Clear()
    Panel4.Enabled = False
    Timer2.Enabled = True

    BTNSelesai.Enabled = True
    btnSebelumnya.Enabled = True
  End Sub


  Private Sub BTNSelesai_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelesai.Click
    If ListBox1.Text = "" Or TextBox1.Text = "" Then
      MsgBox("Anda belum memilih nomor soal")
      Exit Sub
    End If

    Try
      'ketika BTNselesai di klik maka....
      Timer2.Enabled = False
      Timer4.Enabled = False
      Dim awal As Date = TimeValue(Label12.Text)
      Dim hasil As TimeSpan = Now - awal
      'hitung durasi pengerjaan soal ujian
      Label14.Text = (String.Format("{0}:{1}:{2}", hasil.Hours, hasil.Minutes, hasil.Seconds))

      'hitung banyaknya nomor soal ujian
      Label15.Text = ListBox1.Items.Count
      Label16.Text = DGV.RowCount - 1

      Call JumlahBenar()
      Call JumlahSalah()
      'jika jumlah benar > jumlah salah maka "LULUS"
      'If Val(Label17.Text) > Val(Label18.Text) Then
      '    Label19.Text = Label17.Text + ("0 Poin")
      'Else
      '    Label19.Text = "GAGAL"
      'End If

      If ComboBox1.SelectedItem = "1   Soal Tipe A" Or ComboBox1.SelectedItem = "3   Soal Tipe C" Then
        Call hitung1()
      Else
        Call hitung2()
      End If

      'simpan semua hasil ujian ke tabel detail jawaban
      For baris As Integer = 0 To DGV.RowCount - 2
        Dim simpandetail As String = "insert into tbldetailjawaban values ('" & Label23.Text & "','" & Microsoft.VisualBasic.Left(ComboBox1.Text, 3) & "','" & DGV.Rows(baris).Cells(0).Value & "','" & DGV.Rows(baris).Cells(1).Value & "','" & DGV.Rows(baris).Cells(2).Value & "','" & DGV.Rows(baris).Cells(3).Value & "')"
        CMD = New OleDbCommand(simpandetail, Conn)
        CMD.ExecuteNonQuery()
      Next

      'simpan summary hasil ujian ke tabel master jawaban
      Dim simpanmaster As String = "insert into tblmasterjawaban values ('" & Label23.Text & "','" & Microsoft.VisualBasic.Left(ComboBox1.Text, 3) & "','" & Label11.Text & "','" & Label12.Text & "','" & Label13.Text & "','" & Label14.Text & "','" & Label15.Text & "','" & Label16.Text & "','" & Label17.Text & "','" & Label18.Text & "','" & Label19.Text & "')"
      CMD = New OleDbCommand(simpanmaster, Conn)
      CMD.ExecuteNonQuery()
      ComboBox1.Enabled = False
      ListBox1.Enabled = False
      BTNJawab.Enabled = False
      BTNSelesai.Enabled = False
      btnSebelumnya.Enabled = False
      ListBox1.Items.Clear()
      TextBox1.Clear()
      Call BersihkanJawaban()
      TabControl1.SelectTab(TabPage1)
      Panel4.Enabled = False
    Catch ex As Exception
      MsgBox(ex.Message)
    End Try


  End Sub

  Sub cekRadioButton(ByVal pNilaiIndex As Integer)
    If (RadioButton1.Checked = True) Then
      nilai(pNilaiIndex) = "A"
    ElseIf (RadioButton2.Checked = True) Then
      nilai(pNilaiIndex) = "B"
    ElseIf (RadioButton3.Checked = True) Then
      nilai(pNilaiIndex) = "C"
    ElseIf (RadioButton4.Checked = True) Then
      nilai(pNilaiIndex) = "D"
    ElseIf (RadioButton5.Checked = True) Then
      nilai(pNilaiIndex) = "E"
    End If
  End Sub

  Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
    detik -= 1
    If detik < 0 Then
      detik = 59
      menit -= 1
      If menit < 0 Then
        menit = 59
        jam -= 1
      End If
    End If

    Label29.Text = Format(jam, "00") & ":" & Format(menit, "00") & ":" & Format(detik, "00")
    If jam = 0 And menit = 0 And detik = 0 Then
      Timer4.Enabled = False

      MsgBox("waktu habis !!")
      Call hitung()
    End If
  End Sub

  Sub hitung()
    Label15.Text = ListBox1.Items.Count
    Label16.Text = DGV.RowCount - 1

    Call JumlahBenar()
    Call JumlahSalah()
    'jika jumlah benar > jumlah salah maka "LULUS"
    If Val(Label17.Text) > Val(Label18.Text) Then
      Label19.Text = "LULUS"
    Else
      Label19.Text = "GAGAL"
    End If

  End Sub

  Sub hitung1()
    Dim nilai As String = Label17.Text
    If nilai >= 70 Then
      Label19.Text = "30 poin"
    ElseIf nilai >= 50 Then
      Label19.Text = "25 poin"
    ElseIf nilai >= 10 Then
      Label19.Text = "15 poin"
    Else
      Label19.Text = "5 poin"
    End If
  End Sub

  Sub hitung2()
    Dim nilai As String = Label17.Text
    If nilai >= 30 Then
      Label19.Text = "30 poin"
    ElseIf nilai >= 20 Then
      Label19.Text = "25 poin"
    ElseIf nilai >= 10 Then
      Label19.Text = "15 poin"
    Else
      Label19.Text = "5 poin"
    End If
  End Sub

  Private Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    CMD = New OleDbCommand("select nomor from tblsoal where IdSimulasi='" & Microsoft.VisualBasic.Left(ComboBox1.Text, 3) & "' order by 1", Conn)
    DR = CMD.ExecuteReader
    ListBox1.Items.Clear()
    Do While DR.Read
      ListBox1.Items.Add(DR.Item("Nomor"))
    Loop
    ListBox1.Focus()
    ComboBox1.Enabled = False


  End Sub


  Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSebelumnya.Click
    Dim ket As String = nilai(0)
    'jika belum memeilih nomor soal
    If ComboBox1.Text = "" Then
      MsgBox("Anda belum memilih soal")
      Exit Sub
    End If
    'jika belum memeilih nomor soal
    If ListBox1.Text = "" Or TextBox1.Text = "" Then
      MsgBox("Anda belum memilih nomor soal")
      Exit Sub
    End If

    If ListBox1.SelectedIndex = "0" Then
      MsgBox("Klik reset jika anda ingin mengulang ujian")
    Else
      ListBox1.SelectedIndex = ListBox1.SelectedIndex - 1
      n2 = ListBox1.SelectedIndex
      tampilKeRadioButton(n2)
    End If
  End Sub

  Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
    CMD = New OleDbCommand("select nomor from tblsoal where IdSimulasi='" & Microsoft.VisualBasic.Left(ComboBox1.Text, 3) & "' order by 1", Conn)
    DR = CMD.ExecuteReader
        ListBox1.Items.Clear()

        Do While DR.Read
            ListBox1.Items.Add(DR.Item("Nomor"))
        Loop
    ListBox1.Focus()
    ComboBox1.Enabled = False
    Panel4.Enabled = True
    ListBox1.SelectedIndex = 0

  End Sub


    'Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    'If ComboBox1.SelectedItem = "test" Then
    '    '    MsgBox("asda")
    '    'End If

    '    Dim pie As String
    '    pie = ComboBox1.SelectedItem
    '    TextBox2.Text = pie

    '    'If ListBox2.SelectedItem = pie Then
    '    '    MsgBox("waduh")
    '    'End If
    '    'If ComboBox1.Text = "2   Soal Tipe B" Then
    '    '    MsgBox("bisa")
    '    'End If
    'End Sub

End Class