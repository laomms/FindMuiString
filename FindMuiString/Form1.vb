Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports AsmResolver.PE.Win32Resources

Public Class Form1
    Private totalTime As Integer
    Private title As String
    Friend Shared MyInstance As Form1
    Enum TypeControl
        typButton = 1
        typStatic = 2
        typOther = 2
    End Enum

    <DllImport("kernel32", SetLastError:=True)>
    Public Shared Function LoadLibrary(ByVal lpFileName As String) As IntPtr
    End Function
    <DllImport("kernel32.dll")>
    Public Shared Function LoadLibraryEx(lpFileName As String, hReservedNull As IntPtr, dwFlags As Integer) As IntPtr
    End Function
    <DllImport("kernel32.dll")>
    Public Shared Function FreeLibrary(ByVal hModule As IntPtr) As Boolean
    End Function
    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function FindResource(ByVal hModule As IntPtr, ByVal lpName As IntPtr, ByVal lpType As IntPtr) As IntPtr
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function LoadResource(ByVal hModule As IntPtr, ByVal hResInfo As IntPtr) As IntPtr
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function SizeofResource(ByVal hModule As IntPtr, ByVal hResInfo As IntPtr) As UInteger
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function LockResource(ByVal hResData As IntPtr) As IntPtr
    End Function
    <DllImport("version.dll")>
    Public Shared Function VerQueryValue(ByVal pBlock() As Byte, ByVal lpSubBlock As String, ByRef lplpBuffer As IntPtr, ByRef puLen As UInteger) As Boolean
    End Function
    <DllImport("version.dll")>
    Public Shared Function GetFileVersionInfoSize(ByVal fileName As String, <[Out]> ByVal dummy As IntPtr) As Integer
    End Function
    <DllImport("version.dll", SetLastError:=True)>
    Public Shared Function GetFileVersionInfo(ByVal lptstrFilename As String, ByVal dwHandleIgnored As Integer, ByVal dwLen As Integer, ByVal lpData() As Byte) As Boolean
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = String.Empty
        TextBox1.ForeColor = Color.Black
        TextBox1.TextAlign = HorizontalAlignment.Left
        Using dlg As New OpenFileDialog With {.AddExtension = True,
                                              .ValidateNames = False,
                                              .CheckFileExists = False,
                                              .CheckPathExists = True,
                                              .FileName = "Folder Selection"
                                              }
            If dlg.ShowDialog = DialogResult.OK Then
                TextBox1.Text = Path.GetDirectoryName(dlg.FileName)
            End If
        End Using
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_DragDrop(sender As Object, e As DragEventArgs) Handles TextBox1.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each filepath In files
            TextBox1.Text = String.Empty
            TextBox1.ForeColor = Color.Black
            TextBox1.TextAlign = HorizontalAlignment.Left
            TextBox1.Text = Path.GetDirectoryName(filepath)
        Next
    End Sub

    Private Sub TextBox1_DragEnter(sender As Object, e As DragEventArgs) Handles TextBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, er As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim arguments As List(Of String) = TryCast(er.Argument, List(Of String))
        Dim folderPath = arguments(0)
        MyInstance.Invoke(New MethodInvoker(Sub() MyInstance.RichTextBox1.Clear()))
        totalTime = 0
        Me.Invoke(New NewTimersStar(AddressOf TimersStar))
        Try
            Dim fileList = SafeFileEnumerator.EnumerateFiles(folderPath, "*.mui", SearchOption.AllDirectories)
            If fileList.Count > 0 Then
                If Directory.Exists(My.Computer.FileSystem.SpecialDirectories.Temp + "\muifile") = False Then Directory.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.Temp + "\muifile")
                For Each muifile In fileList
                    Dim filename = My.Computer.FileSystem.SpecialDirectories.Temp + "\muifile\" + Path.GetFileName(muifile)
                    MyInstance.Invoke(New MethodInvoker(Sub() MyInstance.RichTextBox1.AppendText(Path.GetFileName(muifile) + vbNewLine)))
                    My.Computer.FileSystem.CopyFile(muifile, filename, True)
                    Dim peImages As AsmResolver.PE.IPEImage = AsmResolver.PE.PEImage.FromFile(filename)
                    For Each entry As IResourceDirectory In peImages.Resources.Entries
                        If entry.IsDirectory Then
                            If entry.Id = 6 Then
                                For Each item In entry.Entries
                                    'Dim dataEntry As IResourceData = peImages.Resources.GetDirectory(ResourceType.String).GetDirectory(251).GetData(1033)
                                    Dim stringTables As IResourceDirectory = peImages.Resources.Entries.First(Function(e) e.Id = ResourceType.String)
                                    Dim stringEntry As IResourceDirectory = stringTables.Entries.First(Function(e) e.Id = item.Id)
                                    Dim LangId = GetLangId(filename)
                                    If LangId = 0 Then
                                        LangId = 2052
                                    End If
                                    Dim dataEntry As IResourceData = stringEntry.Entries.First(Function(e) e.Id = LangId)
                                    Dim content() As Byte = CType(dataEntry.Contents, AsmResolver.DataSegment).Data
                                    Dim stringlist As List(Of String) = GetStringResource(Path.GetFileName(muifile), item.Id, content)
                                    If stringlist.Count > 0 Then
                                        MyInstance.Invoke(New MethodInvoker(Sub() MyInstance.RichTextBox1.AppendText(String.Join(vbNewLine, stringlist) + vbNewLine)))
                                        MyInstance.RichTextBox1.Invoke(New MethodInvoker(Sub() MyInstance.RichTextBox1.ScrollToCaret()))
                                    End If
                                Next
                            ElseIf entry.Id = 5 Then
                                For Each item In entry.Entries
                                    Dim DialogTables As IResourceDirectory = peImages.Resources.Entries.First(Function(e) e.Id = ResourceType.Dialog)
                                    Dim DialogEntry As IResourceDirectory = DialogTables.Entries.First(Function(e) e.Id = item.Id)
                                    Dim LangId = GetLangId(filename)
                                    If LangId = 0 Then
                                        LangId = 2052
                                    End If
                                    Dim dataEntry As IResourceData = DialogEntry.Entries.First(Function(e) e.Id = LangId)
                                    Dim content() As Byte = CType(dataEntry.Contents, AsmResolver.DataSegment).Data
                                    Dim stringlist As List(Of String) = ParseDialogString(Path.GetFileName(muifile), item.Id, content)
                                    If stringlist.Count > 0 Then
                                        MyInstance.Invoke(New MethodInvoker(Sub() MyInstance.RichTextBox1.AppendText(String.Join(vbNewLine, stringlist) + vbNewLine)))
                                        MyInstance.RichTextBox1.Invoke(New MethodInvoker(Sub() MyInstance.RichTextBox1.ScrollToCaret()))
                                    End If
                                Next
                            End If
                        End If
                    Next entry

                Next
            End If
        Catch ex As Exception

        End Try
        If Directory.Exists(My.Computer.FileSystem.SpecialDirectories.Temp + "\muifile") = False Then Directory.Delete(My.Computer.FileSystem.SpecialDirectories.Temp + "\muifile")
        Timer1.Stop()
        MyInstance.Invoke(New MethodInvoker(Sub() MyInstance.Text = title))

    End Sub
    Private Function GetLangId(filename As String) As Integer
        Dim hd As IntPtr = Nothing
        Dim size = GetFileVersionInfoSize(filename, hd)
        Dim buffers(size - 1) As Byte
        If Not GetFileVersionInfo(filename, hd, size, buffers) Then
            'Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If

        Dim versionLength As UInteger
        Dim blockbuffer As IntPtr = IntPtr.Zero

        Dim success = VerQueryValue(buffers, "\VarFileInfo\Translation", blockbuffer, versionLength)
        Dim langs() As UInteger
        Try
            If success Then
                langs = New UInteger(versionLength / 4 - 1) {}
                Dim i As Integer = 0
                Dim j As Integer = 0
                Do While j < versionLength
                    langs(i) = (Marshal.ReadInt16(blockbuffer, j) << 16) ' Or Marshal.ReadInt16(blockbuffer, j + 2)
                    i += 1
                    j += 4
                Loop
            Else
                langs = New UInteger() {&H4090000}
            End If
        Catch ex As Exception
            langs = New UInteger() {&H4090000}
        End Try
        Return langs(0)
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then Return
        If BackgroundWorker1.IsBusy = False Then
            Dim arguments = New List(Of String)
            arguments.Add(TextBox1.Text)
            BackgroundWorker1.RunWorkerAsync(arguments)
        End If
    End Sub
    Private ts As TimeSpan
    Private sw As Stopwatch
    Delegate Sub NewTimersStar()
    Private Sub TimersStar()
        sw = New Stopwatch()
        AddHandler Timer1.Tick, AddressOf Timer1_Tick
        Timer1.Interval = 1
        sw.Start()
        Timer1.Start()
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ts = sw.Elapsed
        MyInstance.Invoke(New MethodInvoker(Sub() MyInstance.Text = title + "        " + String.Format("{0:D2}:{1:D2}:{2:D2}:{3:D2}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds \ 10)))

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        title = Me.Text
        MyInstance = Me
        RichTextBox1.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic).SetValue(RichTextBox1, True, Nothing)
    End Sub
    Private Function GetStringResource(ByVal filename As String, ByVal ResId As Integer, ByVal bytesIn() As Byte) As List(Of String)
        Dim result As New List(Of String)
        Dim ms As New MemoryStream()
        If bytesIn.Length > 0 Then
            Dim writer As New BinaryWriter(ms)
            Dim OldBinary() As Char = Encoding.Unicode.GetString(bytesIn).ToCharArray()
            Dim offset As Integer = 0
            For index As Integer = 0 To 15
                If AscW(OldBinary(offset)) <> 0 Then
                    Dim length As Integer = AscW(OldBinary(offset)) ' GetWord()  // String length in characters
                    offset += 1
                    If length > 0 Then
                        Try
                            Dim stringID As Integer = (ResId - 1) * 16 + index
                            writer.Write(stringID)
                            writer.Write(length)
                            Dim inner As Integer = offset
                            Do While inner < offset + length
                                writer.Write(OldBinary(inner))
                                inner += 1
                            Loop
                        Catch ex As Exception

                        End Try
                    End If
                    offset += length
                Else
                    offset += 1
                End If
            Next index
            ms.Position = 0
        End If
        If ms IsNot Nothing AndAlso ms.Length > 0 Then
            Dim reader As New BinaryReader(ms)
            Do While ms.Position < ms.Length
                ' Keep binary reader's order!
                Dim stringID As Integer = reader.ReadInt32()
                Dim length As Integer = reader.ReadInt32()
                Dim buff() As Char = reader.ReadChars(length)
                'Debug.Print(stringID.ToString + ":" + New String(buff))
                result.Add("[" + filename + "]" + ResId.ToString + "->" + stringID.ToString + ":" + New String(buff))
            Loop
        End If

        Return result
    End Function
    Private Function WriteStringResource(ByVal name As Integer, ByVal bytesIn() As Byte, ByVal UpdateString As Dictionary(Of String, String)) As Byte()
        Dim ms As New MemoryStream()
        If bytesIn.Length > 0 Then
            Dim writer As New BinaryWriter(ms, Encoding.Unicode)
            Dim OldBinary() As Char = Encoding.Unicode.GetString(bytesIn).ToCharArray()
            Dim oldOffset As Integer = 0
            Dim newOffset As Integer = 0
            For index As Integer = 0 To 15
                Try
                    If Convert.ToInt32(OldBinary(oldOffset)) <> 0 Then
                        Dim length As Integer = AscW(OldBinary(oldOffset))
                        oldOffset += 1
                        newOffset += 1
                        If length > 0 Then
                            Dim stringID As Integer = (name - 1) * 16 + index
                            If UpdateString.ContainsKey(stringID.ToString()) Then
                                writer.Write(CShort(UpdateString(stringID).Length))
                                writer.Write(Encoding.Unicode.GetBytes(UpdateString(stringID)))
                                newOffset += UpdateString(stringID).Length
                            Else
                                writer.Write(CShort(length))
                                Dim inner As Integer = oldOffset
                                Do While inner < oldOffset + length
                                    writer.Write(OldBinary(inner))
                                    inner += 1
                                Loop
                            End If
                            oldOffset += length
                        End If
                    Else
                        writer.Write(New Byte() {0, 0})
                        oldOffset += 1
                        newOffset += 1
                    End If
                Catch ex As Exception

                End Try
            Next index
            ms.Position = 0
        End If
        Return ms.ToArray()
    End Function
    Private Shared Function ParseDialogString(ByVal filename As String, ByVal ResId As Integer, ByVal bytesIn() As Byte) As List(Of String)
        Dim result As New List(Of String)
        Dim outString As String = ""
        Dim text As String = ""
        Dim len As Integer
        Dim isDialogEx As Boolean
        Dim DlgBoxHeader As Object
#Region "DlgBoxHeader"
        If Enumerable.SequenceEqual(bytesIn.Take(4).ToArray(), New Byte() {&H1, &H0, &HFF, &HFF}) Then 'Dialog or DialogEx
            isDialogEx = True
            DlgBoxHeader = New DialogBoxHeaderEx()
        Else
            isDialogEx = False
            DlgBoxHeader = New DialogBoxHeader()
        End If
        Try
            Dim size As Integer = Marshal.SizeOf(DlgBoxHeader)
            Dim buff As IntPtr = Marshal.AllocHGlobal(size)
            Marshal.Copy(bytesIn.Take(size).ToArray(), 0, buff, size)
            DlgBoxHeader = If(isDialogEx, CType(Marshal.PtrToStructure(buff, GetType(DialogBoxHeaderEx)), DialogBoxHeaderEx), CType(Marshal.PtrToStructure(buff, GetType(DialogBoxHeader)), DialogBoxHeader))
            Marshal.FreeHGlobal(buff)

#End Region
#Region "是否有菜单MenuName"
            Dim bMenuName = bytesIn.Skip(size).ToArray()
            Dim bClassName() As Byte
            If bMenuName(0) = &HFF And bMenuName(1) = &HFF Then
                bClassName = bMenuName.Skip(2).ToArray
            ElseIf bMenuName(0) = 0 And bMenuName(1) = 0 Then
                bClassName = bMenuName.Skip(2).ToArray
            Else
                len = GetUnicodeString(bMenuName, outString)
                bClassName = bMenuName.Skip(len + 2).ToArray()
            End If
#End Region
#Region "是否有窗口类ClassName"
            Dim bCaption() As Byte
            If bClassName(0) = &HFF And bClassName(1) = &HFF Then
                bCaption = bClassName.Skip(2).ToArray()
            ElseIf bClassName(0) = 0 And bClassName(1) = 0 Then
                bCaption = bClassName.Skip(2).ToArray()
            Else
                len = GetUnicodeString(bClassName, outString)
                Dim szClassName = Encoding.Unicode.GetString(bClassName, 0, len)
                bCaption = bClassName.Skip(len + 2).ToArray()
            End If
#End Region
#Region "窗口标题Caption"
            Dim bFontData() As Byte
            If bCaption(0) = &HFF And bCaption(1) = &HFF Then
                bFontData = bCaption.Skip(2).ToArray()
            Else
                len = GetUnicodeString(bCaption, outString)
                text = "[" + filename + "]" + ResId.ToString + ">>" + outString + "->"
                bFontData = bCaption.Skip(len + 2).ToArray()
            End If
#End Region
#Region "窗口字体DialogFont"
            Dim DlgFont As Object
            If isDialogEx Then
                DlgFont = New DialogFontEx()
            Else
                DlgFont = New DialogFont()
            End If
            size = Marshal.SizeOf(DlgFont)
            buff = Marshal.AllocHGlobal(size)
            Marshal.Copy(bFontData.Take(size).ToArray(), 0, buff, size)
            If isDialogEx Then
                DlgFont = CType(Marshal.PtrToStructure(buff, GetType(DialogFontEx)), DialogFontEx)
            Else
                DlgFont = CType(Marshal.PtrToStructure(buff, GetType(DialogFont)), DialogFont)
            End If
            Marshal.FreeHGlobal(buff)
            bFontData = bFontData.Skip(size).ToArray
            If bFontData(0) = &HFF And bFontData(1) = &HFF Then
                bFontData = bFontData.Skip(2).ToArray
            Else
                len = GetUnicodeString(bFontData, outString)
                bFontData = bFontData.Skip(len).ToArray()
            End If
#End Region
#Region "ControlList"
            Dim ctlData As Object
            If isDialogEx Then
                ctlData = New ControlDataEx()
            Else
                ctlData = New ControlData()
            End If
            Dim bItemData() As Byte = bFontData.Skip(2).ToArray()
            'bItemData = (bItemData.Length - bytesIn.Length + 3) / 4 * 4 + bItemData
            For i As Integer = 0 To DlgBoxHeader.DlgItems - 1
                len = (bytesIn.Length - bItemData.Length + 3) \ 4
                bItemData = bytesIn.Skip(len * 4).ToArray

                size = Marshal.SizeOf(ctlData)
                buff = Marshal.AllocHGlobal(size)
                Marshal.Copy(bItemData.Take(size).ToArray(), 0, buff, size)
                If isDialogEx Then
                    ctlData = CType(Marshal.PtrToStructure(buff, GetType(ControlDataEx)), ControlDataEx)
                Else
                    ctlData = CType(Marshal.PtrToStructure(buff, GetType(ControlData)), ControlData)
                End If
                Marshal.FreeHGlobal(buff)
                'Debug.Print("Constol ID:" + ctlData.id.ToString)
                'Dim Style = If(isDialogEx, ctlData.Style, ctlData.ExStyle)
                'Dim styleValue = Style And &HF
                'Debug.Print("style:0x" + CInt(ctlData.style).ToString("x8") + " exstyle:0x" + CInt(ctlData.exstyle).ToString("x8") + " style value:" + styleValue.ToString)
                Dim bClassID() As Byte = bItemData.Skip(size + 2).ToArray()
                If bClassID(0) = &HFF And bClassID(1) = &HFF Then '控件类型
                    bClassID = bClassID.Skip(2).ToArray()
                End If
                If bClassID(0) = &H80 And bClassID(1) = 0 Then
                    Debug.Print("Button")
                ElseIf bClassID(0) = &H81 And bClassID(1) = 0 Then
                    Debug.Print("EditText")
                ElseIf bClassID(0) = &H82 And bClassID(1) = 0 Then
                    Debug.Print("Static")
                ElseIf bClassID(0) = &H83 And bClassID(1) = 0 Then
                    Debug.Print("ListBox")
                ElseIf bClassID(0) = &H84 And bClassID(1) = 0 Then
                    Debug.Print("ScrollBar")
                ElseIf bClassID(0) = &H85 And bClassID(1) = 0 Then
                    Debug.Print("ComboBox")
                ElseIf bClassID(0) <> 0 Then '其他控件
                    len = GetUnicodeString(bClassID, outString)
                    Debug.Print(outString)
                    bClassID = bClassID.Skip(len).ToArray()
                End If
                bClassID = bClassID.Skip(2).ToArray()
                If bClassID(0) = &HFF And bClassID(1) = &HFF Then
                    bClassID = bClassID.Skip(2).ToArray()
                    len = GetUnicodeString(bClassID, outString)
                    result.Add(text + ctlData.id.ToString() + ":" + outString)
                    Debug.Print(text + ctlData.id.ToString() + ":" + outString)
                    bClassID = bClassID.Skip(len).ToArray()
                Else
                    len = GetUnicodeString(bClassID, outString)
                    result.Add(text + ctlData.id.ToString() + ":" + outString)
                    Debug.Print(text + ctlData.id.ToString() + ":" + outString)
                    bClassID = bClassID.Skip(len).ToArray()
                End If
                bClassID = bClassID.Skip(2).ToArray
                Debug.Print(vbNewLine)
                If isDialogEx Then bItemData = bClassID.Skip(2).ToArray
            Next i
        Catch ex As Exception

        End Try

#End Region
        Return result
    End Function
    Private Shared Function GetUnicodeString(ByVal bytesIn() As Byte, ByRef outString As String) As Integer
        Dim result As New List(Of String)
        Dim len As Integer
        outString = ""
        Using ms As New MemoryStream()
            Using writer As New BinaryWriter(ms)
                If bytesIn.Length > 0 Then
                    Dim OldBinary() As Char = Encoding.Unicode.GetString(bytesIn).ToCharArray()
                    For index As Integer = 0 To bytesIn.Length
                        If AscW(OldBinary(index)) <> 0 Then
                            writer.Write(OldBinary(index))
                            'index += 1
                        Else
                            Exit For
                        End If
                    Next
                    ms.Position = 0
                End If
                If ms IsNot Nothing AndAlso ms.Length > 0 Then
                    Dim reader As New BinaryReader(ms)
                    len = ms.Length
                    Do While ms.Position < ms.Length
                        Dim buff() As Char = reader.ReadChars(ms.Length)
                        outString = New String(buff)
                    Loop
                End If
            End Using
        End Using
        Return outString.Length * 2
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If RichTextBox1.Text = "" Or TextBox2.Text = "" Then Return
        'Find(RichTextBox1, TextBox2.Text, Color.Blue)
        Dim startindex As Integer = 0
        startindex = FindMyText(TextBox2.Text.Trim(), start, RichTextBox1.Text.Length)
        If startindex = -1 AndAlso start >= 0 Then
            start = 0
            startindex = FindMyText(TextBox2.Text.Trim(), start, RichTextBox1.Text.Length)
        End If

        If startindex >= 0 Then
            RichTextBox1.Select(startindex, TextBox2.TextLength)
            RichTextBox1.SelectionBackColor = Color.Yellow
            Dim endindex As Integer = TextBox2.Text.Length
            RichTextBox1.Select(startindex, endindex)
            RichTextBox1.ScrollToCaret()
            start = startindex + endindex
        End If
    End Sub

    Private indexOfSearchText = 0
    Private start = 0
    Public Function FindMyText(ByVal txtToSearch As String, ByVal searchStart As Integer, ByVal searchEnd As Integer) As Integer
        If searchStart > 0 AndAlso searchEnd > 0 AndAlso indexOfSearchText >= 0 Then
            RichTextBox1.Undo()
        End If
        Dim retVal As Integer = -1

        If searchStart >= 0 AndAlso indexOfSearchText >= 0 Then
            If searchEnd > searchStart OrElse searchEnd = -1 Then
                indexOfSearchText = RichTextBox1.Find(txtToSearch, searchStart, searchEnd, RichTextBoxFinds.None)
                If indexOfSearchText <> -1 Then
                    retVal = indexOfSearchText
                Else
                    start = 0
                    indexOfSearchText = 0
                End If
            End If
        End If
        Return retVal
    End Function

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        'start = 0
        'indexOfSearchText = 0
        'Dim temp As String = RichTextBox1.Text
        'RichTextBox1.Text = String.Empty
        'RichTextBox1.Text = temp
        'RichTextBox1.SelectionBackColor = Color.White
    End Sub
End Class
