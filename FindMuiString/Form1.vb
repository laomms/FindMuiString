Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports AsmResolver.PE.Win32Resources

Public Class Form1
    Private totalTime As Integer
    Private title As String
    Friend Shared MyInstance As Form1
    Public Enum lpType
        RT_CURSOR = 1 '由硬件支持的光标资源
        RT_BITMAP = 2 '位图资源
        RT_ICON = 3 '由硬件支持的图标资源
        RT_MENU = 4 '菜单资源
        RT_DIALOG = 5 '对话框
        RT_STRING = 6 '字符表入口
        RT_FONTDIR = 7 '字体目录资源
        RT_FONT = 8 '字体资源
        RT_ACCELERATOR = 9 '加速器表
        RT_MESSAGETABLE = 11 '消息表的入口
        RT_GROUP_CURSOR = 12 '与硬件无关的光标资源
        RT_GROUP_ICON = 14 ' 与硬件无关的目标资源
        RT_VERSION = 16 '版本资源
        RT_PLUGPLAY = 19 '即插即用资源
        RT_VXD = 20 'VXD
        RT_ANICURSOR = 21 '动态光标
        RT_ANIICON = 22 '动态图标
        RT_HTML = 23 'HTML文档
        RT_RCDATA = 10 '原始数据或自定义资源
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
        Dim arguments As List(Of Object) = TryCast(er.Argument, List(Of Object))
        Dim folderPath = arguments(0)
        MyInstance.Invoke(New MethodInvoker(Sub() MyInstance.RichTextBox1.Clear()))
        totalTime = 0
        Me.Invoke(New NewTimersStar(AddressOf TimersStar))
        Try
            Dim fileList = SafeFileEnumerator.EnumerateFiles(folderPath, "*.mui", SearchOption.AllDirectories)
            If fileList.Count > 0 Then
                If Directory.Exists(My.Computer.FileSystem.SpecialDirectories.Temp + "\muifile") = False Then Directory.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.Temp + "\muifile")
                For Each muifile In fileList
                    MyInstance.Invoke(New MethodInvoker(Sub() MyInstance.RichTextBox1.AppendText(Path.GetFileName(muifile) + vbNewLine)))
                    My.Computer.FileSystem.CopyFile(muifile, My.Computer.FileSystem.SpecialDirectories.Temp + "\muifile\" + Path.GetFileName(muifile), True)
                    Dim peImages As AsmResolver.PE.IPEImage = AsmResolver.PE.PEImage.FromFile(My.Computer.FileSystem.SpecialDirectories.Temp + "\muifile\" + Path.GetFileName(muifile))
                    For Each entry As IResourceDirectory In peImages.Resources.Entries
                        If entry.IsDirectory Then
                            If entry.Id = 6 Then
                                For Each item In entry.Entries
                                    'Dim dataEntry As IResourceData = peImages.Resources.GetDirectory(ResourceType.String).GetDirectory(251).GetData(1033)
                                    Dim stringTables As IResourceDirectory = peImages.Resources.Entries.First(Function(e) e.Id = ResourceType.String)
                                    Dim stringEntry As IResourceDirectory = stringTables.Entries.First(Function(e) e.Id = item.Id)
                                    Dim dataEntry As IResourceData = stringEntry.Entries.First(Function(e) e.Id = GetLangId(My.Computer.FileSystem.SpecialDirectories.Temp + "\muifile\" + Path.GetFileName(muifile)))
                                    Dim content() As Byte = CType(dataEntry.Contents, AsmResolver.DataSegment).Data
                                    Dim stringlist As List(Of String) = GetStringResource(Path.GetFileName(muifile), item.Id, content)
                                    If stringlist.Count > 0 Then
                                        MyInstance.Invoke(New MethodInvoker(Sub() MyInstance.RichTextBox1.AppendText(String.Join(vbNewLine, stringlist) + vbNewLine)))
                                        MyInstance.RichTextBox1.Invoke(New MethodInvoker(Sub() MyInstance.RichTextBox1.ScrollToCaret()))
                                    End If
                                Next
                            ElseIf entry.Id = 5 Then

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
        If success Then
            langs = New UInteger((versionLength / 4) - 1) {}
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
        Return langs(0)
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then Return
        If BackgroundWorker1.IsBusy = False Then
            Dim arguments As New List(Of Object)
            arguments.Add(TextBox1.Text)
            BackgroundWorker1.RunWorkerAsync(arguments)
        End If
    End Sub

    Delegate Sub NewTimersStar()
    Private Sub TimersStar()
        Timer1.Interval = 1
        Timer1.Enabled = True
        Timer1.Start()
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        On Error Resume Next
        totalTime = totalTime + 1
        MyInstance.Invoke(New MethodInvoker(Sub() MyInstance.Text = title + "        " + Format((totalTime Mod 3600) \ 60, "00") & ":" & Format(totalTime Mod 60, "00")))
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        title = Me.Text
        MyInstance = Me
        Dim peImages = AsmResolver.PE.PEImage.FromFile("shell32.dll.mui")

        For Each entry As IResourceDirectory In peImages.Resources.Entries
            If entry.IsDirectory Then
                Debug.Print("Directory {0}.", entry.Id)
                If entry.Id = 6 Then
                    For Each item In entry.Entries
                        Debug.Print("item.Name {0}.", item.Id)
                    Next
                End If
            Else ' if (entry.IsData)
                Debug.Print("Data {0}.", entry.Owner.Type.ToString)
            End If
        Next entry
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
                Debug.Print(stringID.ToString + ":" + New String(buff))
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

    Private Shared Function ParseDialog(ByVal bytesIn() As Byte) As Byte()
        Dim isDialogEx As Boolean = False
        Dim DlgBoxHeader As Object
#Region "DlgBoxHeader"
        If Enumerable.SequenceEqual(bytesIn.Take(4).ToArray(), New Byte() {&H1, &H0, &HFF, &HFF}) Then 'Dialog or DialogEx
            isDialogEx = True
            DlgBoxHeader = New DialogBoxHeaderEx()
        Else
            isDialogEx = False
            DlgBoxHeader = New DialogBoxHeader()
        End If
        Dim size As Integer = Marshal.SizeOf(DlgBoxHeader)
        Dim buff As IntPtr = Marshal.AllocHGlobal(size)
        Marshal.Copy(bytesIn.Take(size).ToArray(), 0, buff, size)
        DlgBoxHeader = If(isDialogEx, CType(Marshal.PtrToStructure(buff, GetType(DialogBoxHeaderEx)), DialogBoxHeaderEx), CType(Marshal.PtrToStructure(buff, GetType(DialogBoxHeader)), DialogBoxHeader))
        If DlgBoxHeader.ExStyle <> 0 Then
            Console.WriteLine("0x" & DlgBoxHeader.ExStyle.ToString("x8"))
        End If
        bytesIn = bytesIn.Skip(size).ToArray()
        '			#End Region
        '			#Region "MenuName"
        Dim bClassName() As Byte = Nothing
        Dim bMenuName() As Byte = DlgBoxHeader.menuName
        If BitConverter.ToInt16(bMenuName.Take(2).ToArray(), 0) = -1 Then
            bClassName = bytesIn
        ElseIf BitConverter.ToInt16(bMenuName.Take(2).ToArray(), 0) <> 0 Then
            Dim len = GetUnicodeStringLength(bytesIn)
            Dim szMenuName = Encoding.Unicode.GetString(bytesIn, 0, len)
            Console.WriteLine("MenuName:" & szMenuName)
            bClassName = bytesIn.Skip(len).ToArray()
        Else
            bClassName = bytesIn
        End If
#End Region
#Region "ClassName"
        Dim bCaption() As Byte = Nothing
        If BitConverter.ToInt16(bClassName.Take(2).ToArray(), 0) = -1 Then
            bCaption = bClassName.Skip(2).ToArray()
        ElseIf BitConverter.ToInt16(bClassName.Take(2).ToArray(), 0) <> 0 Then
            Dim len = GetUnicodeStringLength(bClassName)
            Dim szClassName = Encoding.Unicode.GetString(bClassName, 0, len)
            Console.WriteLine("ClassName:" & szClassName)
            bCaption = bClassName.Skip(len + 2).ToArray()
        Else
            bCaption = bClassName.Skip(2).ToArray()
        End If
#End Region
#Region "Caption"
        Dim bItemData() As Byte = Nothing
        If BitConverter.ToInt16(bCaption.Take(2).ToArray(), 0) <> 0 Then
            Dim len = GetUnicodeStringLength(bCaption)
            Dim szCaption = Encoding.Unicode.GetString(bCaption, 0, len)
            Console.WriteLine("Caption:" & szCaption)
            bItemData = bCaption.Skip(len + 2).ToArray()
        Else
            bItemData = bCaption.Skip(2).ToArray()
        End If
#End Region
#Region "DialogFont"
        If (DlgBoxHeader.Style And CUInt(Math.Truncate(DialogBoxStyles.DS_SETFONT))) <> 0 Then 'if have font
            Dim DlgFont As New DialogFontEx()
            size = Marshal.SizeOf(DlgFont)
            buff = Marshal.AllocHGlobal(size)
            Marshal.Copy(bItemData.Take(size).ToArray(), 0, buff, size)
            DlgFont = CType(Marshal.PtrToStructure(buff, GetType(DialogFontEx)), DialogFontEx)
            If BitConverter.ToInt16(DlgFont.FontName.Take(2).ToArray(), 0) <> 0 Then
                bItemData = bItemData.Skip(size - 2).ToArray()
                Dim len = GetUnicodeStringLength(bItemData)
                Dim szFontName = Encoding.Unicode.GetString(bItemData, 0, len)
                Console.WriteLine("FontName:" & szFontName)
                bItemData = bItemData.Skip(len).ToArray()
            End If
        End If
#End Region
#Region "ControlList"
        Dim isControl As Integer = 0
        bItemData = bItemData.Skip(4).ToArray()
        For i As Integer = 0 To DlgBoxHeader.DlgItems - 1
            Dim ctlData As New ControlData()
            size = Marshal.SizeOf(ctlData)
            buff = Marshal.AllocHGlobal(size)
            Dim helpID As UInteger = 0
            If isDialogEx Then
                helpID = CUInt(BitConverter.ToInt16(bItemData.Take(2).ToArray(), 0))
                Marshal.Copy(bItemData.Skip(2).Take(size).ToArray(), 0, buff, size)
            Else
                Marshal.Copy(bItemData.Take(size).ToArray(), 0, buff, size)
            End If
            ctlData = CType(Marshal.PtrToStructure(buff, GetType(ControlData)), ControlData)
            Console.WriteLine("IDC_" & ctlData.id.ToString())
            Console.WriteLine(ctlData.x.ToString() & "," & ctlData.y.ToString() & "," & ctlData.cx.ToString() & "," & ctlData.cy.ToString())
            Dim Style = If(isDialogEx, ctlData.ExStyle, ctlData.Style)
            Dim ExStyle = If(isDialogEx, ctlData.Style, ctlData.ExStyle)
            Dim bClassID() As Byte = bItemData.Skip(size).ToArray() 'move to id in struct
            If isDialogEx Then
                bClassID = bClassID.Skip(2).ToArray()
            End If
            Dim NotStyle As UInteger = 0
#Region "ControlStyle"
            Dim bCtlName() As Byte = Nothing
            If BitConverter.ToInt16(bClassID.Take(2).ToArray(), 0) = -1 Then
                Select Case bClassID(2)
                    Case &H80
                        Select Case Style And &HF
                            Case CUInt(Math.Truncate(WindowStyles.BS_PUSHBUTTON))
                                Console.WriteLine("PUSHBUTTON")
                                NotStyle = CUInt(WindowStyles.BS_PUSHBUTTON Or WindowStyles.WS_TABSTOP)
                            Case CUInt(Math.Truncate(WindowStyles.BS_DEFPUSHBUTTON))
                                Console.WriteLine("DEFPUSHBUTTON")
                                NotStyle = CUInt(WindowStyles.BS_DEFPUSHBUTTON Or WindowStyles.WS_TABSTOP)
                            Case CUInt(Math.Truncate(WindowStyles.BS_CHECKBOX))
                                Console.WriteLine("CHECKBOX")
                                NotStyle = CUInt(WindowStyles.BS_CHECKBOX Or WindowStyles.WS_TABSTOP)
                            Case CUInt(Math.Truncate(WindowStyles.BS_AUTOCHECKBOX))
                                Console.WriteLine("AUTOCHECKBOX")
                                NotStyle = CUInt(Math.Truncate(WindowStyles.BS_AUTOCHECKBOX))
                            Case CUInt(Math.Truncate(WindowStyles.BS_RADIOBUTTON))
                                Console.WriteLine("RADIOBUTTON")
                                NotStyle = CUInt(Math.Truncate(WindowStyles.BS_RADIOBUTTON))
                            Case CUInt(Math.Truncate(WindowStyles.BS_3STATE))
                                Console.WriteLine("STATE3" & vbTab)
                                NotStyle = CUInt(Math.Truncate(WindowStyles.BS_3STATE))
                            Case CUInt(Math.Truncate(WindowStyles.BS_AUTO3STATE))
                                Console.WriteLine("AUTO3STATE")
                                NotStyle = CUInt(Math.Truncate(WindowStyles.BS_AUTO3STATE))
                            Case CUInt(Math.Truncate(WindowStyles.BS_GROUPBOX))
                                Console.WriteLine("GROUPBOX")
                                NotStyle = CUInt(Math.Truncate(WindowStyles.BS_GROUPBOX))
                            Case CUInt(Math.Truncate(WindowStyles.BS_AUTORADIOBUTTON))
                                Console.WriteLine("AUTORADIOBUTTON")
                                NotStyle = CUInt(Math.Truncate(WindowStyles.BS_AUTORADIOBUTTON))
                            Case Else
                                Console.WriteLine("CONTROL" & vbTab)
                                NotStyle = 0
                                isControl = 1
                        End Select
                    Case &H81
                        Console.WriteLine("EDITTEXT")
                        NotStyle = CUInt(WindowStyles.ES_LEFT Or WindowStyles.WS_BORDER Or WindowStyles.WS_TABSTOP)
                    Case &H82
                        Select Case Style And &HF
                            Case CUInt(Math.Truncate(WindowStyles.SS_LEFT))
                                Console.WriteLine("LTEXT" & vbTab)
                                NotStyle = CUInt(WindowStyles.SS_LEFT Or WindowStyles.WS_GROUP)
                            Case CUInt(Math.Truncate(WindowStyles.SS_RIGHT))
                                Console.WriteLine("RTEXT" & vbTab)
                                NotStyle = CUInt(WindowStyles.SS_RIGHT Or WindowStyles.WS_GROUP)
                            Case CUInt(Math.Truncate(WindowStyles.SS_CENTER))
                                Console.WriteLine("CTEXT" & vbTab)
                                NotStyle = CUInt(WindowStyles.SS_CENTER Or WindowStyles.WS_GROUP)
                            Case CUInt(Math.Truncate(WindowStyles.SS_ICON))
                                Console.WriteLine("ICON" & vbTab)
                                NotStyle = CUInt(Math.Truncate(WindowStyles.SS_ICON))
                            Case Else
                                Console.WriteLine("CONTROL" & vbTab)
                                NotStyle = 0
                                isControl = 2
                        End Select
                    Case &H83
                        Console.WriteLine("LISTBOX" & vbTab)
                        NotStyle = CUInt(WindowStyles.WS_BORDER Or WindowStyles.LBS_NOTIFY)
                    Case &H84
                        Console.WriteLine("SCROLLBAR")
                        NotStyle = 0
                    Case &H85
                        Console.WriteLine("COMBOBOX")
                        NotStyle = 0
                    Case Else
                        Console.WriteLine("CONTROL" & vbTab)
                        NotStyle = 0
                        isControl = -2
                End Select
                bCtlName = bClassID.Skip(4).ToArray()
            Else
                Dim len = GetUnicodeStringLength(bClassID)
                Dim szCtlName = Encoding.Unicode.GetString(bClassID, 0, len)
                Console.WriteLine("CtlName:" & szCtlName)
                bCtlName = bClassID.Skip(len + 2).ToArray()
            End If
#End Region
            NotStyle = NotStyle Or CUInt(Math.Truncate(WindowStyles.WS_CHILD)) Or CUInt(Math.Truncate(WindowStyles.WS_VISIBLE))
            Style = Style And Not NotStyle
            NotStyle = NotStyle And Not (If(isDialogEx, ctlData.ExStyle, ctlData.Style))
#Region "ControlName"
            Dim bCtlId() As Byte = Nothing
            If BitConverter.ToInt16(bCtlName.Take(2).ToArray(), 0) = -1 Then
                bCtlId = bCtlName.Skip(6).ToArray()
            Else
                If isControl = 0 AndAlso (bClassID(1) = &H81 OrElse bClassID(1) = &H83 OrElse bClassID(1) = &H84 OrElse bClassID(1) = &H85) Then
                    bCtlId = bCtlName.Skip(2).ToArray()
                Else
                    Dim len = GetUnicodeStringLength(bCtlName)
                    Dim szCtlId = Encoding.Unicode.GetString(bCtlName, 0, len)
                    Console.WriteLine("CtlId:" & szCtlId)
                    If len > 2 Then
                        bCtlId = bCtlName.Skip(len + 4).ToArray()
                    Else
                        bCtlId = bCtlName.Skip(2).ToArray()
                    End If
                End If
            End If
#End Region

            Dim pStyle() As String = {"", "button", "static"}
            If isControl <> 0 Then
                If isControl = -1 Then
                    Console.WriteLine(BitConverter.ToInt16(bCtlName.Take(2).ToArray(), 0).ToString())
                ElseIf isControl > 0 Then
                    Console.WriteLine(pStyle(isControl))
                Else
                    Console.WriteLine(bClassID(0).ToString())
                End If
                If Convert.ToInt32(Style) <> 0 OrElse NotStyle = 0 Then
                    Console.WriteLine("0x" & Style.ToString("x8"))
                End If
                If NotStyle <> 0 Then
                    Console.WriteLine("0x" & NotStyle.ToString("x8"))
                End If
                NotStyle = 0
                Style = NotStyle
            End If

            If Style <> 0 OrElse NotStyle <> 0 OrElse ExStyle <> 0 OrElse (isDialogEx AndAlso helpID <> 0) Then
                If Style <> 0 OrElse ((Not NotStyle) <> 0 AndAlso (Not isControl) <> 0) Then
                    Console.WriteLine("0x" & Style.ToString("x8"))
                End If
                If NotStyle <> 0 Then
                    If Style <> 0 Then
                        Console.WriteLine(" |0x" & NotStyle.ToString("x8"))
                    Else
                        Console.WriteLine(",0x" & NotStyle.ToString("x8"))
                    End If
                End If
                If ExStyle <> 0 OrElse (isDialogEx AndAlso helpID <> 0) Then
                    If ExStyle = CUInt(Math.Truncate(WindowExStyles.WS_EX_STATICEDGE)) Then
                        Console.WriteLine(", WS_EX_STATICEDGE")
                    Else
                        Console.WriteLine("0x" & ExStyle.ToString("x8"))
                    End If
                    If isDialogEx AndAlso helpID <> 0 Then
                        Console.WriteLine(ctlData.id.ToString())
                    End If
                End If
            End If
            bItemData = bCtlId.Skip(2).ToArray()
            If isDialogEx Then
                bItemData = bItemData.Skip(2).ToArray()
            End If
        Next i
#End Region
        Return bytesIn
    End Function

    Private Shared Function GetUnicodeStringLength(ByVal byteIn() As Byte) As Integer
        Dim nullterm As Integer = 0
        Do While nullterm < byteIn.Length AndAlso byteIn(nullterm) <> 0
            nullterm += 2
        Loop
        Return nullterm
    End Function

End Class
