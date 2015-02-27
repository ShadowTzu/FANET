Imports System.IO
Imports System.Drawing

''' <summary>
''' Archive multiple files into one file
''' </summary>
''' <remarks></remarks>
Public Class FileArchive
    Implements IDisposable

    'File Structure:
    '-----------------------
    'Magic number (232) : 1 Byte 
    'File count (Uint32) : 4 Bytes

    'Table info : variable size

    'Current position: Uint32

    'Files pasted end to end
    '-----------------------

    Private Magic As Byte = 232
    Private Count_Files As UInt32
    Private Structure Str_TabFiles
        Public Filename As String
        Public Length As UInt32
    End Structure

    Private Tab_Files As List(Of Str_TabFiles)

    Private Tab_FilesExt() As Str_TabFiles

    Private Directory_Path As String

    Private Offset_File As UInt32
#Region "New"
    Public Sub New(ByVal Directory_File As String)
        Me.Directory_Path = Directory_File
        If Not Me.Directory_Path.EndsWith("\") Then
            Me.Directory_Path &= "\"
        End If
        Tab_Files = New List(Of Str_TabFiles)
    End Sub
#End Region

#Region "Add files"
    Public Sub Add_Files(ByVal Filename As String)
        'Check if file exist
        If File.Exists(Directory_Path & Filename) = False Then
            Throw New Exception("File: '" & Filename & "' not found!")
            Exit Sub
        End If

        'stock filename and size
        Dim myFileInfo As New FileInfo(Directory_Path & Filename)
        Dim data_file As Str_TabFiles
        With data_file
            .Filename = Filename
            .Length = CUInt(myFileInfo.Length)
        End With
        Tab_Files.Add(data_file)

        myFileInfo = Nothing
    End Sub
#End Region

#Region "Save Archive"

    ''' <summary>
    ''' Save archive file.
    ''' </summary>
    ''' <param name="Archive_Filename">Location for your archive file</param>
    ''' <remarks></remarks>
    Public Sub Save(ByVal Archive_Filename As String)
        Dim OutFile As BinaryWriter = New BinaryWriter(File.Create(Archive_Filename))

        'Writing magic number
        OutFile.Write(Magic)

        'Writing files count
        Count_Files = CUInt(Tab_Files.Count)
        OutFile.Write(Count_Files)

        'Writing Table
        For i As Integer = 0 To Tab_Files.Count - 1
            OutFile.Write(Tab_Files(i).Filename)
            OutFile.Write(Tab_Files(i).Length)
        Next

        'Writing the current position to determine the position of files
        OutFile.Write(CType(OutFile.Seek(0, SeekOrigin.Current), UInt32))

        'Writing files
        For i As Integer = 0 To Tab_Files.Count - 1
            Dim myfileAct As BinaryReader = New BinaryReader(File.Open(Directory_Path & "\" & Tab_Files(i).Filename, FileMode.Open))
            OutFile.Write(myfileAct.ReadBytes(CInt(Tab_Files(i).Length)))
            myfileAct.Close()
            myfileAct = Nothing
        Next

        OutFile.Close()
        OutFile = Nothing

    End Sub
#End Region

#Region "Extract Archive"
    ''' <summary>
    ''' Extract the archive into a folder
    ''' </summary>
    ''' <param name="Archive_Filename">Location of your archive file</param>
    ''' <param name="Destination">Destination folder</param>
    ''' <remarks></remarks>
    Public Sub Extract(ByVal Archive_Filename As String, ByVal Destination As String)
        If File.Exists(Archive_Filename) = False Then Exit Sub
        Dim Infile As BinaryReader = New BinaryReader(File.Open(Archive_Filename, FileMode.Open))

        If Infile.ReadByte <> Magic Then
            Infile.Close()
            MsgBox("Fichier invalide!")
            Infile = Nothing
            Exit Sub
        End If
        If Not Destination.EndsWith("\") Then
            Destination &= "\"
        End If

        'Create destination directory if it doesn't exist
        If Not Directory.Exists(Destination) Then
            Directory.CreateDirectory(Destination)
        End If
        'Get files count
        Dim NbrFichier As UInt32 = Infile.ReadUInt32

        'Get the name and file size
        ReDim Tab_FilesExt(CInt(NbrFichier - 1))
        For i As Integer = 0 To CInt(NbrFichier - 1)
            Tab_FilesExt(i).Filename = Infile.ReadString()
            Tab_FilesExt(i).Length = Infile.ReadUInt32
        Next

        'Get current position       
        Offset_File = Infile.ReadUInt32

        For i As Integer = 0 To CInt(NbrFichier - 1)
            Dim buffertps(CInt(Tab_FilesExt(i).Length)) As Byte
            buffertps = Infile.ReadBytes(CInt(Tab_FilesExt(i).Length))

            'if the file is in a directory then create it if it does not exist
            If Tab_FilesExt(i).Filename.IndexOf("\", 1) > 0 Then
                Dim dir_test As String = Tab_FilesExt(i).Filename.Substring(0, Tab_FilesExt(i).Filename.IndexOf("\", 1))
                If IO.Directory.Exists(Destination & dir_test) = False Then
                    IO.Directory.CreateDirectory(Destination & dir_test)
                End If
            End If

            'Writing files
            Dim sfile As BinaryWriter = New BinaryWriter(File.Create(Destination & Tab_FilesExt(i).Filename))
            sfile.Write(buffertps)
            sfile.Close()
            sfile = Nothing
        Next

        Infile.Close()
        Infile = Nothing

    End Sub

    ''' <summary>
    ''' Extract only one file from archive
    ''' </summary>
    ''' <param name="Archive_Filename">Location of your archive file</param>
    ''' <param name="Filename">File to extract</param>
    ''' <param name="Destination">Destination folder</param>
    ''' <remarks></remarks>
    Public Sub Extract(ByVal Archive_Filename As String, ByVal Filename As String, ByVal Destination As String)
        If File.Exists(Archive_Filename) = False Then Exit Sub
        Dim Index_file As Integer

        Dim Infile As BinaryReader = New BinaryReader(File.Open(Archive_Filename, FileMode.Open))

        If Infile.ReadByte <> Magic Then
            Infile.Close()
            MsgBox("Fichier invalide!")
            Infile = Nothing
            Exit Sub
        End If
        If Not Destination.EndsWith("\") Then
            Destination &= "\"
        End If

        Dim NbrFichier As UInt32 = Infile.ReadUInt32
        ReDim Tab_FilesExt(CInt(NbrFichier - 1))
        For i As Integer = 0 To CInt(NbrFichier - 1)
            Tab_FilesExt(i).Filename = Infile.ReadString()
            Tab_FilesExt(i).Length = Infile.ReadUInt32
            If Tab_FilesExt(i).Filename.ToLower = Filename Then Index_file = i
        Next

        'Get current position
        Offset_File = Infile.ReadUInt32

        'Shifts the read start index for the file
        Dim IndexCumul As Integer = CInt(Offset_File + 4)
        If Index_file > 0 Then
            For i As Integer = 1 To Index_file
                IndexCumul += CInt(Tab_FilesExt(i - 1).Length)
            Next
        End If
        Infile.Close()
        Infile = Nothing



        'if the file is in a directory then create it if it does not exist
        If Tab_FilesExt(Index_file).Filename.IndexOf("\", 1) > 0 Then
            Dim dir_test As String = Tab_FilesExt(Index_file).Filename.Substring(0, Tab_FilesExt(Index_file).Filename.IndexOf("\", 1))
            If IO.Directory.Exists(Destination & dir_test) = False Then
                IO.Directory.CreateDirectory(Destination & dir_test)
            End If
        End If

        'TODO: do not save buffer in memory, use read/write method every x bytes
        'Filled the buffer with data from the desired file
        Dim infile2 As BinaryReader = New BinaryReader(File.Open(Archive_Filename, FileMode.Open))
        infile2.BaseStream.Seek(IndexCumul, SeekOrigin.Begin)
        Dim buffertps(CInt(Tab_FilesExt(Index_file).Length - 1)) As Byte
        infile2.Read(buffertps, 0, buffertps.Length)

        'Writing the file in the destination folder
        Dim sfile As BinaryWriter = New BinaryWriter(File.Create(Destination & Tab_FilesExt(Index_file).Filename))
        sfile.Write(buffertps)
        sfile.Close()
        sfile = Nothing
        infile2.Close()
        infile2 = Nothing

    End Sub
#End Region

#Region "Get Files"

    ''' <summary>
    ''' recovers the contents of a text file in the archive
    ''' </summary>
    ''' <param name="Archive_Filename">Location of your archive file</param>
    ''' <param name="Filename">Text file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Get_StringFiles(ByVal Archive_Filename As String, ByVal Filename As String) As String
        Dim streamR As StreamReader = New StreamReader(Get_StreamFiles(Archive_Filename, Filename))
        Return streamR.ReadToEnd
    End Function

    ''' <summary>
    ''' recovers an image in the archive
    ''' </summary>
    ''' <param name="Archive_Filename">Location of your archive file</param>
    ''' <param name="Filename">Image file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Get_ImageFiles(ByVal Archive_Filename As String, ByVal Filename As String) As Image
        Return Image.FromStream(Get_StreamFiles(Archive_Filename, Filename))
    End Function


    ''' <summary>
    ''' recovers a stream from a file in the archive
    ''' </summary>
    ''' <param name="Archive_Filename">Location of your archive file</param>
    ''' <param name="Filename"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Get_StreamFiles(ByVal Archive_Filename As String, ByVal Filename As String) As Stream
        Dim memstr As MemoryStream = New MemoryStream(Get_Files(Archive_Filename, Filename))
        Return memstr
    End Function

    ''' <summary>
    ''' recovers a byte array from a file in the archive
    ''' </summary>
    ''' <param name="Archive_Filename">Location of your archive file</param>
    ''' <param name="Filename"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Get_Files(ByVal Archive_Filename As String, ByVal Filename As String) As Byte()
        If File.Exists(Archive_Filename) = False Then Return Nothing
        Dim Index_file As Integer

        Dim Infile As BinaryReader = New BinaryReader(File.Open(Archive_Filename, FileMode.Open))

        If Infile.ReadByte <> Magic Then
            Infile.Close()
            MsgBox("Fichier invalide!")
            Infile = Nothing
            Return Nothing
        End If
        Dim NbrFichier As UInt32 = Infile.ReadUInt32
        ReDim Tab_FilesExt(CInt(NbrFichier - 1))
        For i As Integer = 0 To CInt(NbrFichier - 1)
            Tab_FilesExt(i).Filename = Infile.ReadString()
            Tab_FilesExt(i).Length = Infile.ReadUInt32

            'recovers the index of the desired file
            If Tab_FilesExt(i).Filename.ToLower = Filename Then Index_file = i
        Next

        'recovers the current position
        Offset_File = Infile.ReadUInt32

        'Shifts the read start index for the file
        Dim IndexCumul As Integer = CInt(Offset_File + 4)
        If Index_file > 0 Then
            For i As Integer = 1 To Index_file
                IndexCumul += CInt(Tab_FilesExt(i - 1).Length)
            Next
        End If
        Infile.Close()
        Infile = Nothing

        'Filled the buffer with data from the desired file
        Dim infile2 As BinaryReader = New BinaryReader(File.Open(Archive_Filename, FileMode.Open))
        infile2.BaseStream.Seek(IndexCumul, SeekOrigin.Begin)
        Dim buffertps(CInt(Tab_FilesExt(Index_file).Length - 1)) As Byte
        infile2.Read(buffertps, 0, buffertps.Length)
        infile2.Close()
        infile2 = Nothing
        Return buffertps

    End Function
#End Region


#Region "Destructor"
    Private disposedValue As Boolean

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                Tab_Files.Clear()
                Tab_Files = Nothing
            End If
        End If
        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class