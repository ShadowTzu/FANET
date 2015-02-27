Imports FA
Imports System.IO

Module Module1

    Private mFileArchive As FileArchive
    Private Folder As String = Path.GetFullPath(My.Application.Info.DirectoryPath & "\..\..\..\test\Pack_Me")
    Private Folder_Root As String = Path.GetFullPath(Folder & "\..\")

    Sub Main()
        'Clean your previous test
        Clean()

        mFileArchive = New FileArchive(Folder)

        Console.WriteLine("Add files:")
        Dim cFile As String
        For Each files As String In My.Computer.FileSystem.GetFiles(Folder, FileIO.SearchOption.SearchAllSubDirectories)
            cFile = files.Replace(Folder & "\", "")

            Console.WriteLine("    " & cFile)
            mFileArchive.Add_Files(cFile)
        Next

        Console.Write("Save Archive...")
        mFileArchive.Save(Folder_Root & "Packed.tzu")
        Console.WriteLine("Saved!")

        Console.WriteLine(vbCrLf & "Press a key to continue..." & vbCrLf)
        Console.ReadKey(True)

        Console.Write("Extracting...")
        mFileArchive.Extract(Folder_Root & "Packed.tzu", Folder_Root & "Extracted\")
        Console.WriteLine("extracted!")

        Console.WriteLine(vbCrLf & "Press a key to continue..." & vbCrLf)
        Console.ReadKey(True)

    End Sub

    'Clean destination folder and archive just for your test
    Private Sub Clean()
        Console.WriteLine("Cleaning....")

        If File.Exists(Folder_Root & "Packed.tzu") Then
            Console.WriteLine("Erase     Packed.tzu'")
            File.Delete(Folder_Root & "Packed.tzu")
        End If
        If Directory.Exists(Folder_Root & "Extracted") Then
            For Each files As String In My.Computer.FileSystem.GetFiles(Folder_Root & "Extracted", FileIO.SearchOption.SearchAllSubDirectories)
                Console.WriteLine("Erase     " & files.Replace(Folder_Root, ""))
                File.Delete(files.Replace(Folder & "\", ""))
            Next
        End If
       
        Console.WriteLine("Cleaned!")
        Console.WriteLine(vbCrLf & "Press a key to continue..." & vbCrLf)
        Console.ReadKey(True)
    End Sub
End Module
