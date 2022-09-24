Attribute VB_Name = "ModulUtama"
Option Explicit

' *** Modul Utama ***

Public NamaUser As String
Public WinDir As String
Public UserProfile As String
Public FolderVirus As String
Public SeluruhInfo As String

' *** Fungsi Umum ***

Public Function Nama_Aplikasi() As String

    Nama_Aplikasi = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision

End Function

Function FileExist(strPath As String) As Integer

    Dim lngRetVal As Long
    
    On Error Resume Next
    lngRetVal = Len(Dir$(strPath))
    
    If Err Or lngRetVal = 0 Then
        FileExist = False
    Else
        FileExist = True
    End If
    
End Function

Public Sub WriteToFile(ByVal strFilename As String, ByVal strFileContents As String)

    Dim lngFileHandle As Long
    
    On Error Resume Next
    
    lngFileHandle = FreeFile
    
    Open strFilename For Append As #lngFileHandle
    Print #lngFileHandle, strFileContents
    Close #lngFileHandle
    
End Sub

Public Sub Bunuh(ByVal Virusnya As String)

    Dim Bersih As Boolean

    Bersih = False

    Informasi "  » Menghapus " & Virusnya
    
    On Error Resume Next
    
    SetAttr Virusnya, vbNormal
    
    If FileExist(Virusnya) Then
        
        If InStr(Virusnya, "*") <> 0 Then
            Kill Virusnya
            Bersih = True
        End If
       
        If Err.Number <> 0 Then
            If Err.Number = 5 And Err.Number <> 52 Then
                Kill Virusnya
                Bersih = True
            Else
                Informasi "     » " & Err.Number & " : " & Err.Description
                Informasi "     » File tidak dapat dihapus! Coba cek manual file tsb!"
            End If
        Else
            Kill Virusnya
            Bersih = True
        End If
    Else
        Informasi "     » File tidak ada!"
    End If
    
    If Bersih = True Then
        Informasi "     » Virus telah dihapus!"
    End If
        
End Sub

Public Sub HapusReg(SubKey As String, Entry As String)

    Informasi "      » Menghapus Entry Registry " & SubKey & "\" & Entry
    DeleteValue SubKey, Entry
    
End Sub

Public Sub UbahRegDWORD(SubKey As String, Entry As String, Nilai As Long)

    'Informasi "      » Mengubah Entry Registry " & SubKey & "\" & Entry
    SetDWORDValue SubKey, Entry, Nilai
    
End Sub

Public Sub UbahRegString(SubKey As String, Entry As String, Nilai As String)

    'Informasi "      » Mengubah Entry Registry " & Subkey & "\" & Entry
    SetStringValue SubKey, Entry, Nilai
    
End Sub

Public Sub Informasi(ByVal Infonya As String)
    
    SeluruhInfo = SeluruhInfo & Infonya & vbCrLf
    FDestroyer.Info.Text = SeluruhInfo
    FDestroyer.Info.SelStart = Len(SeluruhInfo)
    FDestroyer.Info.Refresh
    WriteToFile NamaLogFile, SeluruhInfo
    
End Sub

Public Sub Hajar_Virus()

    Dim Status As Boolean
    Dim JmlVirMem As Integer
    Dim i As Integer

    NamaUser = Environ("UserName")
    WinDir = Environ("Windir")
    UserProfile = Environ("UserProfile")
    FolderVirus = UserProfile & AppData

    Status = False
    Informasi ""
    SeluruhInfo = ""
    JmlVirMem = 0
    
    ' *** Mulai ***
    ' *** Penghapusan Virus dari Memory ***
    Informasi "*** " & Now & " ***"
    Informasi ""
    Informasi "Pencarian dan Penghapusan Virus di Memory..."
    
    DoEvents
    Do While FindWindow(vbNullString, NamaProsesVirus) <> 0
        JmlVirMem = JmlVirMem + 1
        WindowHandle FindWindow(vbNullString, NamaProsesVirus), 0
        Status = True
    Loop
    
    If Status = True Then
        Informasi "  » " & JmlVirMem & " Virus " & NamaVirus & " ditemukan di memory dan telah dibersihkan (Get Out, Bro!)"
        MsgBox JmlVirMem & " Virus " & NamaVirus & " ditemukan di memory dan telah dibersihkan (Get Out, Bro!)", vbInformation
    Else
        Informasi "  » Virus " & NamaVirus & " tidak ada di memory"
        MsgBox "Virus " & NamaVirus & " tidak ada di memory", vbInformation
    End If
    
    ' *** Perbaikan Registry yang dimodifikasi oleh Virus ***
    
    Informasi ""
    Informasi "Perbaikan Registry..."
    Informasi "  » Mengembalikan menu Folder Option pada Explorer..."
    UbahRegDWORD ExPol, "NoFolderOptions", 0
    Informasi "  » Mengaktifkan kembali Registry Editor..."
    UbahRegDWORD SysPol, "DisableRegistryTools", 0
    Informasi "  » Menyembuhkan Shell Windows..."
    UbahRegString WinLogon, "Shell", "Explorer.exe"
    Informasi "  » Menghapus Shell Alternatif Virus pada Safe Mode... "
    HapusReg CtrSet1, "AlternateShell"
    HapusReg CtrSet2, "AlternateShell"
    Informasi "  » Menyembuhkan Hidden Folder Settings..."
    UbahRegDWORD HideFileExt, "CheckedValue", 1
    UbahRegDWORD HideFileExt, "UnCheckedValue", 0
    UbahRegDWORD HideFileExt, "DefaultValue", 1
    Informasi "  » Menyembuhkan SuperHidden Folder Settings..."
    UbahRegDWORD SuperHidden, "CheckedValue", 0
    UbahRegDWORD SuperHidden, "UnCheckedValue", 1
    UbahRegDWORD SuperHidden, "DefaultValue", 0
    
    ' *** Pembunuhan Virus ***
    Informasi ""
    Informasi "Pembersihan Master Virus..."
    Bunuh WinDir & "\System32\Explorer.exe"
    
    ' *** Penghapusan Virus Alternatif ***
    For i = 0 To FDestroyer.Drive1.ListCount - 1
        DoEvents
        Bunuh Left(FDestroyer.Drive1.List(i), 2) & "\Thumbs.db.com"
    Next
    
    Informasi ""
    Informasi "*** Selesai ***"
    Informasi ""
    Informasi "*** " & Now & " ***"
    Informasi ""
    
    MsgBox "Pembersihan Master Virus " & NamaVirus & " telah selesai." _
        , vbInformation
    
    ' *** Selesai ***
    
End Sub
