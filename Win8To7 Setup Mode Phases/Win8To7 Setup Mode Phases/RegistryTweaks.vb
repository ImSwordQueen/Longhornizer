Public Class RegistryTweaks
    ' ID: {DWORD: {Key, Value, Data}, String: {Key, Value, Data}, Delete: {Key, Value, Data}}
    Public SystemTweaks As New Dictionary(Of String, List(Of List(Of String)))
    Public UserTweaks As New Dictionary(Of String, List(Of List(Of String)))
    Public UserClassTweaks As New Dictionary(Of String, List(Of List(Of String)))
    Private sibdir As String
    Private openshelldir As String


    Public Sub init()
        'Deal with misc. paths
        openshelldir = windrive + "Program Files\Classic Shell"

        PopulateSystemTweaks()
        PopulateUserTweaks()
    End Sub


    Public Sub PopulateSystemTweaks() 'Here is where you can dump the keys to change on a system-wide basis
        Dim itemstodelete As New List(Of List(Of String))
        Dim dwordchanges As New List(Of List(Of String))
        Dim stringchanges As New List(Of List(Of String))
        Dim multistringchanges As New List(Of List(Of String))
        Dim hexchanges As New List(Of List(Of String))

        SystemTweaks.Add("DWORD", dwordchanges)
        SystemTweaks.Add("String", stringchanges)
        SystemTweaks.Add("MultiString", multistringchanges)
        SystemTweaks.Add("Binary", hexchanges)
        SystemTweaks.Add("Delete", itemstodelete)
    End Sub

    Sub PopulateUserTweaks() 'Here is where you can dump the keys to change on a user-wide basis
        Dim itemstodelete As New List(Of List(Of String))
        Dim dwordchanges As New List(Of List(Of String))
        Dim stringchanges As New List(Of List(Of String))
        Dim multistringchanges As New List(Of List(Of String))
        Dim hexchanges As New List(Of List(Of String))
        Dim itemstodeleteclass As New List(Of List(Of String))
        Dim dwordchangesclass As New List(Of List(Of String))
        Dim stringchangesclass As New List(Of List(Of String))
        Dim multistringchangesclass As New List(Of List(Of String))
        Dim hexchangesclass As New List(Of List(Of String))

        'Configure Classic Explorer
        dwordchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicExplorer", "ShowedToolbar", "1"))
        stringchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicExplorer\", "TreeStyle", "XPSimple"))
        dwordchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicExplorer", "ShowCaption", "0"))
        dwordchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicExplorer", "ShowUpButton", "DontShow"))
        stringchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicExplorer\", "TreeStyle", "XPSimple"))
        dwordchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicExplorer", "ShowFreeSpace", "1"))
        dwordchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicExplorer", "ShareOverlay", "0"))
        dwordchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicExplorer", "EnableMore", "1"))
        dwordchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicExplorer", "FileExplorer", "0"))

        'Configure ClassicShell
        stringchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicStartMenu", "WinKey", "WindowsMenu"))
        stringchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicStartMenu", "MouseClick", "WindowsMenu"))
        dwordchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicStartMenu", "EnableExit", "1"))
        dwordchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicStartMenu", "EnableExplorer", "1"))
        dwordchanges.Add(ToList("HKCU\Software\IvoSoft\ClassicStartMenu", "HideUserPic", "1"))

        UserTweaks.Add("DWORD", dwordchanges)
        UserTweaks.Add("String", stringchanges)
        UserTweaks.Add("MultiString", multistringchanges)
        UserTweaks.Add("Binary", hexchanges)
        UserTweaks.Add("Delete", itemstodelete)
        UserClassTweaks.Add("DWORD", dwordchangesclass)
        UserClassTweaks.Add("String", stringchangesclass)
        UserClassTweaks.Add("MultiString", multistringchangesclass)
        UserClassTweaks.Add("Binary", hexchangesclass)
        UserClassTweaks.Add("Delete", itemstodeleteclass)
    End Sub





    Public Sub SetSZ(ByVal Key As String, ByVal Value As String, ByVal Data As String, Optional ByVal wait As Boolean = True)
        Dim newKey As String
        Dim newValue As String
        Dim newData As String
        If Key.Contains(" ") Then
            newKey = """" + Key + """"
        Else
            newKey = Key
        End If
        If Value.Contains(" ") Then
            newValue = """" + Value + """"
        Else
            newValue = Value
        End If
        If Data.Contains(" ") Then
            newData = """" + Data + """"
        Else
            newData = Data
        End If
        Dim slashd As String
        If Not String.IsNullOrEmpty(Data) And Not Data = "" Then
            slashd = "/d " + newData
        Else
            slashd = ""
        End If
        Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD " + newKey + " /v " + newValue + " /t REG_SZ " + slashd + " /f", AppWinStyle.Hide, wait)
    End Sub

    Public Sub SetEXPANDSZ(ByVal Key As String, ByVal Value As String, ByVal Data As String, Optional ByVal wait As Boolean = True)
        Dim newKey As String
        Dim newValue As String
        Dim newData As String
        If Key.Contains(" ") Then
            newKey = """" + Key + """"
        Else
            newKey = Key
        End If
        If Value.Contains(" ") Then
            newValue = """" + Value + """"
        Else
            newValue = Value
        End If
        If Data.Contains(" ") Then
            newData = """" + Data + """"
        Else
            newData = Data
        End If
        Dim slashd As String
        If Not String.IsNullOrEmpty(Data) And Not Data = "" Then
            slashd = "/d " + newData
        Else
            slashd = ""
        End If
        Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD " + newKey + " /v " + newValue + " /t REG_EXPAND_SZ " + slashd + " /f", AppWinStyle.Hide, wait)
    End Sub

    Public Sub SetMultiSZ(ByVal Key As String, ByVal Value As String, ByVal Data As String, Optional ByVal wait As Boolean = True)
        Dim newKey As String
        Dim newValue As String
        Dim newData As String
        If Key.Contains(" ") Then
            newKey = """" + Key + """"
        Else
            newKey = Key
        End If
        If Value.Contains(" ") Then
            newValue = """" + Value + """"
        Else
            newValue = Value
        End If
        If Not Data.Contains("""") Then 'Inverted logic here to prevent Open Shell configs from failing, I guess
            newData = """" + Data + """"
        Else
            newData = Data
        End If
        Dim slashd As String
        If Not String.IsNullOrEmpty(Data) And Not Data = "" Then
            slashd = "/d " + newData
        Else
            slashd = ""
        End If
        Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD " + newKey + " /v " + newValue + " /t REG_MULTI_SZ " + slashd + " /f", AppWinStyle.Hide, wait)
    End Sub

    Public Function ConvertToBinHex(ByVal Data As String)
        'Convert data to hex:
        Dim result1 As String = ""
        Dim charcount As Integer = 0
        For Each c As Char In Data
            result1 = result1 + c
            charcount += 1
            If charcount = 2 Then
                result1 = result1 + ","
                charcount = 0
            End If
        Next
        If result1.EndsWith(",") Then
            result1 = result1.Remove(result1.Length - 1, 1)
        End If

        Dim result2() As Byte = Array.ConvertAll(result1.Split(","), Function(b) Convert.ToByte(b, 16))
        Return result2
    End Function

    Public Sub SetBinary(ByVal Key As String, ByVal Value As String, ByVal Data As String, Optional ByVal wait As Boolean = True)
        Dim newKey As String
        Dim newValue As String
        Dim newData As String
        If Key.Contains(" ") Then
            newKey = """" + Key + """"
        Else
            newKey = Key
        End If
        If Value.Contains(" ") Then
            newValue = """" + Value + """"
        Else
            newValue = Value
        End If
        If Data.Contains(" ") Then
            newData = """" + Data + """"
        Else
            newData = Data
        End If
        Dim slashd As String
        If Not String.IsNullOrEmpty(Data) And Not Data = "" Then
            slashd = "/d " + newData
        Else
            slashd = ""
        End If
        Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD " + newKey + " /v " + newValue + " /t REG_BINARY " + slashd + " /f", AppWinStyle.Hide, wait)
    End Sub

    Public Function ConvertToDWORD(ByVal Data As String)
        'Convert data to dword:
        Dim result As String = ""
        Dim charcount As Integer = Hex(Data).ToString().Length()
        While result.Length() < (8 - charcount)
            result = result + "0"
        End While
        result = result + Hex(Data).ToString()
        Return result
    End Function

    Public Sub SetDWORD(ByVal Key As String, ByVal Value As String, ByVal Data As String, Optional ByVal wait As Boolean = True)
        Dim newKey As String
        Dim newValue As String
        Dim newData As String
        If Key.Contains(" ") Then
            newKey = """" + Key + """"
        Else
            newKey = Key
        End If
        If Value.Contains(" ") Then
            newValue = """" + Value + """"
        Else
            newValue = Value
        End If
        If Data.Contains(" ") Then
            newData = """" + Data + """"
        Else
            newData = Data
        End If
        Dim slashd As String
        If Not String.IsNullOrEmpty(Data) And Not Data = "" Then
            slashd = "/d " + newData
        Else
            slashd = ""
        End If
        Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD " + newKey + " /v " + newValue + " /t REG_DWORD " + slashd + " /f", AppWinStyle.Hide, wait)
    End Sub

    Public Sub SetDefaultStr(ByVal Key As String, ByVal Data As String, Optional ByVal wait As Boolean = True)
        Dim newKey As String
        Dim newData As String
        If Key.Contains(" ") Then
            newKey = """" + Key + """"
        Else
            newKey = Key
        End If
        If Data.Contains(" ") Then
            newData = """" + Data + """"
        Else
            newData = Data
        End If
        Dim slashd As String
        If Not String.IsNullOrEmpty(Data) And Not Data = "" Then
            slashd = "/d " + newData
        Else
            slashd = ""
        End If
        Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD " + newKey + " /ve " + slashd + " /f", AppWinStyle.Hide, wait)
    End Sub

    Public Sub DeleteKey(ByVal Key As String, ByVal Value As String, Optional ByVal wait As Boolean = True)
        Dim newKey As String
        Dim newValue As String
        If Key.Contains(" ") Then
            newKey = """" + Key + """"
        Else
            newKey = Key
        End If
        If Value.Contains(" ") Then
            newValue = """" + Value + """"
        Else
            newValue = Value
        End If

        Shell(windir + "\" + sysprefix + "\cmd.exe /c reg DELETE " + newKey + " /v " + newValue + " /f", AppWinStyle.Hide, wait)
    End Sub

    Function translate(ByVal input As String)
        Return input.Replace("%WINDIR%", windir).Replace("%OPENSHELL%", openshelldir)
    End Function


    Function ToList(ByVal Key As String, ByVal Value As String, ByVal Data As String)
        Dim inputs() As String = {Key, Value, Data}
        Dim result As List(Of String) = New List(Of String)(inputs)

        Return result
    End Function
End Class
