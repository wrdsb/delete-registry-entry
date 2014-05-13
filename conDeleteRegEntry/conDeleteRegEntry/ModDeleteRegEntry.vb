Imports Microsoft.Win32
Imports System
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Reflection

Module ModDeleteRegEntry

    Sub Main()
        Dim debug As Boolean = False
        Dim partialPath As Boolean = False
        Dim needHelp As Boolean = False
        Dim onlyOnce As Boolean = False
        'In some languages cmdArgs(0) is the name of the program being called. ie conDeleteRegEntry.exe
        'This does not appear to be the case in VB
        ' Zero byte log file is automatically created
        ' To test abort make orig_path.txt readonly
        Dim log As CLogWriter = New CLogWriter("C:\backup\conDeleteRegEntry.log", True)
        log.Print("------ New Log Entry -----", "------")
        log.Print("", "")

        Dim cmdA As String() = Environment.GetCommandLineArgs()
        Dim cmdArgs() As String = Split(cmdA(1), ",")
        Dim searchString As String = Chr(34)
        'There is a problem passing in quoted parameters hence this next step
        Dim idx As Integer = 0
        Dim temp_string As String
        Dim x As Integer
        Dim argMax As Integer = cmdArgs.Length - 1
        For idx = 0 To argMax
            temp_string = cmdArgs(idx)
            x = InStr(cmdArgs(idx), searchString, CompareMethod.Text)
            If x > 0 Then
                cmdArgs(idx) = Left(cmdArgs(idx), Len(cmdArgs(idx)) - 1)
            End If
        Next
        For idx = 0 To argMax
            If UCase(Trim(cmdArgs(idx))) = "/D" Then debug = True
            If UCase(Trim(cmdArgs(idx))) = "/P" Then partialPath = True
            If (UCase(Trim(cmdArgs(idx))) = "/?") Or (UCase(Trim(cmdArgs(idx))) = "?") Then
                needHelp = True
            End If
            If (UCase(Trim(cmdArgs(idx))) = "/O") And (Not partialPath) Then onlyOnce = True
        Next
        If debug Then log.Print(" Trimmed cmdArgs ", cmdArgs)
        If debug Then log.Print(" cmdA ", cmdA)
        If debug Then log.Print(" cmdArgs ", cmdArgs)
        If debug Then log.Print(" argMax ", argMax)
        Dim iValidArgs As Integer = 0
        Dim path_array() As String
        Dim dt As DateTime = DateTime.Now
        Dim date_modified As String = dt.ToString("yyyy-d-MM : hh:m:s t")
        If debug Then log.Print(" date_modified ", date_modified)
        For idx = 0 To argMax
            If Len(Trim(cmdArgs(idx))) > 0 Then
                iValidArgs += 1
            End If
        Next
        If needHelp Then
            Console.WriteLine(" ConDeleteRegEntry Parameters ")
            Console.WriteLine(" Enter all parameters within double quotes and separated by a comma ")
            Console.WriteLine(" /d means write a debug log ")
            Console.WriteLine(" /p means accept partial paths ")
            Console.WriteLine(" partial paths means that if the parameter is found anywhere in the path entry ")
            Console.WriteLine("         then the entire path entry will be deleted ")
            Console.WriteLine(" default is only fully qualified paths will be deleted ")
            Console.WriteLine(" partial paths and fully qualified paths cannot be mixed ")
            Console.WriteLine(" /o means the parameterized paths will exist only once in the path ")
            Console.WriteLine(" /o cannot be used with /p ")
        Else
            If iValidArgs > 0 Then

                Dim readKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment", True)

                Dim v As Object = readKey.GetValue("Path", Nothing, RegistryValueOptions.DoNotExpandEnvironmentNames)

                path_array = Split(v.ToString, ";")
                If debug Then log.Print(" date_modified " & " Orig path_array ", path_array)
                Dim temp_path As String = "C:\"
                If Directory.Exists("C:\backup") Then
                    temp_path = "C:\backup\"
                Else
                    Try
                        MkDir("C:\backup")
                        temp_path = "C:\backup\"
                    Catch ex As Exception
                        ' if anything here the file defaults to the root of C
                    End Try
                End If

                Dim bSuccess As Boolean = write_values(temp_path & "orig_path.txt", date_modified & " - " & v.ToString)

                If bSuccess Then
                    Dim midCtr As Integer = 0
                    Dim midMax As Integer = path_array.Length - 1
                    Dim discarded_paths(midMax) As String
                    Dim sa As String
                    Dim sb As String
                    Dim occurranceCtr As Integer = 0
                    Dim newEntries(argMax) As String
                    'For idx = 0 To argMax
                    '    If Left(Trim(cmdArgs(idx)), 1) = "/" Then
                    '        newEntries(idx) = ""
                    '    Else
                    '        newEntries(idx) = cmdArgs(idx)
                    '    End If
                    'Next
                    'Loop through command arguments
                    For idx = 0 To argMax
                        '                    temp_string = cmdArgs(idx)
                        '                    Console.WriteLine("Temp String " & temp_string)
                        If occurranceCtr = 0 And idx > 0 And onlyOnce Then
                            'Need a trailing entry for argMax entry
                            If Left(Trim(cmdArgs(idx - 1)), 1) <> "/" Then
                                newEntries(idx - 1) = cmdArgs(idx - 1)
                            End If
                        Else
                            occurranceCtr = 0
                        End If
                        If debug Then log.Print(" In the Loop : idx = ", idx)
                        For midCtr = 0 To midMax
                            'skip switch parameters
                            If Left(Trim(cmdArgs(idx)), 1) = "/" Then Exit For
                            sa = UCase(Trim(path_array(midCtr)))
                            If Right(sa, 1) = "\" Then
                                sa = Left(sa, Len(sa) - 1)
                            End If
                            sb = UCase(Trim(cmdArgs(idx)))
                            If Right(sb, 1) = "\" Then
                                sb = Left(sb, Len(sb) - 1)
                            End If
                            If debug Then log.Print(" In the second Loop : midCtr = ", midCtr)
                            If debug Then log.Print(" In the second Loop : sa = ", sa)
                            If debug Then log.Print(" In the second Loop : sb = ", sb)
                            If partialPath Then
                                x = InStr(sa, sb, CompareMethod.Text)
                                If x > 0 Then
                                    discarded_paths(midCtr) = path_array(midCtr)
                                    path_array(midCtr) = ""
                                    If debug Then log.Print(" In the second Loop : sb was found in sa discard_paths = ", discarded_paths(midCtr))
                                Else
                                    If debug Then log.Print(" sb was not found in SA : sa = ", sa)
                                End If
                            Else
                                If sa = sb Then
                                    If (onlyOnce) And occurranceCtr = 0 Then
                                        occurranceCtr += 1
                                        If debug Then log.Print(" In the second loop: saved 1st occurrance of path", sb)
                                    Else
                                        discarded_paths(midCtr) = path_array(midCtr)
                                        path_array(midCtr) = ""
                                        If debug Then log.Print(" In the second Loop : sa equals sb discard_paths = ", discarded_paths(midCtr))
                                    End If
                                Else
                                    If debug Then log.Print(" sa did not equal sb : sa = ", sa)

                                End If

                            End If

                        Next
                    Next
                    If occurranceCtr = 0 And idx > argMax And onlyOnce Then
                        'Need a trailing entry for argMax entry
                        newEntries(argMax) = cmdArgs(argMax)
                        If debug Then log.Print(" Occurrance Ctr ", occurranceCtr)
                        If debug Then log.Print(" New Entries array ", newEntries)
                        If debug Then log.Print(" OnlyOnce : ", onlyOnce)
                    End If
                    Dim iCtr As Integer = 0
                    For midCtr = 0 To midMax
                        If Len(Trim(path_array(midCtr))) > 0 Then
                            iCtr += 1
                        End If
                    Next
                    Dim compressed_Path_array(iCtr - 1) As String
                    iCtr = 0
                    For midCtr = 0 To midMax
                        If Len(Trim(path_array(midCtr))) > 0 Then
                            compressed_Path_array(iCtr) = path_array(midCtr)
                            iCtr += 1
                        End If
                    Next
                    Dim newEntriesAdded As String = ""
                    If onlyOnce And occurranceCtr = 0 Then
                        For idx = 0 To argMax
                            If Len(Trim(newEntries(idx))) > 0 Then
                                newEntriesAdded = newEntriesAdded & ";" & newEntries(idx)
                            End If
                        Next
                    End If
                    Dim finished_path As String
                    If onlyOnce Then
                        If Len(Trim(newEntriesAdded)) > 0 Then
                            finished_path = Join(compressed_Path_array, ";") & newEntriesAdded
                        Else
                            finished_path = Join(compressed_Path_array, ";")
                        End If
                    Else
                        finished_path = Join(compressed_Path_array, ";")
                    End If
                    If debug Then log.Print("newEntriesAdded : ", newEntriesAdded)
                    If debug Then log.Print("compressed_path_array : ", compressed_Path_array)
                    If debug Then log.Print("finished_path : ", finished_path)
                    If Left(finished_path, 1) = ";" Then
                        finished_path = Right(finished_path, Len(finished_path) - 1)
                    End If
                    If Right(finished_path, 1) = ";" Then
                        finished_path = Left(finished_path, Len(finished_path) - 1)
                    End If
                    bSuccess = write_values(temp_path & "altered_path.txt", date_modified & " - " & finished_path)
                    Dim trash_path As String = Join(discarded_paths, ";")
                    bSuccess = write_values(temp_path & "discarded_path.txt", date_modified & " - " & trash_path)
                    Try
                        readKey.SetValue("path", finished_path, RegistryValueKind.ExpandString)
                    Catch ex As Exception
                    Finally
                        readKey.Close()
                    End Try
                Else
                    Console.WriteLine("Saving of Original Path Failed!....Process Aborted")
                    If debug Then log.Print(" Saving of Original Path Failed!", " Process Aborted ")
                End If
            Else
                Console.WriteLine("No Valid Arguments")
            End If
            End If

    End Sub
    Function write_values(ByVal file_name As String, ByVal msg As String) As Boolean
        Dim bRtnVal As Boolean = False

        Dim fn As Integer = FreeFile()
        Dim myFile As String = Dir(file_name)
        Try
            If Len(Trim(myFile)) > 0 Then
                'file exists
                FileOpen(fn, file_name, OpenMode.Append)
            Else
                FileOpen(fn, file_name, OpenAccess.Write)
            End If
            WriteLine(fn, msg)
            bRtnVal = True
        Catch ex As Exception
        Finally
            FileClose(fn)
        End Try
        Return bRtnVal
    End Function
End Module
