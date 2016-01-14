Imports System
Imports System.IO
Imports System.Xml
Imports Microsoft.VisualBasic
Imports System.Security.Permissions
Imports System.Printing

'14.1.201 version 5.0.1
'trotter 0 byte pdf bypass
'reduced wait time

'All Paths ends with "\"

Public Class S2P
    Dim Jobonoff As New Boolean 'Master switch
    Dim xMinute As New System.Timers.Timer 'Ticker
    Dim xArray As ArrayList

    Dim avalfld As New ArrayList 'Running Inbox folders

    Dim Pcount As Integer = 0
    Dim Runtimexml As New XmlDocument
    Dim Runtimexmlpath As String
    Dim PrintProcess As Boolean = False 'Control Single Print Process



    Private Sub InitRuntime() 'runtime startup log
        Dim rttimestamp = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim xmlerr As String = ""

        Try
            My.Computer.FileSystem.CopyFile(My.Settings.runtimeroot + "template\temp.xml", My.Settings.runtimeroot + "rt" + rttimestamp + ".xml", False) 'log template
        Catch ex As Exception
            xmlerr = ex.ToString
        End Try
        If Not xmlerr = "" Then
            System.Threading.Thread.Sleep(2000)
            rttimestamp = DateTime.Now.ToString("yyyyMMddHHmmss")
            My.Computer.FileSystem.CopyFile(My.Settings.runtimeroot + "template\temp.xml", My.Settings.runtimeroot + "rt" + rttimestamp + ".xml", False)
        End If

        If My.Computer.FileSystem.FileExists(My.Settings.runtimeroot + "rt" + rttimestamp + ".xml") Then
            Runtimexmlpath = My.Settings.runtimeroot + "rt" + rttimestamp + ".xml"
            Runtimexml.Load(Runtimexmlpath)
            Dim runtimenode As XmlElement = Runtimexml.GetElementById("runtimecfg")
            runtimenode.SetAttribute("date", DateTime.Now.ToString("yyyyMMdd"))
            runtimenode.SetAttribute("time", DateTime.Now.ToString("HHmmss"))
            Runtimexml.Save(Runtimexmlpath)
            xremovefolder()
            xchkFolders()

        End If
    End Sub
    Private Sub xchkFolders()

        For Each mFld As String In My.Settings.fldNameList
            If My.Computer.FileSystem.DirectoryExists(My.Settings.fldPath + mFld) Then
                appendfolder(My.Settings.fldPath + mFld, "online")
            Else
                appendfolder(My.Settings.fldPath + mFld, "offline")
            End If
        Next
    End Sub
    Private Sub appendfolder(path As String, status As String)
        Runtimexml.Load(Runtimexmlpath)
        Dim xflds As XmlElement = Runtimexml.GetElementById("folders")
        Dim xfld As XmlElement = Runtimexml.CreateElement("folder")
        xfld.SetAttribute("path", path)
        xfld.SetAttribute("status", status)
        xflds.AppendChild(xfld)
        Runtimexml.Save(Runtimexmlpath)
    End Sub

    Private Sub xremovefolder()
        Runtimexml.Load(Runtimexmlpath)
        Dim allflds As XmlElement = Runtimexml.GetElementById("folders")
        For Each xfld As XmlElement In allflds.ChildNodes
            fLogger(xfld.GetAttribute("Path").ToString)
            allflds.RemoveChild(xfld)
        Next
        Runtimexml.Save(Runtimexmlpath)
    End Sub

    Private Sub chkFolders() 'check folders availability before each timer ticks
        Dim updateavalfld As New ArrayList
        updateavalfld.Clear()
        For Each mFld As String In My.Settings.fldNameList
            If My.Computer.FileSystem.DirectoryExists(My.Settings.fldPath + mFld) Then
                updateavalfld.Add(My.Settings.fldPath + mFld) 'write to avail list
            Else
                fLogger(My.Settings.fldPath + mFld & " is now offline") 'log error
            End If
        Next
        avalfld = updateavalfld 'update global avail list
    End Sub
    Private Sub Trotter() 'Trotter Job : Moves files from Inboxes to PrintBox

        Dim rng As New Random
        Dim number As Integer = rng.Next(1, 9999)
        Dim digits As String = number.ToString("0000")
        Dim dirname As String
        Dim fileInfoObj As FileInfo

        For Each mfld As String In avalfld
            'dirname = "test"
            If My.Computer.FileSystem.DirectoryExists(mfld) Then
                dirname = Path.GetFileName(mfld)

                Dim fileNames = My.Computer.FileSystem.GetFiles(mfld, FileIO.SearchOption.SearchTopLevelOnly, "*.pdf") 'PDF only
                For Each filename As String In fileNames
                    If My.Computer.FileSystem.FileExists(filename) Then
                        fileInfoObj = Nothing
                        fileInfoObj = My.Computer.FileSystem.GetFileInfo(filename)
                        If fileInfoObj.Length <> 0 Then
                            If fileInfoObj.LastWriteTime.AddSeconds(10) < DateTime.Now Then 'Files aged 10 seconds or more
                                Try
                                    number = rng.Next(1, 9999)
                                    digits = number.ToString("0000") 'append rand to timestamp
                                    My.Computer.FileSystem.MoveFile(filename, My.Settings.printBox + DateTime.Now.ToString("yyyyMMddHHmmss") + digits + "-" + dirname + "-" + Path.GetFileName(filename).Replace("-", "_"), False)
                                    fLogger("filename " + filename + "moved to " + My.Settings.printBox + DateTime.Now.ToString("yyyyMMddHHmmss") + digits + "-" + dirname + "-" + Path.GetFileName(filename).Replace("-", "_"))
                                    Pcount = Pcount + 1
                                Catch ex As Exception
                                    fLogger("filename " + filename + " cannot be moved. " + ex.Message)
                                End Try
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub sPrinter() 'printer job: print from PBOX and move file to backup
        If PrintProcess = False Then 'One Print process only
            LockPP() 'Lock Thread
            Dim RetryQ As New List(Of String())
            Dim RetryBKQ As New List(Of String)
            Dim RetryBKRQ As New List(Of String)
            Dim SuccessQ As New List(Of String)
            Dim FailedQ As New List(Of String)
            Dim rng As New Random
            Dim rNumber As Integer = rng.Next(1, 9999)
            Dim rDigits As String = rNumber.ToString("0000")
            Dim backupfld As String = My.Settings.backuproot + DateTime.Now.ToString("yyyyMMdd") + "\"
            Dim OUloc As String = ""
            Dim OUbackupfld As String = ""
            If Not My.Computer.FileSystem.DirectoryExists(backupfld) Then 'Create Daily Backup Folder
                My.Computer.FileSystem.CreateDirectory(backupfld)
            End If
            Dim pErr As String = ""
            Dim fileNames = My.Computer.FileSystem.GetFiles(My.Settings.printBox, FileIO.SearchOption.SearchTopLevelOnly, "*.pdf") 'Current List of files in PBOX

            'Start Print Process
            For Each filename As String In fileNames
                If My.Computer.FileSystem.FileExists(filename) And My.Computer.FileSystem.GetFileInfo(filename).CreationTime.AddSeconds(10) < DateTime.Now Then 'Files aged 10 seconds or more
                    'Check Lock
                    If FileIsLocked(filename, FileAccess.ReadWrite) = False Then 'step 1
                        NameStamp(filename) 'Watermark before print
                        System.Threading.Thread.Sleep(1500) 'Wait for Process complete
                        If FileIsLocked(filename, FileAccess.ReadWrite) = False Then 'step 2
                            Try
                                pdfPrint(filename, pErr)
                                pLogger("printing " + filename + " Job sent to print queue")
                                System.Threading.Thread.Sleep(6000)
                                If CheckPrintSucess(filename) Then
                                    pLogger("printing " + filename + " Checked Print PDF No Error")
                                    SuccessQ.Add(filename) 'append filename to list of success prints
                                Else
                                    'added for retry with extended timer 10 sec
                                    pLogger("waiting 10 more secs...")
                                    System.Threading.Thread.Sleep(10000)

                                    If CheckPrintSucess(filename) Then
                                        pLogger("printing " + filename + " Checked Print PDF No Error")
                                        SuccessQ.Add(filename) 'append filename to list of success prints
                                    Else
                                        'added for retry with extended timer 20 sec
                                        pLogger("waiting 20 more secs...")
                                        System.Threading.Thread.Sleep(20000)
                                        If CheckPrintSucess(filename) Then
                                            pLogger("printing " + filename + " Checked Print PDF No Error")
                                            SuccessQ.Add(filename) 'append filename to list of success prints
                                        Else
                                            'added for retry with extended timer 40 sec
                                            pLogger("waiting 40 more secs...")
                                            System.Threading.Thread.Sleep(40000)
                                            If CheckPrintSucess(filename) Then
                                                pLogger("printing " + filename + " Checked Print PDF No Error")
                                                SuccessQ.Add(filename) 'append filename to list of success prints
                                            Else
                                                'added for retry with extended timer 80 sec
                                                pLogger("waiting 80 more secs...")
                                                System.Threading.Thread.Sleep(80000)
                                                If CheckPrintSucess(filename) Then
                                                    pLogger("printing " + filename + " Checked Print PDF No Error")
                                                    SuccessQ.Add(filename) 'append filename to list of success prints
                                                Else
                                                    'added for retry with extended timer 80 sec
                                                    pLogger("waiting 80 more secs...")
                                                    System.Threading.Thread.Sleep(80000)
                                                    If CheckPrintSucess(filename) Then
                                                        pLogger("printing " + filename + " Checked Print PDF No Error")
                                                        SuccessQ.Add(filename) 'append filename to list of success prints
                                                    Else
                                                        RetryQ.Add({2, filename}) 'append filename to list of retry prints, error level reprint
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                pLogger("printing " + filename + "failed - " + ex.Message)
                            End Try
                        Else
                            RetryQ.Add({2, filename}) 'append filename to list of retry prints, error levelreprint
                            pLogger("printing " + filename + " locked.")
                        End If
                    Else
                        RetryQ.Add({1, filename}) 'append filename to list of retry prints, error level stamp and print
                        pLogger("printing " + filename + " locked.")
                    End If
                Else
                    FailedQ.Add(filename) 'append filename to list of failed print, no retry
                    pLogger("printing " + filename + " not exist or too young.")
                End If
            Next
            'retry Process start 2 seconds after
            System.Threading.Thread.Sleep(2000)
            For Each retryfile As String() In RetryQ
                If FileIsLocked(retryfile(1), FileAccess.ReadWrite) = False Then
                    If retryfile(0) = 1 Then 'level 1
                        NameStamp(retryfile(1)) 'Watermark before print
                        System.Threading.Thread.Sleep(1500) 'Wait for Process complete
                        Try
                            pdfPrint(retryfile(1), pErr)
                            pLogger("Retry printing " + retryfile(1) + " Job sent to print queue")
                            System.Threading.Thread.Sleep(2000)
                            If CheckPrintSucess(retryfile(1)) Then
                                pLogger("Retry printing " + retryfile(1) + " Checked Print PDF No Error")
                                SuccessQ.Add(retryfile(1)) 'append filename to list of success prints
                            Else
                                FailedQ.Add(retryfile(1)) 'append filename to list of failed print
                            End If
                        Catch ex As Exception
                            pLogger("Retry printing " + retryfile(1) + " failed - " + ex.Message)
                        End Try
                    ElseIf retryfile(0) = 2 Then 'level 2
                        Try
                            pdfPrint(retryfile(1), pErr)
                            pLogger("Retry printing " + retryfile(1) + " Job sent to print queue")
                            System.Threading.Thread.Sleep(2000)
                            If CheckPrintSucess(retryfile(1)) Then
                                pLogger("Retry printing " + retryfile(1) + " Checked Print PDF No Error")
                                SuccessQ.Add(retryfile(1)) 'append filename to list of success prints
                            Else
                                FailedQ.Add(retryfile(1)) 'append filename to list of failed print
                            End If
                        Catch ex As Exception
                            pLogger("Retry printing " + retryfile(1) + " failed - " + ex.Message)
                        End Try

                    End If
                Else
                    pLogger("Retry printing " + retryfile(1) + " failed - FILE IS LOCKED")
                End If
            Next

            'Clean up PBOX and Move files to OU backup
            For Each successfile As String In SuccessQ
                If FileIsLocked(successfile, FileAccess.ReadWrite) = False Then
                    Try 'move to backup folder
                        My.Computer.FileSystem.MoveFile(successfile, backupfld + Path.GetFileName(successfile), False) ' Move file to Backup from PBOX
                    Catch ex As Exception 'add to retry Q
                        pLogger("Local Backup to " + backupfld + "failed - " + ex.Message)
                        RetryBKQ.Add(successfile) 'append filename to list of retry backup
                    End Try

                    System.Threading.Thread.Sleep(2000) 'wait for move
                    If My.Computer.FileSystem.FileExists(backupfld + Path.GetFileName(successfile)) Then
                        'Start OU Backup
                        OUloc = Path.GetFileName(successfile).Substring(Path.GetFileName(successfile).IndexOf("-") + 1, (Path.GetFileName(successfile).LastIndexOf("-") - Path.GetFileName(successfile).IndexOf("-")) - 1)
                        OUbackupfld = My.Settings.OUbkpath + OUloc + My.Settings.OUbksuffix + "\" + DateTime.Now.ToString("yyyyMMdd") + "\"
                        'fLogger(OUbackupfld)
                        If Not My.Computer.FileSystem.DirectoryExists(OUbackupfld) Then

                            Try
                                pLogger("creating folder for backup")
                                My.Computer.FileSystem.CreateDirectory(OUbackupfld)
                            Catch ex As Exception
                                pLogger("Create Remote Backup Folder for" + OUbackupfld + " failed - " + ex.Message)
                            End Try

                        End If
                        If My.Computer.FileSystem.DirectoryExists(OUbackupfld) Then

                            Try
                                pLogger("moving " + backupfld + Path.GetFileName(successfile) + " to " + OUbackupfld)
                                My.Computer.FileSystem.CopyFile(backupfld + Path.GetFileName(successfile), OUbackupfld + Path.GetFileName(successfile), False)
                            Catch ex As Exception
                                pLogger("Remote Backup to " + OUbackupfld + " failed - " + ex.Message)
                                RetryBKRQ.Add(backupfld + Path.GetFileName(successfile)) 'append filename to list of retry backup
                            End Try

                        End If
                        'End OU Backup
                    End If
                Else
                    System.Threading.Thread.Sleep(2000)
                    RetryBKQ.Add(successfile) 'append filename to list of retry backup
                    pLogger("backup " + successfile + "locked.")
                End If
            Next

            'Handle failed files...LOCAL BACKUP.
            For Each RetryBKfile As String In RetryBKQ
                If FileIsLocked(RetryBKfile, FileAccess.ReadWrite) = False Then
                    Try 'try
                        rNumber = rng.Next(1, 9999)
                        rDigits = rNumber.ToString("0000") 'append rand to timestamp
                        My.Computer.FileSystem.MoveFile(RetryBKfile, backupfld + "r_" + rDigits + "_" + Path.GetFileName(RetryBKfile), False) ' Move file to Backup from PBOX
                    Catch ex As Exception
                        pLogger("Retry Local Backup to " + RetryBKfile + " failed - " + ex.Message)
                        Try 'try rename
                            My.Computer.FileSystem.RenameFile(RetryBKfile, Path.GetFileName(RetryBKfile) + ".FAILED")
                        Catch nex As Exception
                            pLogger("Rename failed file" + RetryBKfile + " failed - " + nex.Message)
                        End Try
                    End Try

                    System.Threading.Thread.Sleep(2000)
                    If My.Computer.FileSystem.FileExists(backupfld + Path.GetFileName(RetryBKfile)) Then
                        'Start OU Backup
                        OUloc = Path.GetFileName(RetryBKfile).Substring(Path.GetFileName(RetryBKfile).IndexOf("-") + 1, (Path.GetFileName(RetryBKfile).LastIndexOf("-") - Path.GetFileName(RetryBKfile).IndexOf("-")) - 1)
                        OUbackupfld = My.Settings.OUbkpath + OUloc + My.Settings.OUbksuffix + "\" + DateTime.Now.ToString("yyyyMMdd") + "\"

                        If Not My.Computer.FileSystem.DirectoryExists(OUbackupfld) Then

                            Try
                                pLogger("creating folder for backup")
                                My.Computer.FileSystem.CreateDirectory(OUbackupfld)
                            Catch ex As Exception
                                pLogger("Create Remote Backup Folder for" + OUbackupfld + " failed - " + ex.Message)
                            End Try

                        End If
                        If My.Computer.FileSystem.DirectoryExists(OUbackupfld) Then
                            Try
                                pLogger("copying " + backupfld + Path.GetFileName(RetryBKfile) + " to " + OUbackupfld)
                                My.Computer.FileSystem.CopyFile(backupfld + Path.GetFileName(RetryBKfile), OUbackupfld + Path.GetFileName(RetryBKfile), False)
                            Catch ex As Exception
                                pLogger("Remote Backup to " + OUbackupfld + " failed - " + ex.Message)
                                RetryBKRQ.Add(backupfld + Path.GetFileName(RetryBKfile)) 'append filename to list of retry backup
                            End Try

                        End If
                        'End OU Backup
                    End If
                Else
                    pLogger("Retry Local Backup to " + RetryBKfile + " failed - FILE LOCKED")
                    Try 'try rename
                        My.Computer.FileSystem.RenameFile(RetryBKfile, Path.GetFileName(RetryBKfile) + ".FAILED")
                    Catch nex As Exception
                        pLogger("Rename failed file" + RetryBKfile + " failed - " + nex.Message)
                    End Try
                End If
            Next
            '...

            '..............OU BACKUP RETRY............................
            For Each RetryBKRfile As String In RetryBKRQ
                System.Threading.Thread.Sleep(2000)
                If My.Computer.FileSystem.FileExists(backupfld + Path.GetFileName(RetryBKRfile)) Then
                    'Start OU Backup
                    OUloc = Path.GetFileName(RetryBKRfile).Substring(Path.GetFileName(RetryBKRfile).IndexOf("-") + 1, (Path.GetFileName(RetryBKRfile).LastIndexOf("-") - Path.GetFileName(RetryBKRfile).IndexOf("-")) - 1)
                    OUbackupfld = My.Settings.OUbkpath + OUloc + My.Settings.OUbksuffix + "\" + DateTime.Now.ToString("yyyyMMdd") + "\"

                    If Not My.Computer.FileSystem.DirectoryExists(OUbackupfld) Then

                        Try
                            pLogger("creating folder for backup")
                            My.Computer.FileSystem.CreateDirectory(OUbackupfld)
                        Catch ex As Exception
                            pLogger("Create Remote Backup Folder for" + OUbackupfld + " failed - " + ex.Message)
                        End Try

                    End If
                    If My.Computer.FileSystem.DirectoryExists(OUbackupfld) Then
                        Try
                            pLogger("copying " + backupfld + Path.GetFileName(RetryBKRfile) + " to " + OUbackupfld)
                            My.Computer.FileSystem.CopyFile(backupfld + Path.GetFileName(RetryBKRfile), OUbackupfld + "r_" + rDigits + "_" + Path.GetFileName(RetryBKRfile), False)
                        Catch ex As Exception
                            pLogger("RETRY Remote Backup to " + OUbackupfld + " failed - " + ex.Message)
                        End Try

                    End If
                    'End OU Backup
                End If
            Next
            '..............OU BACKUP RETRY............................


            Try
                ReleasePP() 'Release Thread
                pLogger("PP Released")
            Catch ex As Exception
                pLogger("PP Release failed" + ex.Message)
            End Try
            ReleasePP() 'Release Thread

        Else
            pLogger("PP Already Locked")

        End If

    End Sub
    Private Sub LockPP()
        PrintProcess = True
    End Sub
    Private Sub ReleasePP()
        PrintProcess = False
    End Sub
    Private Function CheckPrintSucess(filename As String) As Boolean ' check file print success

        'connect to print server and check status
        Dim jobArr As String = "0"
        Dim jobErr As String = "0"
        Dim myPrintServer1 As New PrintServer(My.Settings.printServer1)
        Dim printQ As PrintQueue = myPrintServer1.GetPrintQueue(My.Settings.printQueue1)
        printQ.Refresh()
        Dim pjobs As PrintJobInfoCollection = printQ.GetPrintJobInfoCollection()

        For Each pjob As PrintSystemJobInfo In pjobs
            If pjob.Name.ToString.Contains(filename) Then
                jobArr = "1"
                pLogger("PrintJob Arrived: " + pjob.Name.ToString + " - " + pjob.JobStatus.ToString + " - " + pjob.TimeJobSubmitted.AddMinutes(480) + " - " + pjob.NumberOfPages.ToString + " - " + pjob.Submitter.ToString)
            End If
        Next

        'Check Daily Folder for Job Errors
        Dim logfld As String = My.Settings.printboxlogfld + DateTime.Now.ToString("yyyyMMdd") + "\"
        Dim logfile As String = logfld + Path.GetFileName(filename) + ".txt"
        Dim logError As String

        If My.Computer.FileSystem.FileExists(logfile) Then
            logError = File.ReadAllText(logfile)
            If logError = "" Then
                pLogger("printing " + filename + " ,sumatra give good log")
                jobErr = "0"
            Else
                pLogger("printing " + filename + " ,sumatra gives bad log- " + logError)
                jobErr = "1"

            End If
        Else
            pLogger("printing " + filename + " ,sumatra gives no log")
            jobErr = "1"

        End If

        If jobArr = "1" And jobErr = "0" Then
            Return True
        Else
            If jobArr <> "1" Then
                pLogger("Not Arrived")
            End If
            Return False
        End If

    End Function
    Private Sub pdfPrint(pfile As String, pErrmsg As String) 'revamp
        Try
            Dim logfld As String = My.Settings.printboxlogfld + DateTime.Now.ToString("yyyyMMdd") + "\"
            If Not My.Computer.FileSystem.DirectoryExists(logfld) Then 'Create Daily log Folder
                My.Computer.FileSystem.CreateDirectory(logfld)
            End If
            Dim stderrstr As String

            Dim Process1 As New Process

            pLogger(My.Settings.pdfviewerpath + " -silent -print-to " + My.Settings.printerName1 + " " + pfile)

            Dim psi As New ProcessStartInfo(My.Settings.pdfviewerpath, " -silent -print-to " + My.Settings.printerName1 + " " + """" + pfile + """")
            psi.RedirectStandardError = True
            psi.CreateNoWindow = True
            psi.UseShellExecute = False
            psi.ErrorDialog = False
            Process1.StartInfo() = psi
            Process1.Start()
            stderrstr = Process1.StandardError.ReadToEnd()
            Process1.WaitForExit()
            PDFstderrLogger(stderrstr, logfld + Path.GetFileName(pfile) + ".txt")
        Catch ex As Exception
            pLogger(ex.ToString)
        End Try

    End Sub
    Private Sub NameStamp(pfile As String) 'Stamp file name in footer using cpdf
        Try
            Dim Process1 As New Process
            Dim psi As New ProcessStartInfo(My.Settings.cpdfpath, " -rotate 90 " + pfile + " AND -add-text " + Path.GetFileName(pfile) + " -bottom 20 -font-size 10 -prerotate AND -rotate 270 -o " + """" + pfile + """")
            psi.CreateNoWindow = True
            psi.UseShellExecute = False
            Process1.StartInfo() = psi
            Process1.Start()
        Catch ex As Exception
            fLogger("NameStamp: " + pfile + " - " + ex.ToString)
        End Try
    End Sub

    Private Sub fLogger(fmsg As String)
        Dim fStrm As FileStream
        Dim fName, fLog As String

        fLog = fmsg
        fName = My.Settings.logfldroot + "trot/" + DateTime.Today.ToString("ddMMyyyy") + ".txt"

        If My.Computer.FileSystem.FileExists(fName) Then
            fStrm = New FileStream(fName, FileMode.Append)
        Else
            fStrm = New FileStream(fName, FileMode.OpenOrCreate)
        End If

        Dim fWriter As StreamWriter = New StreamWriter(fStrm)

        Try
            fWriter.WriteLine(DateTime.Now.ToString("MM/dd/yyyy  HH:mm:ss") + " -" + fmsg)

        Catch ex As Exception
            'MsgBox("!cannot write log!")
        End Try

        fWriter.Flush()
        fWriter.Close()
        fStrm.Close()

    End Sub
    Private Sub PDFstderrLogger(fmsg As String, logpath As String)
        Dim fStrm As FileStream

        fStrm = New FileStream(logpath, FileMode.OpenOrCreate)

        Dim fWriter As StreamWriter = New StreamWriter(fStrm)

        Try
            fWriter.Write(fmsg)

        Catch ex As Exception
            'MsgBox("!cannot write log!")
        End Try

        fWriter.Flush()
        fWriter.Close()
        fStrm.Close()

    End Sub

    Private Sub pLogger(pmsg As String)
        Dim fStrm As FileStream
        Dim fName, pLog As String

        pLog = pmsg

        fName = My.Settings.logfldroot + "print/" + DateTime.Today.ToString("ddMMyyyy") + ".txt"

        If My.Computer.FileSystem.FileExists(fName) Then
            fStrm = New FileStream(fName, FileMode.Append)
        Else
            fStrm = New FileStream(fName, FileMode.OpenOrCreate)
        End If

        Dim fWriter As StreamWriter = New StreamWriter(fStrm)

        Try
            fWriter.WriteLine(DateTime.Now.ToString("MM/dd/yyyy  HH:mm:ss") + " -" + pmsg)

        Catch ex As Exception
            'MsgBox("!cannot write log!")
        End Try

        fWriter.Flush()
        fWriter.Close()
        fStrm.Close()

    End Sub
    Private Function FileIsLocked(ByVal filename As String, _
    ByVal file_access As FileAccess) As Boolean
        Const FILE_LOCKED As Integer = 57

        On Error Resume Next
        Dim fs As New FileStream(filename, FileMode.Open, _
            file_access)
        Dim err_number As Integer = Err.Number
        fs.Close()

        On Error GoTo 0
        If err_number = 0 Then Return False
        If err_number = FILE_LOCKED Then Return True
        Err.Raise(err_number)
    End Function


    Private Sub updatecount()
        Runtimexml.Load(Runtimexmlpath)
        For Each xpcount As XmlElement In Runtimexml.GetElementsByTagName("processed")
            xpcount.SetAttribute("count", Pcount.ToString)
        Next
        Runtimexml.Save(Runtimexmlpath)

    End Sub

    Public Sub timeroff()
        If Jobonoff = True Then
            xMinute.Enabled = False
            'MsgBox("Job Stopped")
            Jobonoff = False
        End If
    End Sub

    Private Sub OnTimerEvent(ByVal [source] As Object, ByVal e As EventArgs)
        'Global Timer Tick
        pLogger("TimerStart")
        chkFolders()
        Trotter()
        System.Threading.Thread.Sleep(12000)
        sPrinter()
        updatecount()
        pLogger("TimerEnd")
    End Sub 'Global Timer

    Public Sub init()

        If Jobonoff = False Then
            InitRuntime()
            xMinute.Interval = My.Settings.taskfrq
            xMinute.Enabled = True
            AddHandler xMinute.Elapsed, AddressOf OnTimerEvent
            'chkFolders()
            'MsgBox("Job Started")
            Jobonoff = True
            'Button1.Enabled = True
            'bnMoniAll.Enabled = False
        End If
    End Sub
End Class