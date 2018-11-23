Imports System
Imports System.Diagnostics
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Outlook = Microsoft.Office.Interop.Outlook

Module Module1

    Private Declare Auto Function ShowWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    Private Declare Auto Function GetConsoleWindow Lib "kernel32.dll" () As IntPtr
    Private Const SW_HIDE As Integer = 0

    Dim application As Outlook.Application

    Dim sHideConsole As String
    Dim sDestinationFolder As String
    Dim sDestinationMailbox As String
    Dim sVerboseOutput As String
    Dim sSourceFolder As String
    Dim sSourceMailbox As String
    Dim sSyncLabel As String
    Dim sSyncDelimiter As String
    Dim sDeleteAfterSync As String
    Dim iDeleteSyncedOlderThanDays As Integer

    Dim oFolderSource As Outlook.Folder
    Dim oFolderDestination As Outlook.Folder

    Dim oMailItemSource As Outlook.MailItem

    Dim ItemsProcessed As Integer = 0
    Dim ItemsCopied As Integer = 0
    Dim ItemsDeleted As Integer = 0

    Sub Main()
        TryToConnectToRunningOutlook()
        SetInitialParameters()
        If sHideConsole = "Yes" Then
            HideWindow()
        End If
        ScanFolderSource()

        Log("", True)
        Log("ItemsProcessed: " & ItemsProcessed, True)        
        Log("ItemsCopied: " & ItemsCopied, True)
        Log("ItemsDeleted: " & ItemsDeleted, True)

    End Sub

    Sub HideWindow()
        Dim hWndConsole As IntPtr
        hWndConsole = GetConsoleWindow()
        ShowWindow(hWndConsole, SW_HIDE)
    End Sub

    Sub TryToConnectToRunningOutlook()
        If Process.GetProcessesByName("OUTLOOK").Count() > 0 Then
            application = DirectCast(Marshal.GetActiveObject("Outlook.Application"), Outlook.Application)
        Else
            application = New Outlook.Application()
            Dim ns As Outlook.NameSpace = application.GetNamespace("MAPI")
            ns.Logon("", "", Missing.Value, Missing.Value)
            ns = Nothing
        End If
    End Sub

    Sub SetInitialParameters()
        Dim sRegistryFolder As String = "HKEY_CURRENT_USER\Software\OutlookSyncFolders\" & My.Application.CommandLineArgs.First()
        sHideConsole = My.Computer.Registry.GetValue(sRegistryFolder, "HideConsole", Nothing)
        sDestinationFolder = My.Computer.Registry.GetValue(sRegistryFolder, "DestinationFolder", Nothing)
        sDestinationMailbox = My.Computer.Registry.GetValue(sRegistryFolder, "DestinationMailbox", Nothing)
        sVerboseOutput = My.Computer.Registry.GetValue(sRegistryFolder, "VerboseOutput", Nothing)
        sSourceMailbox = My.Computer.Registry.GetValue(sRegistryFolder, "SourceMailbox", Nothing)
        sSourceFolder = My.Computer.Registry.GetValue(sRegistryFolder, "SourceFolder", Nothing)
        sSyncLabel = My.Computer.Registry.GetValue(sRegistryFolder, "SyncLabel", Nothing)
        sSyncDelimiter = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\International", "sList", Nothing)
        sDeleteAfterSync = My.Computer.Registry.GetValue(sRegistryFolder, "DeleteAfterSync", Nothing)
        iDeleteSyncedOlderThanDays = My.Computer.Registry.GetValue(sRegistryFolder, "DeleteSyncedOlderThanDays", Nothing)

        oFolderSource = application.Session.Folders.Item(sSourceMailbox).Folders.Item(sSourceFolder)
        oFolderDestination = application.Session.Folders.Item(sDestinationMailbox).Folders.Item(sDestinationFolder)
    End Sub

    Sub ScanFolderSource()
        Log("", True)
        Dim oMailItemsSource As Outlook.Items = oFolderSource.Items
        For i = oMailItemsSource.Count To 1 Step -1
            Try
                oMailItemSource = oMailItemsSource(i)
                If TypeOf oMailItemSource Is Outlook.MailItem And oMailItemSource.Sent Then
                    ProcessMailItem()
                End If
            Catch ex As Exception
            End Try
        Next
    End Sub

    Sub ProcessMailItem()
        Log(oMailItemSource.SentOn & " " & oMailItemSource.Subject, True)

        ClearSyncLabel()
        If Not OriginMailItemIsFoundedInDestination() Then
            CopyMailItem()
        End If

        If sDeleteAfterSync = "Yes" And IsNumeric(iDeleteSyncedOlderThanDays) = True Then
            DeleteOldMailItem()
        End If

        ItemsProcessed = ItemsProcessed + 1
        Log("", True)
    End Sub

    Sub DeleteOldMailItem()
        Dim lDaysElapsed As Long = DateDiff(DateInterval.Day, oMailItemSource.SentOn, Date.Today)
        Log("age: " & lDaysElapsed & " days", True)
        If lDaysElapsed > iDeleteSyncedOlderThanDays And OriginMailItemIsFoundedInDestination() Then
            oMailItemSource.Delete()
            ItemsDeleted = ItemsDeleted + 1
            Log("MailItem has been deleted", True)
        End If
    End Sub

    Sub CopyMailItem()
        Dim oMailItemCopy As Outlook.MailItem
        oMailItemCopy = oMailItemSource.Copy
        oMailItemCopy.Move(oFolderDestination)
        SetSyncLabel()
        ItemsCopied = ItemsCopied + 1
        Log("MailItem has been copied", True)        
    End Sub

    Sub SetSyncLabel()
        If Not SyncLabelExists() Then
            oMailItemSource.Categories = oMailItemSource.Categories & sSyncDelimiter & sSyncLabel
            oMailItemSource.Save()
        End If
    End Sub

    Sub ClearSyncLabel()
        If SyncLabelExists() Then
            oMailItemSource.Categories = oMailItemSource.Categories.Replace(sSyncLabel, "")
            oMailItemSource.Categories = oMailItemSource.Categories.Replace(sSyncDelimiter & sSyncDelimiter, sSyncDelimiter)
            oMailItemSource.Save()
        End If
    End Sub

    Function SyncLabelExists()
        Dim bAnswer As Boolean
        If oMailItemSource.Categories Is Nothing Then
            bAnswer = False
            Log("MailItem has no label", True)
        ElseIf oMailItemSource.Categories.Contains(sSyncLabel) Then
            bAnswer = True
            Log("MailItem has label", True)
        Else
            bAnswer = False
            Log("MailItem has no label", True)
        End If
        Return bAnswer
    End Function

    Function OriginMailItemIsFoundedInDestination()

        Dim blnMailItemIsFounded As Boolean = False
        Dim dMailItemSentOn As Date
        Dim sMailItemSentOnBefore As String
        Dim sMailItemSentOnAfter As String
        Dim sMailItemSenderEmail As String
        Dim sMailItemTo As String
        Dim sMailItemCC As String
        Dim sMailItemSubject As String

        Dim oMailItemSimilar As Outlook.MailItem

        Dim sMailItemsSearchCriterias As String
        Dim oMailItemsSearchResult As Outlook.Items

        Dim PR_SEARCH_KEY As String = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"
        Dim strMessageOriginId As String
        Dim strMessageSimilarId As String

        dMailItemSentOn = oMailItemSource.SentOn.ToUniversalTime
        sMailItemSentOnBefore = FormatDateTime(CStr(dMailItemSentOn.AddMinutes(-1)), DateFormat.ShortDate) + " " + FormatDateTime(CStr(dMailItemSentOn.AddMinutes(-1)), DateFormat.ShortTime)
        sMailItemSentOnAfter = FormatDateTime(CStr(dMailItemSentOn.AddMinutes(1)), DateFormat.ShortDate) + " " + FormatDateTime(CStr(dMailItemSentOn.AddMinutes(1)), DateFormat.ShortTime)
        sMailItemSenderEmail = Replace(oMailItemSource.SenderEmailAddress, "'", "''")
        sMailItemTo = Replace(oMailItemSource.To, "'", "''")
        sMailItemCC = Replace(oMailItemSource.CC, "'", "''")
        sMailItemSubject = Replace(oMailItemSource.Subject, "'", "''")

        sMailItemsSearchCriterias = "@SQL=(" & Chr(34) & "urn:schemas:httpmail:date" & Chr(34) & " > " & "'" & sMailItemSentOnBefore & "'" & _
                                    " AND " & Chr(34) & "urn:schemas:httpmail:date" & Chr(34) & " < " & "'" & sMailItemSentOnAfter & "'" & _
                                    " AND " & Chr(34) & "urn:schemas:httpmail:fromemail" & Chr(34) & " = " & "'" & sMailItemSenderEmail & "'" & _
                                    " AND " & Chr(34) & "urn:schemas:httpmail:displayto" & Chr(34) & " = " & "'" & sMailItemTo & "'" & _
                                    " AND " & Chr(34) & "urn:schemas:httpmail:displaycc" & Chr(34) & " = " & "'" & sMailItemCC & "'" & _
                                    " AND " & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34)

        If sMailItemSubject = "" Then
            sMailItemsSearchCriterias = sMailItemsSearchCriterias & " IS NULL "
        Else
            sMailItemsSearchCriterias = sMailItemsSearchCriterias & " = " & "'" & sMailItemSubject & "'"
        End If
        sMailItemsSearchCriterias = sMailItemsSearchCriterias & ")"

        oMailItemsSearchResult = oFolderDestination.Items.Restrict(sMailItemsSearchCriterias)

        If oMailItemsSearchResult.Count > 0 Then
            For Each oMailItemSimilar In oMailItemsSearchResult
                strMessageOriginId = oMailItemSource.PropertyAccessor.BinaryToString(oMailItemSource.PropertyAccessor.GetProperty(PR_SEARCH_KEY))
                strMessageSimilarId = oMailItemSimilar.PropertyAccessor.BinaryToString(oMailItemSimilar.PropertyAccessor.GetProperty(PR_SEARCH_KEY))
                If strMessageSimilarId = strMessageOriginId Then
                    blnMailItemIsFounded = True
                    SetSyncLabel()
                    Exit For
                End If
            Next
        End If

        If blnMailItemIsFounded Then
            Log("MailImtem is founded in destination", True)
        End If

        Return blnMailItemIsFounded
    End Function

    Sub Log(Message As String, Optional NewLine As Boolean = False)
        Select Case sVerboseOutput
            Case "Debug"
                Debug.Print(Message)
            Case "Console"
                If NewLine Then
                    Console.WriteLine(Message)
                Else
                    Console.Write(Message)
                End If

            Case "Both"
                Debug.Print(Message)
                If NewLine Then
                    Console.WriteLine(Message)
                Else
                    Console.Write(Message)
                End If
        End Select
    End Sub

End Module



