Imports OutlookAddIn1.modMain.globalVar
Imports OutlookAddIn1.FrmMain
Imports System.Drawing
Module modMain
    Public Class globalVar

        Public Shared mailboxBtnRetrive As Boolean = False
        Public Shared rootPath As String = "C:\Users\" & Environ("UserName") & "\Desktop"

        Public Shared Global_ckbxAutoLoad As Boolean = False
        Public Shared Global_ckbxSaveasMSG As Boolean = False
        Public Shared Global_ckbDownloadAttachment As Boolean = False
        Public Shared Global_txbxRetriveFrom As String = ""
        Public Shared Global_txbxMoveTo As String = ""
        Public Shared Global_ckbxExact As Boolean = False
        Public Shared Global_ckbxCaseSensitive As Boolean = False
        Public Shared Global_txbx_FromEmailFilter As String = ""
        Public Shared Global_txbx_SubjectFilter As String = ""
        Public Shared Global_txbx_AttchContans As String = ""
        Public Shared Global_dtFrom As String = ""
        Public Shared Global_ckbxAllFrom As Boolean = False
        Public Shared Global_dtTo As String = ""
        Public Shared Global_ckbxAllTo As Boolean = False
        Public Shared Global_txbx_folderLoc As String = ""
        Public Shared Global_chkBx_ReplaceExistingFile As Boolean = False
        Public Shared Global_cmbx_SaveFileAs As String = ""
        Public Shared Global_txbxNewName As String = ""
        Public Shared Global_autoEmailCount As Long = 0
        Public Shared Global_emailextractCount As Long = 0

        Public Shared ns As Outlook.NameSpace
    End Class


    Public Function checkSubscription(auto As Boolean) As Boolean

        If auto = True Then
            If Global_autoEmailCount > 10 Then
                MsgBox("you can extract only 10 auto emails in basic version." & vbNewLine &
                    "buy lpro version for $30 (lifetime licence) by emailing deepak9delhi@gmail.com " & vbNewLine &
                    "or visit www.fiverr.com/DeepakLohia", vbInformation)
                checkSubscription = False
                Exit Function
            Else
                Global_autoEmailCount = Global_autoEmailCount + 1
                checkSubscription = True
            End If
        Else
            If Global_emailextractCount > 5 Then
                MsgBox("you can run automation only 5 times in basic version." & vbNewLine &
                    "buy pro version for $30 (lifetime licence) by emailing deepak9delhi@gmail.com " & vbNewLine &
                    " or visit www.fiverr.com/DeepakLohia", vbInformation)
                checkSubscription = False
                Exit Function
            Else
                Global_emailextractCount = Global_emailextractCount  + 1
                checkSubscription = True
            End If

        End If

    End Function
    Function extractAttachment(Optional autoCheck As Boolean = False,
                               Optional singleEmail As Boolean = False,
                               Optional email As Object = Nothing) As Boolean



        If modMain.checkSubscription(singleEmail) = False Then Exit Function


        Dim attachmentLocation As String = Global_txbx_folderLoc
        Dim Filter_subjectContains As String = Global_txbx_SubjectFilter
        Dim Filter_emailFrom As String = Global_txbx_FromEmailFilter
        Dim Filter_attachmentContains As String = Global_txbx_AttchContans
        Dim fromDate As String = Global_dtFrom
        Dim toDate As String = Global_dtTo
        Dim fileSaveAs As String = Global_cmbx_SaveFileAs
        Dim fileNewName As String = Global_txbxNewName

        Dim log As String
        Dim errFound As Boolean

        If Global_txbxRetriveFrom = "None" Then
            log = "please choose retrive from mailbox"
            If Global_ckbxAutoLoad = False Then MsgBox(log, vbCritical)
            errFound = True
            GoTo chkFailed
        End If

        If Global_ckbDownloadAttachment = False And Global_ckbxSaveasMSG = False Then
            log = "please choose ""download Attachment / Save email as .msg"" or both"
            If Global_ckbxAutoLoad = False Then MsgBox(log, vbInformation)
            errFound = True
            GoTo chkFailed
        End If

        If Len(Dir(Global_txbx_folderLoc, vbDirectory)) = 0 Then
            log = "Invalid attachment save location " & Global_txbx_folderLoc
            If Global_ckbxAutoLoad = False Then MsgBox(log, vbCritical)
            errFound = True
            GoTo chkFailed
        End If

        fromDate = Trim(fromDate)
        toDate = Trim(toDate)

        If (IsDate(fromDate) = True And IsDate(toDate)) = True Then
            If DateValue(fromDate) > DateValue(toDate) = True Then
                log = "From Date: can not be more than To Date: "
                If Global_ckbxAutoLoad = False Then MsgBox(log, vbCritical)
                errFound = True
                GoTo chkFailed
            End If
        End If

chkFailed:

        If errFound = True Then
            extractAttachment = False
            Exit Function
        Else
            extractAttachment = True        'if user has enabled auto check
            If autoCheck = True Then Exit Function
        End If

        Dim OlApp As Outlook.Application
        Dim ns As Outlook.NameSpace
        Dim OlFolderRetrive As Outlook.Folder = Nothing
        Dim OlFolderMove As Outlook.Folder = Nothing
        Dim OlItems As Outlook.Items
        Dim OlMail As Object
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim xAttachmentName As String = ""
        Dim myExt As String
        Dim strFullFlNm As String
        Dim fromDateDt As Date
        Dim toDateDt As Date
        Dim strMLBx() As String
        Dim filterCount() As String
        Dim strAttachmentNm As String = ""
        Dim FilesCounter As Long
        Dim strExtractionSumm As String = ""
        Dim totEmailsExtracted As Long
        Dim totAttachmentsExtracted As Long
        Dim fileSaved As Boolean = False
        Dim msgEmailCount As Long = 0

        Dim xFilePathName As String
        Dim xSenderName As String
        Dim xEmailFrom As String
        Dim xSubject As String
        Dim xAttachmentCount As Long
        Dim fDate As Boolean = False
        Dim tDate As Boolean = False
        Dim xTitle As String
        Dim xSubTitle As String
        Dim filterMatch As Boolean = False
        Dim totalEmailFound As Long
        Dim totalAttachmentCount As Long

        Try
            OlApp = GetObject(, "Outlook.Application")
        Catch
            OlApp = New Outlook.Application()
        End Try

        ns = OlApp.GetNamespace("MAPI")
        'GETTING MAILBOX NAMES 

        If InStr(Global_txbxRetriveFrom, ">>") > 0 Then
            strMLBx = Split(Global_txbxRetriveFrom, ">>")
            Select Case UBound(strMLBx)
                Case 0
                    OlFolderRetrive = ns.Folders(strMLBx(0))
                Case 1
                    OlFolderRetrive = ns.Folders(strMLBx(0)).Folders(strMLBx(1))
                Case Else
                    OlFolderRetrive = ns.Folders(strMLBx(0)).Folders(strMLBx(1)).Folders(strMLBx(2))
            End Select
        Else
            OlFolderRetrive = ns.Folders(Global_txbxRetriveFrom)
        End If

        'setting Move folder
        If InStr(Global_txbxMoveTo, ">>") > 0 Then
            strMLBx = Split(Global_txbxRetriveFrom, ">>")
            Select Case UBound(strMLBx)
                Case 0
                    OlFolderMove = ns.Folders(strMLBx(0))
                Case 1
                    OlFolderMove = ns.Folders(strMLBx(0)).Folders(strMLBx(1))
                Case Else
                    OlFolderMove = ns.Folders(strMLBx(0)).Folders(strMLBx(1)).Folders(strMLBx(2))
            End Select
        End If

        If IsDate(fromDate) Then
            fromDateDt = fromDate
            fDate = True
        End If

        If IsDate(toDate) Then
            toDateDt = toDate
            tDate = True
        End If

        OlItems = OlFolderRetrive.Items
        OlItems.Sort("[ReceivedTime]")

        'LOOP START
        For j = OlItems.Count To 1 Step -1
            fileSaved = False

            If singleEmail = True Then
                OlMail = email
            Else
                OlMail = OlItems(j)
            End If

            If OlMail.class <> 43 Then GoTo skipEmail   'if its not an email

            xSenderName = OlMail.senderName
            xAttachmentCount = OlMail.Attachments.Count

            If Global_ckbxCaseSensitive = False Then
                xEmailFrom = LCase(OlMail.SenderEmailAddress)
                xSubject = LCase(OlMail.Subject)

                Filter_emailFrom = LCase(Filter_emailFrom)
                Filter_subjectContains = LCase(Filter_subjectContains)
                Filter_attachmentContains = LCase(Filter_attachmentContains)
            Else
                xEmailFrom = OlMail.SenderEmailAddress
                xSubject = OlMail.Subject
            End If

            If Global_ckbxExact = False Then
                Filter_emailFrom = "*" & Filter_emailFrom & "*"
                Filter_emailFrom = Filter_emailFrom.Replace(",", "*,*")
                Filter_emailFrom = Filter_emailFrom.Replace("**", "*")

                Filter_subjectContains = "*" & Filter_subjectContains & "*"
                Filter_subjectContains = Filter_subjectContains.Replace(",", "*,*")
                Filter_subjectContains = Filter_subjectContains.Replace("**", "*")

                Filter_attachmentContains = "*" & Filter_attachmentContains & "*"
                Filter_attachmentContains = Filter_attachmentContains.Replace(",", "*,*")
                Filter_attachmentContains = Filter_attachmentContains.Replace("**", "*")
            End If

            'EMAIL FROM FILTER
            filterMatch = False
            Global_txbx_FromEmailFilter = Filter_emailFrom

            If InStr(Filter_emailFrom, ",") > 0 Then
                filterCount = Split(Filter_emailFrom, ",")

                For i = 0 To UBound(filterCount)
                    If xEmailFrom Like filterCount(i) Then    'condition meets
                        filterMatch = True
                        Exit For
                    End If
                Next i
            ElseIf xEmailFrom Like Filter_emailFrom Then
                filterMatch = True
            End If
            If filterMatch = False Then GoTo skipEmail

            'SUBJECT FILTER
            filterMatch = False
            Global_txbx_SubjectFilter = Filter_subjectContains

            If InStr(Filter_subjectContains, ",") > 0 Then
                filterCount = Split(Filter_subjectContains, ",")

                For i = 0 To UBound(filterCount)
                    If xSubject Like filterCount(i) Then    'condition meets
                        filterMatch = True
                        Exit For
                    End If
                Next i
            ElseIf xSubject Like Filter_subjectContains Then
                filterMatch = True
            End If
            If filterMatch = False Then GoTo skipEmail

            'ATTACHMENT CONTAINS FILTER
            filterMatch = False
            Global_txbx_AttchContans = Filter_attachmentContains

            'if there are no attachments
            If Global_ckbxSaveasMSG = False And
                Global_ckbDownloadAttachment = True And
                xAttachmentCount = 0 Then
                GoTo skipEmail
            ElseIf xAttachmentName Like Filter_attachmentContains Then
                filterMatch = True
            End If
            If filterMatch = False Then GoTo skipEmail

            'DATE FILTERS
            If fDate = True And tDate = True Then
                If Format(OlMail.ReceivedTime, "Long Date") >= Format(fromDate, "Long Date") And
                        Format(OlMail.ReceivedTime, "Long Date") <= Format(toDate, "Long Date") Then
                    GoTo labelProceed
                Else
                    GoTo skipEmail
                End If

                'if greater than from date
            ElseIf fDate = True And tDate = False Then
                If Format(OlMail.ReceivedTime, "Long Date") >= Format(fromDate, "Long Date") Then
                    GoTo labelProceed
                Else
                    GoTo skipEmail
                End If

                'if smaller than to date
            ElseIf fDate = False And tDate = True Then
                If Format(OlMail.ReceivedTime, "Long Date") <= Format(toDate, "Long Date") Then
                    GoTo labelProceed
                Else
                    GoTo skipEmail
                End If
            End If

labelProceed:

            If Global_ckbDownloadAttachment = True And xAttachmentCount > 0 Then
                strAttachmentNm = ""
                FilesCounter = 0

                'loop through all attachments
                For i = 1 To xAttachmentCount
                    If Global_ckbxCaseSensitive = False Then xAttachmentName = LCase(xAttachmentName)
                    myExt = Split(xAttachmentName, ".")(1)

                    'ATTACHMENT FILTER 
                    filterMatch = False
                    If InStr(Filter_attachmentContains, ",") > 0 Then
                        filterCount = Split(Filter_attachmentContains, ",")

                        For k = 0 To UBound(filterCount)
                            If xAttachmentName Like filterCount(i) Then    'condition meets
                                filterMatch = True
                                Exit For
                            End If
                        Next k
                    ElseIf xAttachmentName Like Filter_attachmentContains Then
                        filterMatch = True
                    End If
                    If filterMatch = False Then GoTo skipAttachment

                    strFullFlNm = getCleanFilePath(Global_txbx_folderLoc,
                                                        xAttachmentName,
                                                       myExt,
                                                       xSenderName,
                                                       xEmailFrom,
                                                       fileSaveAs,
                                                       fileNewName,
                                                       Global_chkBx_ReplaceExistingFile,
                                                       False)

                    OlMail.Attachments.Item(i).SaveAsFile(strFullFlNm)
                    totalAttachmentCount = totalAttachmentCount + 1
                    'fileSaved = True
                    strAttachmentNm = strAttachmentNm & vbNewLine & "(file name: " & Replace(strFullFlNm, Global_txbx_folderLoc & "\", "") & ")"
                    FilesCounter = FilesCounter + 1
                    totAttachmentsExtracted = totAttachmentsExtracted + 1
skipAttachment:
                Next i
            End If

            'saving message as .msg
            If Global_ckbxSaveasMSG = True Then
                xFilePathName = getCleanFilePath(Global_txbx_folderLoc,
                                                 xSubject,
                                                 "msg",
                                                 xSenderName,
                                                 xEmailFrom,
                                                 fileSaveAs,
                                                 fileNewName,
                                                 Global_chkBx_ReplaceExistingFile,
                                                 True)
                OlMail.SaveAs(xFilePathName)
                msgEmailCount = msgEmailCount + 1
                'fileSaved = True
            End If

            'move attachment if there were some emails
            If (msgEmailCount > 0 Or totalAttachmentCount > 0) And Global_txbxMoveTo <> "None" Then OlMail.Move(OlFolderMove)

            'create summary
            If strAttachmentNm <> "" Then
                If Global_ckbDownloadAttachment = True Then
                    strExtractionSumm = strExtractionSumm & vbNewLine & "*********************" & vbNewLine & "Email Subject: " & OlMail.Subject &
                                vbNewLine & "Attachment(s)= " & FilesCounter & " " & strAttachmentNm
                End If
                totEmailsExtracted = totEmailsExtracted + 1
            End If
skipEmail:
            If singleEmail = True Then GoTo exitLoop : 

            totalEmailFound = totalEmailFound + 1
        Next j

exitLoop:
        'clean up
        OlFolderRetrive = Nothing
        OlFolderMove = Nothing
        OlItems = Nothing
        OlMail = Nothing
        OlApp = Nothing

        If singleEmail = True Then Exit Function 'if one time loop

        xTitle = "Total Email(s) Found :=" & totalEmailFound
        If Global_ckbxSaveasMSG = True Then xTitle = vbNewLine & "Email(s) Saved :=" & msgEmailCount

        If Global_txbxMoveTo = "None" And totalAttachmentCount > 0 Then
            xTitle = xTitle & vbNewLine & "Email with Attachments(s)"
        ElseIf Global_txbxMoveTo <> "None" And totalAttachmentCount > 0 Then
            xTitle = xTitle & vbNewLine & "Email with Attachments(s) & Moved"
        End If

        xSubTitle = ""
        If Global_chkBx_ReplaceExistingFile = True Then xSubTitle = vbNewLine & "(Duplicate(s) Replaced) "
        xSubTitle = xSubTitle & vbNewLine & "Moved to mailbox :=" & Global_txbxMoveTo & vbNewLine & strExtractionSumm


        If msgEmailCount > 0 And totalAttachmentCount > 0 Then
            MsgBox(xTitle & " :=" & totEmailsExtracted & " , Total Attachment(s) :=" & totAttachmentsExtracted & xSubTitle, vbInformation)
            Call Shell("explorer.exe " & Global_txbx_folderLoc, vbNormalFocus)
        ElseIf msgEmailCount > 0 And totalAttachmentCount = 0 Then
            MsgBox(xTitle, vbInformation)
            Call Shell("explorer.exe " & Global_txbx_folderLoc, vbNormalFocus)
        Else
            If Global_ckbxSaveasMSG = True And Global_ckbDownloadAttachment = False Then
                MsgBox("No emails found with selected criteria", vbInformation)
            ElseIf Global_ckbxSaveasMSG = False And Global_ckbDownloadAttachment = True Then
                MsgBox("No attachments found in email(s) with selected criteria", vbInformation)
            Else
                MsgBox("No emails/attachments found with selected criteria", vbInformation)
            End If
        End If

    End Function

    Private Function getCleanFilePath(flPath As String,
                                    flnm As String,
                                    ext As String,
                                    senderName As String,
                                    senderEmail As String,
                                    fileSaveAs As String,
                                    fileNewName As String,
                                    replaceExisting As Boolean,
                                    ignoreDropdown As Boolean) As String

        Dim j As Integer
        Dim strFlNm As String = flnm

        strFlNm = Replace(strFlNm, "." & ext, "") 'removing extention

        If ignoreDropdown = False Then
            If fileSaveAs = "Original Name" Then
                strFlNm = flnm
            ElseIf fileSaveAs = "New Name" Then
                strFlNm = fileNewName
            ElseIf fileSaveAs = "Original_ + New Name" Then
                strFlNm = strFlNm & "_" & fileNewName
            ElseIf fileSaveAs = "New_ + Original Name" Then
                strFlNm = fileNewName & "_" & strFlNm
            ElseIf fileSaveAs = "SenderName" Then
                strFlNm = senderName
            ElseIf fileSaveAs = "SenderName_ + Original Name" Then
                strFlNm = senderName & "_" & flnm
            ElseIf fileSaveAs = "SenderEmail" Then
                strFlNm = senderEmail
            ElseIf fileSaveAs = "SenderEmail_ + Original Name" Then
                strFlNm = senderEmail & "_" & flnm
            Else
                strFlNm = fileNewName
            End If
        End If

        'removing unwanted data
        strFlNm = Replace(strFlNm, ":", "")
        strFlNm = Replace(strFlNm, "<", "")
        strFlNm = Replace(strFlNm, ">", "")
        strFlNm = Replace(strFlNm, "/", "")
        strFlNm = Replace(strFlNm, "\", "")
        strFlNm = Replace(strFlNm, """", "")
        strFlNm = Replace(strFlNm, "?", "")
        strFlNm = Replace(strFlNm, "*", "")

        getCleanFilePath = flPath & "\" & strFlNm & "." & ext

        If replaceExisting = False And Dir(getCleanFilePath) <> "" Then
            j = 1
            Do While Dir(flPath & "\" & strFlNm & "_" & j & "." & ext) <> ""
                If j > 50 Then Exit Do 'to handle infinite loop , we will check file names only  50 times
                j = j + 1
            Loop
            getCleanFilePath = flPath & "\" & strFlNm & "_" & j & "." & ext
        End If

    End Function
End Module
