' MIT License
' 
' Copyright (c) [Year] [Your Name or Your Organization]
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Option Explicit




Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim varEntryIDs, varEntryID
    Dim objItem As Object
    Dim objMeetingItem As Outlook.MeetingItem
    Dim objApptItem As Outlook.AppointmentItem
    Dim objRecipients As Outlook.Recipients
    Dim objRecipient As Outlook.Recipient
    Dim oMembers As Outlook.AddressEntries
    Dim oMember As Outlook.AddressEntry
    Dim sourceEmail As String

    Dim found As Boolean
    Dim recipientFilePath As String
    Dim olAddrEntry As Outlook.AddressEntry
    recipientFilePath = "C:\recipient_" & Format(Now(), "yyyy_MM_dd_hh_mm_ss") & ".txt"
    Dim i As Integer, j As Integer, dRecCnt As Integer, dMemCnt As Integer
    Dim targetEmail As String
    targetEmail = "tgt@tgt.com"
    sourceEmail = "src@src.com"

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim fileStream As TextStream

    ' Here the actual file is created And opened For write access
    Set fileStream = fso.CreateTextFile(recipientFilePath)

    varEntryIDs = Split(EntryIDCollection, ",")
    
    For Each varEntryID In varEntryIDs
        
        If Application.Session.GetItemFromID(varEntryID) Is Nothing Then
        
            'Sample step
            sourceEmail = sourceEmail
        
        Else
            
            Set objItem = Application.Session.GetItemFromID(varEntryID)
    
            'If meetings Then only try To copy
            If TypeName(objItem) = "MeetingItem" Then
            
                'Setting the value as False initially
                found = False
    
                Set objMeetingItem = objItem
                Set objRecipients = objMeetingItem.Recipients
                
                For Each objRecipient In objRecipients
                    Set olAddrEntry = objRecipient.AddressEntry
                    
                    If olAddrEntry.AddressEntryUserType = olExchangeDistributionListAddressEntry Then
                        ' Handle distribution list (for simplicity, only top-level addresses)
                        If GetDistributionListMembers(olAddrEntry, sourceEmail) = True Then
                            found = True
                            Exit For
                        End If
                        
                    Else
                        If LCase(olAddrEntry.GetExchangeUser().PrimarySmtpAddress) = LCase(sourceEmail) Then
                            found = True
                            Exit For
                        End If
                    End If
                Next objRecipient
                
                If found Then
                    Set objApptItem = objMeetingItem.GetAssociatedAppointment(True)
                    ' Call a custom Function Or handle copying here
                    Call CopyToCalendar(objApptItem, targetEmail)
                End If
                
                
            End If
        End If
    Next varEntryID
    ' Close it, so it is Not locked anymore
    fileStream.Close

    ' Explicitly setting objects To Nothing should Not be necessary in most cases, but If
    ' you're writing macros For Microsoft Access, you may want To uncomment the following
    ' two lines (see https://stackoverflow.com/a/517202/2822719 For details):
    Set fileStream = Nothing
    Set fso = Nothing


End Sub

Sub CopyToCalendar(ByVal appt As Outlook.AppointmentItem, targetEmail As String)
    ' This is just a placeholder Function To show how you might handle the copying
    Dim newItem As Outlook.AppointmentItem
    On Error GoTo ErrorHandler
        Set newItem = appt.Copy
        newItem.Subject = newItem.Subject
        newItem.Sensitivity = olPrivate
        newItem.Move Application.GetNamespace("MAPI").Folders(targetEmail).Folders("Calendar")
        MsgBox "Copied Meeting - " & newItem.Subject, vbExclamation
     Exit Sub

ErrorHandler:
        ' Code To handle the error goes here
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
     Resume Next

End Sub

Function GetDistributionListMembers(distListEntry As Outlook.AddressEntry, sourceEmail As String) As Boolean
    Dim olMembers As Outlook.AddressEntries
    Dim olMember As Outlook.AddressEntry
    
    Set olMembers = distListEntry.GetExchangeDistributionList.Members
    GetDistributionListMembers = False
    
    For Each olMember In olMembers
        If olMember.AddressEntryUserType = olExchangeDistributionListAddressEntry Then
            If GetDistributionListMembers(olMember, sourceEmail) = True Then
                GetDistributionListMembers = True
                Exit For
            End If
        Else
            If LCase(olMember.GetExchangeUser().PrimarySmtpAddress) = LCase(sourceEmail) Then
                GetDistributionListMembers = True
                Exit For
            End If
        End If
    Next olMember
End Function

