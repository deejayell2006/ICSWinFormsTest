Imports System.Text.RegularExpressions
Imports System.Xml
Imports TaskManager.Core
Imports TaskManager.Data

Public Class ContactListing
    Private contacts As New List(Of Contact)()

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SQLiteHelper.SetupConnection()

        GetContacts()
    End Sub

    Private Sub ButtonAdd_Click(sender As Object, e As EventArgs) Handles ButtonAdd.Click
        Using addContactForm As New AddContact()
            If addContactForm.ShowDialog() = DialogResult.OK Then
                GetContacts()
            End If
        End Using
    End Sub

    Private Sub ButtonEdit_Click(sender As Object, e As EventArgs) Handles ButtonEdit.Click

        EditContact()

    End Sub

    Private Sub ButtonImport_Click(sender As Object, e As EventArgs) Handles ButtonImport.Click

        ImportFile()

    End Sub

    Private Sub EditContact()

        If grdContacts.SelectedRows.Count() = 0 Then
            Return
        End If

        Dim selectedContact As Contact = grdContacts.SelectedRows(0).DataBoundItem
        Using editContactForm As New EditContact()
            editContactForm.BindDataObject(selectedContact)

            If editContactForm.ShowDialog() = DialogResult.OK Then
                GetContacts()
            End If
        End Using

    End Sub

    Private Sub ImportFile()

        Dim fd = New OpenFileDialog()

        fd.Title = "Import Contacts"
        fd.InitialDirectory = My.Application.Info.DirectoryPath
        fd.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        fd.RestoreDirectory = True

        If fd.ShowDialog() <> DialogResult.OK Then
            Return
        End If

        Dim importFileName = fd.FileName
        Dim newContacts = New List(Of Contact)

        '** Import the elements from the XML file.
        Try
            Dim importDoc = New XmlDocument()
            importDoc.Load(importFileName)

            Dim contactNodes = importDoc.SelectNodes("/contacts/contact")

            For Each contactNode In contactNodes
                Dim newContact As New Contact() With {
                    .Name = contactNode.Item("name").InnerText,
                    .Email = contactNode.Item("email").InnerText,
                    .Phone = contactNode.Item("phone").InnerText
                }
                newContacts.Add(newContact)
            Next

        Catch ex As Exception
            MessageBox.Show($"Could not import file:{vbCrLf}{ importFileName }{vbCrLf}{vbCrLf}The error message is:{vbCrLf}{ ex.Message }",
                            "File Import Failed",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            Return
        End Try

        ImportContacts(newContacts)

        GetContacts()

    End Sub

    Private Sub ImportContacts(contacts As List(Of Contact))

        '** Check for an empty file.
        If contacts Is Nothing Or Not contacts.Any() Then
            MessageBox.Show($"No contact records found in file.",
                            "File Import Failed",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            Return
        End If

        '** Perform some basic validation.
        Dim importErrors = New List(Of String)

        For Each contact In contacts
            '** Email Address is required.
            If String.IsNullOrWhiteSpace(contact.Email) Then
                importErrors.Add($"[Record #{contacts.IndexOf(contact)}] Email address is required.")
            End If
            '** Name is required.
            If String.IsNullOrWhiteSpace(contact.Name) Then
                importErrors.Add($"[Record #{contacts.IndexOf(contact)}] Name is required.")
            End If
            '** Phone must be numeric.
            If Not Regex.IsMatch(contact.Phone, "^[0-9 ]+$") Then
                importErrors.Add($"[Record #{contacts.IndexOf(contact)}] Phone must be digits only.")
            End If
        Next

        If importErrors.Any() Then
            MessageBox.Show($"Could not import the new Contact records.{vbCrLf}{vbCrLf}The data contains the following errors:{vbCrLf}{ String.Join($"{vbCrLf}", importErrors) }",
                            "File Import Failed",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            Return
        End If

        Try
            Using contactRepo As New ContactRepository()
                For Each contact In contacts
                    contactRepo.InsertContact(contact)
                Next
            End Using

        Catch ex As Exception
            MessageBox.Show($"Could not create new Contact records.{vbCrLf}{vbCrLf}The error message is:{vbCrLf}{ ex.Message }",
                            "File Import Failed",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            Return
        End Try

        '** Show success message.
        MessageBox.Show($"Imported {contacts.Count} new Contact records.",
                        "File Import Completed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

    End Sub

    Private Sub GetContacts()
        Using contactRepo As New ContactRepository()
            contacts = contactRepo.GetContacts()
            grdContacts.DataSource = contacts
        End Using
    End Sub
End Class
