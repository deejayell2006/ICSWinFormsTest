Imports System.ComponentModel
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

    Private Sub GetContacts()
        Using contactRepo As New ContactRepository()
            contacts = contactRepo.GetContacts()
            grdContacts.DataSource = contacts
        End Using
    End Sub
End Class
