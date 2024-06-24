Imports TaskManager.Core
Imports TaskManager.Data

Public Class EditContact

    Private _contact As Contact

    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        SaveDataObject()
        ClearTextBoxes()
        Close()

    End Sub

    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles ButtonCancel.Click
        ClearTextBoxes()
        Close()
    End Sub

    Public Sub BindDataObject(contact As Contact)

        _contact = contact

        txtName.Text = _contact.Name
        txtEmail.Text = _contact.Email
        txtPhone.Text = _contact.Phone

    End Sub

    Private Sub SaveDataObject()

        '** Update the contact properties from the UI.
        _contact.Name = txtName.Text
        _contact.Email = txtEmail.Text
        _contact.Phone = txtPhone.Text

        '** Save changes.
        Using contactRepo As New ContactRepository()
            contactRepo.UpdateContact(_contact.Id, _contact)
        End Using

    End Sub

    Private Sub ClearTextBoxes()
        txtName.Clear()
        txtEmail.Clear()
        txtPhone.Clear()
    End Sub

End Class