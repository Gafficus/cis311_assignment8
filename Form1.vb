Imports System.Data.SqlClient
Public Class Form1
    Dim DBAdaptOwners As SqlDataAdapter
    Dim DBAdaptPets As SqlDataAdapter
    Dim dsOwners As New DataSet
    Dim dsPets As New DataSet
    'Connectio information for the database
    Public strDBName As String = My.Application.Info.DirectoryPath & "\VetOffice.mdf"
    Public strConnString As String = "Server=(localdb)\ProjectsV13;" &
             "Database=VetOffice;Integrated Security=SSPI;AttachDbFileName=" &
             strDBName
    Dim sqlConn As New SqlConnection(strConnString)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strSQL As String
        strSQL = "Select * from Owners"
        DBAdaptOwners = New SqlDataAdapter(strSQL, strConnString)
        DBAdaptOwners.Fill(dsOwners, "Owners")
        txtID.DataBindings.Add(New Binding("Text", dsOwners, "Owners.Id"))
        txtName.DataBindings.Add(New Binding("Text", dsOwners, "Owners.OwnerName"))
        txtStreet.DataBindings.Add(New Binding("Text", dsOwners, "Owners.OwnerAddress"))
        txtCity.DataBindings.Add(New Binding("Text", dsOwners, "Owners.OwnerCity"))
        txtState.DataBindings.Add(New Binding("Text", dsOwners, "Owners.OwnerState"))
        txtZip.DataBindings.Add(New Binding("Text", dsOwners, "Owners.OwnerZipCode"))
        txtPhone.DataBindings.Add(New Binding("Text", dsOwners, "Owners.OwnerPhoneNumber"))
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnSave_Click            -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when Save is pressesed, it will-
    '- commit the changes made to the dataset to the database.  -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        BindingContext(dsOwners, "Owners").EndCurrentEdit()
        sqlConn.Open()
        DBAdaptOwners.Update(dsOwners, "Owners")
        sqlConn.Close()
        dsOwners.AcceptChanges()
        flipTxtEnabled()
        flipBtnSaveCancelShow()
        flipBtnNavigation()
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnCancel_Click          -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when Cancel is pressed, it will-
    '- cancel the changes made to the dataset.                  -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        BindingContext(dsOwners, "Owners").CancelCurrentEdit()
        dsOwners.RejectChanges()
        BindingContext(dsOwners, "Owners").Position =
            BindingContext(dsOwners, "Owners").Count - 1
        flipTxtEnabled()
        flipBtnSaveCancelShow()
        flipBtnNavigation()
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnNext_Click            -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the >> button is pressed  -
    '- it will naviagte to the next item of the data set and-
    '- in doing so will display the next record of the DB      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        BindingContext(dsOwners, "Owners").Position =
                       (BindingContext(dsOwners, "Owners").Position + 1)
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnLast_Click            -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the |> button is pressed  -
    '- it will naviagte to the last item of the data set and-
    '- in doing so will display the last record of the DB      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        BindingContext(dsOwners, "Owners").Position =
                       (BindingContext(dsOwners, "Owners").Count - 1)
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnPrevious_Click           -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the << button is pressed  -
    '- it will naviagte to the previous item of the data set and-
    '- in doing so will display the first record of the DB      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
        BindingContext(dsOwners, "Owners").Position =
                       (BindingContext(dsOwners, "Owners").Position - 1)
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnFirst_Click           -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the <| button is pressed  -
    '- it will naviagte to the first item of the data set and   -
    '- in doing so will display the first record of the DB      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        BindingContext(dsOwners, "Owners").Position = 0
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnAdd_Click             -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the Add button is clicked -
    '- it will move to a new spot in the data set, enable the   -
    '- text boxes necessary for creating a new owner, disable   -
    '- the navigation buttons, and display the save and cancel  -
    '- buttons.                                                 -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim cmdBuilder As SqlCommandBuilder
        cmdBuilder = New SqlCommandBuilder(DBAdaptOwners)
        DBAdaptOwners.InsertCommand = cmdBuilder.GetInsertCommand
        flipTxtEnabled()
        flipBtnSaveCancelShow()
        flipBtnNavigation()
        BindingContext(dsOwners, "Owners").AddNew()
        'txtID.Enabled = False
        'txtID.Text = BindingContext(dsOwners, "Owners").Count + 1
        'BindingContext(dsOwners, "Owners").Position =
        '               (BindingContext(dsOwners, "Owners").Count - 1)

    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnDelete_Click          -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the Delete button is      -
    '- clicked. This will prompt the user if they want to delete-
    '- currently selected record from the dataset               -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     -
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim cmdBuilder As SqlCommandBuilder
        cmdBuilder = New SqlCommandBuilder(DBAdaptOwners)
        DBAdaptOwners.InsertCommand = cmdBuilder.GetDeleteCommand
        If MessageBox.Show("Are you sure you want to delete?", "Delete Record", MessageBoxButtons.YesNo) _
                        = DialogResult.Yes Then
            BindingContext(dsOwners, "Owners").RemoveAt(BindingContext(dsOwners, "Owners").Position)
            dsOwners.AcceptChanges()
        End If
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnUpdate_Click          -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the Update button is      -
    '- clicked it enable the text boxes necessary for modifying -
    '-an owner, disable the navigation buttons, and display the -
    '-save and cancel buttons                                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     -
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Dim cmdBuilder As SqlCommandBuilder
        cmdBuilder = New SqlCommandBuilder(DBAdaptOwners)
        DBAdaptOwners.InsertCommand = cmdBuilder.GetUpdateCommand
        flipTxtEnabled()
        flipBtnSaveCancelShow()
        flipBtnNavigation()
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: flipTxtEnabled           -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called by other subroutines, it will  -
    '- flip wheterh teh text boxes in the owner group box are   -
    '- editable.                                                -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (none)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub flipTxtEnabled()
        'We do not want to reqrite all teh textbox names
        'grab all text boxes inside the grpOwner
        'Ignore the ID field, that should not be self entered
        Dim textBoxes = From text In grpOwner.Controls
                        Where text.GetType = GetType(TextBox)
                        Select text
        For Each textBox In textBoxes
            textBox.Enabled = Not textBox.Enabled
        Next
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: flipBtnSaveCancelShow    -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called by other subroutines, it will  -
    '- flip what the visible value for btnSave and btnCanel     -
    '- This is used to hide the buttons when not needed         -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (none)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub flipBtnSaveCancelShow()
        btnSave.Visible = Not btnSave.Visible
        btnCancel.Visible = Not btnCancel.Visible
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: flipBtnNavigation        -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called by other subroutines, it will  -
    '- flip what buttons are available to press                 -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (none)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub flipBtnNavigation()
        btnFirst.Enabled = Not btnFirst.Enabled
        btnPrevious.Enabled = Not btnPrevious.Enabled
        btnAdd.Enabled = Not btnAdd.Enabled
        btnDelete.Enabled = Not btnDelete.Enabled
        btnUpdate.Enabled = Not btnUpdate.Enabled
        btnNext.Enabled = Not btnNext.Enabled
        btnLast.Enabled = Not btnLast.Enabled
        btnFirst.Visible = Not btnFirst.Visible
        btnPrevious.Visible = Not btnPrevious.Visible
        btnAdd.Visible = Not btnAdd.Visible
        btnDelete.Visible = Not btnDelete.Visible
        btnUpdate.Visible = Not btnUpdate.Visible
        btnNext.Visible = Not btnNext.Visible
        btnLast.Visible = Not btnLast.Visible
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: flipBtnPetUpdate        -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called by other subroutines, it will  -
    '- flip whether upadte pet information can be clicked       -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (none)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub flipBtnPetUpdate()
        btnPetUpdate.Visible = Not btnPetUpdate.Visible
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnPetUpdate_Click       -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine will add the added data to the pets data -
    '- set added to the pets database.                          -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- cmdBuilder - allows for the use of insert on the database-
    '------------------------------------------------------------
    Private Sub btnPetUpdate_Click(sender As Object, e As EventArgs) Handles btnPetUpdate.Click
        Dim cmdBuilder As SqlCommandBuilder
        cmdBuilder = New SqlCommandBuilder(DBAdaptPets)
        DBAdaptPets.InsertCommand = cmdBuilder.GetInsertCommand
        BindingContext(dsPets, "Pets").EndCurrentEdit()
        sqlConn.Open()
        DBAdaptPets.Update(dsPets, "Pets")
        sqlConn.Close()
        dsPets.AcceptChanges()
        flipBtnPetUpdate()

    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: dvgPetRefresh            -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine will refresh teh dcgPet information      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- strSQL - holds a string of the sql command to use        -
    '------------------------------------------------------------
    Private Sub dvgPetRefresh() Handles txtID.TextChanged
        btnPetUpdate.Visible = False
        Dim strSQL As String
        strSQL = "Select * from Pets Where OwnerId = " & txtID.Text
        dsPets.Clear()
        dsPets.AcceptChanges()
        DBAdaptPets = New SqlDataAdapter(strSQL, strConnString)
        DBAdaptPets.Fill(dsPets, "Pets")
        dvgPets.DataSource = dsPets.Tables("Pets")
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: dvgPetRefresh            -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 30 Mar 2020	                -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine will flip whether pet data can be upadted-
    '- into the database.                                       -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub dvgPets_UserAddedRow(sender As Object, e As DataGridViewRowEventArgs) Handles dvgPets.UserAddedRow
        If Not btnPetUpdate.Visible Then
            flipBtnPetUpdate()
        End If
    End Sub
End Class
