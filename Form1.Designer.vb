<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.grpOwner = New System.Windows.Forms.GroupBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnLast = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnPrevious = New System.Windows.Forms.Button()
        Me.btnFirst = New System.Windows.Forms.Button()
        Me.lblID = New System.Windows.Forms.Label()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.txtZip = New System.Windows.Forms.TextBox()
        Me.txtState = New System.Windows.Forms.TextBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtStreet = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.lblPhone = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.dvgPets = New System.Windows.Forms.DataGridView()
        Me.btnPetUpdate = New System.Windows.Forms.Button()
        Me.grpOwner.SuspendLayout()
        CType(Me.dvgPets, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpOwner
        '
        Me.grpOwner.Controls.Add(Me.btnCancel)
        Me.grpOwner.Controls.Add(Me.btnSave)
        Me.grpOwner.Controls.Add(Me.btnLast)
        Me.grpOwner.Controls.Add(Me.btnNext)
        Me.grpOwner.Controls.Add(Me.btnUpdate)
        Me.grpOwner.Controls.Add(Me.btnDelete)
        Me.grpOwner.Controls.Add(Me.btnAdd)
        Me.grpOwner.Controls.Add(Me.btnPrevious)
        Me.grpOwner.Controls.Add(Me.btnFirst)
        Me.grpOwner.Controls.Add(Me.lblID)
        Me.grpOwner.Controls.Add(Me.txtID)
        Me.grpOwner.Controls.Add(Me.txtPhone)
        Me.grpOwner.Controls.Add(Me.txtZip)
        Me.grpOwner.Controls.Add(Me.txtState)
        Me.grpOwner.Controls.Add(Me.txtCity)
        Me.grpOwner.Controls.Add(Me.txtStreet)
        Me.grpOwner.Controls.Add(Me.txtName)
        Me.grpOwner.Controls.Add(Me.lblPhone)
        Me.grpOwner.Controls.Add(Me.Label2)
        Me.grpOwner.Controls.Add(Me.lblName)
        Me.grpOwner.Location = New System.Drawing.Point(12, 12)
        Me.grpOwner.Name = "grpOwner"
        Me.grpOwner.Size = New System.Drawing.Size(776, 317)
        Me.grpOwner.TabIndex = 0
        Me.grpOwner.TabStop = False
        Me.grpOwner.Text = "Owner Information"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(364, 268)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 19
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        Me.btnCancel.Visible = False
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(247, 268)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 18
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        Me.btnSave.Visible = False
        '
        'btnLast
        '
        Me.btnLast.Location = New System.Drawing.Point(526, 217)
        Me.btnLast.Name = "btnLast"
        Me.btnLast.Size = New System.Drawing.Size(75, 23)
        Me.btnLast.TabIndex = 17
        Me.btnLast.Text = "|>"
        Me.btnLast.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(445, 216)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(75, 23)
        Me.btnNext.TabIndex = 16
        Me.btnNext.Text = ">>"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(364, 217)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 15
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(283, 217)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 14
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(202, 217)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 23)
        Me.btnAdd.TabIndex = 13
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnPrevious
        '
        Me.btnPrevious.Location = New System.Drawing.Point(121, 217)
        Me.btnPrevious.Name = "btnPrevious"
        Me.btnPrevious.Size = New System.Drawing.Size(75, 23)
        Me.btnPrevious.TabIndex = 12
        Me.btnPrevious.Text = "<<"
        Me.btnPrevious.UseVisualStyleBackColor = True
        '
        'btnFirst
        '
        Me.btnFirst.Location = New System.Drawing.Point(40, 217)
        Me.btnFirst.Name = "btnFirst"
        Me.btnFirst.Size = New System.Drawing.Size(75, 23)
        Me.btnFirst.TabIndex = 11
        Me.btnFirst.Text = "<|"
        Me.btnFirst.UseVisualStyleBackColor = True
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.Location = New System.Drawing.Point(612, 46)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(21, 13)
        Me.lblID.TabIndex = 10
        Me.lblID.Text = "ID:"
        '
        'txtID
        '
        Me.txtID.Enabled = False
        Me.txtID.Location = New System.Drawing.Point(657, 43)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(100, 20)
        Me.txtID.TabIndex = 9
        Me.txtID.Tag = "ID"
        '
        'txtPhone
        '
        Me.txtPhone.Enabled = False
        Me.txtPhone.Location = New System.Drawing.Point(93, 154)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(314, 20)
        Me.txtPhone.TabIndex = 8
        '
        'txtZip
        '
        Me.txtZip.Enabled = False
        Me.txtZip.Location = New System.Drawing.Point(305, 112)
        Me.txtZip.Name = "txtZip"
        Me.txtZip.Size = New System.Drawing.Size(100, 20)
        Me.txtZip.TabIndex = 7
        '
        'txtState
        '
        Me.txtState.Enabled = False
        Me.txtState.Location = New System.Drawing.Point(199, 112)
        Me.txtState.Name = "txtState"
        Me.txtState.Size = New System.Drawing.Size(100, 20)
        Me.txtState.TabIndex = 6
        '
        'txtCity
        '
        Me.txtCity.Enabled = False
        Me.txtCity.Location = New System.Drawing.Point(93, 112)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(100, 20)
        Me.txtCity.TabIndex = 5
        '
        'txtStreet
        '
        Me.txtStreet.Enabled = False
        Me.txtStreet.Location = New System.Drawing.Point(93, 86)
        Me.txtStreet.Name = "txtStreet"
        Me.txtStreet.Size = New System.Drawing.Size(314, 20)
        Me.txtStreet.TabIndex = 4
        '
        'txtName
        '
        Me.txtName.Enabled = False
        Me.txtName.Location = New System.Drawing.Point(93, 40)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(314, 20)
        Me.txtName.TabIndex = 3
        '
        'lblPhone
        '
        Me.lblPhone.AutoSize = True
        Me.lblPhone.Location = New System.Drawing.Point(6, 154)
        Me.lblPhone.Name = "lblPhone"
        Me.lblPhone.Size = New System.Drawing.Size(81, 13)
        Me.lblPhone.TabIndex = 2
        Me.lblPhone.Text = "Phone Number:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(37, 93)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Address:"
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(47, 40)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(38, 13)
        Me.lblName.TabIndex = 0
        Me.lblName.Text = "Name:"
        '
        'dvgPets
        '
        Me.dvgPets.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dvgPets.Location = New System.Drawing.Point(12, 335)
        Me.dvgPets.Name = "dvgPets"
        Me.dvgPets.Size = New System.Drawing.Size(776, 150)
        Me.dvgPets.TabIndex = 1
        '
        'btnPetUpdate
        '
        Me.btnPetUpdate.Location = New System.Drawing.Point(12, 510)
        Me.btnPetUpdate.Name = "btnPetUpdate"
        Me.btnPetUpdate.Size = New System.Drawing.Size(776, 23)
        Me.btnPetUpdate.TabIndex = 2
        Me.btnPetUpdate.Text = "Update Pet Information"
        Me.btnPetUpdate.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(898, 568)
        Me.Controls.Add(Me.btnPetUpdate)
        Me.Controls.Add(Me.dvgPets)
        Me.Controls.Add(Me.grpOwner)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.grpOwner.ResumeLayout(False)
        Me.grpOwner.PerformLayout()
        CType(Me.dvgPets, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents grpOwner As GroupBox
    Friend WithEvents txtPhone As TextBox
    Friend WithEvents txtZip As TextBox
    Friend WithEvents txtState As TextBox
    Friend WithEvents txtCity As TextBox
    Friend WithEvents txtStreet As TextBox
    Friend WithEvents txtName As TextBox
    Friend WithEvents lblPhone As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents lblName As Label
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents btnLast As Button
    Friend WithEvents btnNext As Button
    Friend WithEvents btnUpdate As Button
    Friend WithEvents btnDelete As Button
    Friend WithEvents btnAdd As Button
    Friend WithEvents btnPrevious As Button
    Friend WithEvents btnFirst As Button
    Friend WithEvents lblID As Label
    Friend WithEvents txtID As TextBox
    Friend WithEvents dvgPets As DataGridView
    Friend WithEvents btnPetUpdate As Button
End Class
