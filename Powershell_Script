# Powershell Project for Help Desk Automation#
############################button1 NEW AD USER#######################
#

##Currently Working on Automated Emails 
## Add Group member Complete - Need to add signature 


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

$button1 = New-Object System.Windows.Forms.Button
$button1.Location = New-Object System.Drawing.Point(25, 75)
$button1.Name = "button1"
$button1.Size = New-Object System.Drawing.Size(250, 41)
$button1.TabIndex = 0
$button1.Text = "New AD User"
$button1.UseVisualStyleBackColor = $true

function new_ad_ui{

    ##Form created from clicking on New AD User Button1
    $NewADAction = New-Object System.Windows.Forms.Form
    $NewADAction.Name = 'New Aduser Form'
    $NewADAction.Text = 'Create User'
    $NewADAction.ClientSize = New-Object System.Drawing.Size(330, 590)
    $NewADAction.StartPosition = 'CenterScreen'


#### Data for Location Combo Box
$locations = @("19 Mile","Benton Harbor","Claw HQ","Claw Kokomo","Claw Warren","Claw Toledo","Groesbeck","Hartford City","JVIS Harper","JVIS HQ","JVIS Shelby 2","JVIS Shelby 3","JVIS Shelby 4","Masonic","Mayco Shelby","Mayco Toledo","Merrill","Nova","SLP","SMI","Stevensville","Transglobal","Valtec","VGE","Warren")
$groups = @("NJT","BH","CLAW","Claw","CLAW","CLAW","NJT","HC","JVIS","JVIS","JVIS","JVIS","JVIS","NJT","MaycoShelby","MaycoToledo","NJT","nova","SLP","SMI","STV","TG","VT","VGE","NJT")
$ouarray = @("NJT","Benton Harbor,OU=JVIS","CLAW", "CLAW", "CLAW","CLAW","NJT","Hartford City","JVIS","JVIS","JVIS","JVIS","JVIS","NJT","Mayco Shelby","Mayco Toledo","NJT","NOVA","SLP","SMI","Stevensville","Transglobal","Valtec","VGE","NJT")
$Bat = @("njt_default.bat","bh_default.bat","claw_default.bat","claw_default.bat","claw_default.bat","Claw_default.bat","njt_default.bat","hc_default.bat","jvis_default.bat","jvis_default.bat","jvis_default.bat", "jvis_default.bat", "jvis_default.bat","njt_default.bat","maycoshelby_default.bat","maycotoledo_default.bat","njt_default.bat","nova_default.bat","slp_default.bat","smi_default.bat","stv_default.bat","tg_default.bat","valtec_default.bat","vge_default.bat","njt_default.bat")




$datatable = New-Object system.Data.DataTable

$col1 = New-Object system.Data.DataColumn "Location",([String])


$datatable.columns.add($col1)
$i=0
foreach($location in $locations){
    $datatable.Rows.Add($location)
}


#create a combobox for Location
$ComboBox = New-Object System.Windows.Forms.ComboBox
$ComboBox.Location  = New-Object System.Drawing.Point(17, 63)
$comboBox.Font = $Font
$comboBox.size = new-object System.Drawing.Size(256, 80)


#clear combo before we bind it
$combobox.Items.Clear()

#bind combobox to datatable
$combobox.ValueMember = "Location"
$combobox.Datasource = $datatable

#add combobox to form
$NewADAction.Controls.Add($combobox)	

### Create a Text Box for First name

$FirstNameTextBox = New-Object System.Windows.Forms.TextBox 
$FirstNameTextBox.Location = New-Object drawing.point (17,127)
$FirstNameTextBox.size = new-object System.Drawing.Size(256,31)
$NewADAction.controls.Add($FirstNameTextBox)

### Create a Text Box for Last name

$LastNameTextBox = New-Object System.Windows.Forms.TextBox 
$LastNameTextBox.Location = New-Object drawing.point (17,189)

$LastNameTextBox.size = new-object System.Drawing.Size(256,31)
$NewADAction.controls.Add($LastNameTextBox)


### Create a Text Box for Username

$UpnTextBox = New-Object System.Windows.Forms.TextBox 
$UpnTextBox.Location = New-Object drawing.point (17,251)
$userinputUP

$UpnTextBox.size = new-object System.Drawing.Size(256,31)
$NewADAction.controls.Add($UpnTextBox)



### Create a DropBox for Department

$Departments = @("Accounting","Engineering","HR","Information Technology","IT Manufacturing","Maintenance","Manufacturing","Materials","Operations","Quality","Safety","Sales","Shipping and Receiving","Tooling")

$i=0


$Deptdatatable = New-Object system.Data.DataTable
$Deptcol1 = New-Object system.Data.DataColumn "Department",([String])
$Deptdatatable.columns.add($Deptcol1)

$i=0
foreach($Department in $Departments){
    $Deptdatatable.Rows.Add($Department)
}

$DeptDropBox = New-Object System.Windows.Forms.Combobox 
$DeptDropBox.Location = New-Object drawing.point (17,313)
$DeptDropBox.Font = $Font

$DeptDropBox.size = new-object System.Drawing.Size(256,33)

$DeptDropBox.Items.Clear()

#bind combobox to Deptdatatable
$DeptDropBox.ValueMember = "Department"
$DeptDropBox.Datasource = $Deptdatatable
$NewADAction.controls.Add($DeptDropBox)



### Create a Text Box for Title

$TitleTextBox = New-Object System.Windows.Forms.TextBox 
$TitleTextBox.Location = New-Object drawing.point (17,377)

$TitleTextBox.size = new-object System.Drawing.Size(256,31)
$NewADAction.controls.Add($TitleTextBox)




########Cancel Button for NewADAction Form
$cancelButton2 = New-Object System.Windows.Forms.Button
$cancelButton2.Location = New-Object System.Drawing.Point(200,500)
$cancelButton2.Size = New-Object System.Drawing.Size(100, 40)
$cancelButton2.Text = 'Cancel'
$cancelButton2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$NewAdAction.CancelButton = $cancelButton2
$NewADAction.Controls.Add($cancelButton2)



### Action for Submit Button on $NEWADAction Form#######
####Submit Button for NewADAction Form
$SubmitButton = New-Object System.Windows.Forms.Button
$SubmitButton.Location = New-Object System.Drawing.Point(12,500)
$SubmitButton.Size = New-Object System.Drawing.Size(100, 40)
$SubmitButton.Text = 'Submit'
#$SubmitButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$NewADAction.Controls.Add($SubmitButton)
$NewADAction.AcceptButton = $SubmitButton




#### LABELS#######

### Putting a label above the dropbox for Location 
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point (12, 35) ### Location of where the label will be
$label.Size = New-Object System.Drawing.Size(160, 25)
$Font = New-Object System.Drawing.Font("Arial",12) ### Formatting text for the label
$label.Font = $Font
$label.Text = "Location" ### Text of label, defined by the parameter that was used when the function is called
$label.ForeColor = 'Black' ### Color of the label text
$NewADAction.Controls.Add($label)

### Putting a label above the TextBox for User First Name 
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(12, 99) ### Location of where the label will be
$label.Size = New-Object System.Drawing.Size(116, 25)
$label.Font = $Font
$label.Text = "First Name" ### Text of label, defined by the parameter that was used when the function is called
$label.ForeColor = 'Black' ### Color of the label text
$NewADAction.Controls.Add($label)

### Putting a label above the TextBox for User Last Name
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(12,161) ### Location of where the label will be
$label.Size = New-Object System.Drawing.Size(115, 25)
$label.Font = $Font
$label.Text = "Last Name" ### Text of label, defined by the parameter that was used when the function is called
$label.ForeColor = 'Black' ### Color of the label text
$NewADAction.Controls.Add($label)



### Putting a label above the text box for username 
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(12,223) ### Location of where the label will be
$label.AutoSize = $True
$label.Font = $Font
$label.Size = New-Object System.Drawing.Size(110, 25)
$label.Text = "Username" ### Text of label, defined by the parameter that was used when the function is called
$label.ForeColor = 'Black' ### Color of the label text
$NewADAction.Controls.Add($label)


### Putting a label above the TextBox for Department
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(12,285) ### Location of where the label will be
$label.AutoSize = $True
$label.Size = New-Object System.Drawing.Size(189, 25)
$label.Font = $Font
$label.Text = "Department" ### Text of label, defined by the parameter that was used when the function is called
$label.ForeColor = 'Black' ### Color of the label text
$NewADAction.Controls.Add($label)


### Putting a label above the TextBox for Title
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(12,349) ### Location of where the label will be
$label.AutoSize = $True
$label.Font = $Font
$label.Size = New-Object System.Drawing.Size(104, 25)
$label.Text = "Title" ### Text of label, defined by the parameter that was used when the function is called
$label.ForeColor = 'Black' ### Color of the label text
$NewADAction.Controls.Add($label)




function get-NewAD {
	$Name = $LastNameTextbox.Text +", "+ $FirstnameTextBox.Text
	$Givenname = $FirstNameTextBox.Text
	$Surname = $LastNameTextbox.Text
	$UserPrincipalName = $UpnTextBox.text+("@na.enterprise.local")
	$Title = $TitleTextBox.text
	$Company = $ComboBox.SelectedValue
    $OU = $ouarray[$ComboBox.SelectedIndex]
	$Path = ("OU=General User Accounts,OU=Users,OU=")+$OU+(",OU=NorthAmerica,DC=na,DC=enterprise,DC=local")
	$pw = ConvertTo-SecureString "Password1" -AsPlainText -Force
	$HomeDirectory = "\\eagle2\users2\%username%"
	$HomeDrive = 'U' 
	$ScriptPath = $Bat[$ComboBox.SelectedIndex]
	$cred = Get-Credential
    $upn = $UpnTextBox.Text
    $department = $DeptDropBox.Selectedvalue
    write-host $Path
	
try{
    New-ADUser -ErrorAction STOP -name $Name -GivenName $Givenname -Surname $Surname -UserPrincipalName $UserPrincipalName -department $department -AccountPassword $pw -title $Title -company $Company -path $Path -SamAccountName $upn  -homedirectory $HomeDirectory -HomeDrive $HomeDrive -ScriptPath $ScriptPath -Credential $cred
} catch {
     [System.Windows.Forms.MessageBox]::Show($_)

  return
}
Try{	Enable-Adaccount -ErrorAction STOP -identity $upn -credential $cred 

	Add-ADGroupMember -ErrorAction STOP -identity "WWW" -members $upn -credential $cred
	set-aduser $upn -ErrorAction STOP -ChangePasswordAtLogon:$True -PasswordNeverExpires $false -Credential $cred

    
    Add-ADGroupMember -ErrorAction STOP -identity $groups[$ComboBox.SelectedIndex] -Members $upn -Credential $cred
    [System.Windows.Forms.MessageBox]::Show("$givenname $surname AD Account Has Been Created")

} catch{
[System.Windows.Forms.MessageBox]::Show($_)}    
    
    #Create U drive 
$fullPath = "\\eagle2\users2\{0}" -f $upn
$driveLetter = "U:"
Try{ 
$User = Get-ADUser -Identity $Upn


if($Null -ne $User) {
    if ($groups[$ComboBox.SelectedIndex] -eq "CLAW") {
        Add-ADGroupMember -identity "CLAW_CTPAT" -members "$upn" -credential $cred
    }
    Set-ADUser $User -HomeDrive $driveLetter -HomeDirectory $fullPath -ea Stop -credential $cred
    $homeShare = New-Item -path $fullPath -ItemType Directory -force -ea Stop 
 
    $acl = Get-Acl $homeShare
 
    $FileSystemRights = [System.Security.AccessControl.FileSystemRights]"FullControl"
    $AccessControlType = [System.Security.AccessControl.AccessControlType]::Allow
    $InheritanceFlags = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
    $PropagationFlags = [System.Security.AccessControl.PropagationFlags]"none"
 
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule ($User.SID, $FileSystemRights, $InheritanceFlags, $PropagationFlags, $AccessControlType)
    $acl.AddAccessRule($AccessRule)
 
    Set-Acl -Path $homeShare -AclObject $acl -ea Stop 
 
    #Write-Host ("HomeDirectory created at {0}" -f $fullPath) 
  }
 }catch{
 write-host $_
}}


$Submitbutton.add_click({
        if(  [string]::IsNullOrEmpty($UpnTextBox.Text) -or [string]::IsNullOrEmpty($TitleTextBox.Text) -or [string]::IsNullOrEmpty($FirstNameTextBox.Text)-or [string]::IsNullOrEmpty($LastNameTextBox.Text) ){
            return
        }
        if(Get-ADUser -Identity $UpnTextBox.text){
            [System.Windows.Forms.MessageBox]::Show("Account Already Exists")
            return 
        }
      
        get-NewAD
    })

    $NewADAction.ShowDialog() 
}

$button1.Add_Click({new_ad_ui})







#
#######################################button2
#

$button2 = New-Object System.Windows.Forms.Button
$button2.Location = New-Object System.Drawing.Point(25, 145)
$button2.Name = "button2"
$button2.Size = New-Object System.Drawing.Size(250, 41)
$button2.TabIndex = 1
$button2.Text = "Termination"
$button2.UseVisualStyleBackColor = $true

function show_term_ui{
##Form created from clicking on Termination Button2
$TermForm = New-Object System.Windows.Forms.Form
$TermForm.Name = 'Termination'
$TermForm.Text = 'Termination'
$TermForm.ClientSize = New-Object System.Drawing.Size(330, 590)
$TermForm.StartPosition = 'CenterScreen'



### Create a Text Box for Term Description

$TermDescTextBox = New-Object System.Windows.Forms.TextBox 
$TermDescTextBox.Location = New-Object drawing.point (7,65)

$TermDescTextBox.size = new-object System.Drawing.Size(289,31)
$TermForm.controls.Add($TermDescTextBox)


### Create a Text Box for Username

$TermUpnTextBox = New-Object System.Windows.Forms.TextBox 
$TermUpnTextBox.Location = New-Object drawing.point (7,115)

$TermUpnTextBox.size = new-object System.Drawing.Size(289,31)
$TermForm.controls.Add($TermUpnTextBox)

###Search Button for Termination Form 
$TermSearchBtn = New-Object System.Windows.Forms.Button
$TermSearchBtn.Location = '12,280'
$TermSearchBtn.Size = '133,36'
$TermSearchBtn.Text = 'Search'
$TermForm.Controls.Add($TermSearchBtn)



### Action for Submit Button on $Termination Form#######
$SubmitButton3 = New-Object System.Windows.Forms.Button
$SubmitButton3.Location = New-Object System.Drawing.Point(12,500)
$SubmitButton3.Size = '100,40'
$SubmitButton3.Text = 'Submit'
#$SubmitButton3.DialogResult = [System.Windows.Forms.DialogResult]::OK
$TermForm.AcceptButton = $SubmitButton3
$TermForm.Controls.Add($SubmitButton3)

########Cancel Button for TermForm Form
$cancelButton3 = New-Object System.Windows.Forms.Button
$cancelButton3.Location = New-Object System.Drawing.Point(210,500)
$cancelButton3.Size = '100,40'
$cancelButton3.Text = 'Cancel'
$cancelButton3.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$TermForm.CancelButton = $cancelButton3
$TermForm.Controls.Add($cancelButton3)

#Temporary Term Checkbox 1
$TermTempTextBox1 = New-Object System.Windows.Forms.checkbox 
$TermTempTextBox1.Location = New-Object drawing.point (40,200)

$TermTempTextBox1.size = new-object System.Drawing.Size(289,31)
$TermForm.controls.Add($TermTempTextBox1)


### Putting a label above the text box for Term username 
$TermUpnlabel = New-Object System.Windows.Forms.Label
$TermUpnlabel.Location = New-Object System.Drawing.Point(7,100) ### Location of where the label will be
$TermUpnlabel.AutoSize = $True
$TermUpnlabel.Size = New-Object System.Drawing.Size(167, 25)
$TermUpnlabel.Font = $Font
$TermUpnlabel.Text = "User Last Name" ### Text of label, defined by the parameter that was used when the function is called
$TermUpnlabel.ForeColor = 'Black' ### Color of the label text
$TermForm.Controls.Add($TermUpnlabel)


#Label for Temporary Term 
$TermTemplabel = New-Object System.Windows.Forms.Label
$TermTemplabel.Location = New-Object System.Drawing.Point(7,185) ### Location of where the label will be
$TermTemplabel.AutoSize = $True
$TermTemplabel.Size = New-Object System.Drawing.Size(167, 25)
$TermTemplabel.Font = $Font
$TermTemplabel.Text = "Is This a Temporary Termination?" ### Text of label, defined by the parameter that was used when the function is called
$TermTemplabel.ForeColor = 'Black' ### Color of the label text
$TermForm.Controls.Add($TermTemplabel)

#Label for TempTerm YES
$TermYeslabel = New-Object System.Windows.Forms.Label
$TermYeslabel.Location = New-Object System.Drawing.Point(7,207) ### Location of where the label will be
$TermYeslabel.AutoSize = $True
$TermYeslabel.Size = New-Object System.Drawing.Size(167, 25)
$TermYeslabel.Font = $Font
$TermYeslabel.Text = "Yes" ### Text of label, defined by the parameter that was used when the function is called
$TermYeslabel.ForeColor = 'Black' ### Color of the label text
$TermForm.Controls.Add($TermYeslabel)


### Label For Term Description
$TermDesclabel = New-Object System.Windows.Forms.Label
$TermDesclabel.Location = New-Object System.Drawing.Point(7,50) ### Location of where the label will be
$TermDesclabel.AutoSize = $True
$TermDesclabel.Font = $Font
$TermDesclabel.Size = New-Object System.Drawing.Size(356, 25)
$TermDesclabel.Text = "Description" ### Text of label, defined by the parameter that was used when the function is called
$TermDesclabel.ForeColor = 'Black' ### Color of the label text
$TermForm.Controls.Add($TermDesclabel)



######Create Term Table to display users#####
$TermTableOutPut = New-Object system.windows.forms.datagridview
$TermTableOutPut.location = New-Object System.Drawing.point(12,320)
$TermTableOutPut.size = New-Object System.Drawing.size (300,150)


$TermForm.controls.Add($TermTableOutPut)
$TermTableOutput.ReadOnly = $TRUE;$TermTableOutput.SelectionMode = 'FullRowSelect';$TermTableOutput.allowusertoaddrows = $false;$TermTableOutput.AutoSizeColumnsMode='AllCells';$TermTableOutput.AllowUserToResizeRows = $false;$TermTableOutput.MultiSelect = $false
$TermTableOutput.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize;$TermTableOutput.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells


####Result from Search Button
Function Get-TermUsers{
    $Filter = $TermUpnTextBox.Text
    $Termusers = @(Get-ADUser -Filter "Name -like '$Filter*'" | Select-Object Name, UserPrincipalName, SamAccountName)

$Term2datatable = [System.Data.DataTable]::new()
$Term2datatable.Columns.AddRange(@("Name","UPN", "SamAccountName"))

foreach($Termuser in $Termusers){
    $Term2datatable.rows.add($Termuser.Name,$Termuser.UserPrincipalName, $Termuser.SamAccountName)
}
    $TermTableOutPut.DataSource = $Term2datatable
    $termTableOutput.SelectionMode = 'FullRowSelect'
}


$TermSearchBtn.add_click({Get-TermUsers})



Function Get-TermSubmitBtn{
$cred = get-credential 
$Termpw = ConvertTo-SecureString "Evict2022" -AsPlainText -Force
#$TermPath =$Chosensite[$TermComboBox.SelectedIndex] 
#$TermOU = $Termouarray[$TermComboBox.SelectedIndex]
$SelectedRow = $TermTableOutPut.SelectedRows.Cells[2].Value
$SelectedRowMessage = $TermTableOutPut.SelectedRows.Cells[0].Value

$desc = ' '
if($NULL -ne $TermDescTextBox.text){
    $desc = $TermDescTextBox.text
}
$userou = get-aduser -Identity $SelectedRow | Select-Object DistinguishedName

$userou = ($userou.DistinguishedName -split ",",3)[2]
$newou = $userou -replace "General User Accounts","Disabled Accounts"
$userGroups = get-aduser -Identity $SelectedRow -properties MemberOf

#write-host $newou

try{
    Set-ADAccountPassword -Identity $SelectedRow -NewPassword $Termpw -credential $cred -ErrorAction Stop
    Set-ADuser -identity $SelectedRow -CannotChangePassword:$true -PasswordNeverExpires:$true -Description $desc -Credential $cred -ErrorAction Stop
    Disable-Adaccount -Identity $SelectedRow -credential $cred -ErrorAction Stop
     
if($TermTempTextBox1.Checked -ne $true){
Foreach ($group in $userGroups.MemberOf) { Remove-ADGroupMember -Identity $group -Members $SelectedRow -credential $cred -confirm: $False -ErrorAction Stop  } }
} catch {[System.Windows.Forms.MessageBox]::Show($_)
}
Try{
Get-ADUser -Identity $SelectedRow | Move-ADObject -TargetPath $newou -Credential $cred -ErrorAction Stop
[System.Windows.Forms.MessageBox]::Show("$SelectedRowMessage has been termed")
} catch {[System.Windows.Forms.MessageBox]::Show($_)

}


}
$SubmitButton3.add_click({

if(  [string]::IsNullOrEmpty($TermUpnTextBox.text) -or [string]::IsNullOrEmpty($TermDescTextBox.text)){
return
}
Get-TermSubmitBtn
})
$TermForm.ShowDialog()
}
$button2.Add_Click({show_term_ui})


#
# button3######################### Add Group
#
$button3 = New-Object System.Windows.Forms.Button
$button3.Location = New-Object System.Drawing.Point(25, 221)
$button3.Name = "button3"
$button3.Size = New-Object System.Drawing.Size(250, 41)
$button3.TabIndex = 2
$button3.Text = "Add Group Member"
$button3.UseVisualStyleBackColor = $true




Function Show_Group_ui{
##Form created from clicking on Add group Button3
$AddGroupForm = New-Object System.Windows.Forms.Form
$AddGroupForm.Name = 'AddGroup'
$AddGroupForm.Text = 'AddGroup'
$AddGroupForm.ClientSize = New-Object System.Drawing.Size(330,590)
$AddGroupForm.StartPosition = 'CenterScreen'



### User name textbox for Add Group
$AddGroupUpnTextBox = New-Object System.Windows.Forms.TextBox 
$AddGroupUpnTextBox.Location = New-Object drawing.point (12,50)

$AddGroupUpnTextBox.size = new-object System.Drawing.Size(289,31)
$AddGroupForm.controls.Add($AddGroupUpnTextBox)

### Textbox for Group
$GroupTextBox = New-Object System.Windows.Forms.TextBox 
$GroupTextBox.Location = New-Object drawing.point (12,110)

$GroupTextBox.size = new-object System.Drawing.Size(289,31)
$AddGroupForm.controls.Add($GroupTextBox)



####Search Button for GroupUser 
$GroupUserSearchBtn = New-Object System.Windows.Forms.Button
$GroupUserSearchBtn.Size = New-Object System.Drawing.Size(133,36)
$GroupUserSearchBtn.AutoSize = $true
$GroupUserSearchBtn.Location = New-Object System.Drawing.Point(12,140)
$GroupUserSearchBtn.Text = 'Search User'
$AddGroupForm.Controls.Add($GroupUserSearchBtn)

####Search Button for Group
$GroupSearchBtn = New-Object System.Windows.Forms.Button
$GroupSearchBtn.Size = New-Object System.Drawing.Size(133,36)
$GroupSearchBtn.Location = New-Object System.Drawing.Point(12,349)
$GroupSearchBtn.AutoSize = $true
$GroupSearchBtn.Text = 'Search Group'
$AddGroupForm.Controls.Add($GroupSearchBtn)




####Submit Button for Add Group Form
$SubmitButton3 = New-Object System.Windows.Forms.Button
$SubmitButton3.Location = New-Object System.Drawing.Point(12,520)
$SubmitButton3.Size = New-Object System.Drawing.Size(100,40)
$SubmitButton3.Text = 'Submit'
$AddGroupForm.AcceptButton = $SubmitButton3
$AddGroupForm.Controls.Add($SubmitButton3)


########Cancel Button for Add Group Form
$AGcancelButton3 = New-Object System.Windows.Forms.Button
$AGcancelButton3.Location = New-Object System.Drawing.Point(210,520)
$AGcancelButton3.Size = New-Object System.Drawing.Size(100,40)
$AGcancelButton3.Text = 'Cancel'
$AGcancelButton3.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$AddGroupForm.CancelButton = $AGcancelButton3
$AddGroupForm.Controls.Add($AGcancelButton3)



### Putting a label above for username add group 
$Addgroupupnlabel = New-Object System.Windows.Forms.Label
$Addgroupupnlabel.Location = New-Object System.Drawing.Point(7,14) ### Location of where the label will be
$Addgroupupnlabel.AutoSize = $True
$Addgroupupnlabel.Font = $Font
$Addgroupupnlabel.Size = New-Object System.Drawing.Size(167, 25)
$Addgroupupnlabel.Text = "User Last Name" ### Text of label, defined by the parameter that was used when the function is called
$Addgroupupnlabel.ForeColor = 'Black' ### Color of the label text
$AddGroupForm.Controls.Add($Addgroupupnlabel)



### Putting a label above for Group Name 
$GroupNamelabel = New-Object System.Windows.Forms.Label
$GroupNamelabel.Location = New-Object System.Drawing.Point(7,80) ### Location of where the label will be
$GroupNamelabel.AutoSize = $True
$GroupNamelabel.Font = $Font
$GroupNamelabel.Size = New-Object System.Drawing.Size(128, 25)
$GroupNamelabel.Text = "Group" ### Text of label, defined by the parameter that was used when the function is called
$GroupNamelabel.ForeColor = 'Black' ### Color of the label text
$AddGroupForm.Controls.Add($GroupNamelabel)





######Create Table to display users for AddGroup#####
$GroupGridView = New-Object system.windows.forms.datagridview
$GroupGridView.location = New-Object System.Drawing.point(12,179)
$GroupGridView.size = New-Object System.Drawing.size (300,168)


$AddGroupForm.controls.Add($GroupGridView)
$GroupGridView.ReadOnly = $TRUE;$GroupGridView.SelectionMode = 'FullRowSelect';$GroupGridView.allowusertoaddrows = $false;$GroupGridView.AutoSizeColumnsMode='AllCells';$GroupGridView.AllowUserToResizeRows = $false;$GroupGridView.MultiSelect = $false
$GroupGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize;$GroupGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells

#Table to display groups

$OutPutGroup = New-Object system.windows.forms.datagridview
$OutPutGroup.location = New-Object System.Drawing.point(12,388)
$OutPutGroup.size = New-Object System.Drawing.size (300,120)


$AddGroupForm.controls.Add($OutPutGroup)
$OutPutGroup.ReadOnly = $TRUE;$OutPutGroup.SelectionMode = 'FullRowSelect';$OutPutGroup.allowusertoaddrows = $false;$OutPutGroup.AutoSizeColumnsMode='AllCells';$OutPutGroup.AllowUserToResizeRows = $false;$OutPutGroup.MultiSelect = $false
$OutPutGroup.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize;$OutPutGroup.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells



Function Get-GroupUser{
    $Filter2 = $AddGroupUpnTextBox.Text
    $GroupUsers = @(Get-ADUser -Filter "Name -like '$Filter2*'" | Select-Object Name, UserPrincipalName, SamAccountName, GivenName, Surname )



$Groupdatatable = [System.Data.DataTable]::new()
$Groupdatatable.Columns.AddRange(@("Name","UPN", "SamAccountName", "FirstName", "LastName"))

foreach($GroupUser in $GroupUsers){
    $Groupdatatable.rows.add($GroupUser.Name,$GroupUser.UserPrincipalName, $GroupUser.SamAccountName, $GroupUser.GivenName, $GroupUser.Surname)
}
    $GroupGridView.DataSource = $Groupdatatable
    $GroupGridView.SelectionMode = 'FullRowSelect'
}

$GroupUserSearchBtn.add_Click({Get-GroupUser})


Function Get-Group{
    $Filter3 = $GroupTextBox.Text
    $Groups = @(Get-ADGroup -Filter "Name -like '$Filter3*'" | Select-Object Name)

$GroupdatatableGroup = [System.Data.DataTable]::new()
$GroupdatatableGroup.Columns.AddRange(@("Name"))

foreach($GroupUserGroup in $Groups){
    $GroupdatatableGroup.rows.add($GroupUserGroup.Name)
}
    $OutPutGroup.DataSource = $GroupdatatableGroup
    $OutPutGroup.SelectionMode = 'FullRowSelect'
}

$GroupSearchBtn.add_Click({Get-Group})


###Action for Group Submit button
Function get-SubmitGroup{

$GU = $GroupGridView.SelectedRows.Cells[2].Value
$Gr = $OutPutGroup.SelectedRows.cells[0].Value
$GUM = $GroupGridView.SelectedRows.Cells[0].Value

$GUMFN = $GroupGridView.SelectedCells[3].value
$GUMLN = $GroupGridView.SelectedCells[4].value

$User = $GUMLN + $GUMFN



Try{
Add-ADGroupMember -ErrorAction STOP -identity "$Gr" -members $GU -credential $cred
[System.Windows.Forms.MessageBox]::Show("$GUM Has Been Added to $Gr, Click Ok To Send Email")

Email
}catch{
[System.Windows.Forms.MessageBox]::Show($_)}
}
Function Email {
    $ol = New-Object -comObject Outlook.Application
    
    # Create the new email
    $mail = $ol.CreateItem(0)

    # Optional, set the subject
    $mail.Subject = "Regarding Your Access Request Ticket"

    $signaturesPath = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Signatures')

    # Optional, set the body
    # Check if the folder exists
    if (Test-Path $signaturesPath) {
        # Get the first HTML file in the signatures folder (assumes it's your default signature)
        $signatureFile = Get-ChildItem $signaturesPath -Filter *.htm | Select-Object -Skip 1

        # Read the content of the signature file
        $signatureContent = Get-Content $signatureFile.FullName -Raw



        # Create the email body with the signature as HTML
 # Specify the path to your image
        $imagePath = "C:\Users\stead\AppData\Roaming\Microsoft\Signatures\image001.png"  # Replace with the actual path to your image

        # Convert the image to Base64
        $imageBytes = [System.IO.File]::ReadAllBytes($imagePath)
        $base64Image = [Convert]::ToBase64String($imageBytes)
        $imageSrc = "data:image/jpeg;base64,$base64Image"

        # Create the email body with the embedded image
        if ($Gr -eq "njt"){

        $mail.HTMLBody = @"
        <html>
            <body>
                <p>Hello $GUMFN $GumLN,</p>
                <p>You now have access to the Mayco drive for the requested folder. Please log off and log back in before accessing. If you have any problems or need this mapped, please respond to this email with a good time I can remote in.</p>
                <img src="$imageSrc" alt="Your Image Alt Text">
                $($signatureContent)
            </body>
        </html>
"@}else{

$mail.HTMLBody = @"
        <html>
            <body>
                <p>Hello $GUMFN $GumLN,</p>
                <p>You now have access to $Gr. Please log off and log back in before accessing. If you have any problems or need this mapped, please respond to this email with a good time I can remote in.</p>
                <img src="$imageSrc" alt="Your Image Alt Text">
                $($signatureContent)
            </body>
        </html>
"@}
        
# Set the recipient
        $mail.To = "$GUMLN, $GUMFN" 

        # Get the new email object
        $inspector = $mail.GetInspector

        # Bring the message window to the front
        $inspector.Activate()
        $mail.display
    } else {
        Write-Host "Outlook signatures folder not found."
    }
}



$SubmitButton3.add_click({

if( [string]::IsNullOrEmpty($AddGroupUpnTextBox.Text) -or 
[string]::IsNullOrEmpty($GroupTextBox.Text)){
return
}

get-SubmitGroup})

$AddGroupForm.ShowDialog()
}

$button3.Add_Click({Show_Group_ui})






#
# button4
#

$button4 = New-Object System.Windows.Forms.Button
$button4.Location = New-Object System.Drawing.Point(25, 301)
$button4.Name = "button4"
$button4.Size = New-Object System.Drawing.Size(250, 41)
$button4.TabIndex = 3
$button4.Text = "Reset Password"
$button4.UseVisualStyleBackColor = $true

Function Show_ResetPW_ui{

##Form created from clicking on Reset Password Button4
$ResetPWForm = New-Object System.Windows.Forms.Form
$ResetPWForm.Name = 'Reset PW'
$ResetPWForm.Text = 'Reset PW'
$ResetPWForm.ClientSize = New-Object System.Drawing.Size(350, 610)
$ResetPWForm.StartPosition = 'CenterScreen'
$ResetPWForm.AutoScale = $false
$ResetPWForm.AutoSize = $false


####Text Box for Reset PW User
$UpnResetPWTextBox = New-Object System.Windows.Forms.TextBox 
$UpnResetPWTextBox.Location = New-Object drawing.point (12, 66)

$UpnResetPWTextBox.size = new-object System.Drawing.Size(289,31)
$ResetPWForm.controls.Add($UpnResetPWTextBox)



####Search Button for ResetPW Form
$ResetPWSearchBtn = New-Object System.Windows.Forms.Button
$ResetPWSearchBtn.Location = New-Object System.Drawing.Point(12,100)
$ResetPWSearchBtn.Size = New-Object System.Drawing.Size(133, 36)
$ResetPWSearchBtn.Text = 'Search'
$ResetPWForm.controls.Add($ResetPWSearchBtn)



### Putting a label above for Reset PW 
$UpnResetPWLabel = New-Object System.Windows.Forms.Label
$UpnResetPWLabel.Location = New-Object System.Drawing.Point(7,38) ### Location of where the label will be
$UpnResetPWLabel.AutoSize = $True
$UpnResetPWLabel.Font = $Font
$UpnResetPWLabel.Size = New-Object System.Drawing.Size(167, 25)
$UpnResetPWLabel.Text = "User Last Name" ### Text of label, defined by the parameter that was used when the function is called
$UpnResetPWLabel.ForeColor = 'Black' ### Color of the label text
$ResetPWForm.Controls.Add($UpnResetPWLabel)




$ResetPWGroupBox = New-Object System.Windows.Forms.GroupBox
$ResetPWGroupBox.Location = New-Object System.Drawing.Point(12,345)
$ResetPWGroupBox.Autosize = $true
$ResetPWGroupBox.Size = New-Object System.Drawing.Size(300,180)
$ResetPWGroupBox.text = "Select Command"


#Create the collection of radio buttons for PW GroupBox
$PWRadioButton1 = New-Object System.Windows.Forms.RadioButton
$PWRadioButton1.Location = New-Object System.Drawing.Point(6,35)
$PWRadioButton1.Size = New-Object System.Drawing.Size(165,29)
$PWRadioButton1.Text = "Reset to Password1, Change at Login"


$PWRadioButton2 = New-Object System.Windows.Forms.RadioButton
$PWRadioButton2.Location = New-Object System.Drawing.Point(6,68)
$PWRadioButton2.Size = New-Object System.Drawing.Size(165,29)
$PWRadioButton2.Text = "Enable Account"



$PWRadioButton3 = New-Object System.Windows.Forms.RadioButton
$PWRadioButton3.Location = New-Object System.Drawing.Point(6,100)
$PWRadioButton3.Size = New-Object System.Drawing.Size(165,29)
$PWRadioButton3.Text = "Reset to Password1, Cannot Change at Login"



$PWRadioButton4 = New-Object System.Windows.Forms.RadioButton
$PWRadioButton4.Location = New-Object System.Drawing.Point(6,133)
$PWRadioButton4.Size = New-Object System.Drawing.Size(165,29)
$PWRadioButton4.Text = "Disable account"


$PWRadioButton5 = New-Object System.Windows.Forms.RadioButton
$PWRadioButton5.Location = New-Object System.Drawing.Point(6,165)
$PWRadioButton5.Size = New-Object System.Drawing.Size(165,29)
$PWRadioButton5.Text = "Reset to Evict2023"

# Add all the GroupBox controls on one line
$ResetPWGroupBox.Controls.AddRange(@($PWRadiobutton1,$PWRadioButton2, $PWRadioButton3, $PWRadioButton4, $PWRadioButton5))
$ResetPWForm.controls.Add($ResetPWGroupBox)


####Submit Button for Reset PW
$SubmitButton4 = New-Object System.Windows.Forms.Button
$SubmitButton4.Location = New-Object System.Drawing.Point(15,558)
$SubmitButton4.Size = '100,40'
$SubmitButton4.Text = 'Submit'
#$SubmitButton4.DialogResult = [System.Windows.Forms.DialogResult]::OK
$ResetPWForm.AcceptButton = $SubmitButton4
$ResetPWForm.Controls.Add($SubmitButton4)

########Cancel Button for Reset PW
$cancelButton4 = New-Object System.Windows.Forms.Button
$cancelButton4.Location = New-Object System.Drawing.Point(210,558)
$cancelButton4.Size = '100,40'
$cancelButton4.Text = 'Cancel'
#$cancelButton4.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$ResetPWForm.CancelButton = $cancelButton4
$ResetPWForm.Controls.Add($cancelButton4)



######Create ResetPW Table to display users#####
$ResetPWGridView = New-Object system.windows.forms.datagridview
$ResetPWGridView.location = New-Object System.Drawing.point(12,150)
$ResetPWGridView.size = New-Object System.Drawing.size (300,203)


$ResetPWForm.controls.Add($ResetPWGridView)
$ResetPWGridView.ReadOnly = $TRUE;$ResetPWGridView.SelectionMode = 'FullRowSelect';$ResetPWGridView.allowusertoaddrows = $false;$ResetPWGridView.AutoSizeColumnsMode='AllCells';$ResetPWGridView.AllowUserToResizeRows = $false;$ResetPWGridView.MultiSelect = $false
$ResetPWGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize;$ResetPWGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
     



Function Get-PWResetUser{
    $Filter3 = $UpnResetPWTextBox.Text
    
    if($NULL -eq $Filter3 -or $Filter3 -eq ''){
        return
    }
    $PWUsers = @(Get-ADUser -Filter "Name -like '$Filter3*'" | Select-Object Name, UserPrincipalName, SamAccountName)

$ResetPWdatatable = [System.Data.DataTable]::new()
$ResetPWdatatable.Columns.AddRange(@("Name","UPN", "SamAccountName"))

foreach($PWUser in $PWUsers){
    $ResetPWdatatable.rows.add($PWUser.Name,$PWUser.UserPrincipalName, $PWUser.SamAccountName)
}
    $ResetPWGridView.DataSource = $ResetPWdatatable
    $ResetPWGridView.SelectionMode = 'FullRowSelect'
}

$ResetPWSearchBtn.add_Click({Get-PWResetUser})



###Action for Radio Buttons


Function ResetPassword{
$RPWUSer = $ResetPWGridView.SelectedRows.Cells[2].Value
$RPWUSerM = $ResetPWGridView.SelectedRows.Cells[0].Value

if($NULL -eq $RPWUSer)
{
    return
}
$Rpw = ConvertTo-SecureString "Password1" -AsPlainText -Force
$cred = Get-Credential
$Evict = ConvertTo-SecureString "Evict2023" -AsPlainText -Force
try{
if($PWRadioButton1.Checked){
Set-ADAccountPassword -identity "$RPWUSer" -newpassword $Rpw -credential $cred -ErrorAction Stop
            Set-ADUser -identity "$RPWUSer" -ChangePasswordAtLogon $True -PasswordNeverExpires $false -credential $cred
            [System.Windows.Forms.MessageBox]::Show("$RPWUSerM has been reset, user must change at login")
            
   }


if($PWRadioButton2.checked){
Unlock-ADAccount -identity "$RPWUSer" -credential $cred -ErrorAction STOP
            Enable-ADAccount -identity "$RPWUSer" -credential $cred
            [System.Windows.Forms.MessageBox]::Show("$RPWUSerM has been enabled and unlocked")
   }

if($PWRadioButton3.checked){
Set-ADAccountPassword -identity "$RPWUSer" -NewPassword $Rpw -credential $cred -ErrorAction STOP 
            Set-ADUser -identity "$RPWUSer" -ChangePasswordAtLogon $False -PasswordNeverExpires $false -credential $cred
            [System.Windows.Forms.MessageBox]::Show("$RPWUSerM has been reset to Password1, user cannot change at login")
}


if($PWRadioButton4.checked){
    Disable-ADAccount -identity "$RPWUSer" -credential $cred -ErrorAction STOP
    [System.Windows.Forms.MessageBox]::Show("$RPWUSerM has been disabled")
}


if($PWRadioButton5.checked){
    set-adaccountpassword -identity "$RPWUSer" -NewPassword $Evict -credential $cred -ErrorAction STOP
    Set-ADUser -Identity "$RPWUSer" -ChangePasswordAtLogon $False -PasswordNeverExpires $true -Credential $cred
    [System.Windows.Forms.MessageBox]::Show("$RPWUSerM has been reset to Evict2023")
}
}catch{
[System.Windows.Forms.MessageBox]::Show($_)
}}


$SubmitButton4.add_click({
if(  [string]::IsNullOrEmpty($UpnResetPWTextBox.text)){
return}
ResetPassword

})


$ResetPWForm.ShowDialog()
}
$button4.Add_Click({Show_ResetPW_ui})




   
#
# Form1
#
$Form1 = New-Object System.Windows.Forms.Form
#$Form1.ClientSize = New-Object System.Drawing.Size(400, 600)
$Form1.ClientSize = New-Object System.Drawing.Size(311, 520)
$Form1.Controls.Add($button4)
$Form1.Controls.Add($button3)
$Form1.Controls.Add($button2)
$Form1.Controls.Add($button1)
$Form1.Name = "Form1"
$Form1.Text = "Form1"
$Form1.StartPosition = 'CenterScreen'
$form1.AutoScale = $true



$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(25, 402)
$cancelButton.Size = New-Object System.Drawing.Size(250, 41)
$cancelButton.Text = 'Cancel'

#$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Form1.CancelButton = $cancelButton
$Form1.Controls.Add($cancelButton)



$Form1.Add_Shown({$Form1.Activate()})
$Form1.ShowDialog()







 









