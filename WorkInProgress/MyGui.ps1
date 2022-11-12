Add-Type -assembly System.Windows.Forms

$main_form = New-Object System.Windows.Forms.Form

$main_form.Text ='GUI for my PoSh script'

$main_form.Width = 600

$main_form.Height = 400

$main_form.AutoSize = $true

$main_form.ShowDialog()





$Label = New-Object System.Windows.Forms.Label

$Label.Text = "AD users"

$Label.Location  = New-Object System.Drawing.Point(0,10)

$Label.AutoSize = $true

$main_form.Controls.Add($Label)






$ComboBox = New-Object System.Windows.Forms.ComboBox

$ComboBox.Width = 300

$Users = get-aduser -filter * -Properties SamAccountName

Foreach ($User in $Users)

{

$ComboBox.Items.Add($User.SamAccountName);

}

$ComboBox.Location  = New-Object System.Drawing.Point(60,10)

$main_form.Controls.Add($ComboBox)




$Label2 = New-Object System.Windows.Forms.Label

$Label2.Text = "Last Password Set:"

$Label2.Location  = New-Object System.Drawing.Point(0,40)

$Label2.AutoSize = $true

$main_form.Controls.Add($Label2)

$Label3 = New-Object System.Windows.Forms.Label

$Label3.Text = ""

$Label3.Location  = New-Object System.Drawing.Point(110,40)

$Label3.AutoSize = $true

$main_form.Controls.Add($Label3)



$Button = New-Object System.Windows.Forms.Button

$Button.Location = New-Object System.Drawing.Size(400,10)

$Button.Size = New-Object System.Drawing.Size(120,23)

$Button.Text = "Check"

$main_form.Controls.Add($Button)



$Button.Add_Click(

{

$Label3.Text =  [datetime]::FromFileTime((Get-ADUser -identity $ComboBox.selectedItem -Properties pwdLastSet).pwdLastSet).ToString('MM dd yy : hh ss')

}

)



