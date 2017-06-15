Add-Type -AssemblyName PresentationFramework

Function Get-SiteCode {
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$title = 'Device renaming script'
$message = 'Enter site code'
$siteCode = [Microsoft.VisualBasic.Interaction]::InputBox($message, $title)
return $siteCode
}


Function Get-WorkstationType {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    $form = New-object System.Windows.Forms.Form
    $form.Width = 300
    $form.Height = 200
    $form.Text = "Workstation Type Selection"
    $form.StartPosition = "CenterScreen"

    $font = New-Object System.Drawing.Font("Times New Roman",12)
    $Form.Font = $Font

    # Create a group that will contain your radio buttons
    $MyGroupBox = New-Object System.Windows.Forms.GroupBox
    $MyGroupBox.Location = '40,10'
    $MyGroupBox.size = '200,100'
    $MyGroupBox.text = "Laptop or Desktop?"

    # Create the collection of radio buttons
    $RadioButton1 = New-Object System.Windows.Forms.RadioButton
    $RadioButton1.Location = '20,30'
    $RadioButton1.size = '150,20'
    $RadioButton1.Checked = $true 
    $RadioButton1.Text = "Laptop"
    $RadioButton2 = New-Object System.Windows.Forms.RadioButton
    $RadioButton2.Location = '20,60'
    $RadioButton2.size = '150,20'
    $RadioButton2.Checked = $false
    $RadioButton2.Text = "Desktop"

    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '110,130'
    $OKButton.Size = '50,20' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK

    # Add all the Form controls on one line 
    $form.Controls.AddRange(@($MyGroupBox,$OKButton))

    # Add all the GroupBox controls on one line
    $MyGroupBox.Controls.AddRange(@($Radiobutton1,$RadioButton2))

    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $form.AcceptButton = $OKButton

    # Activate the form
    $form.Add_Shown({$form.Activate()})    

    # Get the results from the button click
    $dialogResult = $form.ShowDialog()

    # If the OK button is selected
    if ($dialogResult -eq "OK"){
        
        # Check the current state of each radio button and respond accordingly
        if ($RadioButton1.Checked){
            $workstationType = "LT"}
        elseif ($RadioButton2.Checked){
            $workstationType = "DT"}
    }
return $workstationType
}

DO {
    $siteCode = Get-SiteCode
    IF ($siteCode.Length -ne 3){
        [System.Windows.MessageBox]::Show('Site Code must be exactly 3 characters.')
    }
}
Until ($siteCode.Length -eq 3)

$workstationType = Get-WorkstationType

$bios = Get-WmiObject Win32_Bios
$computer = Get-WmiObject Win32_ComputerSystem
$computer.rename($siteCode + "WK" + $workstationType + $bios.SerialNumber)