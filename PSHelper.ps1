<#
    .Synopsis
        Daily Activity Helper
    .Description
        This script will help in performing some daily activity like 'Scheduled Shutdown','Abort Shutdown','Password Manager','Close All Windows' 
#>


#Set-ExecutionPolicy Unrestricted -Force 
import-module .\\AESEncryption\AESEncryption.psm1 -Force
$KeyPath = '.\\AESEncryption\AES.key'

#Password Area
$Office365Password = '76492d1116743f0423413b16050a5345MgB8AHgAcgBmAGgAUABzAHgAbQBKAFcAbwA3AHMAbgBRAHMATQBWAGsAZwB4AGcAPQA9AHwAOABlADAAYwA1ADkAMgA1ADYAYgAxADAAZgBjAGYANwA4AGEAZABkAGQAZQBjADcAMwBkAGYANAA3AGEANAAyADEANgA3ADUAZABiAGIAYwAzADUAZAA0ADMAZgA4ADQAOAA4ADgAYgBmADcANAA4AGIAYQBkAGIAZABlADgAYwA='
$TimesheetPassword = '76492d1116743f0423413b16050a5345MgB8AG0AbABRAEcASwBlAEcAbwArAFMAbgByAHoAVwBxADgAVwBPAHoATgBwAGcAPQA9AHwANgAzAGIANQBmAGEAOAA5AGIAOQBlAGUAOABjAGMAOAA4AGUAZgBiADkAMQA5ADMAMgBkADcAMAAyADAANwA4ADcAMABjADQAOQBkAGQAZAAxADIANQBmADMANQA1AGMAOAAyADYAZQBmAGUAYQAwADUAYQA1ADUAZQA4ADAAYgA='
$TaxPassword = '76492d1116743f0423413b16050a5345MgB8ADUATABaAHkASgB0AEQAVABIAEcASwB5AG8AOABUADcAagB3AGIAVQB3AEEAPQA9AHwANgAxADUAMQBkAGQANAA0ADAAMgBiADMAYgBlADIAZQAzADkAMQBhAGIAMgAxADUAZgBmAGUAZgA1AGUANQAyAA=='
$PersonalPassword = '76492d1116743f0423413b16050a5345MgB8AHkAUABKAFMAUQA2AHQAcABxAHcAagAvADIATgBCAEYAZwBnAFkAMABzAEEAPQA9AHwAYwA2AGIAMgA5ADIAYQAwAGMAMAAzAGQAOAAwAGIAOAAyAGMAZAAyADkAZQA5ADkAMgBkAGEAOAA2ADMAYQA0AGUAYQBhAGQANwA4ADkAMgAxADMAMQBkAGIAOAAyAGUAYQAyADQAYQBkAGUAYgA4ADcAOQBlAGYAMAA3AGIANAA='


$Script                          = '.\\PSHelper.ps1'
$ScriptContent                   = Get-Content $Script

#Form Area
$TestRunner                      = New-Object System.Windows.Forms.Form
$HeaderFont                      = New-Object System.Drawing.Font('Times New Roman',40,[System.Drawing.FontStyle]::Bold)
$BoldFont                        = New-Object System.Drawing.Font('Times New Roman',18,[System.Drawing.FontStyle]::Bold)
$ItalicFont                      = New-Object System.Drawing.Font('Times New Roman',14,[System.Drawing.FontStyle]::Italic)
$RegularFont                     = New-Object System.Drawing.Font('Times New Roman',18,[System.Drawing.FontStyle]::Regular)
$StrikeoutFont                   = New-Object System.Drawing.Font('Times New Roman',18,[System.Drawing.FontStyle]::Strikeout)
$UnderlineFont                   = New-Object System.Drawing.Font('Times New Roman',18,[System.Drawing.FontStyle]::Underline)


$TestRunner.WindowState          = 'Maximized'
$TestRunner.AutoSize             = 'True'
$TestRunner.text                 = 'PSHelper'
$TestRunner.BackColor            = '#ffffff'
$Title                           = New-Object system.Windows.Forms.Label 
$Title.Font                      = $HeaderFont
$Title.text                      = 'PSHelper'
$Title.AutoSize                  = $true
$Title.width                     = 50
$Title.height                    = 10
$Title.location                  = New-Object System.Drawing.Point(20,20)
$Description                     = New-Object system.Windows.Forms.Label
$Description.Font                = $ItalicFont
$Description.text                = 'Welcome,How are you doing! Please Select your action'
$Description.AutoSize            = $false
$Description.width               = 450
$Description.height              = 50
$Description.location            = New-Object System.Drawing.Point(27,100)
$TestRunner.controls.AddRange(@($Title,$Description))
$TestRunner.FormBorderStyle      = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$TestRunner.BackColor = 'cadetblue'


$descriptions                    = @('Select Item','Scheduled Shutdown','Abort Shutdown','Password Manager','Close All Windows','Alexa')

$ComboBox1_SelectedIndexChanged=
    {
        Switch ($comboBox1.text)
       {
            'Scheduled Shutdown'
            {
                $envnames = @('Select Value in Hours','1','2','3','4')
            }
            'Abort Shutdown'
            {
                $envnames = @('Are You Sure','Yes','No')
            }
            'Scheduled Restart'
            {
                $envnames = @('Questionare','Placeholder','Report')
            }
            'Password Manager'
            {
                $envnames = @('Select App','Office365','Timesheet','Tax','Personal')
            }
            'Close All Windows'
            {
                $envnames = @('Are You Sure','Yes','No')
            }
            'Alexa'
            {
                $envnames = @('Male','Female')
            }
            default
            {
                $envnames = @()
            }
        }
$comboBox2.Remove_SelectedIndexChanged($ComboBox2_SelectedIndexChanged)
$comboBox2.DataSource = $envnames
$ComboBox2.add_SelectedIndexChanged($ComboBox2_SelectedIndexChanged)
    }


$ComboBox2_SelectedIndexChanged=
    {
        Switch ($comboBox1.text)
       {
            'Password Manager'
            {
                $envnames2 = @('Save Password','Get Password')
            }
            default
            {
                $envnames2 = @()
            }
        }
$comboBox3.Remove_SelectedIndexChanged($ComboBox3_SelectedIndexChanged)
$comboBox3.DataSource = $envnames2
$ComboBox3.add_SelectedIndexChanged($ComboBox3_SelectedIndexChanged)
    }


#ComboBox1 Label Area
$comboBox1Label                     = New-Object System.Windows.Forms.Label
$comboBox1Label.text                = 'Action*'
$comboBox1Label.Font                = $RegularFont
$comboBox1Label.Size                = New-Object System.Drawing.Size(200, 50)
$comboBox1Label.location            = New-Object System.Drawing.Point(25,150)
$TestRunner.Controls.Add($comboBox1Label)

#ComboBox1 Area
$comboBox1                          = New-Object System.Windows.Forms.ComboBox
$comboBox1.BackColor                = 'mistyrose'
$comboBox1.Font                     = $RegularFont
$comboBox1.Location                 = New-Object System.Drawing.Point(225,150)
$comboBox1.Size                     = New-Object System.Drawing.Size(235, 50)
$comboBox1.DataSource               = $descriptions
$ComboBox1.add_SelectedIndexChanged($ComboBox1_SelectedIndexChanged)
$TestRunner.Controls.Add($comboBox1)

#ComboBox2 Label Area
$comboBox2Label = New-Object System.Windows.Forms.Label
$comboBox2Label.text                = 'Sub-action*'
$comboBox2Label.Font                = $RegularFont
$comboBox2Label.Size                = New-Object System.Drawing.Size(150, 25)
$comboBox2Label.location            = New-Object System.Drawing.Point(25,200)
$TestRunner.Controls.Add($comboBox2Label)

#ComboBox2 Area
$comboBox2 = New-Object System.Windows.Forms.ComboBox
$comboBox2.BackColor                = 'mistyrose'
$comboBox2.Font                     = $RegularFont
$comboBox2.Location = New-Object System.Drawing.Point(225,200)
$comboBox2.Size = New-Object System.Drawing.Size(235, 50)
$ComboBox2.add_SelectedIndexChanged($ComboBox2_SelectedIndexChanged)
$TestRunner.Controls.Add($comboBox2)


#ComboBox3 Label Area
$comboBox3Label = New-Object System.Windows.Forms.Label
$comboBox3Label.text                = 'Option'
$comboBox3Label.Font                = $RegularFont
$comboBox3Label.Size                = New-Object System.Drawing.Size(150, 25)
$comboBox3Label.location            = New-Object System.Drawing.Point(25,250)
$TestRunner.Controls.Add($comboBox3Label)

#ComboBox3 Area
$comboBox3 = New-Object System.Windows.Forms.ComboBox
$comboBox3.BackColor                = 'mistyrose'
$comboBox3.Font                     = $RegularFont
$comboBox3.Location = New-Object System.Drawing.Point(225,250)
$comboBox3.Size = New-Object System.Drawing.Size(235, 50)
$ComboBox3.add_SelectedIndexChanged($ComboBox3_SelectedIndexChanged)
$TestRunner.Controls.Add($comboBox3)

#ComboBox3 Info Area
$comboBox3InfoLabel                     = New-Object System.Windows.Forms.Label
$comboBox3InfoLabel.ForeColor           = 'black'
$comboBox3InfoLabel.Font                = $ItalicFont
$comboBox3InfoLabel.text                = '(This Field is applicable for Action type Password Manager)'
$comboBox3InfoLabel.Size                = New-Object System.Drawing.Size(500, 25)
$comboBox3InfoLabel.location            = New-Object System.Drawing.Point(460,255)
$TestRunner.Controls.Add($comboBox3InfoLabel)


#Value Label Area
$comboBox2Label = New-Object System.Windows.Forms.Label
$comboBox2Label.Font                = $RegularFont
$comboBox2Label.text                = 'Value'
$comboBox2Label.Size                = New-Object System.Drawing.Size(150, 25)
$comboBox2Label.location            = New-Object System.Drawing.Point(25,300)
$TestRunner.Controls.Add($comboBox2Label)

#Value Area
$PwdBox = New-Object System.Windows.Forms.TextBox
$PwdBox.BackColor                = 'mistyrose'
$PwdBox.Font                     = $RegularFont
$PwdBox.Location = New-Object System.Drawing.Point(225,300)
$PwdBox.Size = New-Object System.Drawing.Size(235, 50)
$TestRunner.Controls.Add($PwdBox)

#Value Info Area
$comboBox3InfoLabel = New-Object System.Windows.Forms.Label
$comboBox3InfoLabel.ForeColor           = 'black'
$comboBox3InfoLabel.Font                = $ItalicFont
$comboBox3InfoLabel.text                = '(This Field is applicable for Action type Password Manager)'
$comboBox3InfoLabel.Size                = New-Object System.Drawing.Size(500, 25)
$comboBox3InfoLabel.location            = New-Object System.Drawing.Point(460,310)
$TestRunner.Controls.Add($comboBox3InfoLabel)

#Listbox Area
$OutputBox = New-Object System.Windows.Forms.TextBox
$OutputBox.BackColor  = 'mistyrose'
$OutputBox.multiline = $true
$OutputBox.width = 760
$OutputBox.height = 300
$OutputBox.Location = New-Object System.Drawing.Point(1000,135)
$OutputBox.Font = $BoldFont
$TestRunner.Controls.Add($OutputBox) 

#Execute Button Area
$ExecuteBtn                       = New-Object system.Windows.Forms.Button
$ExecuteBtn.BackColor             = 'lime'
$ExecuteBtn.text                  = 'Execute'
$ExecuteBtn.location              = New-Object System.Drawing.Point(20,405)
$ExecuteBtn.Font                  = $BoldFont
$ExecuteBtn.Size                  = New-Object System.Drawing.Size(150, 35)
$ExecuteBtn.ForeColor             = '#000'
$ExecuteBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Yes
$TestRunner.AcceptButton   = $ExecuteBtn
$TestRunner.Controls.Add($ExecuteBtn)

#GetDetails Button Area
$DetailslBtn                       = New-Object system.Windows.Forms.Button
$DetailslBtn.BackColor             = 'gold'
$DetailslBtn.text                  = 'Get Details'
$DetailslBtn.location              = New-Object System.Drawing.Point(190,405)
$DetailslBtn.Font                  = $BoldFont
$DetailslBtn.Size                  = New-Object System.Drawing.Size(150, 35)
$DetailslBtn.ForeColor             = '#000'
$TestRunner.Controls.Add($DetailslBtn)

#Cancel Button Area
$CancelBtn                       = New-Object system.Windows.Forms.Button
$CancelBtn.BackColor             = 'red'
$CancelBtn.text                  = 'Cancel'
$CancelBtn.location              = New-Object System.Drawing.Point(360,405)
$CancelBtn.Font                  = $BoldFont
$CancelBtn.Size                  = New-Object System.Drawing.Size(150, 35)
$CancelBtn.ForeColor             = '#000'
$CancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$TestRunner.CancelButton   = $CancelBtn
$TestRunner.Controls.Add($CancelBtn)


#Code for Get Details
$DetailslBtn.Add_Click({

   if($comboBox3.SelectedItem -eq 'Get Password')
        {
            
            Switch ($comboBox2.text)
            {
                'Office365'
                {
                   $Password = Get-Decrypt -SecureStringAES $Office365Password -KeyPath $KeyPath 
                }
                'Timesheet'
                {
                   $Password = Get-Decrypt -SecureStringAES $TimesheetPassword -KeyPath $KeyPath 
                }
                'Tax'
                {
                   $Password = Get-Decrypt -SecureStringAES $TaxPassword -KeyPath $KeyPath 
                }
                'Personal'
                {
                   $Password = Get-Decrypt -SecureStringAES $PersonalPassword -KeyPath $KeyPath 
                }
                default
                {
                    
                }
           }
        }
   #$WshShell = new-object -com "Wscript.Shell"
   #$Return = $WshShell.Popup("Your Password is $Password")
   $OutputBox.Text = $Password
})


$TestRunner.ShowDialog() 
 
    #Scheduled Shutdown
    if($comboBox1.SelectedItem -eq 'Scheduled Shutdown')
    {

        if($comboBox2.SelectedItem -eq '1' -or $comboBox2.SelectedItem -eq '2' -or $comboBox2.SelectedItem -eq '3' -or $comboBox2.SelectedItem -eq '4')
        {
            $Time = [int]$comboBox2.SelectedItem
            $TimeInSecond = $Time*3600
            shutdown -s -f -t $TimeInSecond
        }
    }

    #Abort Shutdown
    if($comboBox1.SelectedItem -eq 'Abort Shutdown')
    {

        if($comboBox2.SelectedItem -eq 'Yes')
        {
            shutdown -a
        }
    }

    #Password Manager Logic
    if($comboBox1.SelectedItem -eq 'Password Manager')
    {
        if($comboBox3.SelectedItem -eq 'Save Password')
        {
            $PT_Password = $PwdBox.Text
            $EncryptedPassword = Get-Encrypt -PlainTextPassword $PT_Password -KeyPath $KeyPath
            Switch ($comboBox2.text) 
            {
                'Office365'
                {

                   $ScriptModified = $ScriptContent.Replace('76492d1116743f0423413b16050a5345MgB8AHgAcgBmAGgAUABzAHgAbQBKAFcAbwA3AHMAbgBRAHMATQBWAGsAZwB4AGcAPQA9AHwAOABlADAAYwA1ADkAMgA1ADYAYgAxADAAZgBjAGYANwA4AGEAZABkAGQAZQBjADcAMwBkAGYANAA3AGEANAAyADEANgA3ADUAZABiAGIAYwAzADUAZAA0ADMAZgA4ADQAOAA4ADgAYgBmADcANAA4AGIAYQBkAGIAZABlADgAYwA=',$EncryptedPassword) 
                   $ScriptModified | Set-Content -Path $Script
                }
                'Timesheet'
                {
                   $ScriptModified = $ScriptContent.Replace('76492d1116743f0423413b16050a5345MgB8AG0AbABRAEcASwBlAEcAbwArAFMAbgByAHoAVwBxADgAVwBPAHoATgBwAGcAPQA9AHwANgAzAGIANQBmAGEAOAA5AGIAOQBlAGUAOABjAGMAOAA4AGUAZgBiADkAMQA5ADMAMgBkADcAMAAyADAANwA4ADcAMABjADQAOQBkAGQAZAAxADIANQBmADMANQA1AGMAOAAyADYAZQBmAGUAYQAwADUAYQA1ADUAZQA4ADAAYgA=',$EncryptedPassword) 
                   $ScriptModified | Set-Content -Path $Script
                }
                'Tax'
                {
                   $ScriptModified = $ScriptContent.Replace('76492d1116743f0423413b16050a5345MgB8ADUATABaAHkASgB0AEQAVABIAEcASwB5AG8AOABUADcAagB3AGIAVQB3AEEAPQA9AHwANgAxADUAMQBkAGQANAA0ADAAMgBiADMAYgBlADIAZQAzADkAMQBhAGIAMgAxADUAZgBmAGUAZgA1AGUANQAyAA==',$EncryptedPassword) 
                   $ScriptModified | Set-Content -Path $Script
                }
                'Personal'
                {
                    $ScriptModified = $ScriptContent.Replace('76492d1116743f0423413b16050a5345MgB8AHkAUABKAFMAUQA2AHQAcABxAHcAagAvADIATgBCAEYAZwBnAFkAMABzAEEAPQA9AHwAYwA2AGIAMgA5ADIAYQAwAGMAMAAzAGQAOAAwAGIAOAAyAGMAZAAyADkAZQA5ADkAMgBkAGEAOAA2ADMAYQA0AGUAYQBhAGQANwA4ADkAMgAxADMAMQBkAGIAOAAyAGUAYQAyADQAYQBkAGUAYgA4ADcAOQBlAGYAMAA3AGIANAA=',$EncryptedPassword) 
                    $ScriptModified | Set-Content -Path $Script
                }
                default
                {
                    
                }
           }
        }
    }

     #Close All Windows
    if($comboBox1.SelectedItem -eq 'Close All Windows')
    {
        $softwarelist = 'notepad|firefox|iexplore|chrome|chromedriver|excel|word|OUTLOOK|Eclipse'
        get-process |
            Where-Object {$_.ProcessName -match $softwarelist} |
            stop-process -force

         $a = (New-Object -comObject Shell.Application).Windows() |
         ? { $_.FullName -ne $null} |
         ? { $_.FullName.toLower().Endswith('\explorer.exe') } 
         $a | % {  $_.Quit() }
    }

    if($comboBox1.SelectedItem -eq 'Alexa')
    {
        if($comboBox2.SelectedItem -eq 'Male')
        {
            $TextToSpeak = $OutputBox.Text 
            Add-Type -AssemblyName "System.Speech"
            $Speech = New-Object System.Speech.Synthesis.SpeechSynthesizer
            $Voices = $Speech.GetInstalledVoices()
            $Speech.SelectVoice( $Voices[-2].VoiceInfo.Name ) 
            $Speech.Speak( $TextToSpeak )
        }
        if($comboBox2.SelectedItem -eq 'Female')
        {
            $TextToSpeak = $OutputBox.Text 
            Add-Type -AssemblyName "System.Speech"
            $Speech = New-Object System.Speech.Synthesis.SpeechSynthesizer
            $Voices = $Speech.GetInstalledVoices()
            $Speech.SelectVoice( $Voices[-1].VoiceInfo.Name ) 
            $Speech.Speak( $TextToSpeak )
        }
    }

