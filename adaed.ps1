#----------[Clear All variables]------------
Get-Variable -Exclude PWD, *Preference | Remove-Variable -ErrorAction 0

#----------[Variables]------------
$span_width = 10
$frame_size = 20
$width = 740
$height = 410
$window_width = $width - $frame_size
$autocomplete = (Get-ADUser -Filter * -SearchBase 'OU=External,OU=SAdmins,DC=example,DC=domain' | Where-Object { $_.enabled -like $true }).sAMAccountName

#------------[Logic/Script/Functions]------------
function enable_search_button {
	if ($this.Text -and $InputSearchTB.Text) {
		$SearchButton.Enabled = $true
	} else {
		clearall
		$SearchButton.Enabled = $False
		$returnStatus.Text = "Ready."
		$SaveButton.Enabled = $False
	}
}

function message($Level, $Text) {
	$date = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
	if ($Level -like "info") {
		$OutputSearchRTB.Text = "[ INFO ] " + $date + " [" + $InputSearchTB.Text + "]: " + $Text + "`r`n" + $OutputSearchRTB.Text
	} elseif ($Level -like "error") {
		$OutputSearchRTB.Text = "[ ERROR ] " + $date + " [" + $InputSearchTB.Text + "]: " + $Text + "`r`n" + $OutputSearchRTB.Text
	} elseif ($Level -eq "ok") {
		$OutputSearchRTB.Text = "[ OK ] " + $date + " [" + $InputSearchTB.Text + "]: " + $Text + "`r`n" + $OutputSearchRTB.Text
	} elseif ($Level -eq "line") {
		$OutputSearchRTB.Text = "__________________________________________________`r`n" + $OutputSearchRTB.Text
	} elseif ($Level -eq "system") {
		$OutputSearchRTB.Text = "[ SYSTEM ] " + $date + "$Text" + "`r`n" + $OutputSearchRTB.Text
	} else {
		$OutputSearchRTB.Text = "[ DEBUG ] " + $date + "[" + $InputSearchTB.Text + "]: Something went wrong. Contact with Administrator." + "`r`n" + $OutputSearchRTB.Text
	}
}

function user_data {
	$returnStatus.Text 				= ("Loading... Please Wait.")
	$FullNameLL.Text 				= "Full Name:"
	$MailLL.Text 					= "Email:"
	$DescriptionLL.text 			= "Description:"
	$AccountExpirationDateLL.Text 	= "Expire at:"
	$InfoLL.Text 					= "WARNING! Press `"Update`" buton to set account expiration date after 14 business days counting from today."

	$user = $null
	$user = Get-ADUser -Filter { sAMAccountName -like $InputSearchTB.Text } -Properties * -SearchBase 'OU=External,OU=SAdmins,DC=example,DC=domain' | Select-Object sAMAccountName, displayName, description, mail, AccountExpirationDate -ErrorAction SilentlyContinue
		
	$global:sAMAccountName 			= $user.sAMAccountName
	$global:displayName 			= $user.displayName
	$global:description 			= $user.description
	$global:mail 					= $user.mail
	$global:AccountExpirationDate 	= $user.AccountExpirationDate
	
	If ($user) {
		$InputSearchTB.Text 			= $sAMAccountName
		$FullNameLLVal.Text 			= $displayName
		$DescriptionLLVal.Text 			= $description
		$MailLLVal.Text 				= $mail
		$AccountExpirationDateLLVal.Text= (Get-Date $AccountExpirationDate -Format "yyyy-MM-dd")

		$SaveButton.Enabled = $true
		$returnStatus.Text = ("Successfuly loaded user: " + $FullNameLLVal.Text + " (login: " + $sAMAccountName + ")")
	} else {
		message -Level "info" -Text ("User does not exist in Active Directory Domain. Try again...")
		clearall
		$returnStatus.Text = ("Ready.")
	}
	return $user
}

function get_user_data_textbox {
	if ($_.KeyCode -eq "Enter") {
		if (![string]::IsNullOrEmpty($InputSearchTB.Text)) {
			user_data
		} else {
			$SaveButton.Enabled = $False
		}
	}
}


function get_user_data_button {
	if (![string]::IsNullOrEmpty($InputSearchTB.Text)) {
		user_data
	} else {
		$SaveButton.Enabled = $False
	}
}

function set_expiration {
    Param (
        [Parameter(Mandatory=$True)]
        [Int]$days
    )

    $date = (Get-Date).AddDays($days)
    
    if (!($days -and "" -eq $days)) {
        foreach ($day in $date) {
            $counter=0
            if ($day.DayOfWeek -eq 'Saturday') {
                $counter+=2
                $set_monday = ($day.AddDays($counter))
                Write-Output $set_monday
            } elseif ($day.dayofweek -eq 'Sunday') {
                $counter+=1
                $set_monday = ($day.AddDays($counter))
                Write-Output $set_monday
            } elseif ($day.dayofweek -eq 'Tuesday') {
                $counter+=6
                $set_monday = ($day.AddDays($counter))
                Write-Output $set_monday
            } elseif ($day.dayofweek -eq 'Wednesday') {
                $counter+=5
                $set_monday = ($day.AddDays($counter))
                Write-Output $set_monday
            } elseif ($day.dayofweek -eq 'Thursday') {
                $counter+=4
                $set_monday = ($day.AddDays($counter))
                Write-Output $set_monday
            } elseif ($day.dayofweek -eq 'Friday') {
                $counter+=3
                $set_monday = ($day.AddDays($counter))
                Write-Output $set_monday
            } else {
				# for monday
                Write-Output $day
            }
        }
    } else {
        Write-Output "$days argument cannot be empty"
    }
}

function SetAttrUser($UserLogin) {
	if ($SetDate){
		Clear-Variable -Name SetDate -Scope Global
	}
	$global:SetDate = (set_expiration -days 14)
    Set-ADAccountExpiration -Identity $UserLogin -DateTime "$SetDate" -WhatIf
}

function save() {
	$InputSearchTB_Val = $InputSearchTB.Text
	If ($InputSearchTB_Val) {

		$SaveButton.Enabled = $false
		$returnStatus.Text 	= "Saving... Please wait."
		
		#### General ###
		#--- Exp. Date ---#
		if (!($AccountExpirationDate -like $SetDate)) {
			SetAttrUser -UserLogin $InputSearchTB_Val
			message -Level "ok" -Text (($AccountExpirationDateLL.Text) + " '$AccountExpirationDate' => '" + $SetDate + "'")
		} else {
			message -Level "error" -Text (($AccountExpirationDateLL.Text) + " '$AccountExpirationDate' => '" + $SetDate + "'")
		}
		
		$SaveButton.Enabled = $true
		get_user_data_button
	} else {
		$returnStatus.Text = "No user found."
	}
}

function clearall{
	$InputSearchTB.Text 			= $global:string
	$FullNameLL.Text 				= $global:string
	$FullNameLLVal.Text				= $global:string
	$DescriptionLL.Text 			= $global:string
	$DescriptionLLVal.Text 			= $global:string
	$MailLL.Text 					= $global:string
	$MailLLVal.Text 				= $global:string
	$AccountExpirationDateLL.Text 	= $global:string
	$AccountExpirationDateLLVal.Text= $global:string
	$InfoLL.Text					= $global:string
}

#------------[PowerShell Form Initialisations]------------
Add-Type -AssemblyName PresentationFramework
[void]([System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms"))
[void]([System.Reflection.Assembly]::LoadWithPartialName("System.Drawing"))
[System.Windows.Forms.Application]::EnableVisualStyles()


#### Form start here ###############################################################
$form                                       = New-Object system.Windows.Forms.form
$form.ClientSize                            = New-Object System.Drawing.Point($width,$height)
$form.Font                                  = New-Object System.Drawing.Font('Segoe UI',10)
$form.StartPosition                         = "CenterScreen"
$form.text                                  = "[ver. 1.0.0] ADAED - Active Directory Account Expiration Date by Bogumil Kraszewski"
$form.TopMost                               = $false
$form.KeyPreview                            = $true
$form.AutoSize                              = $false
$iconBase64                                 = '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'
$iconBytes                                  = [Convert]::FromBase64String($iconBase64)
$stream                                     = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$form.Icon                                  = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
$form.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$form.Close()}})
$form.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$SearchButton}})

### Input area #####################################################################
$InputSearchLL                              = New-Object system.Windows.Forms.Label
$InputSearchLL.location                     = New-Object System.Drawing.Point($span_width,10)
$InputSearchLL.Size                         = New-Object System.Drawing.Size(80,25)
$InputSearchLL.Font                         = New-Object System.Drawing.Font('Segoe UI',10)
$InputSearchLL.text                         = "Login:"
$InputSearchLL.AutoSize                     = $false
$InputSearchLL.TextAlign                    = 'MiddleRight'
$InputSearchLL.BorderStyle                  = 'None'
$form.Controls.Add($InputSearchLL)

$InputSearchTB                              = New-Object system.Windows.Forms.TextBox
$InputSearchTB.location                     = New-Object System.Drawing.Point(90,10)
$InputSearchTB.Size                         = New-Object System.Drawing.Size(225,20)
$InputSearchTB.Font                         = New-Object System.Drawing.Font('Segoe UI',10)
$InputSearchTB.multiline                    = $false
$InputSearchTB.Name                         = 'InputSearchTB'
$InputSearchTB.AutoCompleteSource           = 'CustomSource'
$InputSearchTB.AutoCompleteMode             = 'SuggestAppend'
$InputSearchTB.AutoCompleteCustomSource.AddRange($autocomplete)
$InputSearchTB.Add_TextChanged({enable_search_button})
$InputSearchTB.Add_KeyDown({get_user_data_textbox})
$form.Controls.Add($InputSearchTB)

$FullNameLL                                 = New-Object system.Windows.Forms.Label
$FullNameLL.location                        = New-Object System.Drawing.Point($span_width,40)
$FullNameLL.Size                            = New-Object System.Drawing.Size(80,20)
$FullNameLL.Font                            = New-Object System.Drawing.Font('Segoe UI',10)
$FullNameLL.TextAlign                       = 'MiddleRight'
$FullNameLL.BorderStyle                     = 'None'
$form.Controls.Add($FullNameLL)
$FullNameLLVal                              = New-Object system.Windows.Forms.Label
$FullNameLLVal.location						= New-Object System.Drawing.Point(($span_width + 80),40)
$FullNameLLVal.Size							= New-Object System.Drawing.Size(300,20)
$FullNameLLVal.Font							= New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$FullNameLLVal.TextAlign					= 'MiddleLeft'
$FullNameLLVal.ForeColor					= 'Blue'
$FullNameLLVal.BorderStyle					= 'None'
$form.Controls.Add($FullNameLLVal)

$MailLL                                     = New-Object system.Windows.Forms.Label
$MailLL.location                            = New-Object System.Drawing.Point($span_width,60)
$MailLL.Size                                = New-Object System.Drawing.Size(80,20)
$MailLL.Font                                = New-Object System.Drawing.Font('Segoe UI',10)
$MailLL.TextAlign                           = 'MiddleRight'
$MailLL.BorderStyle                         = 'None'
$form.Controls.Add($MailLL)
$MailLLVal                                  = New-Object system.Windows.Forms.Label
$MailLLVal.location                         = New-Object System.Drawing.Point(($span_width + 80),60)
$MailLLVal.Size                             = New-Object System.Drawing.Size(300,20)
$MailLLVal.Font                             = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$MailLLVal.TextAlign                        = 'MiddleLeft'
$MailLLVal.ForeColor                        = 'Blue'
$MailLLVal.BorderStyle                      = 'None'
$form.Controls.Add($MailLLVal)

$DescriptionLL                            	= New-Object system.Windows.Forms.Label
$DescriptionLL.location                   	= New-Object System.Drawing.Point($span_width,80)
$DescriptionLL.Size                       	= New-Object System.Drawing.Size(80,20)
$DescriptionLL.Font                       	= New-Object System.Drawing.Font('Segoe UI',10)
$DescriptionLL.TextAlign                  	= 'MiddleRight'
$DescriptionLL.BorderStyle                	= 'None'
$form.Controls.Add($DescriptionLL)
$DescriptionLLVal							= New-Object system.Windows.Forms.Label
$DescriptionLLVal.location					= New-Object System.Drawing.Point(($span_width + 80),80)
$DescriptionLLVal.Size                      = New-Object System.Drawing.Size(300,20)
$DescriptionLLVal.Font                      = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$DescriptionLLVal.TextAlign                 = 'MiddleLeft'
$DescriptionLLVal.ForeColor                 = 'Blue'
$DescriptionLLVal.BorderStyle				= 'None'
$form.Controls.Add($DescriptionLLVal)

$AccountExpirationDateLL					= New-Object system.Windows.Forms.Label
$AccountExpirationDateLL.location			= New-Object System.Drawing.Point($span_width,100)
$AccountExpirationDateLL.Size				= New-Object System.Drawing.Size(80,20)
$AccountExpirationDateLL.Font				= New-Object System.Drawing.Font('Segoe UI',10)
$AccountExpirationDateLL.TextAlign			= 'MiddleRight'
$AccountExpirationDateLL.BorderStyle		= 'None'
$form.Controls.Add($AccountExpirationDateLL)
$AccountExpirationDateLLVal					= New-Object system.Windows.Forms.Label
$AccountExpirationDateLLVal.location		= New-Object System.Drawing.Point(($span_width + 80),100)
$AccountExpirationDateLLVal.Size			= New-Object System.Drawing.Size(300,20)
$AccountExpirationDateLLVal.Font			= New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$AccountExpirationDateLLVal.TextAlign		= 'MiddleLeft'
$AccountExpirationDateLLVal.ForeColor		= 'Blue'
$AccountExpirationDateLLVal.BorderStyle		= 'None'
$form.Controls.Add($AccountExpirationDateLLVal)

$InfoLL										= New-Object system.Windows.Forms.Label
$InfoLL.location							= New-Object System.Drawing.Point(($span_width),140)
$InfoLL.Size								= New-Object System.Drawing.Size($window_width,20)
$InfoLL.Font								= New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$InfoLL.TextAlign							= 'MiddleLeft'
$InfoLL.ForeColor							= 'Red'
$InfoLL.BorderStyle							= 'None'
$form.Controls.Add($InfoLL)

#### Output textboxes #################################################################
$OutputSearchRTB                            = New-Object System.Windows.Forms.RichTextBox
$OutputSearchRTB.Location                   = New-Object System.Drawing.Point($span_width,180)
$OutputSearchRTB.Size                       = New-Object System.Drawing.Size($window_width,150)
$OutputSearchRTB.Font                       = New-Object System.Drawing.Font("Courier", "8")
$OutputSearchRTB.Anchor                     = 'top,right,left'
$OutputSearchRTB.MultiLine                  = $true
$OutputSearchRTB.Wordwrap                   = $true
$OutputSearchRTB.ReadOnly                   = $true
$OutputSearchRTB.ScrollBars                 = "Vertical"
$OutputSearchRTB.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#fff")
$form.Controls.Add($OutputSearchRTB)

# ----------[Button Clicks - Find the user when the search button is pressed]------------
$SearchButton                               = New-Object system.Windows.Forms.Button
$SearchButton.location                      = New-Object System.Drawing.Point(313,9)
$SearchButton.Size                          = New-Object System.Drawing.Size(80,27)
$SearchButton.Font                          = New-Object System.Drawing.Font('Segoe UI',10)
$SearchButton.text                          = "Search"
$SearchButton.Enabled                       = $False
$SearchButton.FlatStyle                     = 'System'
$SearchButton.Add_Click({get_user_data_button})
$form.Controls.Add($SearchButton)

#----------[Button Clicks - Set the user attributes when the set button is pressed]------------
$SaveButton                                 = New-Object system.Windows.Forms.Button
$SaveButton.location                        = New-Object System.Drawing.Point(10,350)
$SaveButton.Size                            = New-Object System.Drawing.Size(80,25)
$SaveButton.Font                            = New-Object System.Drawing.Font('Segoe UI',10)
$SaveButton.text                            = "Update"
$SaveButton.Enabled                         = $False
$SaveButton.Anchor                          = 'Left,Top'
$SaveButton.Add_Click({save})
$form.Controls.Add($SaveButton)

#----------[Button Clicks - Close application when the Quit button is pressed]------------
$QuitButton                                 = New-Object system.Windows.Forms.Button
$QuitButton.location                        = New-Object System.Drawing.Point(115,350)
$QuitButton.Size                            = New-Object System.Drawing.Size(80,25)
$QuitButton.Font                            = New-Object System.Drawing.Font('Segoe UI',10)
$QuitButton.text                            = "Quit"
$QuitButton.Anchor                          = 'Left,Top'
$QuitButton.Add_Click({$form.Close()})
$form.Controls.Add($QuitButton)

#----------[Button Clicks - Clear all values/data when the Clearr All button is pressed]------------
$ClearAllButton                             = New-Object system.Windows.Forms.Button
$ClearAllButton.location                    = New-Object System.Drawing.Point(650,350)
$ClearAllButton.Size                        = New-Object System.Drawing.Size(80,25)
$ClearAllButton.Font                        = New-Object System.Drawing.Font('Segoe UI',10)
$ClearAllButton.text                        = "Clear All"
$ClearAllButton.Anchor                      = 'Right,Top'
$ClearAllButton.Add_Click({clearall})
$form.Controls.Add($ClearAllButton)

#### Return Status Bar ################################################################
$returnStatus                               = New-Object System.Windows.Forms.StatusBar
$returnStatus.Text                          = "Please enter user login. Example for user Joe Doe: jdoe"
$form.Controls.Add($returnStatus)

$form.Add_Shown({$InputSearchTB.Select()})

# Check if Module is installed
Import-Module -Name ActiveDirectory
$ADModuleVerify = Get-Module -Name ActiveDirectory
if (!$ADModuleVerify){
    message -Level "system" -Text ("ActiveDirectory module is not imported! Please contact with your System Administrator.")
}

$form.ShowDialog()  | Out-Null
