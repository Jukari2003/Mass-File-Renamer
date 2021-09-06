################################################################################
#                            Mass File Renamer                                 #
#                     Written By: MSgt Anthony Brechtel                        #
#                                                                              #
################################################################################
######Load Assemblies###########################################################
clear-host
Add-Type -AssemblyName 'System.Windows.Forms'
Add-Type -AssemblyName 'System.Drawing'
Add-Type -AssemblyName 'PresentationFramework'
[System.Windows.Forms.Application]::EnableVisualStyles();
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

################################################################################
######Global Variables##########################################################
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Set-Location $dir

$script:version = "1.0"
$script:program_title = "Mass File Renamer"
$script:hash = ""
$script:preview_job = "";
$script:real_or_simulation = 0
$script:settings = @{}
$script:max_files = 100000000
$script:output_file = "";
$script:start_time = Get-Date

$script:target_directory = "C:\users\Jukari\Desktop\Test"
$script:target_directory = "Browse or Enter a file path"
$script:drill_down_folders = 0
$script:include_file_names = 1
$script:include_folder_names = 0
$script:include_extensions = 1;
$script:format_titles_automatically = 0;
$script:case_sensitive = 0;
$script:must_contain = "";
$script:action = "Replace Characters:"
$script:word1 = ""
$script:word2 = ""

##Idle Timer
if($Script:Timer){$Script:Timer.Dispose();}
$Script:Timer = New-Object System.Windows.Forms.Timer
$Script:Timer.Interval = 1000
$Script:CountDown = 1

################################################################################
######Main######################################################################
function main
{
    ##################################################################################
    ###########Main Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Location = "200, 200"
    $Form.Font = "Copperplate Gothic,8.1"
    $Form.FormBorderStyle = "FixedDialog"
    $Form.ForeColor = "Black"
    $Form.BackColor = "#434343"
    $Form.Text = "  $script:program_title"
    $Form.Width = 1000 #1245
    $Form.Height = 800

    
    ##################################################################################
    ###########Title Main
    $y_pos = 15
    $title1            = New-Object System.Windows.Forms.Label   
    $title1.Font       = New-Object System.Drawing.Font("Copperplate Gothic Bold",21,[System.Drawing.FontStyle]::Regular)
    $title1.Text       = $script:program_title
    $title1.TextAlign  = "MiddleCenter"
    $title1.Width      = $Form.Width
    $title1.height     = 35
    $title1.ForeColor  = "white"
    $title1.Location   = New-Object System.Drawing.Size((($Form.width / 2) - ($Form.width / 2)),$y_pos)
    $Form.Controls.Add($title1)

    ##################################################################################
    ###########Title Written By
    $y_pos = $y_pos + 30
    $title2            = New-Object System.Windows.Forms.Label
    $title2.Font       = New-Object System.Drawing.Font("Copperplate Gothic",7.5,[System.Drawing.FontStyle]::Regular)
    $title2.Text       = "Written by: Anthony Brechtel`nVer $script:version"
    $title2.TextAlign  = "MiddleCenter"
    $title2.ForeColor  = "darkgray"
    $title2.Width      = $Form.Width
    $title2.Height     = 40
    $title2.Location   = New-Object System.Drawing.Size((($Form.width / 2) - ($Form.width / 2)),$y_pos)
    $Form.Controls.Add($title2)

    ##################################################################################
    ###########Root Directory Label
    $y_pos = $y_pos + 45
    $root_directory_label = New-Object System.Windows.Forms.Label
    $root_directory_label.Location = New-Object System.Drawing.Point(15, $y_pos)
    $root_directory_label.Size = "250, 23"
    $root_directory_label.ForeColor = "White" 
    $root_directory_label.Text = "Root Directory:"
    $root_directory_label.TextAlign = "MiddleRight"
    $root_directory_label.Font = [Drawing.Font]::New("Times New Roman", 12)
    $Form.Controls.Add($root_directory_label)

    ##################################################################################
    ###########Target Box
    $target_box = New-Object System.Windows.Forms.TextBox
    $target_box.Location = New-Object System.Drawing.Point(($root_directory_label.location.x + $root_directory_label.width + 3),($y_pos))
    $target_box.font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $target_box.width = 385
    $target_box.Height = 40
    $target_box.Text = $script:target_directory
    $target_box.Add_Click({
        if($this.Text -eq "Browse or Enter a file path")
        {
            $this.Text = ""
        }
    })
    $target_box.Add_lostFocus({
        if($this.text -eq "Browse or Enter a file path")
        {

        }
        elseif(($this.text -eq "") -or ($this.text -eq $null))
        {
            $this.text = "Browse or Enter a file path"
            #$message = "Path must point to a valid folder."
            #[System.Windows.MessageBox]::Show($message,"Error",'Ok','Error') 
        }
        elseif(!(Test-Path -literalpath $this.text))
        {
            $this.text = "Browse or Enter a file path"
            #$message = "Path must point to a valid folder."
            #[System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')   
        }
        elseif(!((Get-Item -literalpath $this.text) -is [System.IO.DirectoryInfo]))   
        {
            $this.text = "Browse or Enter a file path"
            #$message = "Path must point to a valid folder."
            #[System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
        }
        
    })
    $target_box.Add_TextChanged({
    $this.text = $this.text -replace "^`"|`"$"
        [string]$script:target_directory = $this.text
        if(($script:target_directory -ne $null) -and ($script:target_directory -ne ""))
        {
            if(Test-Path -literalpath $target_box.text)
            {
                $Form.Controls.Add($scan_target_button)
                
            }
            else
            {
                $Form.Controls.Remove($scan_target_button)
            }
        }
        else
        {
            $Form.Controls.Remove($scan_target_button)
        }
    })
    $Form.Controls.Add($target_box)

    ##################################################################################
    ###########Browse Button   
    $browse_button = New-Object System.Windows.Forms.Button
    $browse_button.Location= New-Object System.Drawing.Size(($target_box.location.x + $target_box.width + 3),($y_pos + 1))
    $browse_button.BackColor = "#606060"
    $browse_button.ForeColor = "White"
    $browse_button.Width=70
    $browse_button.Height=22
    $browse_button.Text='Browse'
    $browse_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    $browse_button.Add_Click(
    {    
		    $script:target_directory = prompt_for_folder
            if(($script:target_directory -ne $Null) -and ($script:target_directory -ne "") -and ((Test-Path -literalpath $script:target_directory) -eq $True))
            {
                $target_box.Text="$script:target_directory"
                
            }
    })
    $Form.Controls.Add($browse_button)
    
    ##################################################################################
    ###########Drill Down Into Folders Label
    $y_pos = $y_pos + 30
    $drill_down_folders_label = New-Object System.Windows.Forms.Label
    $drill_down_folders_label.Location = New-Object System.Drawing.Point(15, $y_pos)
    $drill_down_folders_label.Size = "250, 23"
    $drill_down_folders_label.ForeColor = "White" 
    $drill_down_folders_label.Text = "Drill Into Folders:"
    $drill_down_folders_label.TextAlign = "MiddleRight"
    $drill_down_folders_label.Font = [Drawing.Font]::New("Times New Roman", 12)
    $Form.Controls.Add($drill_down_folders_label)

    ##################################################################################
    ###########Drill Down Into Folders Checkbox
    $drill_down_folders_checkbox = new-object System.Windows.Forms.checkbox
    $drill_down_folders_checkbox.Location = new-object System.Drawing.Size(($drill_down_folders_label.location.x + $drill_down_folders_label.width + 3),$y_pos);
    $drill_down_folders_checkbox.Size = new-object System.Drawing.Size(100,30)  
    $drill_down_folders_checkbox.name = $script:drill_down_folders           
    if($script:drill_down_folders -eq "0")
    {
        $drill_down_folders_checkbox.Checked = $false
        $drill_down_folders_checkbox.text = "Disabled"
        $drill_down_folders_checkbox.ForeColor                = "Red"
    }
    else
    {
        $drill_down_folders_checkbox.Checked = $true
        $drill_down_folders_checkbox.text = "Enabled"
        $drill_down_folders_checkbox.ForeColor                = "Lime"
    }
    $drill_down_folders_checkbox.Add_CheckStateChanged({
        if($this.Checked -eq $true)
        {
            $this.text = "Enabled"
            $this.ForeColor                = "Lime"
            $script:drill_down_folders     = 1;
        }
        else
        {
            $this.text = "Disabled"
            $this.ForeColor                = "Red"
            $script:drill_down_folders  = 0;
        }
        
    })
    $Form.controls.Add($drill_down_folders_checkbox);

    ##################################################################################
    ###########Include File Names Checkbox Label
    $y_pos = $y_pos + 30
    $include_file_names_label = New-Object System.Windows.Forms.Label
    $include_file_names_label.Location = New-Object System.Drawing.Point(15, $y_pos)
    $include_file_names_label.Size = "250, 23"
    $include_file_names_label.ForeColor = "White" 
    $include_file_names_label.Text = "Include File Names:"
    $include_file_names_label.TextAlign = "MiddleRight"
    $include_file_names_label.Font = [Drawing.Font]::New("Times New Roman", 12)
    $Form.Controls.Add($include_file_names_label)

    ##################################################################################
    ###########Include File Names Checkbox Label
    $include_file_names_checkbox = new-object System.Windows.Forms.checkbox
    $include_file_names_checkbox.Location = new-object System.Drawing.Size(($include_file_names_label.location.x + $include_file_names_label.width + 3),$y_pos);
    $include_file_names_checkbox.Size = new-object System.Drawing.Size(100,30)  
    $include_file_names_checkbox.name = $script:include_file_names           
    if($script:include_file_names  -eq "0")
    {
        $include_file_names_checkbox.Checked = $false
        $include_file_names_checkbox.text = "Disabled"
        $include_file_names_checkbox.ForeColor                = "Red"
    }
    else
    {
        $include_file_names_checkbox.Checked = $true
        $include_file_names_checkbox.text = "Enabled"
        $include_file_names_checkbox.ForeColor                = "Lime"
    }
    $include_file_names_checkbox.Add_CheckStateChanged({
        if($this.Checked -eq $true)
        {
            $this.text = "Enabled"
            $this.ForeColor                = "Lime"
            $script:include_file_names     = 1;
        }
        else
        {
            $this.text = "Disabled"
            $this.ForeColor                = "Red"
            $script:include_file_names     = 0;
        }
        
    })
    $Form.controls.Add($include_file_names_checkbox);

    ##################################################################################
    ###########Include Folder Names Label
    $y_pos = $y_pos + 30
    $include_folder_names_label = New-Object System.Windows.Forms.Label
    $include_folder_names_label.Location = New-Object System.Drawing.Point(15, $y_pos)
    $include_folder_names_label.Size = "250, 23"
    $include_folder_names_label.ForeColor = "White" 
    $include_folder_names_label.Text = "Include Folder Names:"
    $include_folder_names_label.TextAlign = "MiddleRight"
    $include_folder_names_label.Font = [Drawing.Font]::New("Times New Roman", 12)
    $Form.Controls.Add($include_folder_names_label)

    ##################################################################################
    ###########Include Folder Names Checkbox
    $include_folder_names_checkbox = new-object System.Windows.Forms.checkbox
    $include_folder_names_checkbox.Location = new-object System.Drawing.Size(($include_folder_names_label.location.x + $include_folder_names_label.width + 3),$y_pos);
    $include_folder_names_checkbox.Size = new-object System.Drawing.Size(100,30)  
    $include_folder_names_checkbox.name = $script:include_folder_names          
    if($script:include_folder_names  -eq "0")
    {
        $include_folder_names_checkbox.Checked = $false
        $include_folder_names_checkbox.text = "Disabled"
        $include_folder_names_checkbox.ForeColor                = "Red"
    }
    else
    {
        $include_folder_names_checkbox.Checked = $true
        $include_folder_names_checkbox.text = "Enabled"
        $include_folder_names_checkbox.ForeColor                = "Lime"
    }
    $include_folder_names_checkbox.Add_CheckStateChanged({
        if($this.Checked -eq $true)
        {
            $this.text = "Enabled"
            $this.ForeColor                = "Lime"
            $script:include_folder_names   = 1;
        }
        else
        {
            $this.text = "Disabled"
            $this.ForeColor                = "Red"
            $script:include_folder_names   = 0;
        }
        
    })
    $Form.controls.Add($include_folder_names_checkbox);

    ##################################################################################
    ###########Include Extentions Label
    $y_pos = $y_pos + 30
    $include_extenstions_label = New-Object System.Windows.Forms.Label
    $include_extenstions_label.Location = New-Object System.Drawing.Point(15, $y_pos)
    $include_extenstions_label.Size = "250, 23"
    $include_extenstions_label.ForeColor = "White" 
    $include_extenstions_label.Text = "Include File Extensions:"
    $include_extenstions_label.TextAlign = "MiddleRight"
    $include_extenstions_label.Font = [Drawing.Font]::New("Times New Roman", 12)
    $Form.Controls.Add($include_extenstions_label)

    ##################################################################################
    ###########Include Extentions Checkbox
    $include_extenstions_checkbox = new-object System.Windows.Forms.checkbox
    $include_extenstions_checkbox.Location = new-object System.Drawing.Size(($include_extenstions_label.location.x + $include_extenstions_label.width + 3),$y_pos);
    $include_extenstions_checkbox.Size = new-object System.Drawing.Size(100,30)  
    $include_extenstions_checkbox.name = $script:include_extensions      
    
    if($script:include_extensions    -eq "0")
    {
        $include_extenstions_checkbox.Checked = $false
        $include_extenstions_checkbox.text = "Disabled"
        $include_extenstions_checkbox.ForeColor                = "Red"
    }
    else
    {
        $include_extenstions_checkbox.Checked = $true
        $include_extenstions_checkbox.text = "Enabled"
        $include_extenstions_checkbox.ForeColor                = "Lime"
    }
    $include_extenstions_checkbox.Add_CheckStateChanged({
        if($this.Checked -eq $true)
        {
            $this.text = "Enabled"
            $this.ForeColor                = "Lime"
            $script:include_extensions     = 1;
        }
        else
        {
            $this.text = "Disabled"
            $this.ForeColor                = "Red"
            $script:include_extensions     = 0;
        }
        
    })
    $Form.controls.Add($include_extenstions_checkbox);

    ##################################################################################
    ###########Format Titles Automatically Label
    $y_pos = $y_pos + 30
    $Format_titles_automatically_label = New-Object System.Windows.Forms.Label
    $Format_titles_automatically_label.Location = New-Object System.Drawing.Point(15, $y_pos)
    $Format_titles_automatically_label.Size = "250, 23"
    $Format_titles_automatically_label.ForeColor = "White" 
    $Format_titles_automatically_label.Text = "Format Titles Automatically:"
    $Format_titles_automatically_label.TextAlign = "MiddleRight"
    $Format_titles_automatically_label.Font = [Drawing.Font]::New("Times New Roman", 12)
    $Form.Controls.Add($Format_titles_automatically_label)

    ##################################################################################
    ###########Format Titles Automatically Checkbox
    $format_titles_automatically_checkbox = new-object System.Windows.Forms.checkbox
    $format_titles_automatically_checkbox.Location = new-object System.Drawing.Size(($Format_titles_automatically_label.location.x + $Format_titles_automatically_label.width + 3),$y_pos);
    $format_titles_automatically_checkbox.Size = new-object System.Drawing.Size(100,30)  
    $format_titles_automatically_checkbox.name = $script:format_titles_automatically        
    if($script:format_titles_automatically -eq "0")
    {
        $format_titles_automatically_checkbox.Checked = $false
        $format_titles_automatically_checkbox.text = "Disabled"
        $format_titles_automatically_checkbox.ForeColor                = "Red"
    }
    else
    {
        $format_titles_automatically_checkbox.Checked = $true
        $format_titles_automatically_checkbox.text = "Enabled"
        $format_titles_automatically_checkbox.ForeColor                = "Lime"
    }
    $format_titles_automatically_checkbox.Add_CheckStateChanged({
        if($this.Checked -eq $true)
        {
            $this.text = "Enabled"
            $this.ForeColor                     = "Lime"
            $script:format_titles_automatically = 1;
        }
        else
        {
            $this.text = "Disabled"
            $this.ForeColor                     = "Red"
            $script:format_titles_automatically = 0;
        }
        
    })
    $Form.controls.Add($format_titles_automatically_checkbox);

    ##################################################################################
    ###########Case Sensitive Label
    $y_pos = $y_pos + 30
    $case_sensitive_label = New-Object System.Windows.Forms.Label
    $case_sensitive_label.Location = New-Object System.Drawing.Point(15, $y_pos)
    $case_sensitive_label.Size = "250, 23"
    $case_sensitive_label.ForeColor = "White" 
    $case_sensitive_label.Text = "Case Sensitive:"
    $case_sensitive_label.TextAlign = "MiddleRight"
    $case_sensitive_label.Font = [Drawing.Font]::New("Times New Roman", 12)
    $Form.Controls.Add($case_sensitive_label)

    ##################################################################################
    ###########Case Sensitive Checkbox
    $case_sensitive_checkbox = new-object System.Windows.Forms.checkbox
    $case_sensitive_checkbox.Location = new-object System.Drawing.Size(($case_sensitive_label.location.x + $case_sensitive_label.width + 3),$y_pos);
    $case_sensitive_checkbox.Size = new-object System.Drawing.Size(100,30)  
    $case_sensitive_checkbox.name = $script:case_sensitive       
    if($script:format_titles_automatically -eq "0")
    {
        $case_sensitive_checkbox.Checked = $false
        $case_sensitive_checkbox.text = "Disabled"
        $case_sensitive_checkbox.ForeColor                = "Red"
    }
    else
    {
        $case_sensitive_checkbox.Checked = $true
        $case_sensitive_checkbox.text = "Enabled"
        $case_sensitive_checkbox.ForeColor                = "Lime"
    }
    $case_sensitive_checkbox.Add_CheckStateChanged({
        if($this.Checked -eq $true)
        {
            $this.text = "Enabled"
            $this.ForeColor                     = "Lime"
            $script:case_sensitive              = 1;
        }
        else
        {
            $this.text = "Disabled"
            $this.ForeColor                     = "Red"
            $script:case_sensitive              = 0;
        }
        
    })
    $Form.controls.Add($case_sensitive_checkbox);


    ##################################################################################
    ###########Must Contain Label
    $y_pos = $y_pos + 30
    $must_contain_label = New-Object System.Windows.Forms.Label
    $must_contain_label.Location = New-Object System.Drawing.Point(15, $y_pos)
    $must_contain_label.Size = "250, 23"
    $must_contain_label.ForeColor = "White" 
    $must_contain_label.Text = "File or Folder Must Contain:"
    $must_contain_label.TextAlign = "MiddleRight"
    $must_contain_label.Font = [Drawing.Font]::New("Times New Roman", 12)
    $Form.Controls.Add($must_contain_label)

    ##################################################################################
    ###########Must Contain TextBox
    $must_contain_textbox = New-Object System.Windows.Forms.TextBox
    $must_contain_textbox.Location = New-Object System.Drawing.Point(($must_contain_label.location.x + $must_contain_label.width + 3),($y_pos))
    $must_contain_textbox.width = 250
    $must_contain_textbox.Height = 40
    $must_contain_textbox.Text = $script:must_contain
    $must_contain_textbox.font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $must_contain_textbox.Add_TextChanged({
        $script:must_contain = $this.text
        
    })
    $Form.Controls.Add($must_contain_textbox)

    ##################################################################################
    ###########Action Label
    $y_pos = $y_pos + 30
    $action_label           = New-Object System.Windows.Forms.Label
    $action_label.Location  = New-Object System.Drawing.Point(15, $y_pos)
    $action_label.Size      = "250, 23"
    $action_label.ForeColor = "White" 
    $action_label.Text      = "Action:"
    $action_label.TextAlign = "MiddleRight"
    $action_label.Font      = [Drawing.Font]::New("Times New Roman", 12)
    $Form.Controls.Add($action_label)

    ##################################################################################
    ###########Action Combo
    $word1_textbox                  = New-Object System.Windows.Forms.TextBox
    $word2_textbox                  = New-Object System.Windows.Forms.TextBox
    $word2_label                    = New-Object System.Windows.Forms.Label
    $action_combo                   = New-Object System.Windows.Forms.ComboBox	
    $action_combo.width = 250
    $action_combo.autosize = $false
    $action_combo.Anchor = 'top,right'
    $action_combo.font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $action_combo.Location = New-Object System.Drawing.Point(($action_label.location.x + $action_label.width + 3),($y_pos))
    $action_combo.DropDownStyle = "DropDownList"
    $action_combo.AccessibleName = "";
    $action_combo.Items.Add("Replace Characters:") | Out-Null
    $action_combo.Items.Add("Append Beginning:") | Out-Null
    $action_combo.Items.Add("Append End:") | Out-Null
    $action_combo.Items.Add("Append After:") | Out-Null
    $action_combo.Items.Add("Append Before:") | Out-Null
    $action_combo.Items.Add("Replace Beginning Characters:") | Out-Null
    $action_combo.Items.Add("Replace End Characters:") | Out-Null 
    $action_combo.Items.Add("Delete Everything After:") | Out-Null
    $action_combo.Items.Add("Delete Everything Before:") | Out-Null
    $action_combo.Items.Add("Insert at Position:") | Out-Null
    $action_combo.Add_SelectedValueChanged({
        if($this.SelectedItem -eq "Replace Characters:")
        {
            $word1_textbox.Visible = $true
            $word2_textbox.Visible = $true
            $word2_label.Visible = $true
            $script:action = "Replace Characters:"
        }
        elseif($this.SelectedItem -eq "Insert at Position:")
        {
            $word1_textbox.Visible = $true
            $word2_textbox.Visible = $true
            $word2_label.Visible = $true
            $script:action = "Insert at Position:"
        }
        elseif($this.SelectedItem -eq "Append Beginning:")
        {
            $word1_textbox.Visible = $false
            $word2_textbox.Visible = $true
            $word2_label.Visible = $true
            $script:action = "Append Beginning:"
        }
        elseif($this.SelectedItem -eq "Append End:")
        {
            $word1_textbox.Visible = $false
            $word2_textbox.Visible = $true
            $word2_label.Visible = $true
            $script:action = "Append End:"
        }
        elseif($this.SelectedItem -eq "Append After:")
        {
            $word1_textbox.Visible = $true
            $word2_textbox.Visible = $true
            $word2_label.Visible = $true
            $script:action = "Append After:"
        }
        elseif($this.SelectedItem -eq "Append Before:")
        {
            $word1_textbox.Visible = $true
            $word2_textbox.Visible = $true
            $word2_label.Visible = $true
            $script:action = "Append Before:"
        }
        elseif($this.SelectedItem -eq "Replace Beginning Characters:")
        {
            $word1_textbox.Visible = $true
            $word2_textbox.Visible = $true
            $word2_label.Visible = $true
            $script:action = "Replace Beginning Characters:"
        }
        elseif($this.SelectedItem -eq "Replace End Characters:")
        {
            $word1_textbox.Visible = $true
            $word2_textbox.Visible = $true
            $word2_label.Visible = $true
            $script:action = "Replace End Characters:"
        }
        elseif($this.SelectedItem -eq "Delete Everything After:")
        {
            $word1_textbox.Visible = $true
            $word2_textbox.Visible = $false
            $word2_label.Visible = $false
            $script:action = "Delete Everything After:"
        }
        elseif($this.SelectedItem -eq "Delete Everything Before:")
        {
            $word1_textbox.Visible = $true
            $word2_textbox.Visible = $false
            $word2_label.Visible = $false
            $script:action = "Delete Everything Before:"
        }
        
    })
    $action_combo.SelectedItem = $script:action
    $Form.Controls.Add($action_combo)

    ##################################################################################
    ###########Word 1 TextBox
    $word1_textbox.Location = New-Object System.Drawing.Point(($action_combo.location.x + $action_combo.width + 3),($y_pos -1))
    $word1_textbox.width = 200
    $word1_textbox.Height = 40
    $word1_textbox.Text = $script:word1
    $word1_textbox.font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $word1_textbox.Add_TextChanged({
        $script:word1 = $this.text
        
    })
    $Form.Controls.Add($word1_textbox)

    ##################################################################################
    ###########Word 2 Label
    $y_pos = $y_pos + 30
    $word2_label.Location = New-Object System.Drawing.Point(15, $y_pos)
    $word2_label.Size = "250, 23"
    $word2_label.ForeColor = "White" 
    $word2_label.Text = "With Characters:"
    $word2_label.TextAlign = "MiddleRight"
    $word2_label.Font = [Drawing.Font]::New("Times New Roman", 12)
    $Form.Controls.Add($word2_label)

    ##################################################################################
    ###########Word 2 TextBox
    $word2_textbox.Location = New-Object System.Drawing.Point(($word2_label.location.x + $word2_label.width + 3),($y_pos))
    $word2_textbox.width = 250
    $word2_textbox.Height = 40
    $word2_textbox.Text = $script:word2
    $word2_textbox.font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $word2_textbox.Add_TextChanged({
        $script:word2 = $this.text
        
    })
    $Form.Controls.Add($word2_textbox)


    ##################################################################################
    ###########Preview Label
    $y_pos = $y_pos + 50;
    $preview_label            = New-Object System.Windows.Forms.Label   
    $preview_label.Font       = New-Object System.Drawing.Font("Copperplate Gothic Bold",15,[System.Drawing.FontStyle]::Regular)
    $preview_label.Text       = "Preview"
    $preview_label.TextAlign  = "MiddleCenter"
    $preview_label.Width      = $Form.Width
    $preview_label.height     = 35
    $preview_label.ForeColor  = "white"
    $preview_label.Location   = New-Object System.Drawing.Size((($Form.width / 2) - ($Form.width / 2)),$y_pos)
    $Form.Controls.Add($preview_label)

    ##################################################################################
    ###########Separator
    $y_pos = $y_pos + 35;
    $separator_bar                             = New-Object system.Windows.Forms.Label
    $separator_bar.text                        = ""
    $separator_bar.AutoSize                    = $false
    $separator_bar.BorderStyle                 = "fixed3d"
    $separator_bar.Anchor                      = 'top,left'
    $separator_bar.width                       = ($Form.width - 50)
    $separator_bar.height                      = 1
    $separator_bar.location                    = New-Object System.Drawing.Point((($Form.width / 2) - ($separator_bar.width / 2)),$y_pos)
    $separator_bar.TextAlign                   = 'MiddleCenter'
    $Form.controls.Add($separator_bar);

    ##################################################################################
    ###########Preview Box
    $y_pos = $y_pos + 10;
    $preview_box                                = New-Object system.Windows.Forms.TextBox 
    $preview_box.Font                           = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $preview_box.Size                           = New-Object System.Drawing.Size(($Form.width - 50),250)
    $preview_box.Location                       = New-Object System.Drawing.Size((($Form.width / 2) - ($preview_box.width / 2)),$y_pos)    
    $preview_box.ReadOnly                       = $true
    $preview_box.WordWrap                       = $False
    $preview_box.Multiline                      = $True
    $preview_box.BackColor                      = "white"
    $preview_box.ScrollBars                     = "vertical"
    $Form.Controls.Add($preview_box)

    ##################################################################################
    ###########Perform Rename Button
    $y_pos = $y_pos + 255;
    $perform_renames_button = New-Object System.Windows.Forms.Button
    $perform_renames_button.Width=150
    $perform_renames_button.Height=30  
    $perform_renames_button.Location= New-Object System.Drawing.Size((($Form.width / 3) - ($perform_renames_button.width / 2)),$y_pos)  
    $perform_renames_button.BackColor = "#606060"
    $perform_renames_button.ForeColor = "White"
    $perform_renames_button.TextAlign                   = 'MiddleCenter'
    $perform_renames_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    $perform_renames_button.Text='Perform Renames'
    $perform_renames_button.enabled = $false
    $perform_renames_button.Add_Click(
    {    
        $message = "Are you sure you want to perform these renames?`n`n"
        $yesno = [System.Windows.Forms.MessageBox]::Show("$message","!!!WARNING!!!", "YesNo" , "Information" , "Button1")
        if($yesno -eq "Yes")
        {
            $Script:Timer.Stop()
            if($script:preview_job.state -eq "Running")
            {
                Stop-Job -job $script:preview_job
                Remove-Job -job $script:preview_job
            }
		    $script:real_or_simulation = 1;
            processing
            $script:real_or_simulation = 0;
            $Script:Timer.Start()
        }
    })
    $form.Controls.Add($perform_renames_button)

    ##################################################################################
    ###########Show List Button
    $show_list_button = New-Object System.Windows.Forms.Button
    $show_list_button.Width=150
    $show_list_button.Height=30  
    $show_list_button.Location= New-Object System.Drawing.Size((($Form.width / 3) + $perform_renames_button.Width - ($show_list_button.width / 2)),$y_pos)  
    $show_list_button.BackColor = "#606060"
    $show_list_button.ForeColor = "White"
    $show_list_button.TextAlign                   = 'MiddleCenter'
    $show_list_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    $show_list_button.Text='Show Full List'
    $show_list_button.enabled = $false
    $show_list_button.Add_Click(
    {    
          Invoke-Item "$script:output_file"
    })
    $form.Controls.Add($show_list_button)

    ##################################################################################
    ###########Show List Button
    $undo_renames_button = New-Object System.Windows.Forms.Button
    $undo_renames_button.Width=150
    $undo_renames_button.Height=30  
    $undo_renames_button.Location= New-Object System.Drawing.Size((($Form.width / 3) + $perform_renames_button.Width + $show_list_button.Width - ($undo_renames_button.width / 2)),$y_pos)  
    $undo_renames_button.BackColor = "#606060"
    $undo_renames_button.ForeColor = "White"
    $undo_renames_button.TextAlign                   = 'MiddleCenter'
    $undo_renames_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    $undo_renames_button.Text='Undo Renames'
    $undo_renames_button.enabled = $false
    $undo_renames_button.Add_Click(
    {    
        $previous_renames = (Get-Childitem -File -LiteralPath "$dir\Resources\Executed Renames" -Include "*.txt" -Filter "*Executed*")
        $latest_undo_file = ""
        $latest_undo_file_date = ""
        foreach($name in $previous_renames)
        {
            if($latest_undo_file -eq "")
            {
                $latest_undo_file = $name.FullName
                $latest_undo_file_date = $name.CreationTime
            }
            elseif($latest_undo_file_date -lt $name.CreationTime)
            {
                $latest_undo_file = $name.FullName
                $latest_undo_file_date = $name.CreationTime
            }
        } 
        if(Test-Path -literalpath $latest_undo_file)
        {
            $message = "Are you sure you want to undo your latest renames?`n`n"
            $yesno = [System.Windows.Forms.MessageBox]::Show("$message","!!!WARNING!!!", "YesNo" , "Information" , "Button1")
            if($yesno -eq "Yes")
            {
                $perform_renames_button.enabled = $false
                undo_renames $latest_undo_file            
            }     
        }
    })
    $form.Controls.Add($undo_renames_button)


    $Form.ShowDialog()
}
################################################################################
######Idle Timer################################################################
Function Idle_Timer
{
    ##################################################################################
    ###########Build Hash to Determine if Any changes
    $new_hash = [string]$script:target_directory + [string]$script:drill_down_folders + [string]$script:include_file_names + [string]$script:include_folder_names + [string]$script:include_extensions + [string]$script:format_titles_automatically + [string]$script:case_sensitive + [string]$script:must_contain + [string]$script:action + [string]$script:word1 + [string]$script:word2
    if(($new_hash -cne $script:hash) -or (($preview_box.text -eq "") -and ($script:preview_job.state -ne "Running")) -and ($script:target_directory -ne "Browse or Enter a file path") -and ($script:target_directory -ne ""))
    {
        $script:hash = $new_hash
        $preview_box.text = "Loading..."
        $script:start_time = Get-Date
        $show_list_button.enabled = $false
        $perform_renames_button.enabled = $false
        $form.refresh();
        save_settings
        while($script:preview_job.state -eq "Running")
        {
            Stop-Job -job $script:preview_job
            Remove-Job -job $script:preview_job
        }
        
        ##################################################################################
        ###########Create New Simulated Job
        if($script:real_or_simulation -ne 1)
        {
            processing
        }
    }
    ##################################################################################
    ###########Recieve Job
    if($script:preview_job.state -eq "Completed")
    {
        $status = Receive-Job -Job $script:preview_job
        $preview_box.text = $status
        if(Test-Path -literalpath $script:output_file)
        {
            $show_list_button.enabled = $true
            $perform_renames_button.enabled = $true
        }
        $script:preview_job = "";
    }
    elseif($script:preview_job.state -eq "Running" -and ($script:target_directory -ne "Browse or Enter a file path") -and ($script:target_directory -ne ""))
    {
        $duration = (Get-Date) - $script:start_time
        [int]$seconds = $duration.TotalSeconds
        $preview_box.text = "Loading..." + "(" + $seconds + ")"
    }
    if($script:target_directory -eq "Browse or Enter a file path" -or $script:target_directory -eq "")
    {
        $preview_box.text = "Please Select a Directory"
        $show_list_button.enabled = $false
            $perform_renames_button.enabled = $false
    }

    ##################################################################################
    ###########Check for Undo Actions
    $previous_renames = (Get-Childitem -File -LiteralPath "$dir\Resources\Executed Renames" -Include "*.txt" -Filter "*Executed*")
    if(($previous_renames.count) -ge 1)
    {
        $undo_renames_button.enabled = $true
    }
    else
    {
        if($undo_renames_button.enabled -eq $true)
        {
            $undo_renames_button.enabled = $false
        }
    }
    
}
################################################################################
######Processing################################################################
function processing
{
    ######################################################################################
    ###########Prepare Output File
    if($real_or_simulation -eq 1)
    {
        $script:output_file = build_output_file_name "Executed Renames"
        $script:output_file = "$dir\Resources\Executed Renames\$script:output_file"
    }
    else
    {
        $script:output_file = build_output_file_name "Simulation"
        $script:output_file = "$dir\Resources\Simulations\$script:output_file"
    }
    
    ######################################################################################
    ###########Process Renamer Job
    $process_renamer = {
        ##################################################################################
        ###########Process Renamer Job Local Variables
        $real_or_simulation          = $using:real_or_simulation
        $target_directory            = $using:target_directory
        $drill_down_folders          = $using:drill_down_folders
        $include_file_names          = $using:include_file_names
        $include_folder_names        = $using:include_folder_names
        $include_extensions          = $using:include_extensions
        $case_sensitive              = $using:case_sensitive
        $format_titles_automatically = $using:format_titles_automatically
        $must_contain                = $using:must_contain
        $action                      = $using:action
        $word1                       = $using:word1
        $word2                       = $using:word2
        $max_files                   = $using:max_files
        $output_file                 = $using:output_file

        $number_of_changes_found     = 0

        $return_block = "";
        $go = 1;

        
        ##################################################################################
        ###########Error Checking
        if($target_directory -eq "Browse or Enter a file path")
        {
            $go = 0;
        }
        elseif(!(Test-Path -literalpath "$target_directory") -or ("$target_directory" -eq "") -or ($target_directory -eq $null))
        {
            $return_block = $return_block + "Path Error: Invalid Directory`r`n"
            $go = 0;
        }    
        if($go -eq 1)
        {
            if($action -eq "Replace Characters:")
            {
                if($word1 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing Replace Characters Parameter 1`r`n"
                    $go = 0;
                }
                if($word1 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
                if($word2 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
            }
            elseif($action -eq "Insert at Position:")
            {
                if(!($word1 -match "^\d+$"))
                {
                    $return_block = $return_block + "Action Error: Parameter 1 must contain a number`r`n"
                    $go = 0;
                }
                if($word2 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
            }
            elseif($action -eq "Append Beginning:")
            {
                if($word2 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing Append Beginning Parameter`r`n"
                    $go = 0;
                }
                if($word2 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
            }
            elseif($action -eq "Append End:")
            {
                if($word2 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing Append End Parameter`r`n"
                    $go = 0;
                }
                if($word2 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
            }
            elseif($action -eq "Append After:")
            {
                if($word1 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing Append After Parameter 1`r`n"
                    $go = 0;
                }
                if($word1 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
                if($word2 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing Append After Parameter 2`r`n"
                    $go = 0;
                }
                if($word2 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
            }
            elseif($action -eq "Append Before:")
            {
                if($word1 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing Append Before Parameter 1`r`n"
                    $go = 0;
                }
                if($word1 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
                if($word2 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing Append Before Parameter 2`r`n"
                    $go = 0;
                }
                if($word2 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
            }
            elseif($action -eq "Replace Beginning Characters:")
            {
                if($word1 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing Beginning Characters Parameter 1`r`n"
                    $go = 0;
                }
                if($word1 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
            }
            elseif($action -eq "Replace End Characters:")
            {
                if($word1 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing End Characters Parameter 1`r`n"
                    $go = 0;
                }
                if($word1 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
            }
            elseif($action -eq "Delete Everything After:")
            {
                if($word1 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing Delete Everything After Parameter`r`n"
                    $go = 0;
                }
                if($word1 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
            }
            elseif($action -eq "Delete Everything Before:")
            {
                if($word1 -eq "")
                {
                    $return_block = $return_block + "Action Error: Missing Delete Everything Before Parameter`r`n"
                    $go = 0;
                }
                if($word1 -match '["*:<>?/\\{|}]+')
                {
                    $return_block = $return_block + "Action Error: Invalid Characters`r`n"
                    $go = 0;
                }
            }
        }

        

        if($go -eq 1)
        {
            $writer = [System.IO.StreamWriter]::new($output_file)
            ################################################################
            if($real_or_simulation -eq 1)
            {
                $return_block = $return_block + "Changes Made:`r`n`r`n"
            }
            else
            {
                $return_block = $return_block + "Simulation:`r`n(No Changes Were Made)`r`n`r`n"
            }
            ################################################################
            if($include_folder_names -eq 1)
            {
                ##################################################################################
                ###########Collect All Folders
                $folder_simulation = @{}; #This allows us to project what file paths would be without actually renaming folders (Simulation)
                $folders = "";
                if($drill_down_folders -eq 1)
                {
                    $folders = Get-ChildItem -LiteralPath "$target_directory" -directory -Recurse -ErrorAction SilentlyContinue | Select-Object -First $max_files # | where {! $_.PSIsContainer}
                }
                else
                {
                    $folders = Get-ChildItem -LiteralPath "$target_directory" -directory -ErrorAction SilentlyContinue # | where {! $_.PSIsContainer}
                }
                ##################################################################################
                ###########Process Folders
                foreach($folder in $folders)
                {
                    $folder_name = $folder.name
                    $new_folder = $folder.name 
                    $full_path = $folder.FullName
                    $parent_dir = $folder.Parent.FullName + "\"
                    $sim_path = "";
                    ##################################################################################
                    ############Must Contain
                    if($must_contain -ne "")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $index = $folder_name.IndexOf($must_contain)
                        }
                        else
                        {
                            $index = $folder_name.ToLower().IndexOf($must_contain.ToLower())
                        }
                        if($index -eq -1)
                        {
                            #Move to next Item
                            Continue;
                        }
                    }
                    ##################################################################################
                    ###########Build Folder Simulation (This allows us to project what file paths would be without actually renaming folders (Simulation))
                    $parent_dir_split = $parent_dir -split '\\' 
                    $count = 0;                
                    foreach($leaf in $parent_dir_split)
                    {
                        $sim_path = "";
                        for($i = 0; $i -lt $count; $i++)
                        {
                            if($sim_path -eq "")
                            {
                                $sim_path = $parent_dir_split[$i] #Root Folder No Slash
                            }
                            else
                            {
                                $sim_path = $sim_path + "\" + $parent_dir_split[$i]
                            }
                            if($folder_simulation.Contains($sim_path))
                            {
                                $sim_path = $folder_simulation[$sim_path]
                            }
                        }
                        $count++
                    }
                    
                    ##################################################################################
                    ###########Replace Characters Action (Folders)
                    if($action -eq "Replace Characters:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            
                            $new_folder = $new_folder -creplace "$([regex]::Escape($word1))","$word2"
                            
                        }
                        else
                        {
                            $new_folder = $new_folder -replace "$([regex]::Escape($word1))","$word2"
                        }
                    }
                    ##################################################################################
                    ###########Insert at Position: (Folders)
                    if($action -eq "Insert at Position:")
                    {
                        if($word1 -gt $new_folder.length)
                        {
                            $word1 = $new_folder.length
                        }
                        $new_folder = $new_folder.insert($word1,$word2)
                    }
                    ##################################################################################
                    ###########Append Beginning: (Folders)
                    if($action -eq "Append Beginning:")
                    {
                        $new_folder = $word2 + $new_folder
                    }
                    ##################################################################################
                    ###########Append End: (Folders)
                    if($action -eq "Append End:")
                    {
                        $new_folder = $new_folder + $word2
                    }
                    ##################################################################################
                    ###########Append After: (Folders)
                    if($action -eq "Append After:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $index = $new_folder.IndexOf($word1)
                        }
                        else
                        {
                            $index = $new_folder.ToLower().IndexOf($word1.ToLower())
                        }     
                        if($index -ne -1)
                        {
                           $index = $index + $word1.Length
                           $new_folder = $new_folder.Insert($index,$word2)
                        }
                    }
                    ##################################################################################
                    ###########Append Before: (Folders)
                    if($action -eq "Append Before:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $index = $new_folder.IndexOf($word1)
                        }
                        else
                        {
                            $index = $new_folder.ToLower().IndexOf($word1.ToLower())
                        }
                        if($index -ne -1)
                        {
                           $new_folder = $new_folder.Insert($index,$word2)
                        }
                    }
                    ##################################################################################
                    ###########Replace Beginning Characters: (Folders)
                    if($action -eq "Replace Beginning Characters:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $new_folder = $new_folder -creplace "^$([regex]::Escape($word1))","$word2"
                        }
                        else
                        {
                            $new_folder = $new_folder -replace "^$([regex]::Escape($word1))","$word2"
                        }
                        
                    }
                    ##################################################################################
                    ###########Replace End Characters: (Folders)
                    if($action -eq "Replace End Characters:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $new_folder = $new_folder -creplace "$([regex]::Escape($word1))$","$word2"
                        }
                        else
                        {
                            $new_folder = $new_folder -replace "$([regex]::Escape($word1))$","$word2"
                        }
                        
                    }
                    ##################################################################################
                    ###########Delete Everything After: (Folders)
                    if($action -eq "Delete Everything After:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $index = $new_folder.IndexOf($word1)
                        }
                        else
                        {
                            $index = $new_folder.ToLower().IndexOf($word1.ToLower())
                        }
                        if($index -ne -1)
                        {
                            $index = $index + $word1.Length
                            $new_folder = $new_folder.Substring(0,$index)
                        }
                    }
                    ##################################################################################
                    ###########Delete Everything Before: (Folders)
                    if($action -eq "Delete Everything Before:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $index = $new_folder.IndexOf($word1)
                        }
                        else
                        {
                            $index = $new_folder.ToLower().IndexOf($word1.ToLower())
                        }
                        if($index -ne -1)
                        {
                            $new_folder = $new_folder.Substring($index,($new_folder.Length - $index))
                        }
                    }
                    ##################################################################################
                    ###########Format Titles: (Folders)
                    if($format_titles_automatically -eq 1)
                    {
                        $new_folder = (Get-Culture).TextInfo.ToTitleCase($new_folder)
                    }
                    ##################################################################################
                    ###########Finalize Folder Actions
                    $folder_simulation.Add("$sim_path\$folder_name","$sim_path\$new_folder")
                    #write-host ON $sim_path\$folder_name 
                    #write-host NN $sim_path\$new_folder
                    #write-host

                    $old_name = "$sim_path\$folder_name"
                    $new_name = "$sim_path\$new_folder"

                    $old_name = $old_name -replace '\\\\','\'
                    $new_name = $new_name -replace '\\\\','\'


                    ##Fix Network Paths
                    if($old_name -match "^\\"){$old_name = "\$old_name"}
                    if($new_name -match "^\\"){$new_name = "\$new_name"}

                    if($old_name -cne $new_name)
                    {
                        $number_of_changes_found++
                        $status = "Not Renamed"
                        if($real_or_simulation -eq 1)
                        {
                            if($old_name -eq $new_name)
                            {   #Changing Case Only
                                Rename-Item -LiteralPath $old_name -NewName "$new_name-Temp"
                                Rename-Item -LiteralPath "$new_name-Temp" -NewName $new_name
                                if((Test-Path -LiteralPath "$new_name-Temp"))
                                {
                                    $status = "Rename Failed Temp Transfer"  
                                }
                                else
                                {
                                    $status = "Rename Success"

                                }
                            }
                            else
                            {
                                Rename-Item -LiteralPath $old_name -NewName $new_name
                                if((!(Test-Path -LiteralPath $old_name)) -and (Test-Path -LiteralPath $new_name))
                                {
                                    $status = "Rename Success"
                                }
                                else
                                {
                                    $status = "Rename Failed"  
                                }
                            }
                        }
                        ########################################
                        #Document work
                        if($number_of_changes_found -le 30)
                        {
                            $return_block = $return_block + "Folder: (Change#$number_of_changes_found)`r`n"
                            $return_block = $return_block + "$old_name`r`n"
                            $return_block = $return_block + "$new_name`r`n"
                            $return_block = $return_block + "$status`r`n`r`n"

                            $writer.WriteLine("Folder (Change#$number_of_changes_found)")
                            $writer.WriteLine($old_name)
                            $writer.WriteLine($new_name)
                            $writer.WriteLine("$status")
                            $writer.WriteLine("")
                        }
                        else
                        {
                            $writer.WriteLine("File (Change#$number_of_changes_found)")
                            $writer.WriteLine($old_name)
                            $writer.WriteLine($new_name)
                            $writer.WriteLine("$status")
                            $writer.WriteLine("")
                        }     
                    }
                }#Foreach Folder
            }#Include Folder Names On/Off
            ##################################################################################
            ###########Collect All Files
            if($include_file_names -eq 1)
            {
                $files = "";
                if($drill_down_folders -eq 1)
                {
                    $files = Get-ChildItem -LiteralPath "$target_directory" -file -Recurse -ErrorAction SilentlyContinue | Select-Object -First $max_files
                }
                else
                {
                    $files = Get-ChildItem -LiteralPath "$target_directory" -file -ErrorAction SilentlyContinue
                }
                ##################################################################################
                ###########Process Files
                foreach($file in $files)
                {
                    $file_name = $file.Basename
                    
                    $extension = $file.Extension
                    if($include_extensions -eq 1)
                    {
                        $file_name = $file_name + $extension
                        $extension = "";
                    }
                    $new_file_name = $file_name
                    $full_path = $file.FullName
                    $parent_dir = [string]$file.Directory + "\"
                    $sim_path = [string]$file.Directory

                    ##################################################################################
                    ############Must Contain
                    if($must_contain -ne "")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $index = $file_name.IndexOf($must_contain)
                        }
                        else
                        {
                            $index = $file_name.ToLower().IndexOf($file_name.ToLower())
                        }
                        if($index -eq -1)
                        {
                            #Move to next Item
                            Continue;
                        }
                    }

                    
                    ##################################################################################
                    ###########Build Folder Simulation for Files (This allows us to project what file paths would be without actually renaming folders (Simulation))
                    if($include_folder_names -eq 1)
                    {
                        $parent_dir_split = $parent_dir -split '\\' 
                        $count = 0;                
                        foreach($leaf in $parent_dir_split)
                        {
                            $sim_path = "";
                            for($i = 0; $i -lt $count; $i++)
                            {
                                if($sim_path -eq "")
                                {
                                    $sim_path = $parent_dir_split[$i] #Root Folder No Slash
                                }
                                else
                                {
                                    $sim_path = $sim_path + "\" + $parent_dir_split[$i]
                                }
                                if($folder_simulation.Contains($sim_path))
                                {
                                    $sim_path = $folder_simulation[$sim_path]
                                }
                            }
                            $count++
                        }
                    }
                    $sim_path = $sim_path -replace "\\\\","\\"
                    ##################################################################################
                    ###########Replace Characters Action (Files)
                    if($action -eq "Replace Characters:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $new_file_name = $new_file_name -creplace "$([regex]::Escape("$word1"))","$word2"
                        }
                        else
                        {
                            $new_file_name = $new_file_name -replace "$([regex]::Escape($word1))","$word2"
                        }
                    }
                    ##################################################################################
                    ###########Insert at Position: (Folders)
                    if($action -eq "Insert at Position:")
                    {
                        if($word1 -gt $new_file_name.length)
                        {
                            $word1 = $new_file_name.length
                        }
                        $new_file_name = $new_file_name.insert($word1,$word2)
                    }
                    ##################################################################################
                    ###########Append Beginning: (Files)
                    if($action -eq "Append Beginning:")
                    {
                        $new_file_name = $word2 + $new_file_name
                    }
                    ##################################################################################
                    ###########Append End: (Files)
                    if($action -eq "Append End:")
                    {
                        $new_file_name = $new_file_name + $word2
                    }
                    ##################################################################################
                    ###########Append After: (Files)
                    if($action -eq "Append After:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $index = $new_file_name.IndexOf($word1)
                        }
                        else
                        {
                            $index = $new_file_name.ToLower().IndexOf($word1.ToLower())
                        }     
                        if($index -ne -1)
                        {
                           $index = $index + $word1.Length
                           $new_file_name = $new_file_name.Insert($index,$word2)
                        }
                    }
                    ##################################################################################
                    ###########Append Before: (Files)
                    if($action -eq "Append Before:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $index = $new_file_name.IndexOf($word1)
                        }
                        else
                        {
                            $index = $new_file_name.ToLower().IndexOf($word1.ToLower())
                        }
                        if($index -ne -1)
                        {
                           $new_file_name = $new_file_name.Insert($index,$word2)
                        }
                    }
                    ##################################################################################
                    ###########Replace Beginning Characters: (Files)
                    if($action -eq "Replace Beginning Characters:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $new_file_name = $new_file_name -creplace "^$([regex]::Escape($word1))","$word2"
                        }
                        else
                        {
                            $new_file_name = $new_file_name -replace "^$([regex]::Escape($word1))","$word2"
                        }
                        
                    }
                    ##################################################################################
                    ###########Replace End Characters: (Files)
                    if($action -eq "Replace End Characters:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $new_file_name = $new_file_name -creplace "$([regex]::Escape($word1))$","$word2"
                        }
                        else
                        {
                            $new_file_name = $new_file_name -replace "$([regex]::Escape($word1))$","$word2"
                        }
                        
                    }
                    ##################################################################################
                    ###########Delete Everything After: (Files)
                    if($action -eq "Delete Everything After:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $index = $new_file_name.IndexOf($word1)
                        }
                        else
                        {
                            $index = $new_file_name.ToLower().IndexOf($word1.ToLower())
                        }
                        if($index -ne -1)
                        {
                            $index = $index + $word1.Length
                            $new_file_name = $new_file_name.Substring(0,$index)
                        }
                    }
                    ##################################################################################
                    ###########Delete Everything Before: (Files)
                    if($action -eq "Delete Everything Before:")
                    {
                        if($case_sensitive -eq 1)
                        {
                            $index = $new_file_name.IndexOf($word1)
                        }
                        else
                        {
                            $index = $new_file_name.ToLower().IndexOf($word1.ToLower())
                        }
                        if($index -ne -1)
                        {
                            $new_file_name = $new_file_name.Substring($index,($new_file_name.Length - $index))
                        }
                    }
                    ##################################################################################
                    ###########Format Titles: (Files)
                    if($format_titles_automatically -eq 1)
                    {
                        $new_file_name = (Get-Culture).TextInfo.ToTitleCase($new_file_name)
                    }
                    ##################################################################################
                    ###########Finalize Folder Actions
                    #write-host ON $sim_path\$file_name$extension 
                    #write-host NN $sim_path\$new_file_name$extension
                    #write-host

                    $old_name = "$sim_path\$file_name$extension"
                    $new_name = "$sim_path\$new_file_name$extension"




                    $old_name = $old_name -replace '\\\\','\'
                    $new_name = $new_name -replace '\\\\','\'

                    ##Fix Network Paths
                    if($old_name -match "^\\"){$old_name = "\$old_name"}
                    if($new_name -match "^\\"){$new_name = "\$new_name"}



                    if($old_name -cne $new_name)
                    {
                        $number_of_changes_found++
                        
                        ############################################
                        ######Make Changes
                        $status = "Not Renamed"        
                        if($real_or_simulation -eq 1)
                        {
                            
                            if($old_name -eq $new_name)
                            {   #Changing Case Only
                                Rename-Item -LiteralPath $old_name -NewName "$new_name-Temp"
                                Rename-Item -LiteralPath "$new_name-Temp" -NewName $new_name
                                if((Test-Path -LiteralPath "$new_name-Temp"))
                                {
                                    $status = "Rename Failed Temp Transfer"  
                                }
                                else
                                {
                                    $status = "Rename Success"
                                }
                            }
                            else
                            {
                                Rename-Item -LiteralPath $old_name -NewName $new_name
                                if((!(Test-Path -LiteralPath $old_name)) -and (Test-Path -LiteralPath $new_name))
                                {
                                    $status = "Rename Success"
                                }
                                else
                                {
                                    $status = "Rename Failed"   
                                }
                            }
                        }

                        ########################################
                        #Document work
                        if($number_of_changes_found -le 30)
                        {
                            $return_block = $return_block + "File: (Change#$number_of_changes_found)`r`n"
                            $return_block = $return_block + "$old_name`r`n"
                            $return_block = $return_block + "$new_name`r`n"
                            $return_block = $return_block + "$status`r`n`r`n"

                            $writer.WriteLine("File (Change#$number_of_changes_found)")
                            $writer.WriteLine($old_name)
                            $writer.WriteLine($new_name)
                            $writer.WriteLine("$status")
                            $writer.WriteLine("")
                        }
                        else
                        {
                            $writer.WriteLine("File (Change#$number_of_changes_found)")
                            $writer.WriteLine($old_name)
                            $writer.WriteLine($new_name)
                            $writer.WriteLine("$status")
                            $writer.WriteLine("")
                        }       
                    }
                }
            }#Include Files On/Off
            ############################################
            #Final Information
            $writer.close()
            if($number_of_changes_found -gt 30)
            {
                $return_block = $return_block + "Showing 30 of $number_of_changes_found Changes Found`r`n`r`n"
            }
            elseif($number_of_changes_found -eq 0)
            {
                $return_block = $return_block + "No Files/Folder Meet Requirements`r`n`r`n"
                if(Test-Path -literalpath $output_file)
                {
                    Remove-Item -literalpath $output_file
                }
            }
            else
            {
                $return_block = $return_block + "$number_of_changes_found Changes Found`r`n`r`n"
            }

            
        }#Go = 1


        Write-output $return_block
        #return $return_block
    }#Process Jop

    ##################################################################################
    ###########Start Process Renamer Job
    $script:preview_job = Start-Job -ScriptBlock  $process_renamer  
}
################################################################################
######Prompt for Folder#########################################################
function prompt_for_folder()
{  
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}
################################################################################
######Initial Checks############################################################
function initial_checks
{
    if(!(Test-Path -LiteralPath "$dir\Resources"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources"
    }
    if(!(Test-Path -LiteralPath "$dir\Resources\Required"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources\Required"
    }
    if(!(Test-Path -LiteralPath "$dir\Resources\Executed Renames"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources\Executed Renames"
    }
    if(!(Test-Path -LiteralPath "$dir\Resources\Simulations"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources\Simulations"
    }
    else
    {
        #Clean Directory
        Get-ChildItem -LiteralPath "$dir\Resources\Simulations" -File | Remove-Item
    }
    if(!(Test-Path -LiteralPath "$dir\Resources\Undone Renames"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources\Undone Renames"
    }  
    ###################################################################################
    if(!(Test-Path -LiteralPath "$dir\Resources\Required\Settings.csv"))
    {
        save_settings
    }
}
################################################################################
######CSV Line to Array#########################################################
function csv_line_to_array ($line)
{
    if($line -match "^,")
    {
        $line = ",$line"; 
    }
    Select-String '(?:^|,)(?=[^"]|(")?)"?((?(1)[^"]*|[^,"]*))"?(?=,|$)' -input $line -AllMatches | Foreach { [System.Collections.ArrayList]$line_split = $_.matches -replace '^,|"',''}
    return $line_split
}
#################################################################################
######Load Settings##############################################################
function load_settings
{
    if(Test-Path -literalpath "$dir\Resources\Required\Settings.csv")
    {
        $line_count = 0;
        $reader = [System.IO.File]::OpenText("$dir\Resources\Required\Settings.csv")
        while($null -ne ($line = $reader.ReadLine()))
        {
            $line_count++;
            if($line_count -ne 1)
            {
                ($key,$value) = csv_line_to_array $line
                #write-host $key
                #write-host $value
                if(!($script:settings.containskey($key)))
                {
                    $script:settings.Add($key,$value);
                }
            } 
        }
        $reader.close(); 
        ###################################################
        if($script:settings.contains('TARGET_DIRECTORY'))
        {
            $script:target_directory = $script:settings['TARGET_DIRECTORY']
        }
        if($script:settings.contains('DRILL_DOWN_FOLDERS'))
        {
            $script:drill_down_folders = $script:settings['DRILL_DOWN_FOLDERS']
        }
        if($script:settings.contains('INCLUDE_FILE_NAMES'))
        {
            $script:include_file_names = $script:settings['INCLUDE_FILE_NAMES']
        }
        if($script:settings.contains('INCLUDE_FOLDER_NAMES'))
        {
            $script:include_folder_names = $script:settings['INCLUDE_FOLDER_NAMES']
        }
        if($script:settings.contains('INCLUDE_EXTENSIONS'))
        {
            $script:include_extensions = $script:settings['INCLUDE_EXTENSIONS']
        }
        if($script:settings.contains('FORMAT_TITLES'))
        {
            $script:format_titles_automatically = $script:settings['FORMAT_TITLES']
        }
        if($script:settings.contains('CASE_SENSITIVE'))
        {
            $script:case_sensitive = $script:settings['CASE_SENSITIVE']
        }
        if($script:settings.contains('MUST_CONTAIN'))
        {
            $script:must_contain = $script:settings['MUST_CONTAIN']
        }
        if($script:settings.contains('ACTION'))
        {
            $script:action = $script:settings['ACTION']
        }
        if($script:settings.contains('WORD1'))
        {
            $script:word1 = $script:settings['WORD1']
        }
        if($script:settings.contains('WORD2'))
        {
            $script:word2 = $script:settings['WORD2']
        }
    }
}
#################################################################################
######Save Settings##############################################################
function save_settings
{
    if(Test-Path -literalpath "$dir\Resources\Required\Settings.csv")
    {
        Remove-Item -literalpath "$dir\Resources\Required\Settings.csv"
    }
    $settings_writer = new-object system.IO.StreamWriter("$dir\Resources\Required\Settings.csv",$true)
    $settings_writer.write("PROPERTY,VALUE`r`n");
    $settings_writer.write("TARGET_DIRECTORY,$script:target_directory`r`n");
    $settings_writer.write("DRILL_DOWN_FOLDERS,$script:drill_down_folders`r`n");
    $settings_writer.write("INCLUDE_FILE_NAMES,$script:include_file_names`r`n");
    $settings_writer.write("INCLUDE_FOLDER_NAMES,$script:include_folder_nameS`r`n");
    $settings_writer.write("INCLUDE_EXTENSIONS,$script:include_extensions`r`n");
    $settings_writer.write("FORMAT_TITLES,$script:format_titles_automatically`r`n");
    $settings_writer.write("CASE_SENSITIVE,$script:case_sensitive`r`n");
    $settings_writer.write("MUST_CONTAIN,$script:must_contain`r`n");
    $settings_writer.write("ACTION,$script:action`r`n");
    $settings_writer.write("WORD1,$script:word1`r`n");
    $settings_writer.write("WORD2,$script:word2`r`n");
    $settings_writer.close();
}
################################################################################
function build_output_file_name($type)
{  
    $date = Get-Date -Format G
    [regex]$pattern = " "
    $date = $pattern.replace($date, " @ ", 1);
    $date = $date.replace('/',"-");
    $date = $date.replace(':',".");

    $output = "$type      ($date)" + ".txt";
    return $output
}
######Show Console##############################################################
function Show-Console
{
    param ([Switch]$Show,[Switch]$Hide)
    if (-not ("Console.Window" -as [type])) { 

        Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        '
    }

    if ($Show)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()

        # Hide = 0,
        # ShowNormal = 1,
        # ShowMinimized = 2,
        # ShowMaximized = 3,
        # Maximize = 3,
        # ShowNormalNoActivate = 4,
        # Show = 5,
        # Minimize = 6,
        # ShowMinNoActivate = 7,
        # ShowNoActivate = 8,
        # Restore = 9,
        # ShowDefault = 10,
        # ForceMinimized = 11

        $null = [Console.Window]::ShowWindow($consolePtr, 5)
    }

    if ($Hide)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        #0 hide
        $null = [Console.Window]::ShowWindow($consolePtr, 0)
    }
}
################################################################################
######Undo Renames##############################################################
function undo_renames($undo_file)
{
    $undo_hash = @{};
    $count = 0;
    $reader = New-Object IO.StreamReader $undo_file
    $old_name = ""
    $new_name = ""
    #########################################################################
    #########Extract Contents
    while($null -ne ($line = $reader.ReadLine()))
    {
        if($line -ne "")
        {
            if(($old_name -eq "") -and ($line -match "\\"))
            {
                $old_name = $line
            }
            elseif(($old_name -ne "")-and ($line -match "\\"))
            {
                $new_name = $line
            }
            if($old_name -and $new_name)
            {
                $count++
                #write-host Match $count
                #write-host $old_name
                #write-host $new_name
                #write-host            
                $undo_hash.Add($count,"$old_name::$new_name");
                $old_name = ""
                $new_name = ""
            }
        }
    }
    $reader.Close()
    #########################################################################
    #########Reverse Actions
    $count = 0;
    $undo_output = build_output_file_name "Undone Renames"
    $undo_output = "$dir\Resources\Undone Renames\$undo_output"
    $writer = [System.IO.StreamWriter]::new($undo_output)
    $return_block = $return_block + "Undone Changes`r`n`r`n"
    foreach($match in $undo_hash.GetEnumerator() | sort Key -Descending)
    {
        ($new_name, $old_name) = $match.value -split "::"
        #write-host $match.key
        #write-host $new_name
        #write-host $old_name
        #write-host
        if(Test-Path -LiteralPath $old_name)
        {
            $status = ""
            $count++;
            
            if($old_name -eq $new_name)
            {   
                #Changing Case Only
                Rename-Item -LiteralPath $old_name -NewName "$new_name-Temp"
                Rename-Item -LiteralPath "$new_name-Temp" -NewName $new_name
                if((Test-Path -LiteralPath "$new_name-Temp"))
                {
                    $status = "Rename Failed Temp Transfer"
                }
                else
                {
                    $status = "Rename Success"
                }
            }
            else
            {
                Rename-Item -LiteralPath $old_name -NewName $new_name
                if((!(Test-Path -LiteralPath $old_name)) -and (Test-Path -LiteralPath $new_name))
                {
                    $status = "Rename Success"
                }
                else
                {
                    $status = "Rename Failed"  
                }
            }
            $writer.WriteLine("Undo #$count")
            $writer.WriteLine("$old_name")
            $writer.WriteLine("$new_name")
            $writer.WriteLine("$status")
            $writer.WriteLine("")
            if($count -le 30)
            {
                $return_block = $return_block + "Undo #$count`r`n"
                $return_block = $return_block + "$old_name`r`n"
                $return_block = $return_block + "$new_name`r`n"
                $return_block = $return_block + "$status`r`n`r`n"
            } 
        }
    }
    if($count -gt 30)
    {
        $return_block = $return_block + "Showing 30 of $number_of_changes_found Changes Found`r`n`r`n"
    }
    $writer.close()
    if(Test-Path -literalpath $undo_file)
    {
        ###Remove the changes file
        Remove-Item -literalpath $undo_file
    }
    $script:output_file = $undo_output
    $preview_box.text = $return_block

}
################################################################################
######Main Sequence Start#######################################################
Show-Console -Hide
initial_checks
load_settings
$Script:Timer.Add_Tick({Idle_Timer})
$Script:Timer.Start()
#undo_renames
main | Out-Null

##Bug Fixed 22 Aug 2021
##Fixed Network Paths
## 