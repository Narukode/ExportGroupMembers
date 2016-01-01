#╔════════════════════════════════════╗
#║Load AD module and needed assemblies║
#╚════════════════════════════════════╝

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
Import-Module ActiveDirectory


#╔═════════╗
#║Main Form║
#╚═════════╝

$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Width = 1470
$MainForm.Height = 700
$MainForm.Text = "Group Member Export"
$MainForm.FormBorderStyle = 'FixedDialog'


#╔════════════════════╗
#║Group Search textBox║
#╚════════════════════╝

$textBox = New-Object System.Windows.Forms.TextBox 
$textBox.Location = New-Object System.Drawing.Point(5,5) 
$textBox.Size = New-Object System.Drawing.Size(602,20) 

$MainForm.Controls.Add($textBox)


#╔═════════╗
#║List View║
#╚═════════╝

$listView = New-Object System.Windows.Forms.ListView
$listView.View = 'Details'
$listView.Width = 675
$listView.Height = 625
$listView.Top = 30
$listView.Left = 5

$listView.Columns.Add('Name', 175)
$listView.Columns.Add('Type', 75)
$listView.Columns.Add('Description', 175)
$listView.Columns.Add('OU', 190)

$MainForm.Controls.Add($listView)


#╔═══════╗
#║Buttons║
#╚═══════╝

#>Search Button
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(611,4)
$searchButton.Size = New-Object System.Drawing.Size(70,22)
$searchButton.Text = "Search"

$MainForm.Controls.Add($searchButton)

#>Members Button
$membersButton = New-Object System.Windows.Forms.Button
$membersButton.Location = New-Object System.Drawing.Point(685,300)
$membersButton.Size = New-Object System.Drawing.Size(40,25)
$membersButton.Text = "->"

$MainForm.Controls.Add($membersButton)

#>Export CSV Button
$exportCSVButton = New-Object System.Windows.Forms.Button
$exportCSVButton.Location = New-Object System.Drawing.Point(1360,4)
$exportCSVButton.Size = New-Object System.Drawing.Size(87,20)
$exportCSVButton.Text = "Export CSV"

$MainForm.Controls.Add($exportCSVButton)

#>Export HTML Button
$exportHTMLButton = New-Object System.Windows.Forms.Button
$exportHTMLButton.Location = New-Object System.Drawing.Point(1360,26)
$exportHTMLButton.Size = New-Object System.Drawing.Size(87,20)
$exportHTMLButton.Text = "Export HTML"

$MainForm.Controls.Add($exportHTMLButton)


#╔═════════════╗
#║Datagrid View║
#╚═════════════╝

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size=New-Object System.Drawing.Size(717,605)

$dataGridView.ColumnHeadersVisible = $true
$dataGridView.Left = 730
$dataGridView.Top = 50

$MainForm.Controls.Add($dataGridView)


#╔═══════════════════╗
#║Logon Name CheckBox║
#╚═══════════════════╝

$logon_CkBx = New-Object System.Windows.Forms.CheckBox
$logon_CkBx.Text = "Logon Name"
$logon_CkBx.Checked = $true

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$logon_CkBx.Size = $System_Drawing_Size

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 730
$System_Drawing_Point.Y = 1
$logon_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($logon_CkBx)


#╔══════════════════╗
#║Full Name CheckBox║
#╚══════════════════╝

$full_CkBx = New-Object System.Windows.Forms.CheckBox
$full_CkBx.Text = "Full Name"
$full_CkBx.Checked = $true

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$full_CkBx.Size = $System_Drawing_Size

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 820
$System_Drawing_Point.Y = 1
$full_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($full_CkBx)


#╔═══════════════════╗
#║First Name CheckBox║
#╚═══════════════════╝

$fn_CkBx = New-Object System.Windows.Forms.CheckBox
$fn_CkBx.Text = "First Name"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$fn_CkBx.Size = $System_Drawing_Size

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 910
$System_Drawing_Point.Y = 1
$fn_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($fn_CkBx)


#╔══════════════════╗
#║Last Name CheckBox║
#╚══════════════════╝

$ln_CkBx = New-Object System.Windows.Forms.CheckBox
$ln_CkBx.Text = "Last Name"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$ln_CkBx.Size = $System_Drawing_Size

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1000
$System_Drawing_Point.Y = 1
$ln_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($ln_CkBx)


#╔══════════════╗
#║Email CheckBox║
#╚══════════════╝

$email_CkBx = New-Object System.Windows.Forms.CheckBox
$email_CkBx.Text = "Email"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$email_CkBx.Size = $System_Drawing_Size

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1090
$System_Drawing_Point.Y = 1
$email_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($email_CkBx)


#╔═══════════╗
#║OU CheckBox║
#╚═══════════╝

$ou_CkBx = New-Object System.Windows.Forms.CheckBox
$ou_CkBx.Text = "OU"
$ou_CkBx.Checked = $true

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$ou_CkBx.Size = $System_Drawing_Size

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1180
$System_Drawing_Point.Y = 1
$ou_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($ou_CkBx)


#╔════════════════╗
#║Manager CheckBox║
#╚════════════════╝

$manager_CkBx = New-Object System.Windows.Forms.CheckBox

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$manager_CkBx.Size = $System_Drawing_Size

$manager_CkBx.Text = "Manager"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1270
$System_Drawing_Point.Y = 1
$manager_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($manager_CkBx)


#Row 2
#╔═══════════════════╗
#║Employee # CheckBox║
#╚═══════════════════╝

$emp_CkBx = New-Object System.Windows.Forms.CheckBox
$emp_CkBx.Text = "Emp#"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$emp_CkBx.Size = $System_Drawing_Size

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 730
$System_Drawing_Point.Y = 25
$emp_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($emp_CkBx)


#╔════════════════════╗
#║Description CheckBox║
#╚════════════════════╝

$desc_CkBx = New-Object System.Windows.Forms.CheckBox
$desc_CkBx.Text = "Desc"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$desc_CkBx.Size = $System_Drawing_Size

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 820
$System_Drawing_Point.Y = 25
$desc_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($desc_CkBx)


#╔═══════════════════╗
#║Department CheckBox║
#╚═══════════════════╝

$dept_CkBx = New-Object System.Windows.Forms.CheckBox
$dept_CkBx.Text = "Dept"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$dept_CkBx.Size = $System_Drawing_Size

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 910
$System_Drawing_Point.Y = 25
$dept_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($dept_CkBx)


#╔══════════════╗
#║Title CheckBox║
#╚══════════════╝

$title_CkBx = New-Object System.Windows.Forms.CheckBox
$title_CkBx.Text = "Title"
$title_CkBx.Checked = $true

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$title_CkBx.Size = $System_Drawing_Size

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1000
$System_Drawing_Point.Y = 25
$title_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($title_CkBx)


#╔════════════════╗
#║Enabled CheckBox║
#╚════════════════╝

$enabled_CkBx = New-Object System.Windows.Forms.CheckBox
$enabled_CkBx.Text = "Enabled"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$enabled_CkBx.Size = $System_Drawing_Size


$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1090
$System_Drawing_Point.Y = 25
$enabled_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($enabled_CkBx)


#╔═══════════════╗
#║Office CheckBox║
#╚═══════════════╝

$office_CkBx = New-Object System.Windows.Forms.CheckBox
$office_CkBx.Text = "Office"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$office_CkBx.Size = $System_Drawing_Size


$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1180
$System_Drawing_Point.Y = 25
$office_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($office_CkBx)


#╔════════════════╗
#║Created CheckBox║
#╚════════════════╝

$created_CkBx = New-Object System.Windows.Forms.CheckBox
$created_CkBx.Text = "Created"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 24
$created_CkBx.Size = $System_Drawing_Size


$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1270
$System_Drawing_Point.Y = 25
$created_CkBx.Location = $System_Drawing_Point

$MainForm.Controls.Add($created_CkBx)



#╔════════════════════════════════════╗
#║Pull groups and display in list view║
#╚════════════════════════════════════╝

$searchButton.Add_Click(
        {
            
            # -----------------
            #|Search for Groups|
            # -----------------
            
            $listView.Items.Clear()

            $searchText = $textBox.Text
            
            #>get groups from AD based on search and get the properties needed
            $searchResults = Get-ADGroup -Filter {Name -like $searchText} -Properties Description, CanonicalName
            
            #>put group properties in array
            $nameArray = @($searchResults.Name)
            $categoryArray = @($searchResults.GroupCategory)
            $descriptionArray = @($searchResults.Description)
            $ouArray = @($searchResults.CanonicalName)
            
            # ---------------------------
            #|Place groups into list view|
            # ---------------------------

            #>save the length of the name array to get the number of groups returned
            $maxCount = $nameArray.Length

            $count = 0

            #>array that will hold each list item object
            $listItem = @()

            #>Create new list view object for each group adding in the information from the arrays.
            While ($count -le $maxCount) {
                $newListItemObj = New-Object System.Windows.Forms.ListViewItem($nameArray[$count])
                $listItem += $newListItemObj
                $listItem[$count].SubItems.Add([string] $categoryArray[$count])
                $listItem[$count].SubItems.Add([string] $descriptionArray[$count])
                $listItem[$count].SubItems.Add([string] $ouArray[$count])

                $count = $count + 1

            } #End while loop for building list items
            
            $listView.Items.AddRange($listItem)


        } #End Code Block for searchButton click
    ) #End searchButton click



#╔══════════════════════════════════════════════════════════════════════╗
#║Take selected group, pull users info based on checkboxes, store info  ║
#║display info in grid view                                             ║
#╚══════════════════════════════════════════════════════════════════════╝

$membersButton.Add_Click(
        {
            
            $propArray = @()

            if($logon_CkBx.Checked){$propArray += "SamAccountName"}
            if($full_CkBx.Checked){$propArray += "Name"}
            if($fn_CkBx.CheckSted){$propArray += "givenName"}
            if($ln_CkBx.CheckSted){$propArray += "surname"}
            if($email_CkBx.Checked){$propArray += "EmailAddress"}
            if($ou_CkBx.Checked){$propArray += "CanonicalName"}
            if($manager_CkBx.Checked){$propArray += "Manager"}
            if($emp_CkBx.Checked){$propArray += "EmployeeNumber"}
            if($desc_CkBx.Checked){$propArray += "Description"}
            if($dept_CkBx.Checked){$propArray += "Department"}
            if($title_CkBx.Checked){$propArray += "Title"}
            if($enabled_CkBx.Checked){$propArray += "Enabled"}
            if($office_CkBx.Checked){$propArray += "Office"}
            if($created_CkBx.Checked){$propArray += "Created"}


            $formatArray = @()

            if($logon_CkBx.Checked)  {$formatArray += @{Name = "Logon"; Expression = {$_.SamAccountName}}   }
            if($full_CkBx.Checked)   {$formatArray += @{Name = "Name"; Expression = {$_.Name}}              }
            if($fn_CkBx.CheckSted)   {$formatArray += @{Name = "First Name"; Expression = {$_.givenName}}   }
            if($ln_CkBx.CheckSted)   {$formatArray += @{Name = "Last Name"; Expression = {$_.surname}}      }
            if($email_CkBx.Checked)  {$formatArray += @{Name = "Email"; Expression = {$_.EmailAddress}}     }
            if($ou_CkBx.Checked)     {$formatArray += @{Name = "OU"; Expression = {$_.CanonicalName}}       }
            if($manager_CkBx.Checked){$formatArray += @{Name = "Manager"; Expression = {$_.Manager}}        }
            if($emp_CkBx.Checked)    {$formatArray += @{Name = "Emp#"; Expression = {$_.EmployeeNumber}}    }
            if($desc_CkBx.Checked)   {$formatArray += @{Name = "Description"; Expression = {$_.Description}}}
            if($dept_CkBx.Checked)   {$formatArray += @{Name = "Department"; Expression = {$_.Department}}  }
            if($title_CkBx.Checked)  {$formatArray += @{Name = "Title"; Expression = {$_.Title}}            }
            if($enabled_CkBx.Checked){$formatArray += @{Name = "Enabled"; Expression = {$_.Enabled}}        }
            if($office_CkBx.Checked) {$formatArray += @{Name = "Office"; Expression = {$_.Office}}          }
            if($created_CkBx.Checked){$formatArray += @{Name = "Created"; Expression = {$_.Created}}        }


            $propHash = @{Properties = $propArray}
            $formatHash = @{Property = $formatArray}

            #>Get text that is selected
            [string] $global:selectedGroup = $listView.SelectedItems.Text            

            #>Get group member info based on selected group
            $global:memberInfo = Get-ADGroup $global:selectedGroup | Get-ADGroupMember | Select -ExpandProperty SamAccountName | Get-ADUser @propHash | Select @formatHash

            #Create a list based off returned information then set gridView to use list
            $list = New-Object System.collections.ArrayList
            $list.AddRange($global:memberInfo)
            $dataGridView.DataSource = $list
            $dataGridView.Columns | Foreach-Object{$_.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells}
            

        } #End Code Block for searchButton click
    ) #End searchButton click



$exportCSVButton.Add_Click(
        {
            $fileName = ".\" + $global:selectedGroup + ".csv"
            $global:memberInfo | Export-CSV $fileName -Force -NoTypeInformation
            Invoke-Item ".\"

        }
)



$exportHTMLButton.Add_Click(
        {
            $fileName = ".\" + $global:selectedGroup + ".htm"

            $a = "<style>"
            $a = $a + "BODY{background-color:white;}"
            $a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
            $a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#FA5858}"
            $a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:white}"
            $a = $a + "</style>"
            $global:memberInfo | ConvertTo-HTML -Head $a -body "<H2>$global:selectedGroup</H2>" | Out-File $fileName -Force
            Invoke-Item ".\"

        }
)

 

[void] $MainForm.ShowDialog()
