# The RuneScape username to look up
$username = "Your Username"

# Runemetrics API - https://runescape.wiki/w/Application_programming_interface#Runemetrics
# Does not require an API key as it is a public, unofficial endpoint
$url = "https://apps.runescape.com/runemetrics/profile/profile?user=$username&activities=20"

# Sends an HTTP request to $url and parses the returned JSON into a PowerShell object
function Send-Runemetrics {
    param([string]$url)
    
    try {
        $response = Invoke-RestMethod $url -ErrorAction Stop
        return $response
    } catch {
        # Prepare to show error in terminal
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value_
            $statusDescription = $_.Exception.Response.$statusDescription

            Write-Host "[ERROR] Could not retreive info from API endpoint."
            Write-Host "Status Code: $statusCode"
            Write-Host "Status Description: $statusDescription"

            # Read API error body (if present)
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $responseBody = $reader.ReadToEnd()
            Write-Host "Error Body: $responseBody"
        } else {
            Write-Host "[ERROR] Network error: $($_.Exception.Message)"
        }
    }
}


# Array of skill names (0-28)
$skills = @(
    "Attack","Defence","Strength","Constitution","Ranged",
    "Prayer","Magic","Cooking","Woodcutting","Fletching",
    "Fishing","Firemaking","Crafting","Smithing","Mining",
    "Herblore","Agility","Thieving","Slayer","Farming",
    "Runecrafting","Hunter","Construction","Summoning","Dungeoneering"
    "Divination","Invention","Archaeology","Necromancy"
)

# Load Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Runemetrics GUI"
$form.Size = New-Object System.Drawing.Size(500,410)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
# Dark gray RGB: 45,45,45
$form.BackColor = [System.Drawing.Color]::FromArgb(45,45,45)

# Create a table to define layout
$table = New-Object System.Windows.Forms.TableLayoutPanel
$table.Dock = "Fill"
$table.RowCount = 3
$table.ColumnCount = 2
$table.Padding = New-Object System.Windows.Forms.Padding(10)
$table.AutoSize = $true

# Set equal column widths
$table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
$table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))

# Create a big label to show your name
$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = "Runemetrics"
$nameLabel.Font = New-Object System.Drawing.Font("Segoe UI",24,[System.Drawing.FontStyle]::Bold)
# Make the text white
$nameLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)
$nameLabel.AutoSize = $true

# Create a TextBox to enter a username
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Size = New-Object System.Drawing.Size(200,40)

# Create a button to lookup the username
$usernameButton = New-Object System.Windows.Forms.Button
$usernameButton.Text = "Fetch info"
$usernameButton.Size = New-Object System.Drawing.Size(100,30)
# Make back color brighter and text white
$usernameButton.BackColor = [System.Drawing.Color]::FromArgb(95,95,95)
$usernameButton.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Button click event to send an HTTP request to the API with the username
$usernameButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($textBox.Text)) {
        # Show popup if username is empty
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a username!",
            "Input Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    } else {
        $username = $textBox.Text
        $usernameLabel.Text = "Username: $username"

        # Change text to "..." to show processing
        $overallRankLabel.Text = "Overall Ranking: ..."
        $userStatusLabel.Text = "Status: ..."
        $totalSkillLabel.Text = "Total Skill: ..."
        $totalExperienceLabel.Text = "Total XP: ..."
        $combatLevelLabel.Text = "Combat Level: ..."
        $magicExperienceLabel.Text = "Magic XP: ..."
        $meleeExperienceLabel.Text = "Melee XP: ..."
        $rangedExperienceLabel.Text = "Ranged XP: ..."
        $questsStartedLabel.Text = "Quests Started: ..."
        $questsCompletedLabel.Text = "Quests Completed: ..."
        $questsNotStartedLabel.Text = "Quests Not Started: ..."

        # Send the HTTP request and parse it
        $url = "https://apps.runescape.com/runemetrics/profile/profile?user=$username&activities=20"
        $response = Send-Runemetrics($url)

        # Display the results to the user
        $overallRankLabel.Text = "Overall Ranking: $($response.rank)"

        # Rank is formatted but other numbers are not
        $totalSkillFormatted = "{0:N0}" -f [int]$response.totalskill
        $totalSkillLabel.Text = "Total Skill: $totalSkillFormatted"

        $totalExperienceFormatted = "{0:N0}" -f [int]$response.totalxp
        $totalExperienceLabel.Text = "Total XP: $totalExperienceFormatted"

        $combatLevelLabel.Text = "Combat Level: $($response.combatlevel)"

        $magicExperienceFormatted = "{0:N0}" -f [int]$response.magic
        $magicExperienceLabel.Text = "Magic XP: $magicExperienceFormatted"

        $meleeExperienceFormatted = "{0:N0}" -f [int]$response.melee
        $meleeExperienceLabel.Text = "Melee XP: $meleeExperienceFormatted"

        $rangedExperienceFormatted = "{0:N0}" -f [int]$response.ranged
        $rangedExperienceLabel.Text = "Ranged XP: $rangedExperienceFormatted"

        $questsStartedLabel.Text = "Quests Started: $($response.questsstarted)"
        $questsCompletedLabel.Text = "Quests Completed: $($response.questscomplete)"
        $questsNotStartedLabel.Text = "Quests Not Started: $($response.questsnotstarted)"

        # Display if the user is online or offline should return string "true" or "false"
        if ($response.loggedIn -eq "true") { $userStatusLabel.Text = "Status: Online" }
        elseif ($response.loggedIn -eq "false") { $userStatusLabel.Text = "Status: Offline" }
        else { $userStatusLabel.Text = "Status: Error retrieving status." }

        # Add items to list view (skillvalues is an array of objects)
        for ($i = 0; $i -lt $response.skillvalues.Count; $i++) {
            # Add the skill names column
            $item = New-Object System.Windows.Forms.ListViewItem($skills[$response.skillvalues[$i].id])

            # Add the skill levels column
            $item.SubItems.Add($response.skillvalues[$i].level)

            # Add the rank column
            $skillRankFormatted = "{0:N0}" -f [int]$response.skillvalues[$i].rank
            $item.SubItems.Add($skillRankFormatted)

            # Displays highest skill level and descends
            $listView.Items.Add($item)
        }
    }
})

# Create a button to export the contents of the ListView to a CSV
$downloadCSVButton = New-Object System.Windows.Forms.Button
$downloadCSVButton.Text = "Download CSV"
$downloadCSVButton.Size = New-Object System.Drawing.Size(100,30)
# Make back color brighter and text white
$downloadCSVButton.BackColor = [System.Drawing.Color]::FromArgb(95,95,95)
$downloadCSVButton.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Button click event for the download CSV button
$downloadCSVButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV file (*.csv)|*.csv"

    if ($saveFileDialog.ShowDialog() -eq "OK") {
        Export-ListViewToCsv -ListView $listView -Path $saveFileDialog.FileName
    }
})

# Function that allows for exporting a ListView to a CSV
function Export-ListViewToCsv {
    # Defines the parameters: the listview and the filepath where the csv will be saved
    param (
        [System.Windows.Forms.ListView]$ListView,
        [string]$Path
    )

    # Loop through every row in the listview
    $rows = foreach ($item in $ListView.Items) {
        $obj = [ordered]@{}

        # Loop through each column
        for ($i = 0; $i -lt $ListView.Columns.Count; $i++) {
            # Define the columns: level and skill name
            $columnName = $ListView.Columns[$i].Text
            
            # First column is text, anything else is subitem
            if ($i -eq 0) {
                $value = $item.Text
            }
            else {
                $value = $item.SubItems[$i].Text
            }

            # Add column and value to row object
            $obj[$columnName] = $value
        }

        # Convert it into a powershell object
        [pscustomobject]$obj
    }

    # Finally we can export to CSV
    $rows | Export-Csv $Path -NoTypeInformation
}

# Label to show information the username entered into textbox
$usernameLabel = New-Object System.Windows.Forms.Label
$usernameLabel.Text = $username
$usernameLabel.AutoSize = $true
$usernameLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show user's overall rank (string)
$overallRankLabel = New-Object System.Windows.Forms.Label
$overallRankLabel.Text = "Overall Ranking:"
$overallRankLabel.AutoSize = $true
$overallRankLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show if the player is logged in or offline
$userStatusLabel = New-Object System.Windows.Forms.Label
$userStatusLabel.Text = "Status:"
$userStatusLabel.AutoSize = $true
$userStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show the player's total skill level
$totalSkillLabel = New-Object System.Windows.Forms.Label
$totalSkillLabel.Text = "Total Skill:"
$totalSkillLabel.AutoSize = $true
$totalSkillLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show the player's total experience
$totalExperienceLabel = New-Object System.Windows.Forms.Label
$totalExperienceLabel.Text = "Total XP:"
$totalExperienceLabel.AutoSize = $true
$totalExperienceLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show the player's combat level
$combatLevelLabel = New-Object System.Windows.Forms.Label
$combatLevelLabel.Text = "Combat Level:"
$combatLevelLabel.AutoSize = $true
$combatLevelLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show the player's magic experience
$magicExperienceLabel = New-Object System.Windows.Forms.Label
$magicExperienceLabel.Text = "Magic XP:"
$magicExperienceLabel.AutoSize = $true
$magicExperienceLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show the player's melee experience
$meleeExperienceLabel = New-Object System.Windows.Forms.Label
$meleeExperienceLabel.Text = "Melee XP:"
$meleeExperienceLabel.AutoSize = $true
$meleeExperienceLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show the player's ranged experience
$rangedExperienceLabel = New-Object System.Windows.Forms.Label
$rangedExperienceLabel.Text = "Ranged XP:"
$rangedExperienceLabel.AutoSize = $true
$rangedExperienceLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show the amount of quests the player has started
$questsStartedLabel = New-Object System.Windows.Forms.Label
$questsStartedLabel.Text = "Quests Started:"
$questsStartedLabel.AutoSize = $true
$questsStartedLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show the amount of quests the player has completed
$questsCompletedLabel = New-Object System.Windows.Forms.Label
$questsCompletedLabel.Text = "Quests Completed:"
$questsCompletedLabel.AutoSize = $true
$questsCompletedLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Label to show the amount of quests the player has not started
$questsNotStartedLabel = New-Object System.Windows.Forms.Label
$questsNotStartedLabel.Text = "Quests Not Started:"
$questsNotStartedLabel.AutoSize = $true
$questsNotStartedLabel.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)

# Create a ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.MultiSelect = $false
$listView.Location = New-Object System.Drawing.Point(230,10)
$listView.Size = New-Object System.Drawing.Size(245,350)

# Add Columns (column name, size)
$listView.Columns.Add("Skill", 100)
$listView.Columns.Add("Level", 50)
$listView.Columns.Add("Rank", 70)

# Add controls and table
$table.Controls.Add($nameLabel,0,0)
$table.Controls.Add($textBox,0,1)
$table.Controls.Add($usernameButton,0,2)
$table.Controls.Add($usernameLabel,0,3)
$table.Controls.Add($overallRankLabel,0,4)
$table.Controls.Add($userStatusLabel,0,5)
$table.Controls.Add($totalSkillLabel,0,6)
$table.Controls.Add($totalExperienceLabel,0,7)
$table.Controls.Add($combatLevelLabel,0,8)
$table.Controls.Add($magicExperienceLabel,0,9)
$table.Controls.Add($meleeExperienceLabel,0,10)
$table.Controls.Add($rangedExperienceLabel,0,11)
$table.Controls.Add($questsStartedLabel,0,12)
$table.Controls.Add($questsCompletedLabel,0,13)
$table.Controls.Add($questsNotStartedLabel,0,14)
$table.Controls.Add($downloadCSVButton,0,15)
$form.Controls.Add($listView)
$form.Controls.Add($table)

# Show the form
$form.Topmost = $true
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()