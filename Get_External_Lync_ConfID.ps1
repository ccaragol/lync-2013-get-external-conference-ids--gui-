<##########################################################################  
   
Script Name:		Get External Conference IDs GUI 
   
Contact Info: 
Name	  		C. Anthony Caragol 
LinkedIn  		https://www.linkedin.com/profile/view?id=18013352
  
Description:		By request of a user, show all external conference IDs.

Notes:			Yes, I know it doesn't show the numeric conference ID that
			would be punched in via DTMF.  I can't seem to find this in
			the database.  If you can find it, I'll gladly add it.
		
			Please excuse the sloppy coding for now, I don't use a
			development environment, IDE or ISE.  I use notepad, not
			even Notepad++, just notepad.

Version:
1.001	- 	Minor bugfix, issue with starting in folders with spaces.
 
 
##########################################################################>  
 
function MainFunction { 


	Add-Type -AssemblyName System.Windows.Forms 
	[System.Windows.Forms.Application]::EnableVisualStyles() 
 
 
	$ObjForm = New-Object System.Windows.Forms.Form -Property @{ 
		Width  = 600 
		Height = 600 
		Text = ”Get External Lync Conference IDs” 
		ShowIcon = $false 
		StartPosition = "CenterScreen" 
		BackColor = [System.Drawing.Color]::FromArgb(255,204,212,247)

		MinimumSize = @{ 
			Width  = 530
			Height = 400 
        } 
    } 
 
  
    $CAC_FormClosing = { 
    } 
 
    $CAC_FormSizeChanged = { 
		$dataGridView.Columns[0].Width = ($ObjForm.Width - 110) /2
		$dataGridView.Columns[1].Width = ($ObjForm.Width - 110) /2
    } 
 
	$ObjForm.Add_SizeChanged($CAC_FormSizeChanged) 
	$ObjForm.Add_Closing($CAC_FormClosing) 
	
	$RefreshButton = New-Object System.Windows.Forms.Button
	$RefreshButton.Location = New-Object System.Drawing.Size(425,525)
	$RefreshButton.Size = New-Object System.Drawing.Size(75,25)
	$RefreshButton.Text = "Refresh"
	$RefreshButton.Add_Click({
		$DataGridView.Rows.Clear()
		RefreshData
	})
	$RefreshButton.Anchor = 'Bottom, Right'
	$objForm.Controls.Add($RefreshButton)
         
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(500,525)
	$CancelButton.Size = New-Object System.Drawing.Size(75,25)
	$CancelButton.Text = "Quit"
	$CancelButton.Add_Click({$objForm.Close()})
	$CancelButton.Anchor = 'Bottom, Right'
	$objForm.Controls.Add($CancelButton)

	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,20) 
	$objLabel.Size = New-Object System.Drawing.Size(230,20) 
	$objLabel.Text = "Please enter the FQDN of a front end server:"
	$objForm.Controls.Add($objLabel) 

	$objTextBox = New-Object System.Windows.Forms.TextBox 
	$objTextBox.Location = New-Object System.Drawing.Size(240,20) 
	$objTextBox.Size = New-Object System.Drawing.Size(260,20) 
	$objTextBox.Add_KeyDown({if ($_.KeyCode -eq "Enter") {RefreshData}})
	$objForm.Controls.Add($objTextBox) 

	$dataGridView = New-Object System.Windows.Forms.DataGridView
	$dataGridView.Location = New-Object System.Drawing.Size(10,60) 
	$dataGridView.Size = New-Object System.Drawing.Size(550,400) 
	$dataGridView.Anchor = 'Top, Bottom, Left, Right'


	$dataGridView.ColumnCount = 2
	$dataGridView.Columns[0].Width = 245
	$dataGridView.Columns[1].Width = 245
	$dataGridView.Columns[0].Name = "User"
	$dataGridView.Columns[1].Name = "External Conference ID"

	$objForm.Controls.Add($dataGridView) 
 
    [void] $ObjForm.ShowDialog() 

} 
 
Function RefreshData
{
	$servername = $objTextBox.text + "\rtclocal"
	$database = "rtc"
	Write-host $servername
	$connectionString = "Server=$servername;Database=$database;Integrated Security=SSPI;"
 
	$query = "SELECT r.ResourceId,c.OrganizerId,r.UserAtHost,c.Static,c.ExternalConfId FROM dbo.Resource r INNER JOIN dbo.Conference c ON r.ResourceId=c.OrganizerId WHERE c.Static=1 ORDER BY r.UserAtHost"
 
	$connection = New-Object System.Data.SqlClient.SqlConnection
	$connection.ConnectionString = $connectionString
	$connection.Open()
	$command = $connection.CreateCommand()
	$command.CommandText  = $query
 
	$result = $command.ExecuteReader()
 
	$table = new-object “System.Data.DataTable”
	$table.Load($result)
 
	foreach ($Row in $table) {
		$tableentry = @( $Row.UserAtHost, $Row.ExternalConfId) 
		$dataGridView.Rows.Add($tableentry)
	}
 
	$connection.Close()
}

<#####################################################################  

Force the script to run as Administrator to allow access to local RTC 
instances.   

Ripped from Ben Armstrong's blog:
http://blogs.msdn.com/b/virtual_pc_guy/archive/2010/09/23/a-self-elevating-powershell-script.aspx

######################################################################>  

# Get the ID and security principal of the current user account
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
  
# Get the security principal for the Administrator role
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
  
# Check to see if we are currently running "as Administrator"
if ($myWindowsPrincipal.IsInRole($adminRole))
	{
	# We are running "as Administrator" - so change the title and background color to indicate this
	$Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
	$Host.UI.RawUI.BackgroundColor = "DarkBlue"
	clear-host
	}
else
	{
	# We are not running "as Administrator" - so relaunch as administrator
    
	# Create a new process object that starts PowerShell
	$newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
    
	# Specify the current script path and name as a parameter
	$newProcess.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "' -noexit"
    
	# Indicate that the process should be elevated
	$newProcess.Verb = "runas";
    
	# Start the new process
	[System.Diagnostics.Process]::Start($newProcess);
    
	# Exit from the current, unelevated, process
	exit
	}

<#####################################################################  

Start Main Function  
  
######################################################################>  

MainFunction