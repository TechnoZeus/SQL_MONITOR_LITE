# Init PowerShell Gui
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$b2 = 0
#$global:Sucessfullsession = 0
#$global:outloop = 1
$storecred = get-credential -credential username
$global:username = $storecred.username
$global:password = $storecred.getnetworkcredential().Password
Function CreateGlobalConnection
{ Add-Type -Path 'C:\Program Files (x86)\Microsoft SQL Server\130\SDK\Assemblies\Microsoft.SqlServer.Smo.dll'
  Add-Type -Path 'C:\Program Files (x86)\Microsoft SQL Server\130\SDK\Assemblies\Microsoft.SqlServer.ConnectionInfo.dll'

  Try {
  #$srvname="10.80.44.42,1433"
 $srvname=$objTextBox.text
 #$dbname="master"
 $mySrvConn=new-object Microsoft.SqlServer.Management.Common.ServerConnection
 $mySrvConn.ServerInstance=$srvname
 $mySrvConn.LoginSecure=$false
 #$mySrvConn.Login ="sa"
 #$mySrvConn.Password ="admin@123"
 $mySrvConn.Login = $global:username
 $mySrvConn.Password = $global:password
 $Server=New-Object Microsoft.SqlServer.Management.Smo.Server($mySrvConn)
 $MasterString =$Server.Databases.Item("master"); # Nextproject
 if ($b2 -eq 1)
 {
  $global:Sucessfullsession = $srvname
  $global:outloop = 1
  Write-output "Server connection LOADED " $global:Sucessfullsession | Out-GridView -Title "Success"
  $b2 = 0 ;
 }

 
 return $MasterString
  } 
 Catch
 {  Write-output "Error has occured in server connection: $Error " | Out-GridView -Title "Error Window" }
 }
## V.1.3 Logic
function LivesessionCount

{   
    Param($functionToCall)
         
    
    $script=(Get-Item "function:$functionToCall").ScriptBlock
      [ScriptBlock]::Create($script).Invoke();
    
$C1Form = New-Object System.Windows.Forms.Form
$C1Form.ClientSize         = '500,500'
$C1Form.text               = "Real Time OP"
$C1Form.BackColor          = "#FFFFFF"
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000
$timer.Add_Tick({UpdateProcCount $textBoxDisplay})
$timer.Enabled = $True
$C1Form.Controls.Add($DataGridView1)
  $C1Form.ShowDialog()
  if ($C1Form.DialogResult -eq 'Cancel')  ## sense op of the cancel button was pressed break SP exection in success connection state
    {$checkBox1.CheckState = 0 # reset checkbox else chekbox option 
      $timer.Dispose()
      return; 
     }

    }
   

    
     function UpdateProcCount($textBoxDisplay)
    { 
       Try {
        $db = CreateGlobalConnection
      [String]$sql ="qm";
      $result=$db.ExecuteWithResults($sql);
      #$table=$result.Tables[0];
      #$DataGridView1.DataSource = [System.Collections.ArrayList]([System.Object[]](,$table | Sort-Object))
      $DataGridView1.DataSource = $result.Tables[0];
      $DataGridView1.AutoGenerateColumns = 'true';
        }
       Catch {             
                Write-Output 'ERROR' | Out-GridView                
            }
    }

  
  ##### Get drive space ######

  Function Get_driveusage{
    Param(
    [string[]]$computername 
    )
    
    [string]$g=[char]9608
    $data = Get-WmiObject -class win32_logicaldisk -ComputerName $computername -filter 'drivetype=3' | 
    Select-Object @{Name="Computername";Expression={$_.Systemname}},
    @{Name="Drive";Expression={$_.DeviceID}},
    @{Name="SizeMB";Expression={[int]($_.Size/1MB)}},
    @{Name="FreeMB";Expression={[int]($_.Freespace/1MB)}},
    @{Name="UsedMB";Expression={[math]::round(($_.size - $_.Freespace)/1MB,2)}},
    @{Name="Free%";Expression={[math]::round(($_.Freespace/$_.Size)*100,2)}},
    @{Name="Used Graph"; Expression={ [int]$per=(($_.size - $_.Freespace)/$_.Size)*100
     "$($g)" * $per ; }
     } 
     #send the results to Out-Gridview
     $data | out-gridview -Title "Drive Report"
    }


    # check cup logic for button 8

    ##Function Get_Cpuusage{
    #    Param(
    #    [string[]]$computername
    #    )
#
    #    $Processor = (Get-WmiObject -ComputerName $computername -Class win32_processor -ErrorAction Stop | Measure-Object -Property LoadPercentage -Average | Select-Object Average).Average
    #     Write-Output $Processor | Out-GridView -Title "Drive Report on server" + $computername
#
    #}


    ##########Buttons and contorls logic ###########################
        
#Button 1 Action
$button1_RunOnClick=
{        
$db = CreateGlobalConnection  
[String]$sql ="SP_configure";
$result=$db.ExecuteWithResults($sql);
$table=$result.Tables[0];
$table | Out-GridView -Title ( "Results from server " + $objTextBox.text) 
}
#Button 2 Action
$button2_RunOnClick=
{    $b2 = 1;
     CreateGlobalConnection  
}
#button 3 action 

$button3_RunOnClick=
{  
$db = CreateGlobalConnection  
[String]$sql ="EXEC sp_readerrorlog";
$result=$db.ExecuteWithResults($sql);
$table=$result.Tables[0];
$table | Out-GridView  -Title ( "Results from server " + $objTextBox.text) 
}

#button 4 action
$button4_RunOnClick=
{  
Try {
$srvname=$objTextBox.text
Get-service -ComputerName $srvname  | Where-Object {($_.name -like "MSSQL$*" -or $_.name -like "MSSQLSERVER" -or $_.name -like "SQL Server (*" -or $_.name -like "SQL*")} | Out-GridView  -Title ( "Results from server " + $srvname)
}
Catch { Write-output "Couldn't connect" | Out-GridView }
}

#button 5 action 

$button5_RunOnClick=
{ if ($checkBox1.checked -And $global:Sucessfullsession -eq $objTextBox.text )
  {LivesessionCount -functionToCall UpdateProcCount}
  else {Write-Output "Makesure you have tested connection before running live session" | Out-GridView -Title "Testconnection required"} #Call the counter if check box is ticked
 }

 #space display logic
 # AD user can only check space which have adequate permissions
 $button6_RunOnClick=
{ 
  $srvname=$objTextBox.text;
  Get_driveusage $srvname ;
}

##refresh session button 7 control
$button7_RunOnClick=
{ 
  
$storecred = get-credential -credential username
$global:username = $storecred.username
$global:password = $storecred.getnetworkcredential().Password
    
}

# button 8 contorol

$button8_RunOnClick=
{ 
  $srvname=$objTextBox.text;
  Get_Cpuusage $srvname ;
}


# button 9 contorol

$button9_RunOnClick=
{ 
$db = CreateGlobalConnection  
[String]$sql ="SELECT top 5
[DatabaseName] = CASE [database_id] WHEN 32767
THEN 'Resource DB'
ELSE DB_NAME([database_id]) END,
COUNT_BIG(*) [Pages in Buffer],
COUNT_BIG(*)/128 [Buffer Size in MB]
FROM sys.dm_os_buffer_descriptors
GROUP BY [database_id]
ORDER BY [Pages in Buffer] DESC;";
$result=$db.ExecuteWithResults($sql);
$table=$result.Tables[0];
$table | Out-GridView  -Title ( "Results from server " + $objTextBox.text) 
}


############################# form display below ##############################################
# Create a new form
$LocalPrinterForm= New-Object system.Windows.Forms.Form

# Define the size, title and background color
$LocalPrinterForm.ClientSize         = '1500,300'
$LocalPrinterForm.text               = "SQL DASHBOARD"
$LocalPrinterForm.BackColor          = "#fffffF"
# Create Input 
#Add Label and TextBox
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,25)
$objLabel.Size = New-Object System.Drawing.Size(110,25)
$objLabel.Text = “Enter Server Name”
$LocalPrinterForm.Controls.Add($objLabel)
$objTextBox = New-Object System.Windows.Forms.TextBox
$objTextBox.Location = New-Object System.Drawing.Size(120,25)
$objTextBox.Size = New-Object System.Drawing.Size(150,20)
$LocalPrinterForm.Controls.Add($objTextBox)
#Button 1 
$button1 = New-Object System.Windows.Forms.Button #add button 
$button1.TabIndex = 4
$button1.Name = “button1”
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 23
$button1.Size = $System_Drawing_Size
$button1.UseVisualStyleBackColor = $True
$button1.Text = “Run”
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 300
$System_Drawing_Point.Y = 25
$button1.Location = $System_Drawing_Point
$button1.DataBindings.DefaultDataSourceUpdateMode = 0
$button1.add_Click($button1_RunOnClick)
$LocalPrinterForm.Controls.Add($button1)


#Button 2 
$button2 = New-Object System.Windows.Forms.Button #add button 
$button2.TabIndex = 4
$button2.Name = “button2”
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 100
$System_Drawing_Size.Height = 23
$button2.Size = $System_Drawing_Size
$button2.UseVisualStyleBackColor = $True
$button2.Text = “Test Connection”
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 380
$System_Drawing_Point.Y = 25
$button2.Location = $System_Drawing_Point
$button2.DataBindings.DefaultDataSourceUpdateMode = 0
$button2.add_Click($button2_RunOnClick)
$LocalPrinterForm.Controls.Add($button2)

### PH2 code below ##

#created button 3 for server errorlog check

$button3 = New-Object System.Windows.Forms.Button #add button 
$button3.TabIndex = 4
$button3.Name = “button3”
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 23
$button3.Size = $System_Drawing_Size
$button3.UseVisualStyleBackColor = $True
$button3.Text = “Error log”
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 50
$button3.Location = $System_Drawing_Point
$button3.DataBindings.DefaultDataSourceUpdateMode = 0
$button3.add_Click($button3_RunOnClick)
$LocalPrinterForm.Controls.Add($button3)


### PH3 code below ##
#created button 4  for server services check

$button4 = New-Object System.Windows.Forms.Button #add button 
$button4.TabIndex = 4
$button4.Name = “button4”
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 23
$button4.Size = $System_Drawing_Size
$button4.UseVisualStyleBackColor = $True
$button4.Text = “SQL Service”
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 75
$button4.Location = $System_Drawing_Point
$button4.DataBindings.DefaultDataSourceUpdateMode = 0
$button4.add_Click($button4_RunOnClick)
$LocalPrinterForm.Controls.Add($button4)
 
##V1.3 check box

$checkBox1 = New-Object System.Windows.Forms.CheckBox  
$checkBox1.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 200
$System_Drawing_Size.Height = 25
$checkBox1.Size = $System_Drawing_Size
$checkBox1.TabIndex = 0
$checkBox1.Text = “Live monitor sessions”
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 265
$checkBox1.Location = $System_Drawing_Point
$checkBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox1.Name = “checkBox1”
$LocalPrinterForm.Controls.Add($checkBox1)
$DataGridView1 = New-Object system.Windows.Forms.DataGridView  ## DISPLAY LIVE OP
$DataGridView1.width = 500
$DataGridView1.height = 500
$DataGridView1.AutoSize = $True
$DataGridView1.location = New-Object System.Drawing.Point(10,10)

##V1.3 check box CONTORL BUTTON 5
$button5 = New-Object System.Windows.Forms.Button #add button 
$button5.TabIndex = 4
$button5.Name = “button5”
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 23
$button5.Size = $System_Drawing_Size
$button5.UseVisualStyleBackColor = $True
$button5.Text = “Run Live View”
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 25
$System_Drawing_Point.Y = 240
$button5.Location = $System_Drawing_Point
$button5.DataBindings.DefaultDataSourceUpdateMode = 0
$button5.add_Click($button5_RunOnClick)
$LocalPrinterForm.Controls.Add($button5)

#get drive details v1.7 # gui contorls
$button6 = New-Object System.Windows.Forms.Button #add button 
$button6.TabIndex = 4
$button6.Name = “button6”
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 23
$button6.Size = $System_Drawing_Size
$button6.UseVisualStyleBackColor = $True
$button6.Text = “Drive info”
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 100
$button6.Location = $System_Drawing_Point
$button6.DataBindings.DefaultDataSourceUpdateMode = 0
$button6.add_Click($button6_RunOnClick)
$LocalPrinterForm.Controls.Add($button6)
#get drive details v1.8 # Self reload contorls
$button7 = New-Object System.Windows.Forms.Button #add button 
$button7.TabIndex = 4
$button7.Name = “button7”
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 23
$button7.Size = $System_Drawing_Size
$button7.UseVisualStyleBackColor = $True
$button7.Text = “Check other server”
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 380
$System_Drawing_Point.Y = 270
$button7.Location = $System_Drawing_Point
$button7.DataBindings.DefaultDataSourceUpdateMode = 0
$button7.add_Click($button7_RunOnClick)
$LocalPrinterForm.Controls.Add($button7)

#get drive details v1.8 # Get server CPU 
$button8 = New-Object System.Windows.Forms.Button #add button 
$button8.TabIndex = 4
$button8.Name = “button8”
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 23
$button8.Size = $System_Drawing_Size
$button8.UseVisualStyleBackColor = $True
$button8.Text = “Check CPU”
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 150
$System_Drawing_Point.Y = 100
$button8.Location = $System_Drawing_Point
$button8.DataBindings.DefaultDataSourceUpdateMode = 0
$button8.add_Click($button8_RunOnClick)
$LocalPrinterForm.Controls.Add($button8)

#get drive details v1.8 # Get SQL server Memory DB wise
$button9 = New-Object System.Windows.Forms.Button #add button 
$button9.TabIndex = 4
$button9.Name = “button9”
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 90
$System_Drawing_Size.Height = 23
$button9.Size = $System_Drawing_Size
$button9.UseVisualStyleBackColor = $True
$button9.Text = “Check Memory”
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 300
$System_Drawing_Point.Y = 100
$button9.Location = $System_Drawing_Point
$button9.DataBindings.DefaultDataSourceUpdateMode = 0
$button9.add_Click($button9_RunOnClick)
$LocalPrinterForm.Controls.Add($button9)
# Display the form
$LocalPrinterForm.ShowDialog()