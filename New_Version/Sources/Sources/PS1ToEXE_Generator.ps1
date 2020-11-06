#========================================================================
#
# Tool Name	: PS1 To EXE Generator
# Author 	: Damien VAN ROBAEYS
# Website	: http://www.systanddeploy.com/
# Twitter	: https://twitter.com/syst_and_deploy
#
#========================================================================

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  				| out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 				| out-null
[System.Reflection.Assembly]::LoadFrom('MahApps.Metro.dll')       				| out-null
[System.Reflection.Assembly]::LoadFrom('MahApps.Metro.IconPacks.dll')      | out-null  

function LoadXml ($global:filename)
{
    $XamlLoader=(New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}

# Load MainWindow
$XamlMainWindow=LoadXml("PS1ToEXE_Generator.xaml")
$Reader=(New-Object System.Xml.XmlNodeReader $XamlMainWindow)
$Form=[Windows.Markup.XamlReader]::Load($Reader)

[System.Windows.Forms.Application]::EnableVisualStyles()

$browse_exe = $Form.findname("browse_exe") 
$exe_sources_textbox = $Form.findname("exe_sources_textbox") 
$exe_name = $Form.findname("exe_name") 
$icon_sources_textbox = $Form.findname("icon_sources_textbox") 
$browse_icon = $Form.findname("browse_icon") 
$Build = $Form.findname("Build") 
$Choose_ps1 = $Form.findname("Choose_ps1") 
$Change_Theme = $Form.findname("Change_Theme") 
$browse_EXE_Path = $Form.findname("browse_EXE_Path") 
$exe_Final_Path_textbox = $Form.findname("exe_Final_Path_textbox") 

$object = New-Object -comObject Shell.Application  

$openfiledialog1 = New-Object 'System.Windows.Forms.OpenFileDialog'
$openfiledialog1.DefaultExt = "ico"
$openfiledialog1.Filter = "Applications (*.ico) |*.ico"
$openfiledialog1.ShowHelp = $True
$openfiledialog1.filename = "Search for ICO files"
$openfiledialog1.title = "Select an icon"

$Choose_ps1.IsEnabled = $false
$Build.IsEnabled = $false
$exe_name.IsEnabled = $false
$exe_sources_textbox.IsEnabled = $false
$icon_sources_textbox.IsEnabled = $false	

$browse_EXE_Path.IsEnabled = $false
$exe_Final_Path_textbox.IsEnabled = $false	

$browse_icon.IsEnabled = $false
$icon_sources_textbox.IsEnabled = $false													

$Global:Current_Folder =(get-location).path 

$Change_Theme.Add_Click({
	$Theme = [MahApps.Metro.ThemeManager]::DetectAppStyle($form)	
	$Script:my_theme = ($Theme.Item1).name	
	If($my_theme -eq "BaseLight")
		{		
			[MahApps.Metro.ThemeManager]::ChangeAppStyle($form, $Theme.Item2, [MahApps.Metro.ThemeManager]::GetAppTheme("BaseDark"));			
		}
	ElseIf($my_theme -eq "BaseDark")
		{							
			[MahApps.Metro.ThemeManager]::ChangeAppStyle($form, $Theme.Item2, [MahApps.Metro.ThemeManager]::GetAppTheme("BaseLight"));			
		}		
})		
		
$browse_exe.Add_Click({		
	$folder = $object.BrowseForFolder(0, $message, 0, 0) 
	If ($folder -ne $null) 
		{ 		
			$Script:EXE_folder = $folder.self.Path 
			$exe_sources_textbox.Text = $EXE_folder	
			
			$Check_Sources_Folder_Content = Get-ChildItem $EXE_folder -Recurse |? { $_.PSIsContainer } | Measure-Object | select -Expand Count
			If($Check_Sources_Folder_Content -gt 0)
				{
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Form, "Warning", "Your project folder contains subfolders.`nThe tool won't include subfolders.")										
				}			

			$Script:Folder_name = Split-Path -leaf -path $EXE_folder			
										
			$browse_EXE_Path.IsEnabled = $true
			$exe_Final_Path_textbox.IsEnabled = $true

			$Dir_EXE_Folder = get-childitem $EXE_folder -recurse
			$List_All_PS1 = $Dir_EXE_Folder | where {$_.extension -eq ".ps1"}				
			foreach ($ps1 in $List_All_PS1)
				{
					$Choose_ps1.Items.Add($ps1)	
					$Script:EXE_PS1_To_Run = $Choose_ps1.SelectedItem	
					$Script:PS1_Full_Path = $PS1.FullName					
				}	
			$Global:PS1_Path = "$EXE_folder\$EXE_PS1_To_Run"
				
			$Choose_ps1.add_SelectionChanged({
				$Script:EXE_PS1_To_Run = $Choose_ps1.SelectedItem
				$Global:PS1_Path = "$EXE_folder\$EXE_PS1_To_Run"
			})	
		}
})	


$browse_EXE_Path.Add_Click({
	Add-Type -AssemblyName System.Windows.Forms
	$Folder_Object = New-Object System.Windows.Forms.FolderBrowserDialog
	[void]$Folder_Object.ShowDialog()	
	$Script:EXE_Export_Folder = $Folder_Object.SelectedPath	
	$exe_Final_Path_textbox.Text = $EXE_Export_Folder
	
	If($EXE_Export_Folder -ne $null)
		{
			$Choose_ps1.IsEnabled = $true
			$Build.IsEnabled = $true
			$exe_name.IsEnabled = $true	
			$browse_icon.IsEnabled = $true
			$icon_sources_textbox.IsEnabled = $true			
		}

})


$browse_icon.Add_Click({	
	If($openfiledialog1.ShowDialog() -eq 'OK')
		{	
			$icon_sources_textbox.Text = $openfiledialog1.FileName
			$Global:EXE_Icon_To_Set = $openfiledialog1.FileName 
			copy-item $EXE_Icon_To_Set $EXE_folder -force
		}	
})	

$Build.Add_Click({		
	$EXE_File_Name = $exe_name.Text.ToString()	
	If ($exe_name.Text -eq "") 
		{
			[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Form, "Oops :-)", "Type an EXE name")									
		}
	Else
		{			
			Try
				{
					& .\make-exe.exe -file $PS1_Path -embed -silent
					sleep 3			
					$PS1_To_Run = (Get-Item $PS1_Path).Name
					$PS1_To_Rename = $PS1_To_Run.replace(".ps1","")
					$Script:EXE_Full_Path = "$EXE_folder\$PS1_To_Rename.exe"
					
					Rename-item $EXE_Full_Path "$EXE_File_Name.exe" -force						
					Move-item "$EXE_folder\$EXE_File_Name.exe" $EXE_Export_Folder
					
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Form, "Success :-)", "Your EXE has been created")		
					
					GCI $env:temp | where {$_.name -like "*Make-EXE*"} | remove-item -recurse -Force
				}
			Catch
				{
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Form, "Oops :-)", "Your EXE has not been created")										
				}		
		}
})



$Form.ShowDialog() | Out-Null