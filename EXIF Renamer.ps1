#---------------------------------------------------------------------------------------------------------------------------
# EXIF Renamer
#  © 2025 Remus Rigo
# v1.3.20250726                                                              [System.Windows.Forms.MessageBox]::Show("Debug")

Clear-Host
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$appTitle = "EXIF Renamer v1.3 by Remus Rigo"
$ExifToolPath = "D:\Tools\exiftool\exiftool.exe"

$supportedExtensions = @(".3fr", ".arw", ".cr2", ".crw", ".dng", ".erf", ".kdk", ".orf", ".mef", ".mos", ".mrw", ".nef", ".orf",
   ".pef", ".raf", ".raw", ".rw2", ".srw", "x3f")

#---------------------------------------------------------------------------------------------------------------------------
# check EXIFtool path

if (-not (Test-Path $ExifToolPath))
{
   $ExifToolPathTXT = Join-Path -Path $PSScriptRoot -ChildPath "exiftool.txt"
   if (Test-Path $ExifToolPathTXT)
   {
      if (((Get-Item $ExifToolPathTXT).Length) -gt 0)
      {
         $tempExifToolPath = (Get-Content $ExifToolPathTXT)
         if (Test-Path $tempExifToolPath)
         {
            
            $ExifToolPath = $tempExifToolPath
         }
         else
         {
            $ExifToolPath = $ull
            [System.Windows.Forms.MessageBox]::Show("Path not found")
         }
      }
      else
      {
         $ExifToolPath = $ull
         [System.Windows.Forms.MessageBox]::Show("Empty file, no path to EXIFtool.exe")
      }
   }
   else
   {
      $ExifToolPath = $ull
   }

   if ($ExifToolPath -eq $ull)
   {
      $dlgFolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
      $dlgFolderBrowser.Description = "Select EXIFTool folder:"
      $dlgFolderBrowser.ShowNewFolderButton = $false

      if ($dlgFolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
      {
         $tempExifToolPath = Join-Path -Path $dlgFolderBrowser.SelectedPath -ChildPath "exiftool.exe"
         if (Test-Path $tempExifToolPath)
         {
            $ExifToolPath = $tempExifToolPath
            Set-Content -Path $ExifToolPathTXT -Value $ExifToolPath
         }
         else
         {
            [System.Windows.Forms.MessageBox]::Show("EXIFtool.exe not found")
         }
      }
      else
      {
         [System.Windows.Forms.MessageBox]::Show("No folder selected")
      }
   }
}

#---------------------------------------------------------------------------------------------------------------------------
# Form: Main

$frmMain = New-Object System.Windows.Forms.Form
$frmMain.AutoScroll = $true
$frmMain.FormBorderStyle = "FixedSingle"
$frmMain.MaximizeBox = $false
$frmMain.MinimizeBox = $true
$frmMain.Size = New-Object System.Drawing.Size(([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width*0.75), ([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height*0.75))
$frmMain.StartPosition = "CenterScreen"
$frmMain.Text = $appTitle

#---------------------------------------------------------------------------------------------------------------------------
# ListView: lvFiles

$lvFiles = New-Object System.Windows.Forms.ListView
$lvFiles.AllowDrop=$true
$lvFiles.CheckBoxes=$true
$lvFiles.Columns.Add("Path", 500)
$lvFiles.Columns.Add("Original File Name", 200)
$lvFiles.Columns.Add("New File Name", 250)
$lvFiles.FullRowSelect=$true
$lvFiles.Location = New-Object System.Drawing.Point(3,3)
$lvFiles.Size = New-Object System.Drawing.Size(($frmMain.Width*0.75),($frmMain.Height-76))
$lvFiles.View = [System.Windows.Forms.View]::Details
$lvFiles.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop))
    {
        $_.Effect = [Windows.Forms.DragDropEffects]::Copy
    }
    else
    {
        $_.Effect = [Windows.Forms.DragDropEffects]::None
    }
})
$lvFiles.Add_DragDrop({
   $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
   foreach ($file in $files)
   {
      $ext=[System.IO.Path]::GetExtension($file).ToLower()
      if ($supportedExtensions -contains $ext) #($file -like "*.arw" -or $file -like "*.nef")
      {
         $item = New-Object Windows.Forms.ListViewItem
         $item.Checked=$true
         $item.Text = [System.IO.Path]::GetDirectoryName($file)
         $item.SubItems.Add([System.IO.Path]::GetFileName($file))
         $item.SubItems.Add("")
         $lvFiles.Items.Add($item)
      }
   }
})

#---------------------------------------------------------------------------------------------------------------------------
# Button: Load

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Location = New-Object System.Drawing.Point(5, ($lvFiles.Height+5))
$btnLoad.Size = New-Object System.Drawing.Size(90,30)
$btnLoad.Text = "Load files"

#---------------------------------------------------------------------------------------------------------------------------
# Button: Process
$btnProcess = New-Object System.Windows.Forms.Button
$btnProcess.Location = New-Object System.Drawing.Point(($btnLoad.Location.X+$btnLoad.Width+7), ($btnLoad.Location.Y))
$btnProcess.Size = New-Object System.Drawing.Size(90,30)
$btnProcess.Text = "Process Data"

#---------------------------------------------------------------------------------------------------------------------------
# Button: Rename
$btnRename = New-Object System.Windows.Forms.Button
$btnRename.Location = New-Object System.Drawing.Point(($btnProcess.Location.X+$btnProcess.Width+7), ($btnProcess.Location.Y))
$btnRename.Size = New-Object System.Drawing.Size(90,30)
$btnRename.Text = "Rename"

#---------------------------------------------------------------------------------------------------------------------------
# Button: Clear
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(($lvFiles.Width-$btnClear.Width-10), ($btnProcess.Location.Y))
$btnClear.Size = New-Object System.Drawing.Size(90,30)
$btnClear.Text = "Clear"

#---------------------------------------------------------------------------------------------------------------------------
# Functions


function LoadItems()
{
   $dlgFolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
   $dlgFolderBrowser.Description = "Select folder:"
   $dlgFolderBrowser.ShowNewFolderButton = $false

   if ($dlgFolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
   {
      $files = Get-ChildItem -Path $dlgFolderBrowser.SelectedPath -File
      foreach ($file in $files)
      {
         $ext=[System.IO.Path]::GetExtension($file).ToLower()
         if ($supportedExtensions -contains $ext)
         {
            $item = New-Object Windows.Forms.ListViewItem
            $item.Checked=$true
            $item.Text = $file.DirectoryName #[System.IO.Path]::GetDirectoryName(
            $item.SubItems.Add([System.IO.Path]::GetFileName($file))
            $item.SubItems.Add("")
            $lvFiles.Items.Add($item)
         }
      }
   }
   else
   {
      [System.Windows.Forms.MessageBox]::Show("No folder selected")
   }
}

#---------------------------------------------------------------------------------------------------------------------------

function ProcessItems()
{
   $btnRename.Enabled = $flase
   $btnRename.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 200)
   $btnClear.Enabled = $flase

   foreach ($item in $lvFiles.Items)
   {
      $imgFile = $item.Text + "\" + $item.SubItems[1].Text
      $ext=[System.IO.Path]::GetExtension($imgFile).ToLower()
      $shutterCount =  & $ExifToolPath -ShutterCount -s -s -s "$imgFile"
      $creationTimeRaw = & $ExifToolPath -DateTimeOriginal -s -s -s -d "%Y.%m.%d %H.%M.%S" "$imgFile"
      if ($shutterCount -eq "")
      {
         $newFileName="$creationTimeRaw"
      }
      else
      {
         $newFileName = "$creationTimeRaw $shutterCount"
      }
      
      # if last char is space delete it
      if ($newFileName[-1] -eq " ")
      {
         $newFileName=$newFileName.Substring(0, $newFileName.Length - 1)
      }

      if ($newFileName -eq "")
      {
         $item.Checked = $false
         $item.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 200)
      }
      else
      {
         $item.SubItems[2].Text="$newFileName$ext"
      }
    }
    $btnRename.Enabled = $true
    $btnRename.BackColor = [System.Drawing.SystemColors]::ButtonHighlight          
    $btnClear.Enabled = $true
}

#---------------------------------------------------------------------------------------------------------------------------

function RenameItems()
{
   $btnLoad.Enabled = $false
   $btnProcess.Enabled = $false
   $btnClear.Enabled = $false

   foreach ($item in $lvFiles.Items)
   {
      $filePath =$item.Text + "\" + $item.SubItems[1].Text
      $newFileName = $item.SubItems[2].Text

      if ($newFileName -ne "" -and $newFileName -ne "EXIF not found")
      {
         Rename-Item -Path $filePath -NewName $newFileName -Force
         $item.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 200)
      }
   }
   $btnLoad.Enabled = $true
   $btnProcess.Enabled = $true
   $btnClear.Enabled = $true
}

#---------------------------------------------------------------------------------------------------------------------------
# Events

$btnLoad.Add_Click({ LoadItems })

$btnProcess.Add_Click({ ProcessItems })

$btnRename.Add_Click({ RenameItems })

$btnClear.Add_Click({ $lvFiles.Clear() })

#---------------------------------------------------------------------------------------------------------------------------
# Add controls

$frmMain.Controls.AddRange(@($lvFiles, $btnLoad,$btnProcess, $btnRename, $btnClear))

#---------------------------------------------------------------------------------------------------------------------------
# Show main window

$frmMain.Add_Shown({
   $frmMain.Activate()
})

[void] $frmMain.ShowDialog()
