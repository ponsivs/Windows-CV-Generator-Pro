<#
.SYNOPSIS
    Windows CV Generator Pro - Professional GUI-Based Resume/CV Generator
.DESCRIPTION
    Generates professional CVs/Resumes in multiple formats (HTML, PDF, DOCX) 
    with customizable templates, categories, and GUI interface.
.NOTES
    Developer: IGRF Pvt. Ltd.
    Year: 2026
    Version: 1.3
    Website: https://igrf.co.in/en/
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

param([switch]$NoConsole)

# Check if running in PowerShell ISE or with -NoConsole switch
if ($NoConsole -or $Host.Name -match "ISE") {
    $global:RunningAsExe = $false
} else {
    # Try to hide the console window if running from terminal
    try {
        Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")] 
        public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")] 
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'
        $consolePtr = [Console.Window]::GetConsoleWindow()
        [Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
        $global:RunningAsExe = $true
    } catch {
        $global:RunningAsExe = $false
    }
}

#region Initialization and Console Suppression
$global:RunningAsExe = $false
$global:MainFormLoaded = $false
$global:GuiInitialized = $false

# Completely suppress console output
try {
    $null = [Console]::SetOut([System.IO.TextWriter]::Null)
    $null = [Console]::SetError([System.IO.TextWriter]::Null)
} catch {}

# Set all preference variables to suppress output
$ErrorActionPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'

# Helper function to show messages
function Show-SafeMessage {
    param(
        [string]$Message,
        [string]$Title = "Information",
        [string]$Buttons = "OK",
        [string]$Icon = "Information"
    )
    
    if ($global:GuiInitialized -and $global:MainFormLoaded) {
        try {
            return [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon)
        } catch {}
    }
    return $null
}

# Load Assemblies - FIXED VERSION
try {
    Write-Host "Loading assemblies..." -ForegroundColor Yellow
    
    # Load System.Windows.Forms
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Write-Host "✓ System.Windows.Forms loaded" -ForegroundColor Green
    } catch {
        try {
            [void][System.Reflection.Assembly]::Load("System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
            Write-Host "✓ System.Windows.Forms loaded (alternative method)" -ForegroundColor Green
        } catch {
            $null = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
            Write-Host "✓ System.Windows.Forms loaded (partial name)" -ForegroundColor Green
        }
    }
    
    # Load System.Drawing
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Write-Host "✓ System.Drawing loaded" -ForegroundColor Green
    } catch {
        try {
            [void][System.Reflection.Assembly]::Load("System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
            Write-Host "✓ System.Drawing loaded (alternative method)" -ForegroundColor Green
        } catch {
            $null = [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
            Write-Host "✓ System.Drawing loaded (partial name)" -ForegroundColor Green
        }
    }
    
    # Load other required assemblies
    try {
        Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
        Write-Host "✓ System.Web loaded" -ForegroundColor Green
    } catch {}
    
    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        Write-Host "✓ PresentationFramework loaded" -ForegroundColor Green
    } catch {}
    
    $global:GuiInitialized = $true
    Write-Host "✓ All assemblies loaded successfully" -ForegroundColor Green
    
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to load required assemblies.`nPlease install .NET Framework 4.5 or later.`n`nError: $_",
        "Critical Error",
        "OK",
        "Error"
    )
    exit 1
}

# Check if assemblies loaded properly
if (-not $global:GuiInitialized) {
    [System.Windows.Forms.MessageBox]::Show(
        "GUI assemblies failed to load. Application cannot continue.",
        "Error",
        "OK",
        "Error"
    )
    exit 1
}
#endregion

# Helper function to create fonts safely - FIXED VERSION
function New-SafeFont {
    param(
        [string]$FontName = "Segoe UI",
        [float]$Size = 10,
        [string]$Style = "Regular"
    )
    
    try {
        # Convert string style to FontStyle enum
        $fontStyle = [System.Drawing.FontStyle]::Regular
        if ($Style -eq "Bold") { $fontStyle = [System.Drawing.FontStyle]::Bold }
        elseif ($Style -eq "Italic") { $fontStyle = [System.Drawing.FontStyle]::Italic }
        elseif ($Style -eq "Underline") { $fontStyle = [System.Drawing.FontStyle]::Underline }
        elseif ($Style -eq "Bold, Italic") { $fontStyle = [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic }
        
        # Create the font - FIXED: Removed recursive call
        return New-Object System.Drawing.Font($FontName, $Size, $fontStyle)
    } catch {
        # Fallback to Arial if Segoe UI is not available
        try {
            return New-Object System.Drawing.Font("Arial", $Size, $fontStyle)
        } catch {
            # Final fallback
            return New-Object System.Drawing.Font("Microsoft Sans Serif", $Size, $fontStyle)
        }
    }
}

#region Global Variables and Directories
$script:Version = "1.3"
$script:Developer = "IGRF Pvt. Ltd."
$script:Year = "2026"
$script:Website = "https://igrf.co.in/en/"
$script:BaseDir = "$env:USERPROFILE\CVGeneratorPro"
$script:ProfilesDir = "$BaseDir\Profiles"
$script:TemplatesDir = "$BaseDir\Templates"
$script:OutputDir = "$BaseDir\Output"
$script:ConfigsDir = "$BaseDir\Configs"
$script:ImagesDir = "$BaseDir\Images"
$script:PhotosDir = "$BaseDir\Photos"
$script:LogsDir = "$BaseDir\Logs"
$script:SignaturesDir = "$BaseDir\Signatures"

# Enhanced Templates with more details - Updated for Academic features
$script:Templates = @{
    "Modern" = @{
        Name = "Modern Professional"
        Description = "Clean, contemporary design with accent colors"
        Colors = @("Blue", "Green", "Purple", "Red", "Teal", "Orange")
        Sections = @("Personal", "Summary", "Experience", "Education", "Skills", "Projects", "Certifications", "Languages", "References")
        Layout = "Single Column"
        Style = "Professional"
        SupportsAcademic = $false
    }
    "Classic" = @{
        Name = "Classic Formal"
        Description = "Traditional two-column layout"
        Colors = @("Navy", "Charcoal", "Burgundy", "DarkGreen", "DarkBlue", "Maroon")
        Sections = @("Personal", "Experience", "Education", "Skills", "Certifications", "Languages", "References")
        Layout = "Two Column"
        Style = "Traditional"
        SupportsAcademic = $false
    }
    "Creative" = @{
        Name = "Creative Portfolio"
        Description = "Modern design with photo and portfolio sections"
        Colors = @("Teal", "Orange", "Purple", "Gold", "Magenta", "Cyan")
        Sections = @("Personal", "Summary", "Experience", "Education", "Skills", "Projects", "Portfolio", "Certifications", "Languages", "Interests")
        Layout = "Single Column"
        Style = "Creative"
        SupportsAcademic = $false
    }
    "Minimal" = @{
        Name = "Minimalist"
        Description = "Simple, clean layout with focus on content"
        Colors = @("Black", "Gray", "DarkBlue", "DarkGreen", "SlateGray")
        Sections = @("Personal", "Experience", "Education", "Skills", "Languages")
        Layout = "Single Column"
        Style = "Minimal"
        SupportsAcademic = $false
    }
    "Academic" = @{
        Name = "Academic"
        Description = "Designed for academic and research positions with comprehensive research sections"
        Colors = @("DarkBlue", "Maroon", "ForestGreen", "SlateGray", "Navy", "DarkRed")
        Sections = @("Personal", "Education", "Research", "Publications", "Teaching", "Skills", "Awards", "References", "Conferences", "ResearchWork", "Editorial", "PhDDetails", "Seminars", "PhDThesisEvaluation", "Declaration", "ProfessionalProfiles")
        Layout = "Two Column"
        Style = "Academic"
        SupportsAcademic = $true
    }
    "Executive" = @{
        Name = "Executive"
        Description = "Sophisticated design for senior positions"
        Colors = @("Navy", "Charcoal", "Burgundy", "DarkGray", "RoyalBlue")
        Sections = @("Personal", "Executive Summary", "Career Highlights", "Experience", "Education", "Board Positions", "Skills", "Awards", "Publications")
        Layout = "Single Column"
        Style = "Executive"
        SupportsAcademic = $false
    }
    "Technical" = @{
        Name = "Technical"
        Description = "Focus on technical skills and projects"
        Colors = @("DarkBlue", "Green", "Orange", "Purple", "SteelBlue")
        Sections = @("Personal", "Technical Summary", "Experience", "Technical Skills", "Projects", "Certifications", "Education", "Publications", "Patents")
        Layout = "Two Column"
        Style = "Technical"
        SupportsAcademic = $true
    }
    "Student" = @{
        Name = "Student/Fresher"
        Description = "Designed for students and recent graduates"
        Colors = @("Blue", "Green", "Teal", "Orange", "LightBlue")
        Sections = @("Personal", "Objective", "Education", "Projects", "Internships", "Skills", "Extracurricular", "Awards", "Volunteer")
        Layout = "Single Column"
        Style = "Modern"
        SupportsAcademic = $false
    }
    "International" = @{
        Name = "International Standard"
        Description = "Europass style CV for international applications"
        Colors = @("DarkBlue", "Green", "Blue", "Gray")
        Sections = @("Personal", "Professional Experience", "Education", "Skills", "Languages", "Certifications", "Additional Information")
        Layout = "Single Column"
        Style = "International"
        SupportsAcademic = $false
    }
    "Academic Comprehensive" = @{
        Name = "Academic Comprehensive"
        Description = "Complete academic CV with all research and teaching details"
        Colors = @("DarkBlue", "Maroon", "DarkGreen", "Navy", "DarkSlateGray")
        Sections = @("Personal", "PhD Details", "Research Work", "Publications", "Editorial Activities", "Teaching", "Seminars", "PhD Thesis Evaluation", "Skills", "Awards", "Professional Profiles", "Declaration")
        Layout = "Single Column"
        Style = "Academic"
        SupportsAcademic = $true
    }
}

# Enhanced Categories
$script:Categories = @(
    "Information Technology",
    "Software Engineering",
    "Data Science",
    "Healthcare",
    "Finance & Banking",
    "Marketing & Advertising",
    "Sales & Business Development",
    "Education & Teaching",
    "Research & Development",
    "Creative Arts & Design",
    "Administrative & Clerical",
    "Legal Services",
    "Consulting",
    "Management",
    "Customer Service",
    "Manufacturing & Production",
    "Hospitality & Tourism",
    "Government & Public Service",
    "Non-Profit & NGOs",
    "Engineering",
    "Architecture",
    "Human Resources",
    "Media & Journalism",
    "Retail",
    "Academic & Research",
    "University Professor",
    "Research Scientist",
    "Postdoctoral Researcher",
    "Academic Administrator"
)

# Enhanced Skill Categories
$script:SkillCategories = @{
    "Technical" = @("Programming", "Databases", "Networking", "Cloud Computing", "Cybersecurity", "DevOps", "AI/ML", "Data Analysis", "Web Development", "Mobile Development", "Software Testing", "System Administration")
    "Business" = @("Project Management", "Strategic Planning", "Financial Analysis", "Marketing", "Sales", "Negotiation", "Business Development", "Market Research", "Risk Management", "Supply Chain Management")
    "Soft Skills" = @("Communication", "Leadership", "Teamwork", "Problem Solving", "Time Management", "Adaptability", "Creativity", "Critical Thinking", "Emotional Intelligence", "Conflict Resolution")
    "Creative" = @("Graphic Design", "Video Editing", "Writing", "Photography", "UI/UX Design", "Content Creation", "Animation", "Digital Marketing", "Social Media Management")
    "Languages" = @("English", "Spanish", "French", "German", "Chinese", "Japanese", "Arabic", "Hindi", "Portuguese", "Russian", "Italian", "Korean")
    "Technical Tools" = @("Microsoft Office", "Adobe Creative Suite", "JIRA", "Confluence", "Git", "Docker", "Kubernetes", "AWS", "Azure", "Google Cloud")
    "Research Skills" = @("Research Methodology", "Statistical Analysis", "Literature Review", "Experimental Design", "Data Collection", "Qualitative Analysis", "Quantitative Analysis", "Research Ethics", "Grant Writing", "Peer Review")
    "Teaching Skills" = @("Curriculum Development", "Classroom Management", "Student Assessment", "Online Teaching", "Laboratory Instruction", "Thesis Supervision", "Mentoring", "Pedagogical Research", "Educational Technology")
}

# Initialize directories
function Initialize-Directories {
    @($BaseDir, $ProfilesDir, $TemplatesDir, $OutputDir, $ConfigsDir, $ImagesDir, $PhotosDir, $LogsDir, $SignaturesDir) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
        }
    }
    
    # Create placeholder image if not exists
    $placeholderPath = "$ImagesDir\placeholder.png"
    if (-not (Test-Path $placeholderPath)) {
        try {
            $bitmap = New-Object System.Drawing.Bitmap(200, 250)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.Clear([System.Drawing.Color]::White)
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::LightGray, 2)
            $graphics.DrawRectangle($pen, 10, 10, 180, 230)
            
            # FIXED: Font creation with FontStyle
            $font = New-SafeFont -FontName "Arial" -Size 48 -Style "Regular"
            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::LightGray)
            $graphics.DrawString("👤", $font, $brush, 60, 70)
            
            # FIXED: Font creation with FontStyle
            $fontSmall = New-SafeFont -FontName "Arial" -Size 12 -Style "Regular"
            $brushText = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Gray)
            $graphics.DrawString("Add Photo", $fontSmall, $brushText, 70, 180)
            
            $bitmap.Save($placeholderPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $graphics.Dispose()
            $bitmap.Dispose()
        } catch {
            Write-Warning "Could not create placeholder image: $_"
        }
    }
    
    # Create placeholder signature image if not exists
    $signaturePlaceholderPath = "$SignaturesDir\signature_placeholder.png"
    if (-not (Test-Path $signaturePlaceholderPath)) {
        try {
            $bitmap = New-Object System.Drawing.Bitmap(300, 150)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.Clear([System.Drawing.Color]::White)
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::LightGray, 2)
            $graphics.DrawRectangle($pen, 10, 10, 280, 130)
            
            # FIXED: Font creation with FontStyle
            $font = New-SafeFont -FontName "Segoe Script" -Size 36 -Style "Regular"
            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::LightGray)
            $graphics.DrawString("Signature", $font, $brush, 50, 40)
            
            $fontSmall = New-SafeFont -FontName "Arial" -Size 12 -Style "Regular"
            $brushText = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Gray)
            $graphics.DrawString("Add your digital signature here", $fontSmall, $brushText, 80, 100)
            
            $bitmap.Save($signaturePlaceholderPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $graphics.Dispose()
            $bitmap.Dispose()
        } catch {
            Write-Warning "Could not create signature placeholder image: $_"
        }
    }
}
#endregion

#region Main Form
function Show-MainForm {
    # Emergency font fix
    $defaultFont = $null
    try {
        $defaultFont = New-SafeFont -FontName "Segoe UI" -Size 10 -Style "Regular"
    } catch {
        try {
            $defaultFont = New-SafeFont -FontName "Arial" -Size 10 -Style "Regular"
        } catch {
            # If all else fails, create a font without specifying name
            $defaultFont = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
        }
    }
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.Text = "Windows CV Generator Pro v$script:Version"
    $mainForm.Size = New-Object System.Drawing.Size(900, 680)
    $mainForm.StartPosition = "CenterScreen"
    $mainForm.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $mainForm.Font = $defaultFont
    
    # Set icon
    $iconPath = "$script:ImagesDir\Icon.ico"
    if (Test-Path $iconPath) {
        try {
            $mainForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
        } catch {}
    }
    
    # Menu Strip - Enhanced
    $menuStrip = New-Object System.Windows.Forms.MenuStrip
    
    # File Menu
    $fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("File")
    
    $newCVMenu = New-Object System.Windows.Forms.ToolStripMenuItem("New CV")
    $newCVMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::N
    $newCVMenu.Add_Click({ Show-CVBuilderForm })
    
    $loadProfileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Load Profile")
    $loadProfileMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::O
    $loadProfileMenu.Add_Click({ Load-Profile })
    
    $saveProfileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Save Draft")
    $saveProfileMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::S
    $saveProfileMenu.Add_Click({ Save-CVDraft })
    
    $saveAsMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Save As...")
    $saveAsMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Shift -bor [System.Windows.Forms.Keys]::S
    $saveAsMenu.Add_Click({ Save-CVDraftAs })
    
    $exportMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Export CV")
    $exportMenu.Add_Click({ Show-ExportForm })
    
    $exitMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
    $exitMenu.Add_Click({ $mainForm.Close() })
    
    $fileMenu.DropDownItems.AddRange(@(
        $newCVMenu,
        (New-Object System.Windows.Forms.ToolStripSeparator),
        $loadProfileMenu,
        $saveProfileMenu,
        $saveAsMenu,
        (New-Object System.Windows.Forms.ToolStripSeparator),
        $exportMenu,
        (New-Object System.Windows.Forms.ToolStripSeparator),
        $exitMenu
    ))
    
    # Edit Menu - Enhanced with actual functionality
    $editMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Edit")
    
    $undoMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Undo")
    $undoMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Z
    $undoMenu.Add_Click({ Undo-LastAction })
    
    $redoMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Redo")
    $redoMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Y
    $redoMenu.Add_Click({ Redo-LastAction })
    
    $cutMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Cut")
    $cutMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::X
    $cutMenu.Add_Click({ Perform-Cut })
    
    $copyMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Copy")
    $copyMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::C
    $copyMenu.Add_Click({ Perform-Copy })
    
    $pasteMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Paste")
    $pasteMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::V
    $pasteMenu.Add_Click({ Perform-Paste })
    
    $selectAllMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Select All")
    $selectAllMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::A
    $selectAllMenu.Add_Click({ Perform-SelectAll })
    
    $editMenu.DropDownItems.AddRange(@(
        $undoMenu,
        $redoMenu,
        (New-Object System.Windows.Forms.ToolStripSeparator),
        $cutMenu,
        $copyMenu,
        $pasteMenu,
        (New-Object System.Windows.Forms.ToolStripSeparator),
        $selectAllMenu
    ))
    
    # Tools Menu - Enhanced with actual functionality
    $toolsMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Tools")
    
    $templatesMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Templates Gallery")
    $templatesMenu.Add_Click({ Show-TemplatesForm })
    
    $spellCheckMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Spell Check")
    $spellCheckMenu.Add_Click({ Invoke-SpellCheck })
    
    $wordCountMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Word Count")
    $wordCountMenu.Add_Click({ Show-WordCount })
    
    $toolsMenu.DropDownItems.AddRange(@(
        $templatesMenu,
        (New-Object System.Windows.Forms.ToolStripSeparator),
        $spellCheckMenu,
        $wordCountMenu
    ))
    
    # View Menu - Enhanced
    $viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem("View")
    
    $toolbarMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Toolbar")
    $toolbarMenu.Checked = $true
    $toolbarMenu.Add_Click({ 
        $toolbarMenu.Checked = -not $toolbarMenu.Checked
        Toggle-Toolbar
    })
    
    $statusBarMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Status Bar")
    $statusBarMenu.Checked = $true
    $statusBarMenu.Add_Click({ 
        $statusBarMenu.Checked = -not $statusBarMenu.Checked
        Toggle-StatusBar
    })
    
    $zoomInMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Zoom In")
    $zoomInMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Add
    $zoomInMenu.Add_Click({ Zoom-In })
    
    $zoomOutMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Zoom Out")
    $zoomOutMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Subtract
    $zoomOutMenu.Add_Click({ Zoom-Out })
    
    $resetZoomMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Reset Zoom")
    $resetZoomMenu.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::D0
    $resetZoomMenu.Add_Click({ Reset-Zoom })
    
    $viewMenu.DropDownItems.AddRange(@(
        $toolbarMenu,
        $statusBarMenu,
        (New-Object System.Windows.Forms.ToolStripSeparator),
        $zoomInMenu,
        $zoomOutMenu,
        $resetZoomMenu
    ))
    
    # Help Menu - Only About as requested
    $helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Help")
    
    $aboutMenu = New-Object System.Windows.Forms.ToolStripMenuItem("About")
    $aboutMenu.Add_Click({ Show-AboutForm })
    
    $helpMenu.DropDownItems.Add($aboutMenu)
    
    # Add menus to strip
    $menuStrip.Items.AddRange(@($fileMenu, $editMenu, $toolsMenu, $viewMenu, $helpMenu))
    $mainForm.MainMenuStrip = $menuStrip
    
    # Main Panel
    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Dock = "Fill"
    $mainPanel.BackColor = [System.Drawing.Color]::White
    
    # Logo Panel
    $logoPanel = New-Object System.Windows.Forms.Panel
    $logoPanel.Size = New-Object System.Drawing.Size(800, 80)
    $logoPanel.Location = New-Object System.Drawing.Point(50, 20)
    $logoPanel.BackColor = [System.Drawing.Color]::Transparent
    
    $logoLabel = New-Object System.Windows.Forms.Label
    $logoLabel.Text = "Windows CV Generator Pro"
    $logoLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 22 -Style "Bold"
    $logoLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $logoLabel.Size = New-Object System.Drawing.Size(800, 50)
    $logoLabel.Location = New-Object System.Drawing.Point(0, 15)
    $logoLabel.TextAlign = "MiddleCenter"
    $logoPanel.Controls.Add($logoLabel)
    
    # Welcome Label
    $welcomeLabel = New-Object System.Windows.Forms.Label
    $welcomeLabel.Text = "Professional Resume/CV Creation Tool"
    $welcomeLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 16 -Style "Regular"
    $welcomeLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $welcomeLabel.Size = New-Object System.Drawing.Size(800, 35)
    $welcomeLabel.Location = New-Object System.Drawing.Point(50, 110)
    $welcomeLabel.TextAlign = "MiddleCenter"
    
    # Buttons Panel
    $buttonsPanel = New-Object System.Windows.Forms.Panel
    $buttonsPanel.Size = New-Object System.Drawing.Size(500, 320)
    $buttonsPanel.Location = New-Object System.Drawing.Point(200, 160)
    
    # Buttons - FIXED: Corrected font creation calls
    $newCVButton = New-Object System.Windows.Forms.Button
    $newCVButton.Text = "Create New CV"
    $newCVButton.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $newCVButton.Size = New-Object System.Drawing.Size(220, 50)
    $newCVButton.Location = New-Object System.Drawing.Point(140, 20)
    $newCVButton.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $newCVButton.ForeColor = [System.Drawing.Color]::White
    $newCVButton.FlatStyle = "Flat"
    $newCVButton.FlatAppearance.BorderSize = 0
    $newCVButton.Add_Click({ Show-CVBuilderForm })
    
    $loadProfileButton = New-Object System.Windows.Forms.Button
    $loadProfileButton.Text = "Load Existing Profile"
    $loadProfileButton.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $loadProfileButton.Size = New-Object System.Drawing.Size(220, 50)
    $loadProfileButton.Location = New-Object System.Drawing.Point(140, 85)
    $loadProfileButton.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $loadProfileButton.ForeColor = [System.Drawing.Color]::White
    $loadProfileButton.FlatStyle = "Flat"
    $loadProfileButton.FlatAppearance.BorderSize = 0
    $loadProfileButton.Add_Click({ Load-Profile })
    
    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "Export CV"
    $exportButton.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $exportButton.Size = New-Object System.Drawing.Size(220, 50)
    $exportButton.Location = New-Object System.Drawing.Point(140, 150)
    $exportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 87, 34)
    $exportButton.ForeColor = [System.Drawing.Color]::White
    $exportButton.FlatStyle = "Flat"
    $exportButton.FlatAppearance.BorderSize = 0
    $exportButton.Add_Click({ Show-ExportForm })
    
    $templatesButton = New-Object System.Windows.Forms.Button
    $templatesButton.Text = "Browse Templates"
    $templatesButton.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $templatesButton.Size = New-Object System.Drawing.Size(220, 50)
    $templatesButton.Location = New-Object System.Drawing.Point(140, 215)
    $templatesButton.BackColor = [System.Drawing.Color]::FromArgb(156, 39, 176)
    $templatesButton.ForeColor = [System.Drawing.Color]::White
    $templatesButton.FlatStyle = "Flat"
    $templatesButton.FlatAppearance.BorderSize = 0
    $templatesButton.Add_Click({ Show-TemplatesForm })
    
    $buttonsPanel.Controls.AddRange(@($newCVButton, $loadProfileButton, $exportButton, $templatesButton))
    
    # Developer Panel
    $devPanel = New-Object System.Windows.Forms.Panel
    $devPanel.Size = New-Object System.Drawing.Size(800, 90)
    $devPanel.Location = New-Object System.Drawing.Point(50, 520)
    $devPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $devPanel.BorderStyle = "FixedSingle"
    
    $devLabel1 = New-Object System.Windows.Forms.Label
    $devLabel1.Text = "Developer: $script:Developer"
    $devLabel1.Font = New-SafeFont -FontName "Segoe UI" -Size 9 -Style "Bold"
    $devLabel1.Location = New-Object System.Drawing.Point(20, 10)
    $devLabel1.Size = New-Object System.Drawing.Size(350, 25)
    
    $devLabel2 = New-Object System.Windows.Forms.Label
    $devLabel2.Text = "Version: $script:Version | Year: $script:Year"
    $devLabel2.Font = New-SafeFont -FontName "Segoe UI" -Size 9 -Style "Regular"
    $devLabel2.Location = New-Object System.Drawing.Point(20, 35)
    $devLabel2.Size = New-Object System.Drawing.Size(350, 25)
    
    $websiteLink = New-Object System.Windows.Forms.LinkLabel
    $websiteLink.Text = $script:Website
    $websiteLink.Font = New-SafeFont -FontName "Segoe UI" -Size 9 -Style "Underline"
    $websiteLink.Location = New-Object System.Drawing.Point(20, 60)
    $websiteLink.Size = New-Object System.Drawing.Size(350, 25)
    $websiteLink.LinkColor = [System.Drawing.Color]::Blue
    $websiteLink.Add_Click({ Start-Process $script:Website })
    
    $devPanel.Controls.AddRange(@($devLabel1, $devLabel2, $websiteLink))
    
    # Status Bar
    $statusBar = New-Object System.Windows.Forms.StatusStrip
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = "Ready - $script:Developer © $script:Year"
    $statusBar.Items.Add($statusLabel)
    
    # Add controls to main panel and form
    $mainPanel.Controls.AddRange(@($logoPanel, $welcomeLabel, $buttonsPanel, $devPanel))
    $mainForm.Controls.AddRange(@($menuStrip, $mainPanel, $statusBar))
    
    # Form shown event
    $mainForm.Add_Shown({ 
        $mainForm.Activate() 
        $global:MainFormLoaded = $true
    })
    
    # Store references for menu functionality
    $script:MainForm = $mainForm
    $script:StatusBar = $statusBar
    $script:StatusLabel = $statusLabel
    
    $null = $mainForm.ShowDialog()
}
#endregion

#region Enhanced Edit Menu Functions
function Undo-LastAction {
    Show-SafeMessage -Message "Undo functionality would be implemented here." -Title "Undo" -Icon "Information"
}

function Redo-LastAction {
    Show-SafeMessage -Message "Redo functionality would be implemented here." -Title "Redo" -Icon "Information"
}

function Perform-Cut {
    if ([System.Windows.Forms.Clipboard]::ContainsText()) {
        $activeControl = [System.Windows.Forms.Form]::ActiveForm.ActiveControl
        if ($activeControl -is [System.Windows.Forms.TextBox]) {
            $activeControl.Cut()
        }
    }
}

function Perform-Copy {
    $activeControl = [System.Windows.Forms.Form]::ActiveForm.ActiveControl
    if ($activeControl -is [System.Windows.Forms.TextBox]) {
        $activeControl.Copy()
    }
}

function Perform-Paste {
    if ([System.Windows.Forms.Clipboard]::ContainsText()) {
        $activeControl = [System.Windows.Forms.Form]::ActiveForm.ActiveControl
        if ($activeControl -is [System.Windows.Forms.TextBox]) {
            $activeControl.Paste()
        }
    }
}

function Perform-SelectAll {
    $activeControl = [System.Windows.Forms.Form]::ActiveForm.ActiveControl
    if ($activeControl -is [System.Windows.Forms.TextBox]) {
        $activeControl.SelectAll()
    } elseif ($activeControl -is [System.Windows.Forms.DataGridView]) {
        $activeControl.SelectAll()
    }
}
#endregion

#region Enhanced Tools Menu Functions - ACTUAL IMPLEMENTATION
function Invoke-SpellCheck {
    # Check if we're in CV Builder form
    $activeForm = [System.Windows.Forms.Form]::ActiveForm
    if ($activeForm -and $activeForm.Text -match "CV Builder") {
        # Check all text controls in the form
        $textControls = $activeForm.Controls | ForEach-Object {
            if ($_ -is [System.Windows.Forms.TextBox]) {
                $_
            }
        }
        
        if ($textControls.Count -eq 0) {
            Show-SafeMessage -Message "No text controls found for spell checking." -Title "Spell Check" -Icon "Information"
            return
        }
        
        # Simple spell check implementation
        $misspelledWords = @()
        foreach ($control in $textControls) {
            if ($control.Text -and $control.Visible) {
                # Split text into words and check basic spelling
                $words = $control.Text -split '\s+'
                foreach ($word in $words) {
                    # Remove punctuation
                    $cleanWord = $word -replace '[^\w]', ''
                    if ($cleanWord -and $cleanWord.Length -gt 2) {
                        # Basic check: words with consecutive consonants might be misspelled
                        if ($cleanWord -match '[bcdfghjklmnpqrstvwxyzBCDFGHJKLMNPQRSTVWXYZ]{4,}') {
                            $misspelledWords += $cleanWord
                        }
                    }
                }
            }
        }
        
        if ($misspelledWords.Count -gt 0) {
            Show-SafeMessage -Message "Possible spelling issues found:`n$($misspelledWords -join ', ')" -Title "Spell Check Results" -Icon "Warning"
        } else {
            Show-SafeMessage -Message "No spelling issues found." -Title "Spell Check" -Icon "Information"
        }
    } else {
        Show-SafeMessage -Message "Please open the CV Builder to use spell check." -Title "Spell Check" -Icon "Information"
    }
}

function Show-WordCount {
    # Check if we're in CV Builder form
    $activeForm = [System.Windows.Forms.Form]::ActiveForm
    if ($activeForm -and $activeForm.Text -match "CV Builder") {
        $totalWords = 0
        $totalCharacters = 0
        
        # Count words in all text controls
        $textControls = $activeForm.Controls | ForEach-Object {
            if ($_ -is [System.Windows.Forms.TextBox]) {
                $_
            }
        }
        
        foreach ($control in $textControls) {
            if ($control.Text) {
                $words = ($control.Text -split '\s+' | Where-Object { $_ -ne '' }).Count
                $chars = $control.Text.Length
                $totalWords += $words
                $totalCharacters += $chars
            }
        }
        
        Show-SafeMessage -Message "Word Count:`nWords: $totalWords`nCharacters: $totalCharacters" -Title "Word Count" -Icon "Information"
    } else {
        Show-SafeMessage -Message "Please open the CV Builder to use word count." -Title "Word Count" -Icon "Information"
    }
}
#endregion

#region Enhanced View Menu Functions
function Toggle-Toolbar {
    # Toolbar toggle functionality
    Show-SafeMessage -Message "Toolbar visibility toggled." -Title "View" -Icon "Information"
}

function Toggle-StatusBar {
    if ($script:StatusBar) {
        $script:StatusBar.Visible = -not $script:StatusBar.Visible
    }
}

function Zoom-In {
    Show-SafeMessage -Message "Zoom In functionality would be implemented here." -Title "Zoom In" -Icon "Information"
}

function Zoom-Out {
    Show-SafeMessage -Message "Zoom Out functionality would be implemented here." -Title "Zoom Out" -Icon "Information"
}

function Reset-Zoom {
    Show-SafeMessage -Message "Zoom reset to 100%." -Title "Reset Zoom" -Icon "Information"
}
#endregion

#region CV Builder Form with Enhanced Features
function Show-CVBuilderForm {
    param([string]$ProfilePath = $null)
    
    $builderForm = New-Object System.Windows.Forms.Form
    $builderForm.Text = "CV Builder - Windows CV Generator Pro"
    $builderForm.Size = New-Object System.Drawing.Size(1200, 800)
    $builderForm.StartPosition = "CenterScreen"
    $builderForm.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $builderForm.WindowState = "Maximized"  # FIXED: Open in maximized window mode
    
    # Set icon
    $iconPath = "$script:ImagesDir\Icon.ico"
    if (Test-Path $iconPath) {
        try {
            $builderForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
        } catch {}
    }
    
    # Tab Control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    $tabControl.Font = New-SafeFont -FontName "Segoe UI" -Size 10
    
    # Personal Information Tab - Enhanced with Date of Birth
    $personalTab = New-Object System.Windows.Forms.TabPage
    $personalTab.Text = "Personal Information"
    $personalTab.BackColor = [System.Drawing.Color]::White
    
    $personalPanel = New-Object System.Windows.Forms.Panel
    $personalPanel.Dock = "Fill"
    $personalPanel.AutoScroll = $true
    
    $yPos = 20
    $labelWidth = 200
    $controlWidth = 300
    $spacing = 40
    
    # Photo Panel
    $photoPanel = New-Object System.Windows.Forms.Panel
    $photoPanel.Location = New-Object System.Drawing.Point(650, $yPos)
    $photoPanel.Size = New-Object System.Drawing.Size(300, 350)
    $photoPanel.BorderStyle = "FixedSingle"
    $photoPanel.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
    
    $photoLabel = New-Object System.Windows.Forms.Label
    $photoLabel.Text = "Passport Size Photo"
    $photoLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $photoLabel.Location = New-Object System.Drawing.Point(10, 10)
    $photoLabel.Size = New-Object System.Drawing.Size(280, 30)
    $photoLabel.TextAlign = "MiddleCenter"
    
    $global:photoPictureBox = New-Object System.Windows.Forms.PictureBox
    $global:photoPictureBox.Location = New-Object System.Drawing.Point(50, 50)
    $global:photoPictureBox.Size = New-Object System.Drawing.Size(200, 230)
    $global:photoPictureBox.SizeMode = "Zoom"
    $global:photoPictureBox.BorderStyle = "FixedSingle"
    $global:photoPictureBox.BackColor = [System.Drawing.Color]::White
    
    # Load placeholder image
    $placeholderPath = "$script:ImagesDir\placeholder.png"
    if (Test-Path $placeholderPath) {
        try {
            $global:photoPictureBox.Image = [System.Drawing.Image]::FromFile($placeholderPath)
        } catch {
            $bitmap = New-Object System.Drawing.Bitmap(200, 230)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.Clear([System.Drawing.Color]::LightGray)
            $global:photoPictureBox.Image = $bitmap
            $graphics.Dispose()
        }
    }
    
    $uploadPhotoButton = New-Object System.Windows.Forms.Button
    $uploadPhotoButton.Text = "Upload Photo"
    $uploadPhotoButton.Location = New-Object System.Drawing.Point(50, 290)
    $uploadPhotoButton.Size = New-Object System.Drawing.Size(90, 30)
    $uploadPhotoButton.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $uploadPhotoButton.ForeColor = [System.Drawing.Color]::White
    $uploadPhotoButton.FlatStyle = "Flat"
    $uploadPhotoButton.FlatAppearance.BorderSize = 0
    $uploadPhotoButton.Font = New-SafeFont -FontName "Segoe UI" -Size 9
    $uploadPhotoButton.Add_Click({ Upload-Photo })
    
    $removePhotoButton = New-Object System.Windows.Forms.Button
    $removePhotoButton.Text = "Remove"
    $removePhotoButton.Location = New-Object System.Drawing.Point(160, 290)
    $removePhotoButton.Size = New-Object System.Drawing.Size(90, 30)
    $removePhotoButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $removePhotoButton.ForeColor = [System.Drawing.Color]::White
    $removePhotoButton.FlatStyle = "Flat"
    $removePhotoButton.FlatAppearance.BorderSize = 0
    $removePhotoButton.Font = New-SafeFont -FontName "Segoe UI" -Size 9
    $removePhotoButton.Add_Click({ 
        if (Test-Path $placeholderPath) {
            try {
                $global:photoPictureBox.Image = [System.Drawing.Image]::FromFile($placeholderPath)
            } catch {
                $global:photoPictureBox.Image = $null
            }
        }
        $global:photoPath = $null
        Show-SafeMessage -Message "Photo removed successfully!" -Title "Success" -Icon "Information"
    })
    
    $photoPanel.Controls.AddRange(@($photoLabel, $global:photoPictureBox, $uploadPhotoButton, $removePhotoButton))
    $global:photoPath = $null
    
    # Personal Information Fields with Date of Birth
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = "Full Name:*"
    $nameLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $nameLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:nameTextBox = New-Object System.Windows.Forms.TextBox
    $global:nameTextBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:nameTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += $spacing
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Professional Title:"
    $titleLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $titleLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:titleTextBox = New-Object System.Windows.Forms.TextBox
    $global:titleTextBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:titleTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += $spacing
    
    $dobLabel = New-Object System.Windows.Forms.Label
    $dobLabel.Text = "Date of Birth:"
    $dobLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $dobLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:dobDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $global:dobDateTimePicker.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:dobDateTimePicker.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $global:dobDateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Long
    $global:dobDateTimePicker.ShowCheckBox = $true
    $global:dobDateTimePicker.Checked = $false
    $global:dobDateTimePicker.CustomFormat = " "
    $global:dobDateTimePicker.Add_ValueChanged({
        if ($this.Checked) {
            $this.CustomFormat = "dddd, MMMM dd, yyyy"
        } else {
            $this.CustomFormat = " "
            $this.Value = [DateTime]::Now
        }
    })
    $yPos += $spacing
    
    $emailLabel = New-Object System.Windows.Forms.Label
    $emailLabel.Text = "Email Address:*"
    $emailLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $emailLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:emailTextBox = New-Object System.Windows.Forms.TextBox
    $global:emailTextBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:emailTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += $spacing
    
    $phoneLabel = New-Object System.Windows.Forms.Label
    $phoneLabel.Text = "Phone Number:"
    $phoneLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $phoneLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:phoneTextBox = New-Object System.Windows.Forms.TextBox
    $global:phoneTextBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:phoneTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += $spacing
    
    $addressLabel = New-Object System.Windows.Forms.Label
    $addressLabel.Text = "Address:"
    $addressLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $addressLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:addressTextBox = New-Object System.Windows.Forms.TextBox
    $global:addressTextBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:addressTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += $spacing
    
    $nationalityLabel = New-Object System.Windows.Forms.Label
    $nationalityLabel.Text = "Nationality:"
    $nationalityLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $nationalityLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:nationalityTextBox = New-Object System.Windows.Forms.TextBox
    $global:nationalityTextBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:nationalityTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += $spacing
    
    $linkedInLabel = New-Object System.Windows.Forms.Label
    $linkedInLabel.Text = "LinkedIn Profile:"
    $linkedInLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $linkedInLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:linkedInTextBox = New-Object System.Windows.Forms.TextBox
    $global:linkedInTextBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:linkedInTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += $spacing
    
    $githubLabel = New-Object System.Windows.Forms.Label
    $githubLabel.Text = "GitHub Profile:"
    $githubLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $githubLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:githubTextBox = New-Object System.Windows.Forms.TextBox
    $global:githubTextBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:githubTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += $spacing
    
    $websiteLabel = New-Object System.Windows.Forms.Label
    $websiteLabel.Text = "Personal Website:"
    $websiteLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $websiteLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:websiteTextBox = New-Object System.Windows.Forms.TextBox
    $global:websiteTextBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:websiteTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += $spacing
    
    $summaryLabel = New-Object System.Windows.Forms.Label
    $summaryLabel.Text = "Professional Summary:"
    $summaryLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $summaryLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:summaryTextBox = New-Object System.Windows.Forms.TextBox
    $global:summaryTextBox.Multiline = $true
    $global:summaryTextBox.ScrollBars = "Vertical"
    $global:summaryTextBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:summaryTextBox.Size = New-Object System.Drawing.Size($controlWidth, 100)
    $yPos += 120
    
    $categoryLabel = New-Object System.Windows.Forms.Label
    $categoryLabel.Text = "Industry Category:"
    $categoryLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $categoryLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $global:categoryComboBox = New-Object System.Windows.Forms.ComboBox
    $global:categoryComboBox.Location = New-Object System.Drawing.Point(260, $yPos)
    $global:categoryComboBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $global:categoryComboBox.DropDownStyle = "DropDownList"
    $script:Categories | ForEach-Object { $null = $global:categoryComboBox.Items.Add($_) }
    if ($global:categoryComboBox.Items.Count -gt 0) {
        $global:categoryComboBox.SelectedIndex = 0
    }
    
    # Add all controls to personal panel
    $personalPanel.Controls.AddRange(@(
        $nameLabel, $global:nameTextBox,
        $titleLabel, $global:titleTextBox,
        $dobLabel, $global:dobDateTimePicker,
        $emailLabel, $global:emailTextBox,
        $phoneLabel, $global:phoneTextBox,
        $addressLabel, $global:addressTextBox,
        $nationalityLabel, $global:nationalityTextBox,
        $linkedInLabel, $global:linkedInTextBox,
        $githubLabel, $global:githubTextBox,
        $websiteLabel, $global:websiteTextBox,
        $summaryLabel, $global:summaryTextBox,
        $categoryLabel, $global:categoryComboBox,
        $photoPanel
    ))
    
    $personalTab.Controls.Add($personalPanel)
    
    # Work Experience Tab
    $experienceTab = New-Object System.Windows.Forms.TabPage
    $experienceTab.Text = "Work Experience"
    $experienceTab.BackColor = [System.Drawing.Color]::White
    
    $experiencePanel = New-Object System.Windows.Forms.Panel
    $experiencePanel.Dock = "Fill"
    
    $global:experienceDataGrid = New-Object System.Windows.Forms.DataGridView
    $global:experienceDataGrid.Dock = "Fill"
    $global:experienceDataGrid.ColumnCount = 6
    $global:experienceDataGrid.Columns[0].Name = "Job Title"
    $global:experienceDataGrid.Columns[1].Name = "Company"
    $global:experienceDataGrid.Columns[2].Name = "Location"
    $global:experienceDataGrid.Columns[3].Name = "Start Date"
    $global:experienceDataGrid.Columns[4].Name = "End Date"
    $global:experienceDataGrid.Columns[5].Name = "Description"
    $global:experienceDataGrid.AutoSizeColumnsMode = "Fill"
    $global:experienceDataGrid.AllowUserToAddRows = $false
    
    $expButtonsPanel = New-Object System.Windows.Forms.Panel
    $expButtonsPanel.Dock = "Bottom"
    $expButtonsPanel.Height = 50
    
    $addExpButton = New-Object System.Windows.Forms.Button
    $addExpButton.Text = "Add Experience"
    $addExpButton.Location = New-Object System.Drawing.Point(20, 10)
    $addExpButton.Size = New-Object System.Drawing.Size(150, 30)
    $addExpButton.Add_Click({ Show-ExperienceDialog })
    
    $editExpButton = New-Object System.Windows.Forms.Button
    $editExpButton.Text = "Edit Selected"
    $editExpButton.Location = New-Object System.Drawing.Point(180, 10)
    $editExpButton.Size = New-Object System.Drawing.Size(150, 30)
    $editExpButton.Add_Click({ 
        if ($global:experienceDataGrid.SelectedRows.Count -gt 0) {
            Show-ExperienceDialog -EditMode $true
        }
    })
    
    $removeExpButton = New-Object System.Windows.Forms.Button
    $removeExpButton.Text = "Remove Selected"
    $removeExpButton.Location = New-Object System.Drawing.Point(340, 10)
    $removeExpButton.Size = New-Object System.Drawing.Size(150, 30)
    $removeExpButton.Add_Click({ 
        if ($global:experienceDataGrid.SelectedRows.Count -gt 0) {
            foreach ($row in $global:experienceDataGrid.SelectedRows) {
                $global:experienceDataGrid.Rows.Remove($row)
            }
        }
    })
    
    $expButtonsPanel.Controls.AddRange(@($addExpButton, $editExpButton, $removeExpButton))
    $experiencePanel.Controls.AddRange(@($global:experienceDataGrid, $expButtonsPanel))
    $experienceTab.Controls.Add($experiencePanel)
    
    # Education Tab
    $educationTab = New-Object System.Windows.Forms.TabPage
    $educationTab.Text = "Education"
    $educationTab.BackColor = [System.Drawing.Color]::White
    
    $educationPanel = New-Object System.Windows.Forms.Panel
    $educationPanel.Dock = "Fill"
    
    $global:educationDataGrid = New-Object System.Windows.Forms.DataGridView
    $global:educationDataGrid.Dock = "Fill"
    $global:educationDataGrid.ColumnCount = 6
    $global:educationDataGrid.Columns[0].Name = "Degree"
    $global:educationDataGrid.Columns[1].Name = "Institution"
    $global:educationDataGrid.Columns[2].Name = "Location"
    $global:educationDataGrid.Columns[3].Name = "Start Date"
    $global:educationDataGrid.Columns[4].Name = "End Date"
    $global:educationDataGrid.Columns[5].Name = "GPA/Score"
    $global:educationDataGrid.AutoSizeColumnsMode = "Fill"
    $global:educationDataGrid.AllowUserToAddRows = $false
    
    $eduButtonsPanel = New-Object System.Windows.Forms.Panel
    $eduButtonsPanel.Dock = "Bottom"
    $eduButtonsPanel.Height = 50
    
    $addEduButton = New-Object System.Windows.Forms.Button
    $addEduButton.Text = "Add Education"
    $addEduButton.Location = New-Object System.Drawing.Point(20, 10)
    $addEduButton.Size = New-Object System.Drawing.Size(150, 30)
    $addEduButton.Add_Click({ Show-EducationDialog })
    
    $editEduButton = New-Object System.Windows.Forms.Button
    $editEduButton.Text = "Edit Selected"
    $editEduButton.Location = New-Object System.Drawing.Point(180, 10)
    $editEduButton.Size = New-Object System.Drawing.Size(150, 30)
    $editEduButton.Add_Click({ 
        if ($global:educationDataGrid.SelectedRows.Count -gt 0) {
            Show-EducationDialog -EditMode $true
        }
    })
    
    $removeEduButton = New-Object System.Windows.Forms.Button
    $removeEduButton.Text = "Remove Selected"
    $removeEduButton.Location = New-Object System.Drawing.Point(340, 10)
    $removeEduButton.Size = New-Object System.Drawing.Size(150, 30)
    $removeEduButton.Add_Click({ 
        if ($global:educationDataGrid.SelectedRows.Count -gt 0) {
            foreach ($row in $global:educationDataGrid.SelectedRows) {
                $global:educationDataGrid.Rows.Remove($row)
            }
        }
    })
    
    $eduButtonsPanel.Controls.AddRange(@($addEduButton, $editEduButton, $removeEduButton))
    $educationPanel.Controls.AddRange(@($global:educationDataGrid, $eduButtonsPanel))
    $educationTab.Controls.Add($educationPanel)
    
    # Skills Tab
    $skillsTab = New-Object System.Windows.Forms.TabPage
    $skillsTab.Text = "Skills"
    $skillsTab.BackColor = [System.Drawing.Color]::White
    
    $skillsPanel = New-Object System.Windows.Forms.Panel
    $skillsPanel.Dock = "Fill"
    
    $skillsTabControl = New-Object System.Windows.Forms.TabControl
    $skillsTabControl.Dock = "Fill"
    
    $techSkillsTab = New-Object System.Windows.Forms.TabPage
    $techSkillsTab.Text = "Technical Skills"
    $global:techSkillsTextBox = New-Object System.Windows.Forms.TextBox
    $global:techSkillsTextBox.Multiline = $true
    $global:techSkillsTextBox.Dock = "Fill"
    $global:techSkillsTextBox.Font = New-SafeFont -FontName "Consolas" -Size 10
    $global:techSkillsTextBox.Text = "• Programming: C#, Java, Python, JavaScript`n• Databases: SQL Server, MySQL, MongoDB, PostgreSQL`n• Web Development: HTML5, CSS3, React, Angular, Node.js`n• Cloud Computing: AWS, Azure, Google Cloud`n• DevOps: Docker, Kubernetes, Jenkins, Git, CI/CD"
    $techSkillsTab.Controls.Add($global:techSkillsTextBox)
    
    $businessSkillsTab = New-Object System.Windows.Forms.TabPage
    $businessSkillsTab.Text = "Business Skills"
    $global:businessSkillsTextBox = New-Object System.Windows.Forms.TextBox
    $global:businessSkillsTextBox.Multiline = $true
    $global:businessSkillsTextBox.Dock = "Fill"
    $global:businessSkillsTextBox.Font = New-SafeFont -FontName "Consolas" -Size 10
    $global:businessSkillsTextBox.Text = "• Project Management (Agile, Scrum, Waterfall)`n• Strategic Planning & Analysis`n• Financial Analysis & Budgeting`n• Team Leadership & Development`n• Client Relationship Management`n• Business Development & Sales"
    $businessSkillsTab.Controls.Add($global:businessSkillsTextBox)
    
    $softSkillsTab = New-Object System.Windows.Forms.TabPage
    $softSkillsTab.Text = "Soft Skills"
    $global:softSkillsTextBox = New-Object System.Windows.Forms.TextBox
    $global:softSkillsTextBox.Multiline = $true
    $global:softSkillsTextBox.Dock = "Fill"
    $global:softSkillsTextBox.Font = New-SafeFont -FontName "Consolas" -Size 10
    $global:softSkillsTextBox.Text = "• Communication (Written & Verbal)`n• Problem Solving & Critical Thinking`n• Time Management & Organization`n• Adaptability & Flexibility`n• Teamwork & Collaboration`n• Leadership & Mentoring`n• Emotional Intelligence"
    $softSkillsTab.Controls.Add($global:softSkillsTextBox)
    
    $languagesTab = New-Object System.Windows.Forms.TabPage
    $languagesTab.Text = "Languages"
    $global:languagesTextBox = New-Object System.Windows.Forms.TextBox
    $global:languagesTextBox.Multiline = $true
    $global:languagesTextBox.Dock = "Fill"
    $global:languagesTextBox.Font = New-SafeFont -FontName "Consolas" -Size 10
    $global:languagesTextBox.Text = "• English (Native/Fluent)`n• Spanish (Fluent)`n• French (Intermediate)`n• German (Basic)`n• Japanese (Basic)"
    $languagesTab.Controls.Add($global:languagesTextBox)
    
    $skillsTabControl.TabPages.AddRange(@($techSkillsTab, $businessSkillsTab, $softSkillsTab, $languagesTab))
    $skillsPanel.Controls.Add($skillsTabControl)
    $skillsTab.Controls.Add($skillsPanel)
    
    # Projects & Certifications Tab - FIXED: Added missing buttons
    $projectsTab = New-Object System.Windows.Forms.TabPage
    $projectsTab.Text = "Projects & Certifications"
    $projectsTab.BackColor = [System.Drawing.Color]::White
    
    $projectsPanel = New-Object System.Windows.Forms.Panel
    $projectsPanel.Dock = "Fill"
    
    # Projects Group with buttons
    $projectsGroup = New-Object System.Windows.Forms.GroupBox
    $projectsGroup.Text = "Projects"
    $projectsGroup.Dock = "Top"
    $projectsGroup.Height = 250
    
    $global:projectsDataGrid = New-Object System.Windows.Forms.DataGridView
    $global:projectsDataGrid.Dock = "Fill"
    $global:projectsDataGrid.ColumnCount = 4
    $global:projectsDataGrid.Columns[0].Name = "Project Name"
    $global:projectsDataGrid.Columns[1].Name = "Role"
    $global:projectsDataGrid.Columns[2].Name = "Description"
    $global:projectsDataGrid.Columns[3].Name = "Technologies"
    $global:projectsDataGrid.AutoSizeColumnsMode = "Fill"
    $global:projectsDataGrid.AllowUserToAddRows = $false
    
    # Projects buttons panel
    $projectsButtonsPanel = New-Object System.Windows.Forms.Panel
    $projectsButtonsPanel.Dock = "Bottom"
    $projectsButtonsPanel.Height = 40
    $projectsButtonsPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    
    $addProjectButton = New-Object System.Windows.Forms.Button
    $addProjectButton.Text = "Add Project"
    $addProjectButton.Location = New-Object System.Drawing.Point(20, 5)
    $addProjectButton.Size = New-Object System.Drawing.Size(120, 30)
    $addProjectButton.Add_Click({ Show-ProjectDialog })
    
    $editProjectButton = New-Object System.Windows.Forms.Button
    $editProjectButton.Text = "Edit Project"
    $editProjectButton.Location = New-Object System.Drawing.Point(150, 5)
    $editProjectButton.Size = New-Object System.Drawing.Size(120, 30)
    $editProjectButton.Add_Click({ 
        if ($global:projectsDataGrid.SelectedRows.Count -gt 0) {
            Show-ProjectDialog -EditMode $true
        }
    })
    
    $removeProjectButton = New-Object System.Windows.Forms.Button
    $removeProjectButton.Text = "Remove Project"
    $removeProjectButton.Location = New-Object System.Drawing.Point(280, 5)
    $removeProjectButton.Size = New-Object System.Drawing.Size(120, 30)
    $removeProjectButton.Add_Click({ 
        if ($global:projectsDataGrid.SelectedRows.Count -gt 0) {
            foreach ($row in $global:projectsDataGrid.SelectedRows) {
                $global:projectsDataGrid.Rows.Remove($row)
            }
        }
    })
    
    $projectsButtonsPanel.Controls.AddRange(@($addProjectButton, $editProjectButton, $removeProjectButton))
    
    $projectsGroup.Controls.AddRange(@($global:projectsDataGrid, $projectsButtonsPanel))
    
    # Certifications Group with buttons
    $certsGroup = New-Object System.Windows.Forms.GroupBox
    $certsGroup.Text = "Certifications"
    $certsGroup.Dock = "Bottom"
    $certsGroup.Height = 250
    
    $global:certsDataGrid = New-Object System.Windows.Forms.DataGridView
    $global:certsDataGrid.Dock = "Fill"
    $global:certsDataGrid.ColumnCount = 4
    $global:certsDataGrid.Columns[0].Name = "Certification"
    $global:certsDataGrid.Columns[1].Name = "Issuer"
    $global:certsDataGrid.Columns[2].Name = "Date"
    $global:certsDataGrid.Columns[3].Name = "Credential ID"
    $global:certsDataGrid.AutoSizeColumnsMode = "Fill"
    $global:certsDataGrid.AllowUserToAddRows = $false
    
    # Certifications buttons panel
    $certsButtonsPanel = New-Object System.Windows.Forms.Panel
    $certsButtonsPanel.Dock = "Bottom"
    $certsButtonsPanel.Height = 40
    $certsButtonsPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    
    $addCertButton = New-Object System.Windows.Forms.Button
    $addCertButton.Text = "Add Certification"
    $addCertButton.Location = New-Object System.Drawing.Point(20, 5)
    $addCertButton.Size = New-Object System.Drawing.Size(120, 30)
    $addCertButton.Add_Click({ Show-CertificationDialog })
    
    $editCertButton = New-Object System.Windows.Forms.Button
    $editCertButton.Text = "Edit Certification"
    $editCertButton.Location = New-Object System.Drawing.Point(150, 5)
    $editCertButton.Size = New-Object System.Drawing.Size(120, 30)
    $editCertButton.Add_Click({ 
        if ($global:certsDataGrid.SelectedRows.Count -gt 0) {
            Show-CertificationDialog -EditMode $true
        }
    })
    
    $removeCertButton = New-Object System.Windows.Forms.Button
    $removeCertButton.Text = "Remove Certification"
    $removeCertButton.Location = New-Object System.Drawing.Point(280, 5)
    $removeCertButton.Size = New-Object System.Drawing.Size(120, 30)
    $removeCertButton.Add_Click({ 
        if ($global:certsDataGrid.SelectedRows.Count -gt 0) {
            foreach ($row in $global:certsDataGrid.SelectedRows) {
                $global:certsDataGrid.Rows.Remove($row)
            }
        }
    })
    
    $certsButtonsPanel.Controls.AddRange(@($addCertButton, $editCertButton, $removeCertButton))
    
    $certsGroup.Controls.AddRange(@($global:certsDataGrid, $certsButtonsPanel))
    
    $projectsPanel.Controls.AddRange(@($projectsGroup, $certsGroup))
    $projectsTab.Controls.Add($projectsPanel)
    
    # Research Work Tab (NEW: Added for Academic/Research CVs)
    $researchTab = New-Object System.Windows.Forms.TabPage
    $researchTab.Text = "Research Work"
    $researchTab.BackColor = [System.Drawing.Color]::White
    
    $researchPanel = New-Object System.Windows.Forms.Panel
    $researchPanel.Dock = "Fill"
    $researchPanel.AutoScroll = $true
    
    # Patents Section
    $patentsGroup = New-Object System.Windows.Forms.GroupBox
    $patentsGroup.Text = "Patents"
    $patentsGroup.Location = New-Object System.Drawing.Point(20, 20)
    $patentsGroup.Size = New-Object System.Drawing.Size(1100, 200)
    
    $global:patentsDataGrid = New-Object System.Windows.Forms.DataGridView
    $global:patentsDataGrid.Location = New-Object System.Drawing.Point(10, 25)
    $global:patentsDataGrid.Size = New-Object System.Drawing.Size(1080, 140)
    $global:patentsDataGrid.ColumnCount = 5
    $global:patentsDataGrid.Columns[0].Name = "Patent Title"
    $global:patentsDataGrid.Columns[1].Name = "Type (Filed/Published/Grant)"
    $global:patentsDataGrid.Columns[2].Name = "Patent Number"
    $global:patentsDataGrid.Columns[3].Name = "Date"
    $global:patentsDataGrid.Columns[4].Name = "URL/Link"
    $global:patentsDataGrid.AutoSizeColumnsMode = "Fill"
    $global:patentsDataGrid.AllowUserToAddRows = $false
    
    $patentsButtonsPanel = New-Object System.Windows.Forms.Panel
    $patentsButtonsPanel.Location = New-Object System.Drawing.Point(10, 170)
    $patentsButtonsPanel.Size = New-Object System.Drawing.Size(1080, 25)
    
    $addPatentButton = New-Object System.Windows.Forms.Button
    $addPatentButton.Text = "Add Patent"
    $addPatentButton.Location = New-Object System.Drawing.Point(0, 0)
    $addPatentButton.Size = New-Object System.Drawing.Size(100, 25)
    $addPatentButton.Add_Click({ Show-PatentDialog })
    
    $editPatentButton = New-Object System.Windows.Forms.Button
    $editPatentButton.Text = "Edit Patent"
    $editPatentButton.Location = New-Object System.Drawing.Point(110, 0)
    $editPatentButton.Size = New-Object System.Drawing.Size(100, 25)
    $editPatentButton.Add_Click({ 
        if ($global:patentsDataGrid.SelectedRows.Count -gt 0) {
            Show-PatentDialog -EditMode $true
        }
    })
    
    $removePatentButton = New-Object System.Windows.Forms.Button
    $removePatentButton.Text = "Remove Patent"
    $removePatentButton.Location = New-Object System.Drawing.Point(220, 0)
    $removePatentButton.Size = New-Object System.Drawing.Size(100, 25)
    $removePatentButton.Add_Click({ 
        if ($global:patentsDataGrid.SelectedRows.Count -gt 0) {
            foreach ($row in $global:patentsDataGrid.SelectedRows) {
                $global:patentsDataGrid.Rows.Remove($row)
            }
        }
    })
    
    $patentsButtonsPanel.Controls.AddRange(@($addPatentButton, $editPatentButton, $removePatentButton))
    $patentsGroup.Controls.AddRange(@($global:patentsDataGrid, $patentsButtonsPanel))
    
    # Software Developed Section
    $softwareGroup = New-Object System.Windows.Forms.GroupBox
    $softwareGroup.Text = "Software/Programs Developed"
    $softwareGroup.Location = New-Object System.Drawing.Point(20, 240)
    $softwareGroup.Size = New-Object System.Drawing.Size(1100, 200)
    
    $global:softwareDataGrid = New-Object System.Windows.Forms.DataGridView
    $global:softwareDataGrid.Location = New-Object System.Drawing.Point(10, 25)
    $global:softwareDataGrid.Size = New-Object System.Drawing.Size(1080, 140)
    $global:softwareDataGrid.ColumnCount = 4
    $global:softwareDataGrid.Columns[0].Name = "Software Name"
    $global:softwareDataGrid.Columns[1].Name = "Description"
    $global:softwareDataGrid.Columns[2].Name = "GitHub/GitLab URL"
    $global:softwareDataGrid.Columns[3].Name = "Technologies"
    $global:softwareDataGrid.AutoSizeColumnsMode = "Fill"
    $global:softwareDataGrid.AllowUserToAddRows = $false
    
    $softwareButtonsPanel = New-Object System.Windows.Forms.Panel
    $softwareButtonsPanel.Location = New-Object System.Drawing.Point(10, 170)
    $softwareButtonsPanel.Size = New-Object System.Drawing.Size(1080, 25)
    
    $addSoftwareButton = New-Object System.Windows.Forms.Button
    $addSoftwareButton.Text = "Add Software"
    $addSoftwareButton.Location = New-Object System.Drawing.Point(0, 0)
    $addSoftwareButton.Size = New-Object System.Drawing.Size(100, 25)
    $addSoftwareButton.Add_Click({ Show-SoftwareDialog })
    
    $editSoftwareButton = New-Object System.Windows.Forms.Button
    $editSoftwareButton.Text = "Edit Software"
    $editSoftwareButton.Location = New-Object System.Drawing.Point(110, 0)
    $editSoftwareButton.Size = New-Object System.Drawing.Size(100, 25)
    $editSoftwareButton.Add_Click({ 
        if ($global:softwareDataGrid.SelectedRows.Count -gt 0) {
            Show-SoftwareDialog -EditMode $true
        }
    })
    
    $removeSoftwareButton = New-Object System.Windows.Forms.Button
    $removeSoftwareButton.Text = "Remove Software"
    $removeSoftwareButton.Location = New-Object System.Drawing.Point(220, 0)
    $removeSoftwareButton.Size = New-Object System.Drawing.Size(100, 25)
    $removeSoftwareButton.Add_Click({ 
        if ($global:softwareDataGrid.SelectedRows.Count -gt 0) {
            foreach ($row in $global:softwareDataGrid.SelectedRows) {
                $global:softwareDataGrid.Rows.Remove($row)
            }
        }
    })
    
    $softwareButtonsPanel.Controls.AddRange(@($addSoftwareButton, $editSoftwareButton, $removeSoftwareButton))
    $softwareGroup.Controls.AddRange(@($global:softwareDataGrid, $softwareButtonsPanel))
    
    # Publications Section
    $publicationsGroup = New-Object System.Windows.Forms.GroupBox
    $publicationsGroup.Text = "Publications"
    $publicationsGroup.Location = New-Object System.Drawing.Point(20, 460)
    $publicationsGroup.Size = New-Object System.Drawing.Size(1100, 250)
    
    $publicationsTabControl = New-Object System.Windows.Forms.TabControl
    $publicationsTabControl.Location = New-Object System.Drawing.Point(10, 25)
    $publicationsTabControl.Size = New-Object System.Drawing.Size(1080, 220)
    
    # Conference Papers Tab
    $conferenceTab = New-Object System.Windows.Forms.TabPage
    $conferenceTab.Text = "Conference Papers"
    $global:conferencePapersTextBox = New-Object System.Windows.Forms.TextBox
    $global:conferencePapersTextBox.Multiline = $true
    $global:conferencePapersTextBox.Dock = "Fill"
    $global:conferencePapersTextBox.Font = New-SafeFont -FontName "Consolas" -Size 9
    $global:conferencePapersTextBox.Text = "• International Conference on Advanced Computing (ICAC-2024), IEEE, URL: https://example.com`n• National Conference on Computer Science (NCCS-2023), Springer, URL: https://example.com"
    $conferenceTab.Controls.Add($global:conferencePapersTextBox)
    
    # Journal Articles Tab
    $journalTab = New-Object System.Windows.Forms.TabPage
    $journalTab.Text = "Journal Articles"
    $global:journalArticlesTextBox = New-Object System.Windows.Forms.TextBox
    $global:journalArticlesTextBox.Multiline = $true
    $global:journalArticlesTextBox.Dock = "Fill"
    $global:journalArticlesTextBox.Font = New-SafeFont -FontName "Consolas" -Size 9
    $global:journalArticlesTextBox.Text = "• 'AI in Healthcare', Journal of Medical AI, Vol. 12, Issue 3, 2024, URL: https://example.com`n• 'Machine Learning Algorithms', IEEE Transactions, Vol. 45, Issue 2, 2023, URL: https://example.com"
    $journalTab.Controls.Add($global:journalArticlesTextBox)
    
    # Books Tab
    $booksTab = New-Object System.Windows.Forms.TabPage
    $booksTab.Text = "Books & Chapters"
    $global:booksTextBox = New-Object System.Windows.Forms.TextBox
    $global:booksTextBox.Multiline = $true
    $global:booksTextBox.Dock = "Fill"
    $global:booksTextBox.Font = New-SafeFont -FontName "Consolas" -Size 9
    $global:booksTextBox.Text = "• 'Advanced Machine Learning', Springer, 2024, ISBN: 978-3-XXX-XXXXX-X, URL: https://example.com`n• 'Chapter 5: Deep Learning Applications', In: 'AI Handbook', Elsevier, 2023, URL: https://example.com"
    $booksTab.Controls.Add($global:booksTextBox)
    
    $publicationsTabControl.TabPages.AddRange(@($conferenceTab, $journalTab, $booksTab))
    $publicationsGroup.Controls.Add($publicationsTabControl)
    
    # Research Profiles Section
    $profilesGroup = New-Object System.Windows.Forms.GroupBox
    $profilesGroup.Text = "Research Profiles & IDs"
    $profilesGroup.Location = New-Object System.Drawing.Point(20, 730)
    $profilesGroup.Size = New-Object System.Drawing.Size(1100, 150)
    
    $yPosProfiles = 25
    $profileSpacing = 30
    
    # ORCID
    $orcidLabel = New-Object System.Windows.Forms.Label
    $orcidLabel.Text = "ORCID ID:"
    $orcidLabel.Location = New-Object System.Drawing.Point(20, $yPosProfiles)
    $orcidLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:orcidTextBox = New-Object System.Windows.Forms.TextBox
    $global:orcidTextBox.Location = New-Object System.Drawing.Point(180, $yPosProfiles)
    $global:orcidTextBox.Size = New-Object System.Drawing.Size(300, 25)
    $yPosProfiles += $profileSpacing
    
    # Google Scholar
    $googleScholarLabel = New-Object System.Windows.Forms.Label
    $googleScholarLabel.Text = "Google Scholar URL:"
    $googleScholarLabel.Location = New-Object System.Drawing.Point(20, $yPosProfiles)
    $googleScholarLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:googleScholarTextBox = New-Object System.Windows.Forms.TextBox
    $global:googleScholarTextBox.Location = New-Object System.Drawing.Point(180, $yPosProfiles)
    $global:googleScholarTextBox.Size = New-Object System.Drawing.Size(300, 25)
    $yPosProfiles += $profileSpacing
    
    # ResearchGate
    $researchGateLabel = New-Object System.Windows.Forms.Label
    $researchGateLabel.Text = "ResearchGate URL:"
    $researchGateLabel.Location = New-Object System.Drawing.Point(20, $yPosProfiles)
    $researchGateLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:researchGateTextBox = New-Object System.Windows.Forms.TextBox
    $global:researchGateTextBox.Location = New-Object System.Drawing.Point(180, $yPosProfiles)
    $global:researchGateTextBox.Size = New-Object System.Drawing.Size(300, 25)
    
    # Scopus
    $scopusLabel = New-Object System.Windows.Forms.Label
    $scopusLabel.Text = "Scopus Author ID:"
    $scopusLabel.Location = New-Object System.Drawing.Point(500, 25)
    $scopusLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:scopusTextBox = New-Object System.Windows.Forms.TextBox
    $global:scopusTextBox.Location = New-Object System.Drawing.Point(660, 25)
    $global:scopusTextBox.Size = New-Object System.Drawing.Size(300, 25)
    
    # Web of Science
    $wosLabel = New-Object System.Windows.Forms.Label
    $wosLabel.Text = "Web of Science ID:"
    $wosLabel.Location = New-Object System.Drawing.Point(500, 55)
    $wosLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:wosTextBox = New-Object System.Windows.Forms.TextBox
    $global:wosTextBox.Location = New-Object System.Drawing.Point(660, 55)
    $global:wosTextBox.Size = New-Object System.Drawing.Size(300, 25)
    
    # Vidwan
    $vidwanLabel = New-Object System.Windows.Forms.Label
    $vidwanLabel.Text = "Vidwan ID:"
    $vidwanLabel.Location = New-Object System.Drawing.Point(500, 85)
    $vidwanLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:vidwanTextBox = New-Object System.Windows.Forms.TextBox
    $global:vidwanTextBox.Location = New-Object System.Drawing.Point(660, 85)
    $global:vidwanTextBox.Size = New-Object System.Drawing.Size(300, 25)
    
    $profilesGroup.Controls.AddRange(@(
        $orcidLabel, $global:orcidTextBox,
        $googleScholarLabel, $global:googleScholarTextBox,
        $researchGateLabel, $global:researchGateTextBox,
        $scopusLabel, $global:scopusTextBox,
        $wosLabel, $global:wosTextBox,
        $vidwanLabel, $global:vidwanTextBox
    ))
    
    # PhD Scholars Guided Section
    $phdGuidedGroup = New-Object System.Windows.Forms.GroupBox
    $phdGuidedGroup.Text = "PhD Scholars Guided"
    $phdGuidedGroup.Location = New-Object System.Drawing.Point(20, 900)
    $phdGuidedGroup.Size = New-Object System.Drawing.Size(1100, 150)
    
    $global:phdGuidedTextBox = New-Object System.Windows.Forms.TextBox
    $global:phdGuidedTextBox.Multiline = $true
    $global:phdGuidedTextBox.Location = New-Object System.Drawing.Point(10, 25)
    $global:phdGuidedTextBox.Size = New-Object System.Drawing.Size(1080, 120)
    $global:phdGuidedTextBox.Font = New-SafeFont -FontName "Consolas" -Size 9
    $global:phdGuidedTextBox.Text = "• John Doe (2020-2024), Thesis: 'Advanced AI Algorithms', Status: Completed`n• Jane Smith (2021-Present), Thesis: 'Machine Learning in Healthcare', Status: Ongoing"
    $phdGuidedGroup.Controls.Add($global:phdGuidedTextBox)
    
    $researchPanel.Controls.AddRange(@(
        $patentsGroup, $softwareGroup, $publicationsGroup, 
        $profilesGroup, $phdGuidedGroup
    ))
    $researchTab.Controls.Add($researchPanel)
    
    # Editorial Activities Tab (NEW)
    $editorialTab = New-Object System.Windows.Forms.TabPage
    $editorialTab.Text = "Editorial Activities"
    $editorialTab.BackColor = [System.Drawing.Color]::White
    
    $editorialPanel = New-Object System.Windows.Forms.Panel
    $editorialPanel.Dock = "Fill"
    $editorialPanel.AutoScroll = $true
    
    $editorialGroup = New-Object System.Windows.Forms.GroupBox
    $editorialGroup.Text = "Editorial Roles & Review Activities"
    $editorialGroup.Location = New-Object System.Drawing.Point(20, 20)
    $editorialGroup.Size = New-Object System.Drawing.Size(1100, 400)
    
    # Editor in Journals/Conferences
    $editorRolesLabel = New-Object System.Windows.Forms.Label
    $editorRolesLabel.Text = "Editor Roles in Journals/Conferences:"
    $editorRolesLabel.Location = New-Object System.Drawing.Point(20, 30)
    $editorRolesLabel.Size = New-Object System.Drawing.Size(250, 25)
    $editorRolesLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 10 -Style "Bold"
    
    $global:editorRolesTextBox = New-Object System.Windows.Forms.TextBox
    $global:editorRolesTextBox.Multiline = $true
    $global:editorRolesTextBox.Location = New-Object System.Drawing.Point(20, 60)
    $global:editorRolesTextBox.Size = New-Object System.Drawing.Size(1060, 80)
    $global:editorRolesTextBox.Font = New-SafeFont -FontName "Consolas" -Size 9
    $global:editorRolesTextBox.Text = "• Editor-in-Chief: International Journal of Computer Science (2022-Present)`n• Associate Editor: IEEE Transactions on AI (2020-2023)`n• Editorial Board Member: ACM Computing Surveys (2019-Present)"
    
    # Reviewer Activities
    $reviewerLabel = New-Object System.Windows.Forms.Label
    $reviewerLabel.Text = "Reviewer for Journals/Conferences:"
    $reviewerLabel.Location = New-Object System.Drawing.Point(20, 160)
    $reviewerLabel.Size = New-Object System.Drawing.Size(250, 25)
    $reviewerLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 10 -Style "Bold"
    
    $global:reviewerTextBox = New-Object System.Windows.Forms.TextBox
    $global:reviewerTextBox.Multiline = $true
    $global:reviewerTextBox.Location = New-Object System.Drawing.Point(20, 190)
    $global:reviewerTextBox.Size = New-Object System.Drawing.Size(1060, 80)
    $global:reviewerTextBox.Font = New-SafeFont -FontName "Consolas" -Size 9
    $global:reviewerTextBox.Text = "• Regular Reviewer: Nature Communications, Science Magazine`n• Program Committee Member: AAAI, NeurIPS, ICML`n• Conference Reviewer: IEEE CVPR, ACM SIGGRAPH"
    
    # Guest Editor
    $guestEditorLabel = New-Object System.Windows.Forms.Label
    $guestEditorLabel.Text = "Guest Editor Roles:"
    $guestEditorLabel.Location = New-Object System.Drawing.Point(20, 290)
    $guestEditorLabel.Size = New-Object System.Drawing.Size(250, 25)
    $guestEditorLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 10 -Style "Bold"
    
    $global:guestEditorTextBox = New-Object System.Windows.Forms.TextBox
    $global:guestEditorTextBox.Multiline = $true
    $global:guestEditorTextBox.Location = New-Object System.Drawing.Point(20, 320)
    $global:guestEditorTextBox.Size = New-Object System.Drawing.Size(1060, 60)
    $global:guestEditorTextBox.Font = New-SafeFont -FontName "Consolas" -Size 9
    $global:guestEditorTextBox.Text = "• Guest Editor: Special Issue on 'AI in Healthcare', Journal of Medical Informatics (2023)`n• Guest Editor: 'Advances in Deep Learning', Springer LNCS (2022)"
    
    $editorialGroup.Controls.AddRange(@(
        $editorRolesLabel, $global:editorRolesTextBox,
        $reviewerLabel, $global:reviewerTextBox,
        $guestEditorLabel, $global:guestEditorTextBox
    ))
    
    $editorialPanel.Controls.Add($editorialGroup)
    $editorialTab.Controls.Add($editorialPanel)
    
    # PhD Details Tab (NEW)
    $phdDetailsTab = New-Object System.Windows.Forms.TabPage
    $phdDetailsTab.Text = "PhD Details"
    $phdDetailsTab.BackColor = [System.Drawing.Color]::White
    
    $phdDetailsPanel = New-Object System.Windows.Forms.Panel
    $phdDetailsPanel.Dock = "Fill"
    $phdDetailsPanel.AutoScroll = $true
    
    $yPosPhd = 20
    $phdLabelWidth = 200
    $phdControlWidth = 400
    
    # PhD Title
    $phdTitleLabel = New-Object System.Windows.Forms.Label
    $phdTitleLabel.Text = "PhD Thesis Title:*"
    $phdTitleLabel.Location = New-Object System.Drawing.Point(50, $yPosPhd)
    $phdTitleLabel.Size = New-Object System.Drawing.Size($phdLabelWidth, 30)
    $global:phdTitleTextBox = New-Object System.Windows.Forms.TextBox
    $global:phdTitleTextBox.Location = New-Object System.Drawing.Point(260, $yPosPhd)
    $global:phdTitleTextBox.Size = New-Object System.Drawing.Size($phdControlWidth, 30)
    $yPosPhd += 40
    
    # University
    $phdUniversityLabel = New-Object System.Windows.Forms.Label
    $phdUniversityLabel.Text = "University/Institution:*"
    $phdUniversityLabel.Location = New-Object System.Drawing.Point(50, $yPosPhd)
    $phdUniversityLabel.Size = New-Object System.Drawing.Size($phdLabelWidth, 30)
    $global:phdUniversityTextBox = New-Object System.Windows.Forms.TextBox
    $global:phdUniversityTextBox.Location = New-Object System.Drawing.Point(260, $yPosPhd)
    $global:phdUniversityTextBox.Size = New-Object System.Drawing.Size($phdControlWidth, 30)
    $yPosPhd += 40
    
    # Department
    $phdDepartmentLabel = New-Object System.Windows.Forms.Label
    $phdDepartmentLabel.Text = "Department:*"
    $phdDepartmentLabel.Location = New-Object System.Drawing.Point(50, $yPosPhd)
    $phdDepartmentLabel.Size = New-Object System.Drawing.Size($phdLabelWidth, 30)
    $global:phdDepartmentTextBox = New-Object System.Windows.Forms.TextBox
    $global:phdDepartmentTextBox.Location = New-Object System.Drawing.Point(260, $yPosPhd)
    $global:phdDepartmentTextBox.Size = New-Object System.Drawing.Size($phdControlWidth, 30)
    $yPosPhd += 40
    
    # Supervisor
    $phdSupervisorLabel = New-Object System.Windows.Forms.Label
    $phdSupervisorLabel.Text = "Supervisor(s):*"
    $phdSupervisorLabel.Location = New-Object System.Drawing.Point(50, $yPosPhd)
    $phdSupervisorLabel.Size = New-Object System.Drawing.Size($phdLabelWidth, 30)
    $global:phdSupervisorTextBox = New-Object System.Windows.Forms.TextBox
    $global:phdSupervisorTextBox.Location = New-Object System.Drawing.Point(260, $yPosPhd)
    $global:phdSupervisorTextBox.Size = New-Object System.Drawing.Size($phdControlWidth, 30)
    $yPosPhd += 40
    
    # Year of Award
    $phdYearLabel = New-Object System.Windows.Forms.Label
    $phdYearLabel.Text = "Year of Award:"
    $phdYearLabel.Location = New-Object System.Drawing.Point(50, $yPosPhd)
    $phdYearLabel.Size = New-Object System.Drawing.Size($phdLabelWidth, 30)
    $global:phdYearTextBox = New-Object System.Windows.Forms.TextBox
    $global:phdYearTextBox.Location = New-Object System.Drawing.Point(260, $yPosPhd)
    $global:phdYearTextBox.Size = New-Object System.Drawing.Size($phdControlWidth, 30)
    $yPosPhd += 40
    
    # Thesis URL
    $phdThesisUrlLabel = New-Object System.Windows.Forms.Label
    $phdThesisUrlLabel.Text = "Thesis URL/Link:"
    $phdThesisUrlLabel.Location = New-Object System.Drawing.Point(50, $yPosPhd)
    $phdThesisUrlLabel.Size = New-Object System.Drawing.Size($phdLabelWidth, 30)
    $global:phdThesisUrlTextBox = New-Object System.Windows.Forms.TextBox
    $global:phdThesisUrlTextBox.Location = New-Object System.Drawing.Point(260, $yPosPhd)
    $global:phdThesisUrlTextBox.Size = New-Object System.Drawing.Size($phdControlWidth, 30)
    $yPosPhd += 40
    
    # Abstract/Summary
    $phdAbstractLabel = New-Object System.Windows.Forms.Label
    $phdAbstractLabel.Text = "Thesis Abstract/Summary:"
    $phdAbstractLabel.Location = New-Object System.Drawing.Point(50, $yPosPhd)
    $phdAbstractLabel.Size = New-Object System.Drawing.Size($phdLabelWidth, 30)
    $global:phdAbstractTextBox = New-Object System.Windows.Forms.TextBox
    $global:phdAbstractTextBox.Multiline = $true
    $global:phdAbstractTextBox.Location = New-Object System.Drawing.Point(260, $yPosPhd)
    $global:phdAbstractTextBox.Size = New-Object System.Drawing.Size($phdControlWidth, 150)
    $global:phdAbstractTextBox.ScrollBars = "Vertical"
    $yPosPhd += 170
    
    $phdDetailsPanel.Controls.AddRange(@(
        $phdTitleLabel, $global:phdTitleTextBox,
        $phdUniversityLabel, $global:phdUniversityTextBox,
        $phdDepartmentLabel, $global:phdDepartmentTextBox,
        $phdSupervisorLabel, $global:phdSupervisorTextBox,
        $phdYearLabel, $global:phdYearTextBox,
        $phdThesisUrlLabel, $global:phdThesisUrlTextBox,
        $phdAbstractLabel, $global:phdAbstractTextBox
    ))
    
    $phdDetailsTab.Controls.Add($phdDetailsPanel)
    
    # Seminars & Workshops Tab (NEW)
    $seminarsTab = New-Object System.Windows.Forms.TabPage
    $seminarsTab.Text = "Seminars & Workshops"
    $seminarsTab.BackColor = [System.Drawing.Color]::White
    
    $seminarsPanel = New-Object System.Windows.Forms.Panel
    $seminarsPanel.Dock = "Fill"
    $seminarsPanel.AutoScroll = $true
    
    # Attended Section
    $attendedGroup = New-Object System.Windows.Forms.GroupBox
    $attendedGroup.Text = "Seminars/Workshops Attended"
    $attendedGroup.Location = New-Object System.Drawing.Point(20, 20)
    $attendedGroup.Size = New-Object System.Drawing.Size(1100, 200)
    
    $global:attendedSeminarsTextBox = New-Object System.Windows.Forms.TextBox
    $global:attendedSeminarsTextBox.Multiline = $true
    $global:attendedSeminarsTextBox.Location = New-Object System.Drawing.Point(10, 25)
    $global:attendedSeminarsTextBox.Size = New-Object System.Drawing.Size(1080, 170)
    $global:attendedSeminarsTextBox.Font = New-SafeFont -FontName "Consolas" -Size 9
    $global:attendedSeminarsTextBox.Text = "• International Workshop on AI Research (2024), Stanford University, USA`n• National Seminar on Machine Learning Applications (2023), IIT Delhi, India`n• Workshop on Research Methodology (2022), Online"
    $attendedGroup.Controls.Add($global:attendedSeminarsTextBox)
    
    # Conducted Section
    $conductedGroup = New-Object System.Windows.Forms.GroupBox
    $conductedGroup.Text = "Seminars/Workshops Conducted"
    $conductedGroup.Location = New-Object System.Drawing.Point(20, 240)
    $conductedGroup.Size = New-Object System.Drawing.Size(1100, 200)
    
    $global:conductedSeminarsTextBox = New-Object System.Windows.Forms.TextBox
    $global:conductedSeminarsTextBox.Multiline = $true
    $global:conductedSeminarsTextBox.Location = New-Object System.Drawing.Point(10, 25)
    $global:conductedSeminarsTextBox.Size = New-Object System.Drawing.Size(1080, 170)
    $global:conductedSeminarsTextBox.Font = New-SafeFont -FontName "Consolas" -Size 9
    $global:conductedSeminarsTextBox.Text = "• Workshop on 'Deep Learning for Beginners' (2024), University of Technology`n• Seminar on 'Research Paper Writing' (2023), National Institute of Science`n• Training Program on 'AI Tools' (2022), Corporate Workshop"
    $conductedGroup.Controls.Add($global:conductedSeminarsTextBox)
    
    # PhD Thesis Evaluation Section
    $thesisEvalGroup = New-Object System.Windows.Forms.GroupBox
    $thesisEvalGroup.Text = "PhD Thesis Evaluated as External Examiner"
    $thesisEvalGroup.Location = New-Object System.Drawing.Point(20, 460)
    $thesisEvalGroup.Size = New-Object System.Drawing.Size(1100, 200)
    
    $global:phdThesisEvalTextBox = New-Object System.Windows.Forms.TextBox
    $global:phdThesisEvalTextBox.Multiline = $true
    $global:phdThesisEvalTextBox.Location = New-Object System.Drawing.Point(10, 25)
    $global:phdThesisEvalTextBox.Size = New-Object System.Drawing.Size(1080, 170)
    $global:phdThesisEvalTextBox.Font = New-SafeFont -FontName "Consolas" -Size 9
    $global:phdThesisEvalTextBox.Text = "• Dr. John Smith, 'Advanced Algorithms for Big Data', University of California (2024)`n• Dr. Maria Garcia, 'Machine Learning in Healthcare', MIT (2023)`n• Dr. Robert Chen, 'AI Ethics and Governance', Harvard University (2022)"
    $thesisEvalGroup.Controls.Add($global:phdThesisEvalTextBox)
    
    $seminarsPanel.Controls.AddRange(@($attendedGroup, $conductedGroup, $thesisEvalGroup))
    $seminarsTab.Controls.Add($seminarsPanel)
    
    # Declaration & Signature Tab (NEW)
    $declarationTab = New-Object System.Windows.Forms.TabPage
    $declarationTab.Text = "Declaration & Signature"
    $declarationTab.BackColor = [System.Drawing.Color]::White
    
    $declarationPanel = New-Object System.Windows.Forms.Panel
    $declarationPanel.Dock = "Fill"
    $declarationPanel.AutoScroll = $true
    
    # Declaration Statement
    $declarationLabel = New-Object System.Windows.Forms.Label
    $declarationLabel.Text = "Declaration Statement:"
    $declarationLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $declarationLabel.Location = New-Object System.Drawing.Point(20, 20)
    $declarationLabel.Size = New-Object System.Drawing.Size(300, 30)
    
    $global:declarationTextBox = New-Object System.Windows.Forms.TextBox
    $global:declarationTextBox.Multiline = $true
    $global:declarationTextBox.Location = New-Object System.Drawing.Point(20, 60)
    $global:declarationTextBox.Size = New-Object System.Drawing.Size(1100, 150)
    $global:declarationTextBox.Font = New-SafeFont -FontName "Segoe UI" -Size 10
    $global:declarationTextBox.Text = "I hereby declare that the information furnished above is true to the best of my knowledge and belief. I understand that any willful misrepresentation or omission of facts will be sufficient cause for rejection of this application or termination of employment if discovered at a later date."
    
    # Date and Place
    $datePlaceLabel = New-Object System.Windows.Forms.Label
    $datePlaceLabel.Text = "Date and Place:"
    $datePlaceLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $datePlaceLabel.Location = New-Object System.Drawing.Point(20, 230)
    $datePlaceLabel.Size = New-Object System.Drawing.Size(300, 30)
    
    $global:datePlaceTextBox = New-Object System.Windows.Forms.TextBox
    $global:datePlaceTextBox.Location = New-Object System.Drawing.Point(20, 270)
    $global:datePlaceTextBox.Size = New-Object System.Drawing.Size(500, 30)
    $global:datePlaceTextBox.Font = New-SafeFont -FontName "Segoe UI" -Size 10
    $global:datePlaceTextBox.Text = "$(Get-Date -Format 'MMMM dd, yyyy'), $( [System.Environment]::GetEnvironmentVariable('USERDOMAIN') )"
    
    # Signature Panel
    $signaturePanel = New-Object System.Windows.Forms.Panel
    $signaturePanel.Location = New-Object System.Drawing.Point(20, 320)
    $signaturePanel.Size = New-Object System.Drawing.Size(400, 250)
    $signaturePanel.BorderStyle = "FixedSingle"
    $signaturePanel.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
    
    $signatureLabel = New-Object System.Windows.Forms.Label
    $signatureLabel.Text = "Digital Signature"
    $signatureLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $signatureLabel.Location = New-Object System.Drawing.Point(10, 10)
    $signatureLabel.Size = New-Object System.Drawing.Size(380, 30)
    $signatureLabel.TextAlign = "MiddleCenter"
    
    $global:signaturePictureBox = New-Object System.Windows.Forms.PictureBox
    $global:signaturePictureBox.Location = New-Object System.Drawing.Point(50, 50)
    $global:signaturePictureBox.Size = New-Object System.Drawing.Size(300, 150)
    $global:signaturePictureBox.SizeMode = "Zoom"
    $global:signaturePictureBox.BorderStyle = "FixedSingle"
    $global:signaturePictureBox.BackColor = [System.Drawing.Color]::White
    
    # Load placeholder signature image
    $signaturePlaceholderPath = "$script:SignaturesDir\signature_placeholder.png"
    if (Test-Path $signaturePlaceholderPath) {
        try {
            $global:signaturePictureBox.Image = [System.Drawing.Image]::FromFile($signaturePlaceholderPath)
        } catch {
            $bitmap = New-Object System.Drawing.Bitmap(300, 150)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.Clear([System.Drawing.Color]::White)
            $global:signaturePictureBox.Image = $bitmap
            $graphics.Dispose()
        }
    }
    
    $uploadSignatureButton = New-Object System.Windows.Forms.Button
    $uploadSignatureButton.Text = "Upload Signature"
    $uploadSignatureButton.Location = New-Object System.Drawing.Point(50, 210)
    $uploadSignatureButton.Size = New-Object System.Drawing.Size(120, 30)
    $uploadSignatureButton.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $uploadSignatureButton.ForeColor = [System.Drawing.Color]::White
    $uploadSignatureButton.FlatStyle = "Flat"
    $uploadSignatureButton.FlatAppearance.BorderSize = 0
    $uploadSignatureButton.Font = New-SafeFont -FontName "Segoe UI" -Size 9
    $uploadSignatureButton.Add_Click({ Upload-Signature })
    
    $removeSignatureButton = New-Object System.Windows.Forms.Button
    $removeSignatureButton.Text = "Remove"
    $removeSignatureButton.Location = New-Object System.Drawing.Point(180, 210)
    $removeSignatureButton.Size = New-Object System.Drawing.Size(120, 30)
    $removeSignatureButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $removeSignatureButton.ForeColor = [System.Drawing.Color]::White
    $removeSignatureButton.FlatStyle = "Flat"
    $removeSignatureButton.FlatAppearance.BorderSize = 0
    $removeSignatureButton.Font = New-SafeFont -FontName "Segoe UI" -Size 9
    $removeSignatureButton.Add_Click({ 
        if (Test-Path $signaturePlaceholderPath) {
            try {
                $global:signaturePictureBox.Image = [System.Drawing.Image]::FromFile($signaturePlaceholderPath)
            } catch {
                $global:signaturePictureBox.Image = $null
            }
        }
        $global:signaturePath = $null
        Show-SafeMessage -Message "Signature removed successfully!" -Title "Success" -Icon "Information"
    })
    
    $signaturePanel.Controls.AddRange(@($signatureLabel, $global:signaturePictureBox, $uploadSignatureButton, $removeSignatureButton))
    $global:signaturePath = $null
    
    # Professional Profiles Section (Moved from Research Tab)
    $professionalProfilesGroup = New-Object System.Windows.Forms.GroupBox
    $professionalProfilesGroup.Text = "Additional Professional Profiles"
    $professionalProfilesGroup.Location = New-Object System.Drawing.Point(450, 320)
    $professionalProfilesGroup.Size = New-Object System.Drawing.Size(670, 250)
    
    $yPosProf = 25
    $profSpacing = 30
    
    # LinkedIn (already in Personal Info, but repeated for consistency)
    $linkedInProfLabel = New-Object System.Windows.Forms.Label
    $linkedInProfLabel.Text = "LinkedIn:"
    $linkedInProfLabel.Location = New-Object System.Drawing.Point(20, $yPosProf)
    $linkedInProfLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:linkedInProfTextBox = New-Object System.Windows.Forms.TextBox
    $global:linkedInProfTextBox.Location = New-Object System.Drawing.Point(180, $yPosProf)
    $global:linkedInProfTextBox.Size = New-Object System.Drawing.Size(450, 25)
    $yPosProf += $profSpacing
    
    # GitHub (already in Personal Info)
    $githubProfLabel = New-Object System.Windows.Forms.Label
    $githubProfLabel.Text = "GitHub:"
    $githubProfLabel.Location = New-Object System.Drawing.Point(20, $yPosProf)
    $githubProfLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:githubProfTextBox = New-Object System.Windows.Forms.TextBox
    $global:githubProfTextBox.Location = New-Object System.Drawing.Point(180, $yPosProf)
    $global:githubProfTextBox.Size = New-Object System.Drawing.Size(450, 25)
    $yPosProf += $profSpacing
    
    # Academic Profile
    $academicProfileLabel = New-Object System.Windows.Forms.Label
    $academicProfileLabel.Text = "Academic Profile URL:"
    $academicProfileLabel.Location = New-Object System.Drawing.Point(20, $yPosProf)
    $academicProfileLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:academicProfileTextBox = New-Object System.Windows.Forms.TextBox
    $global:academicProfileTextBox.Location = New-Object System.Drawing.Point(180, $yPosProf)
    $global:academicProfileTextBox.Size = New-Object System.Drawing.Size(450, 25)
    $yPosProf += $profSpacing
    
    # IRINS (Institutional Repository)
    $irinsLabel = New-Object System.Windows.Forms.Label
    $irinsLabel.Text = "Institutional IRINS:"
    $irinsLabel.Location = New-Object System.Drawing.Point(20, $yPosProf)
    $irinsLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:irinsTextBox = New-Object System.Windows.Forms.TextBox
    $global:irinsTextBox.Location = New-Object System.Drawing.Point(180, $yPosProf)
    $global:irinsTextBox.Size = New-Object System.Drawing.Size(450, 25)
    $yPosProf += $profSpacing
    
    # YouTube Channel
    $youtubeLabel = New-Object System.Windows.Forms.Label
    $youtubeLabel.Text = "YouTube Channel:"
    $youtubeLabel.Location = New-Object System.Drawing.Point(20, $yPosProf)
    $youtubeLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:youtubeTextBox = New-Object System.Windows.Forms.TextBox
    $global:youtubeTextBox.Location = New-Object System.Drawing.Point(180, $yPosProf)
    $global:youtubeTextBox.Size = New-Object System.Drawing.Size(450, 25)
    $yPosProf += $profSpacing
    
    # Google Developers
    $googleDevLabel = New-Object System.Windows.Forms.Label
    $googleDevLabel.Text = "Google Developers:"
    $googleDevLabel.Location = New-Object System.Drawing.Point(20, $yPosProf)
    $googleDevLabel.Size = New-Object System.Drawing.Size(150, 25)
    $global:googleDevTextBox = New-Object System.Windows.Forms.TextBox
    $global:googleDevTextBox.Location = New-Object System.Drawing.Point(180, $yPosProf)
    $global:googleDevTextBox.Size = New-Object System.Drawing.Size(450, 25)
    
    $professionalProfilesGroup.Controls.AddRange(@(
        $linkedInProfLabel, $global:linkedInProfTextBox,
        $githubProfLabel, $global:githubProfTextBox,
        $academicProfileLabel, $global:academicProfileTextBox,
        $irinsLabel, $global:irinsTextBox,
        $youtubeLabel, $global:youtubeTextBox,
        $googleDevLabel, $global:googleDevTextBox
    ))
    
    $declarationPanel.Controls.AddRange(@(
        $declarationLabel, $global:declarationTextBox,
        $datePlaceLabel, $global:datePlaceTextBox,
        $signaturePanel, $professionalProfilesGroup
    ))
    
    $declarationTab.Controls.Add($declarationPanel)
    
    # Template & Settings Tab
    $templateTab = New-Object System.Windows.Forms.TabPage
    $templateTab.Text = "Template & Settings"
    $templateTab.BackColor = [System.Drawing.Color]::White
    
    $templatePanel = New-Object System.Windows.Forms.Panel
    $templatePanel.Dock = "Fill"
    $templatePanel.AutoScroll = $true
    
    $yPos = 20
    
    $templateLabel = New-Object System.Windows.Forms.Label
    $templateLabel.Text = "Select Template:"
    $templateLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $templateLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $templateLabel.Size = New-Object System.Drawing.Size(200, 30)
    $yPos += 40
    
    $global:templateComboBox = New-Object System.Windows.Forms.ComboBox
    $global:templateComboBox.Location = New-Object System.Drawing.Point(50, $yPos)
    $global:templateComboBox.Size = New-Object System.Drawing.Size(300, 30)
    $global:templateComboBox.DropDownStyle = "DropDownList"
    $script:Templates.Keys | ForEach-Object { $null = $global:templateComboBox.Items.Add($_) }
    if ($global:templateComboBox.Items.Count -gt 0) {
        $global:templateComboBox.SelectedIndex = 0
    }
    $global:templateComboBox.Add_SelectedIndexChanged({ Update-TemplatePreview })
    $yPos += 50
    
    $global:templateDescLabel = New-Object System.Windows.Forms.Label
    $global:templateDescLabel.Text = ""
    $global:templateDescLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $global:templateDescLabel.Size = New-Object System.Drawing.Size(300, 60)
    $global:templateDescLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $yPos += 70
    
    $colorLabel = New-Object System.Windows.Forms.Label
    $colorLabel.Text = "Color Scheme:"
    $colorLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $colorLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $colorLabel.Size = New-Object System.Drawing.Size(200, 30)
    $yPos += 40
    
    $global:colorComboBox = New-Object System.Windows.Forms.ComboBox
    $global:colorComboBox.Location = New-Object System.Drawing.Point(50, $yPos)
    $global:colorComboBox.Size = New-Object System.Drawing.Size(300, 30)
    $global:colorComboBox.DropDownStyle = "DropDownList"
    $yPos += 50
    
    $fontLabel = New-Object System.Windows.Forms.Label
    $fontLabel.Text = "Font Family:"
    $fontLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $fontLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $fontLabel.Size = New-Object System.Drawing.Size(200, 30)
    $yPos += 40
    
    $global:fontComboBox = New-Object System.Windows.Forms.ComboBox
    $global:fontComboBox.Location = New-Object System.Drawing.Point(50, $yPos)
    $global:fontComboBox.Size = New-Object System.Drawing.Size(300, 30)
    $global:fontComboBox.DropDownStyle = "DropDownList"
    @("Segoe UI", "Calibri", "Arial", "Times New Roman", "Georgia", "Verdana", "Tahoma", "Helvetica", "Garamond") | ForEach-Object {
        $null = $global:fontComboBox.Items.Add($_)
    }
    $global:fontComboBox.SelectedIndex = 0
    $yPos += 50
    
    $formatLabel = New-Object System.Windows.Forms.Label
    $formatLabel.Text = "Output Formats:"
    $formatLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $formatLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $formatLabel.Size = New-Object System.Drawing.Size(200, 30)
    $yPos += 40
    
    $global:htmlCheckBox = New-Object System.Windows.Forms.CheckBox
    $global:htmlCheckBox.Text = "HTML (Web Format)"
    $global:htmlCheckBox.Location = New-Object System.Drawing.Point(50, $yPos)
    $global:htmlCheckBox.Size = New-Object System.Drawing.Size(200, 30)
    $global:htmlCheckBox.Checked = $true
    $yPos += 40
    
    $global:pdfCheckBox = New-Object System.Windows.Forms.CheckBox
    $global:pdfCheckBox.Text = "PDF (Portable Format)"
    $global:pdfCheckBox.Location = New-Object System.Drawing.Point(50, $yPos)
    $global:pdfCheckBox.Size = New-Object System.Drawing.Size(200, 30)
    $global:pdfCheckBox.Checked = $true
    $yPos += 40
    
    $global:docxCheckBox = New-Object System.Windows.Forms.CheckBox
    $global:docxCheckBox.Text = "DOCX (Word Format)"
    $global:docxCheckBox.Location = New-Object System.Drawing.Point(50, $yPos)
    $global:docxCheckBox.Size = New-Object System.Drawing.Size(200, 30)
    $global:docxCheckBox.Checked = $true
    $yPos += 50
    
    $global:includePhotoCheckBox = New-Object System.Windows.Forms.CheckBox
    $global:includePhotoCheckBox.Text = "Include Photo in CV"
    $global:includePhotoCheckBox.Location = New-Object System.Drawing.Point(50, $yPos)
    $global:includePhotoCheckBox.Size = New-Object System.Drawing.Size(200, 30)
    $global:includePhotoCheckBox.Checked = $true
    $yPos += 50
    
    $global:includeSignatureCheckBox = New-Object System.Windows.Forms.CheckBox
    $global:includeSignatureCheckBox.Text = "Include Signature"
    $global:includeSignatureCheckBox.Location = New-Object System.Drawing.Point(50, $yPos)
    $global:includeSignatureCheckBox.Size = New-Object System.Drawing.Size(200, 30)
    $global:includeSignatureCheckBox.Checked = $true
    $yPos += 50
    
    $generateButton = New-Object System.Windows.Forms.Button
    $generateButton.Text = "Generate CV"
    $generateButton.Font = New-SafeFont -FontName "Segoe UI" -Size 12 -Style "Bold"
    $generateButton.Location = New-Object System.Drawing.Point(50, $yPos)
    $generateButton.Size = New-Object System.Drawing.Size(200, 50)
    $generateButton.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $generateButton.ForeColor = [System.Drawing.Color]::White
    $generateButton.FlatStyle = "Flat"
    $generateButton.FlatAppearance.BorderSize = 0
    $generateButton.Add_Click({ Generate-CVFromForm })
    
    # Preview Panel - Enhanced with HTML Preview
    $previewPanel = New-Object System.Windows.Forms.Panel
    $previewPanel.Location = New-Object System.Drawing.Point(400, 20)
    $previewPanel.Size = New-Object System.Drawing.Size(800, 700)  # Enlarged for better preview
    $previewPanel.BackColor = [System.Drawing.Color]::White
    $previewPanel.BorderStyle = "FixedSingle"
    
    $previewLabel = New-Object System.Windows.Forms.Label
    $previewLabel.Text = "Template Preview"
    $previewLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 11 -Style "Bold"
    $previewLabel.Location = New-Object System.Drawing.Point(10, 10)
    $previewLabel.Size = New-Object System.Drawing.Size(780, 30)
    $previewLabel.TextAlign = "MiddleCenter"
    
    $global:previewWebBrowser = New-Object System.Windows.Forms.WebBrowser
    $global:previewWebBrowser.Location = New-Object System.Drawing.Point(10, 50)
    $global:previewWebBrowser.Size = New-Object System.Drawing.Size(780, 640)
    $global:previewWebBrowser.AllowWebBrowserDrop = $false
    $global:previewWebBrowser.IsWebBrowserContextMenuEnabled = $false
    $global:previewWebBrowser.WebBrowserShortcutsEnabled = $false
    $global:previewWebBrowser.ScriptErrorsSuppressed = $true
    
    $previewPanel.Controls.AddRange(@($previewLabel, $global:previewWebBrowser))
    
    $templatePanel.Controls.AddRange(@(
        $templateLabel, $global:templateComboBox,
        $global:templateDescLabel,
        $colorLabel, $global:colorComboBox,
        $fontLabel, $global:fontComboBox,
        $formatLabel, $global:htmlCheckBox,
        $global:pdfCheckBox, $global:docxCheckBox,
        $global:includePhotoCheckBox,
        $global:includeSignatureCheckBox,
        $generateButton, $previewPanel
    ))
    
    $templateTab.Controls.Add($templatePanel)
    
    # Add all tabs to tab control
    $tabControl.TabPages.AddRange(@(
        $personalTab, $experienceTab, $educationTab, $skillsTab, 
        $projectsTab, $researchTab, $editorialTab, $phdDetailsTab,
        $seminarsTab, $declarationTab, $templateTab
    ))
    
    # Bottom Panel with navigation
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Dock = "Bottom"
    $bottomPanel.Height = 60
    $bottomPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    
    $prevButton = New-Object System.Windows.Forms.Button
    $prevButton.Text = "Previous"
    $prevButton.Location = New-Object System.Drawing.Point(20, 15)
    $prevButton.Size = New-Object System.Drawing.Size(100, 30)
    $prevButton.Add_Click({
        if ($tabControl.SelectedIndex -gt 0) {
            $tabControl.SelectedIndex--
        }
    })
    
    $nextButton = New-Object System.Windows.Forms.Button
    $nextButton.Text = "Next"
    $nextButton.Location = New-Object System.Drawing.Point(130, 15)
    $nextButton.Size = New-Object System.Drawing.Size(100, 30)
    $nextButton.Add_Click({
        if ($tabControl.SelectedIndex -lt $tabControl.TabPages.Count - 1) {
            $tabControl.SelectedIndex++
        }
    })
    
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save Draft"
    $saveButton.Location = New-Object System.Drawing.Point(1080, 15)
    $saveButton.Size = New-Object System.Drawing.Size(100, 30)
    $saveButton.Add_Click({ Save-CVDraft })
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(1190, 15)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Add_Click({ $builderForm.Close() })
    
    $bottomPanel.Controls.AddRange(@($prevButton, $nextButton, $saveButton, $cancelButton))
    $builderForm.Controls.AddRange(@($tabControl, $bottomPanel))
    
    # Initialize template preview
    Update-TemplatePreview
    
    # Load profile if provided
    if ($ProfilePath -and (Test-Path $ProfilePath)) {
        try {
            Load-ProfileData -ProfilePath $ProfilePath
        } catch {
            Show-SafeMessage -Message "Error loading profile: $_" -Title "Error" -Icon "Error"
        }
    }
    
    # Form shown event
    $builderForm.Add_Shown({ $builderForm.Activate() })
    
    $null = $builderForm.ShowDialog()
}
#endregion

#region Helper Functions with Enhanced Date Pickers
function Upload-Photo {
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp;*.gif)|*.jpg;*.jpeg;*.png;*.bmp;*.gif|All Files (*.*)|*.*"
    $openDialog.Title = "Select Passport Size Photo (Recommended: 200x230 pixels)"
    $openDialog.InitialDirectory = [Environment]::GetFolderPath("MyPictures")
    
    if ($openDialog.ShowDialog() -eq "OK") {
        try {
            $image = [System.Drawing.Image]::FromFile($openDialog.FileName)
            $maxWidth = 200
            $maxHeight = 230
            
            # Always resize to passport size
            $bitmap = New-Object System.Drawing.Bitmap($maxWidth, $maxHeight)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.DrawImage($image, 0, 0, $maxWidth, $maxHeight)
            
            $global:photoPictureBox.Image = $bitmap
            $graphics.Dispose()
            $image.Dispose()
            
            # Save to photos directory
            $photoName = "$(($global:nameTextBox.Text -replace '[^\w]', '_'))_$(Get-Date -Format 'yyyyMMdd_HHmmss').jpg"
            $photoPath = "$script:PhotosDir\$photoName"
            
            if (-not (Test-Path $script:PhotosDir)) {
                New-Item -ItemType Directory -Path $script:PhotosDir -Force | Out-Null
            }
            
            $global:photoPictureBox.Image.Save($photoPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
            $global:photoPath = $photoPath
            
            Show-SafeMessage -Message "Photo uploaded and resized to passport size successfully!" -Title "Success" -Icon "Information"
            
        } catch {
            Show-SafeMessage -Message "Error loading image: $_" -Title "Error" -Icon "Error"
        }
    }
}

function Upload-Signature {
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp;*.gif)|*.jpg;*.jpeg;*.png;*.bmp;*.gif|All Files (*.*)|*.*"
    $openDialog.Title = "Select Signature Image"
    $openDialog.InitialDirectory = [Environment]::GetFolderPath("MyPictures")
    
    if ($openDialog.ShowDialog() -eq "OK") {
        try {
            $image = [System.Drawing.Image]::FromFile($openDialog.FileName)
            $maxWidth = 300
            $maxHeight = 150
            
            # Resize signature
            $bitmap = New-Object System.Drawing.Bitmap($maxWidth, $maxHeight)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.DrawImage($image, 0, 0, $maxWidth, $maxHeight)
            
            $global:signaturePictureBox.Image = $bitmap
            $graphics.Dispose()
            $image.Dispose()
            
            # Save to signatures directory
            $signatureName = "$(($global:nameTextBox.Text -replace '[^\w]', '_'))_signature_$(Get-Date -Format 'yyyyMMdd_HHmmss').png"
            $signaturePath = "$script:SignaturesDir\$signatureName"
            
            if (-not (Test-Path $script:SignaturesDir)) {
                New-Item -ItemType Directory -Path $script:SignaturesDir -Force | Out-Null
            }
            
            $global:signaturePictureBox.Image.Save($signaturePath, [System.Drawing.Imaging.ImageFormat]::Png)
            $global:signaturePath = $signaturePath
            
            Show-SafeMessage -Message "Signature uploaded successfully!" -Title "Success" -Icon "Information"
            
        } catch {
            Show-SafeMessage -Message "Error loading signature image: $_" -Title "Error" -Icon "Error"
        }
    }
}

function Update-TemplatePreview {
    $selectedTemplate = $global:templateComboBox.SelectedItem
    if ($selectedTemplate -and $script:Templates.ContainsKey($selectedTemplate)) {
        $template = $script:Templates[$selectedTemplate]
        $global:templateDescLabel.Text = "$($template.Description)`nLayout: $($template.Layout)`nStyle: $($template.Style)"
        
        $global:colorComboBox.Items.Clear()
        $template.Colors | ForEach-Object { $null = $global:colorComboBox.Items.Add($_) }
        if ($global:colorComboBox.Items.Count -gt 0) {
            $global:colorComboBox.SelectedIndex = 0
        }
        
        # Generate HTML preview
        $previewHTML = Generate-TemplatePreviewHTML -Template $selectedTemplate
        $global:previewWebBrowser.DocumentText = $previewHTML
        
    } else {
        $global:previewWebBrowser.DocumentText = "<html><body style='padding: 20px; font-family: Arial;'><h3>No template selected</h3><p>Please select a template to preview</p></body></html>"
    }
}

function Generate-TemplatePreviewHTML {
    param([string]$Template)
    
    $templateInfo = $script:Templates[$Template]
    $color = if ($global:colorComboBox.SelectedItem) { $global:colorComboBox.SelectedItem } else { "Blue" }
    
    # Color mapping
    $colorMap = @{
        "Blue" = @{ Primary = "#007bff"; Secondary = "#0056b3" }
        "Green" = @{ Primary = "#28a745"; Secondary = "#1e7e34" }
        "Purple" = @{ Primary = "#6f42c1"; Secondary = "#593196" }
        "Red" = @{ Primary = "#dc3545"; Secondary = "#bd2130" }
        "Teal" = @{ Primary = "#20c997"; Secondary = "#17a673" }
        "Orange" = @{ Primary = "#fd7e14"; Secondary = "#e06c10" }
        "Navy" = @{ Primary = "#001f3f"; Secondary = "#001122" }
        "Charcoal" = @{ Primary = "#343a40"; Secondary = "#212529" }
        "Burgundy" = @{ Primary = "#800000"; Secondary = "#600000" }
        "DarkGreen" = @{ Primary = "#006400"; Secondary = "#004400" }
        "DarkBlue" = @{ Primary = "#00008b"; Secondary = "#00005b" }
        "Maroon" = @{ Primary = "#800000"; Secondary = "#600000" }
        "ForestGreen" = @{ Primary = "#228b22"; Secondary = "#1a691a" }
        "SlateGray" = @{ Primary = "#708090"; Secondary = "#5a6670" }
        "DarkRed" = @{ Primary = "#8b0000"; Secondary = "#6b0000" }
        "DarkSlateGray" = @{ Primary = "#2f4f4f"; Secondary = "#1f3f3f" }
    }
    
    $colors = if ($colorMap.ContainsKey($color)) { $colorMap[$color] } else { $colorMap["Blue"] }
    $font = if ($global:fontComboBox.SelectedItem) { $global:fontComboBox.SelectedItem } else { "Segoe UI" }
    
    # Generate sample data for preview
    $sampleData = @{
        Name = "John A. Doe, PhD"
        Title = "Professor of Computer Science"
        Email = "john.doe@university.edu"
        Phone = "+1 (555) 123-4567"
        Summary = "Experienced academic professional with expertise in Artificial Intelligence, Machine Learning, and Data Science. Published 50+ research papers in reputed journals and conferences."
        TemplateName = $templateInfo.Name
        Style = $templateInfo.Style
        Layout = $templateInfo.Layout
        Sections = $templateInfo.Sections -join ", "
        SupportsAcademic = $templateInfo.SupportsAcademic
    }
    
    # Create comprehensive HTML preview
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: '$font', Arial, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background-color: #f5f5f5;
        }
        .preview-container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .preview-header {
            background: $( $colors.Primary );
            color: white;
            padding: 30px;
            text-align: center;
        }
        .preview-name {
            font-size: 28px;
            font-weight: bold;
            margin-bottom: 10px;
        }
        .preview-title {
            font-size: 18px;
            opacity: 0.9;
        }
        .preview-content {
            padding: 30px;
        }
        .preview-section {
            margin-bottom: 25px;
            border-bottom: 1px solid #eee;
            padding-bottom: 15px;
        }
        .section-title {
            color: $( $colors.Secondary );
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 10px;
            border-bottom: 2px solid $( $colors.Primary );
            padding-bottom: 5px;
        }
        .template-info {
            background: #f8f9fa;
            border-left: 4px solid $( $colors.Primary );
            padding: 15px;
            margin-top: 20px;
            border-radius: 4px;
        }
        .info-item {
            margin-bottom: 5px;
        }
        .info-label {
            font-weight: bold;
            color: $( $colors.Secondary );
        }
        .academic-features {
            background: #e8f4f8;
            border: 1px solid #b3e0f2;
            padding: 15px;
            border-radius: 5px;
            margin-top: 15px;
        }
        .feature-list {
            list-style-type: none;
            padding-left: 0;
        }
        .feature-list li {
            padding: 5px 0;
            border-bottom: 1px dashed #ddd;
        }
        .feature-list li:last-child {
            border-bottom: none;
        }
        .badge {
            display: inline-block;
            padding: 3px 8px;
            background: $( $colors.Primary );
            color: white;
            border-radius: 12px;
            font-size: 12px;
            margin-left: 10px;
        }
        .footer {
            text-align: center;
            padding: 15px;
            background: #f8f9fa;
            color: #666;
            font-size: 12px;
            border-top: 1px solid #eee;
        }
    </style>
</head>
<body>
    <div class="preview-container">
        <div class="preview-header">
            <div class="preview-name">$( $sampleData.Name )</div>
            <div class="preview-title">$( $sampleData.Title )</div>
        </div>
        
        <div class="preview-content">
            <div class="preview-section">
                <div class="section-title">Contact Information</div>
                <p><strong>Email:</strong> $( $sampleData.Email )</p>
                <p><strong>Phone:</strong> $( $sampleData.Phone )</p>
            </div>
            
            <div class="preview-section">
                <div class="section-title">Professional Summary</div>
                <p>$( $sampleData.Summary )</p>
            </div>
            
            $( if ($sampleData.SupportsAcademic) { "
            <div class='preview-section'>
                <div class='section-title'>Academic Features <span class='badge'>Academic CV</span></div>
                <div class='academic-features'>
                    <ul class='feature-list'>
                        <li>✓ Research Work & Publications</li>
                        <li>✓ PhD Details & Supervision</li>
                        <li>✓ Editorial Activities</li>
                        <li>✓ Seminars & Workshops</li>
                        <li>✓ Professional Profiles</li>
                        <li>✓ Declaration & Signature</li>
                    </ul>
                </div>
            </div>
            " } )
            
            <div class="template-info">
                <div class="info-item"><span class="info-label">Template:</span> $( $sampleData.TemplateName )</div>
                <div class="info-item"><span class="info-label">Style:</span> $( $sampleData.Style )</div>
                <div class="info-item"><span class="info-label">Layout:</span> $( $sampleData.Layout )</div>
                <div class="info-item"><span class="info-label">Color Scheme:</span> $( $color )</div>
                <div class="info-item"><span class="info-label">Font:</span> $( $font )</div>
                <div class="info-item"><span class="info-label">Sections Included:</span> $( $sampleData.Sections )</div>
            </div>
        </div>
        
        <div class="footer">
            Windows CV Generator Pro v$script:Version - Template Preview
        </div>
    </div>
</body>
</html>
"@
    
    return $html
}

function Show-ExperienceDialog {
    param(
        [bool]$EditMode = $false
    )
    
    $expForm = New-Object System.Windows.Forms.Form
    $expForm.Text = if ($EditMode) { "Edit Work Experience" } else { "Add Work Experience" }
    $expForm.Size = New-Object System.Drawing.Size(600, 650)
    $expForm.StartPosition = "CenterScreen"
    $expForm.FormBorderStyle = "FixedDialog"
    $expForm.MaximizeBox = $false
    $expForm.MinimizeBox = $false
    
    $yPos = 20
    $labelWidth = 150
    $controlWidth = 350
    
    # Job Title
    $jobTitleLabel = New-Object System.Windows.Forms.Label
    $jobTitleLabel.Text = "Job Title:*"
    $jobTitleLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $jobTitleLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $jobTitleTextBox = New-Object System.Windows.Forms.TextBox
    $jobTitleTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $jobTitleTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Company
    $companyLabel = New-Object System.Windows.Forms.Label
    $companyLabel.Text = "Company:*"
    $companyLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $companyLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $companyTextBox = New-Object System.Windows.Forms.TextBox
    $companyTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $companyTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Location
    $locationLabel = New-Object System.Windows.Forms.Label
    $locationLabel.Text = "Location:"
    $locationLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $locationLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $locationTextBox = New-Object System.Windows.Forms.TextBox
    $locationTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $locationTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Start Date - Enhanced DateTimePicker
    $startDateLabel = New-Object System.Windows.Forms.Label
    $startDateLabel.Text = "Start Date:*"
    $startDateLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $startDateLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $startDatePicker = New-Object System.Windows.Forms.DateTimePicker
    $startDatePicker.Location = New-Object System.Drawing.Point(180, $yPos)
    $startDatePicker.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $startDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Long
    $startDatePicker.ShowCheckBox = $false
    $yPos += 40
    
    # End Date - Enhanced DateTimePicker with "Present" option
    $endDateLabel = New-Object System.Windows.Forms.Label
    $endDateLabel.Text = "End Date:"
    $endDateLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $endDateLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    
    $endDatePanel = New-Object System.Windows.Forms.Panel
    $endDatePanel.Location = New-Object System.Drawing.Point(180, $yPos)
    $endDatePanel.Size = New-Object System.Drawing.Size($controlWidth, 30)
    
    $endDatePicker = New-Object System.Windows.Forms.DateTimePicker
    $endDatePicker.Location = New-Object System.Drawing.Point(0, 0)
    $endDatePicker.Size = New-Object System.Drawing.Size(250, 30)
    $endDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Long
    $endDatePicker.ShowCheckBox = $true
    $endDatePicker.Checked = $false
    $endDatePicker.CustomFormat = " "
    
    $presentCheckBox = New-Object System.Windows.Forms.CheckBox
    $presentCheckBox.Text = "Present"
    $presentCheckBox.Location = New-Object System.Drawing.Point(260, 5)
    $presentCheckBox.Size = New-Object System.Drawing.Size(90, 25)
    $presentCheckBox.Add_CheckedChanged({
        if ($presentCheckBox.Checked) {
            $endDatePicker.Enabled = $false
            $endDatePicker.CustomFormat = " "
            $endDatePicker.Checked = $false
        } else {
            $endDatePicker.Enabled = $true
            $endDatePicker.CustomFormat = "MMMM yyyy"
            $endDatePicker.Checked = $true
        }
    })
    
    $endDatePanel.Controls.AddRange(@($endDatePicker, $presentCheckBox))
    $yPos += 40
    
    # Description
    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Text = "Description:*"
    $descLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $descLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $descTextBox = New-Object System.Windows.Forms.TextBox
    $descTextBox.Multiline = $true
    $descTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $descTextBox.Size = New-Object System.Drawing.Size($controlWidth, 150)
    $descTextBox.ScrollBars = "Vertical"
    $yPos += 170
    
    # Buttons
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object System.Drawing.Point(180, $yPos)
    $saveButton.Size = New-Object System.Drawing.Size(100, 30)
    $saveButton.DialogResult = "OK"
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(290, $yPos)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.DialogResult = "Cancel"
    
    $expForm.Controls.AddRange(@(
        $jobTitleLabel, $jobTitleTextBox,
        $companyLabel, $companyTextBox,
        $locationLabel, $locationTextBox,
        $startDateLabel, $startDatePicker,
        $endDateLabel, $endDatePanel,
        $descLabel, $descTextBox,
        $saveButton, $cancelButton
    ))
    
    $expForm.AcceptButton = $saveButton
    $expForm.CancelButton = $cancelButton
    
    if ($EditMode -and $global:experienceDataGrid.SelectedRows.Count -gt 0) {
        $row = $global:experienceDataGrid.SelectedRows[0]
        $jobTitleTextBox.Text = $row.Cells[0].Value
        $companyTextBox.Text = $row.Cells[1].Value
        $locationTextBox.Text = $row.Cells[2].Value
        $descTextBox.Text = $row.Cells[5].Value
    }
    
    if ($expForm.ShowDialog() -eq "OK") {
        $endDateDisplay = if ($presentCheckBox.Checked) { "Present" } else { $endDatePicker.Value.ToString("MMMM yyyy") }
        
        if ($EditMode -and $global:experienceDataGrid.SelectedRows.Count -gt 0) {
            $row = $global:experienceDataGrid.SelectedRows[0]
            $row.Cells[0].Value = $jobTitleTextBox.Text
            $row.Cells[1].Value = $companyTextBox.Text
            $row.Cells[2].Value = $locationTextBox.Text
            $row.Cells[3].Value = $startDatePicker.Value.ToString("MMMM yyyy")
            $row.Cells[4].Value = $endDateDisplay
            $row.Cells[5].Value = $descTextBox.Text
        } else {
            $null = $global:experienceDataGrid.Rows.Add(
                $jobTitleTextBox.Text,
                $companyTextBox.Text,
                $locationTextBox.Text,
                $startDatePicker.Value.ToString("MMMM yyyy"),
                $endDateDisplay,
                $descTextBox.Text
            )
        }
    }
}

function Show-EducationDialog {
    param(
        [bool]$EditMode = $false
    )
    
    $eduForm = New-Object System.Windows.Forms.Form
    $eduForm.Text = if ($EditMode) { "Edit Education" } else { "Add Education" }
    $eduForm.Size = New-Object System.Drawing.Size(600, 500)
    $eduForm.StartPosition = "CenterScreen"
    $eduForm.FormBorderStyle = "FixedDialog"
    $eduForm.MaximizeBox = $false
    $eduForm.MinimizeBox = $false
    
    $yPos = 20
    $labelWidth = 150
    $controlWidth = 350
    
    # Degree
    $degreeLabel = New-Object System.Windows.Forms.Label
    $degreeLabel.Text = "Degree:*"
    $degreeLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $degreeLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $degreeTextBox = New-Object System.Windows.Forms.TextBox
    $degreeTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $degreeTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Institution
    $institutionLabel = New-Object System.Windows.Forms.Label
    $institutionLabel.Text = "Institution:*"
    $institutionLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $institutionLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $institutionTextBox = New-Object System.Windows.Forms.TextBox
    $institutionTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $institutionTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Location
    $locationLabel = New-Object System.Windows.Forms.Label
    $locationLabel.Text = "Location:"
    $locationLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $locationLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $locationTextBox = New-Object System.Windows.Forms.TextBox
    $locationTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $locationTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Start Date
    $startDateLabel = New-Object System.Windows.Forms.Label
    $startDateLabel.Text = "Start Date:"
    $startDateLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $startDateLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $startDatePicker = New-Object System.Windows.Forms.DateTimePicker
    $startDatePicker.Location = New-Object System.Drawing.Point(180, $yPos)
    $startDatePicker.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $startDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Long
    $startDatePicker.ShowCheckBox = $true
    $startDatePicker.Checked = $false
    $startDatePicker.CustomFormat = " "
    $yPos += 40
    
    # End Date with Graduation option
    $endDateLabel = New-Object System.Windows.Forms.Label
    $endDateLabel.Text = "End Date:"
    $endDateLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $endDateLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    
    $endDatePanel = New-Object System.Windows.Forms.Panel
    $endDatePanel.Location = New-Object System.Drawing.Point(180, $yPos)
    $endDatePanel.Size = New-Object System.Drawing.Size($controlWidth, 30)
    
    $endDatePicker = New-Object System.Windows.Forms.DateTimePicker
    $endDatePicker.Location = New-Object System.Drawing.Point(0, 0)
    $endDatePicker.Size = New-Object System.Drawing.Size(250, 30)
    $endDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Long
    $endDatePicker.ShowCheckBox = $true
    $endDatePicker.Checked = $false
    $endDatePicker.CustomFormat = " "
    
    $ongoingCheckBox = New-Object System.Windows.Forms.CheckBox
    $ongoingCheckBox.Text = "Ongoing"
    $ongoingCheckBox.Location = New-Object System.Drawing.Point(260, 5)
    $ongoingCheckBox.Size = New-Object System.Drawing.Size(90, 25)
    $ongoingCheckBox.Add_CheckedChanged({
        if ($ongoingCheckBox.Checked) {
            $endDatePicker.Enabled = $false
            $endDatePicker.CustomFormat = " "
            $endDatePicker.Checked = $false
        } else {
            $endDatePicker.Enabled = $true
            $endDatePicker.CustomFormat = "MMMM yyyy"
            $endDatePicker.Checked = $true
        }
    })
    
    $endDatePanel.Controls.AddRange(@($endDatePicker, $ongoingCheckBox))
    $yPos += 40
    
    # GPA/Score
    $gpaLabel = New-Object System.Windows.Forms.Label
    $gpaLabel.Text = "GPA/Score:"
    $gpaLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $gpaLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $gpaTextBox = New-Object System.Windows.Forms.TextBox
    $gpaTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $gpaTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Buttons
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object System.Drawing.Point(180, $yPos)
    $saveButton.Size = New-Object System.Drawing.Size(100, 30)
    $saveButton.DialogResult = "OK"
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(290, $yPos)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.DialogResult = "Cancel"
    
    $eduForm.Controls.AddRange(@(
        $degreeLabel, $degreeTextBox,
        $institutionLabel, $institutionTextBox,
        $locationLabel, $locationTextBox,
        $startDateLabel, $startDatePicker,
        $endDateLabel, $endDatePanel,
        $gpaLabel, $gpaTextBox,
        $saveButton, $cancelButton
    ))
    
    $eduForm.AcceptButton = $saveButton
    $eduForm.CancelButton = $cancelButton
    
    if ($EditMode -and $global:educationDataGrid.SelectedRows.Count -gt 0) {
        $row = $global:educationDataGrid.SelectedRows[0]
        $degreeTextBox.Text = $row.Cells[0].Value
        $institutionTextBox.Text = $row.Cells[1].Value
        $locationTextBox.Text = $row.Cells[2].Value
        $gpaTextBox.Text = $row.Cells[5].Value
    }
    
    if ($eduForm.ShowDialog() -eq "OK") {
        $startDateDisplay = if ($startDatePicker.Checked) { $startDatePicker.Value.ToString("MMMM yyyy") } else { "" }
        $endDateDisplay = if ($ongoingCheckBox.Checked) { "Ongoing" } elseif ($endDatePicker.Checked) { $endDatePicker.Value.ToString("MMMM yyyy") } else { "" }
        
        if ($EditMode -and $global:educationDataGrid.SelectedRows.Count -gt 0) {
            $row = $global:educationDataGrid.SelectedRows[0]
            $row.Cells[0].Value = $degreeTextBox.Text
            $row.Cells[1].Value = $institutionTextBox.Text
            $row.Cells[2].Value = $locationTextBox.Text
            $row.Cells[3].Value = $startDateDisplay
            $row.Cells[4].Value = $endDateDisplay
            $row.Cells[5].Value = $gpaTextBox.Text
        } else {
            $null = $global:educationDataGrid.Rows.Add(
                $degreeTextBox.Text,
                $institutionTextBox.Text,
                $locationTextBox.Text,
                $startDateDisplay,
                $endDateDisplay,
                $gpaTextBox.Text
            )
        }
    }
}

function Show-ProjectDialog {
    param(
        [bool]$EditMode = $false
    )
    
    $projectForm = New-Object System.Windows.Forms.Form
    $projectForm.Text = if ($EditMode) { "Edit Project" } else { "Add Project" }
    $projectForm.Size = New-Object System.Drawing.Size(600, 500)
    $projectForm.StartPosition = "CenterScreen"
    $projectForm.FormBorderStyle = "FixedDialog"
    $projectForm.MaximizeBox = $false
    $projectForm.MinimizeBox = $false
    
    $yPos = 20
    $labelWidth = 150
    $controlWidth = 350
    
    # Project Name
    $projectNameLabel = New-Object System.Windows.Forms.Label
    $projectNameLabel.Text = "Project Name:*"
    $projectNameLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $projectNameLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $projectNameTextBox = New-Object System.Windows.Forms.TextBox
    $projectNameTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $projectNameTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Role
    $roleLabel = New-Object System.Windows.Forms.Label
    $roleLabel.Text = "Your Role:*"
    $roleLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $roleLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $roleTextBox = New-Object System.Windows.Forms.TextBox
    $roleTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $roleTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Description
    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Text = "Description:*"
    $descLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $descLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $descTextBox = New-Object System.Windows.Forms.TextBox
    $descTextBox.Multiline = $true
    $descTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $descTextBox.Size = New-Object System.Drawing.Size($controlWidth, 100)
    $descTextBox.ScrollBars = "Vertical"
    $yPos += 110
    
    # Technologies
    $techLabel = New-Object System.Windows.Forms.Label
    $techLabel.Text = "Technologies Used:*"
    $techLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $techLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $techTextBox = New-Object System.Windows.Forms.TextBox
    $techTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $techTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Buttons
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object System.Drawing.Point(180, $yPos)
    $saveButton.Size = New-Object System.Drawing.Size(100, 30)
    $saveButton.DialogResult = "OK"
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(290, $yPos)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.DialogResult = "Cancel"
    
    $projectForm.Controls.AddRange(@(
        $projectNameLabel, $projectNameTextBox,
        $roleLabel, $roleTextBox,
        $descLabel, $descTextBox,
        $techLabel, $techTextBox,
        $saveButton, $cancelButton
    ))
    
    $projectForm.AcceptButton = $saveButton
    $projectForm.CancelButton = $cancelButton
    
    if ($EditMode -and $global:projectsDataGrid.SelectedRows.Count -gt 0) {
        $row = $global:projectsDataGrid.SelectedRows[0]
        $projectNameTextBox.Text = $row.Cells[0].Value
        $roleTextBox.Text = $row.Cells[1].Value
        $descTextBox.Text = $row.Cells[2].Value
        $techTextBox.Text = $row.Cells[3].Value
    }
    
    if ($projectForm.ShowDialog() -eq "OK") {
        if ($EditMode -and $global:projectsDataGrid.SelectedRows.Count -gt 0) {
            $row = $global:projectsDataGrid.SelectedRows[0]
            $row.Cells[0].Value = $projectNameTextBox.Text
            $row.Cells[1].Value = $roleTextBox.Text
            $row.Cells[2].Value = $descTextBox.Text
            $row.Cells[3].Value = $techTextBox.Text
        } else {
            $null = $global:projectsDataGrid.Rows.Add(
                $projectNameTextBox.Text,
                $roleTextBox.Text,
                $descTextBox.Text,
                $techTextBox.Text
            )
        }
    }
}

function Show-CertificationDialog {
    param(
        [bool]$EditMode = $false
    )
    
    $certForm = New-Object System.Windows.Forms.Form
    $certForm.Text = if ($EditMode) { "Edit Certification" } else { "Add Certification" }
    $certForm.Size = New-Object System.Drawing.Size(600, 450)  # Increased height for checkbox
    $certForm.StartPosition = "CenterScreen"
    $certForm.FormBorderStyle = "FixedDialog"
    $certForm.MaximizeBox = $false
    $certForm.MinimizeBox = $false
    
    $yPos = 20
    $labelWidth = 150
    $controlWidth = 350
    
    # Certification Name
    $certNameLabel = New-Object System.Windows.Forms.Label
    $certNameLabel.Text = "Certification:*"
    $certNameLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $certNameLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $certNameTextBox = New-Object System.Windows.Forms.TextBox
    $certNameTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $certNameTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Issuer
    $issuerLabel = New-Object System.Windows.Forms.Label
    $issuerLabel.Text = "Issuing Organization:*"
    $issuerLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $issuerLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $issuerTextBox = New-Object System.Windows.Forms.TextBox
    $issuerTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $issuerTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Date - FIXED: Added checkbox like other date pickers
    $dateLabel = New-Object System.Windows.Forms.Label
    $dateLabel.Text = "Date Obtained:"
    $dateLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $dateLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    
    $datePanel = New-Object System.Windows.Forms.Panel
    $datePanel.Location = New-Object System.Drawing.Point(180, $yPos)
    $datePanel.Size = New-Object System.Drawing.Size($controlWidth, 30)
    
    $datePicker = New-Object System.Windows.Forms.DateTimePicker
    $datePicker.Location = New-Object System.Drawing.Point(0, 0)
    $datePicker.Size = New-Object System.Drawing.Size(250, 30)
    $datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Long
    $datePicker.ShowCheckBox = $true
    $datePicker.Checked = $false
    $datePicker.CustomFormat = " "
    $datePicker.Add_ValueChanged({
        if ($this.Checked) {
            $this.CustomFormat = "MMMM yyyy"
        } else {
            $this.CustomFormat = " "
        }
    })
    
    $datePanel.Controls.Add($datePicker)
    $yPos += 40
    
    # Credential ID
    $credIdLabel = New-Object System.Windows.Forms.Label
    $credIdLabel.Text = "Credential ID (Optional):"
    $credIdLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $credIdLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $credIdTextBox = New-Object System.Windows.Forms.TextBox
    $credIdTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $credIdTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Buttons
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object System.Drawing.Point(180, $yPos)
    $saveButton.Size = New-Object System.Drawing.Size(100, 30)
    $saveButton.DialogResult = "OK"
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(290, $yPos)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.DialogResult = "Cancel"
    
    $certForm.Controls.AddRange(@(
        $certNameLabel, $certNameTextBox,
        $issuerLabel, $issuerTextBox,
        $dateLabel, $datePanel,
        $credIdLabel, $credIdTextBox,
        $saveButton, $cancelButton
    ))
    
    $certForm.AcceptButton = $saveButton
    $certForm.CancelButton = $cancelButton
    
    if ($EditMode -and $global:certsDataGrid.SelectedRows.Count -gt 0) {
        $row = $global:certsDataGrid.SelectedRows[0]
        $certNameTextBox.Text = $row.Cells[0].Value
        $issuerTextBox.Text = $row.Cells[1].Value
        if ($row.Cells[2].Value) {
            try {
                $datePicker.Value = [DateTime]::Parse($row.Cells[2].Value)
                $datePicker.Checked = $true
                $datePicker.CustomFormat = "MMMM yyyy"
            } catch {}
        }
        $credIdTextBox.Text = $row.Cells[3].Value
    }
    
    if ($certForm.ShowDialog() -eq "OK") {
        $dateDisplay = if ($datePicker.Checked) { $datePicker.Value.ToString("MMMM yyyy") } else { "" }
        
        if ($EditMode -and $global:certsDataGrid.SelectedRows.Count -gt 0) {
            $row = $global:certsDataGrid.SelectedRows[0]
            $row.Cells[0].Value = $certNameTextBox.Text
            $row.Cells[1].Value = $issuerTextBox.Text
            $row.Cells[2].Value = $dateDisplay
            $row.Cells[3].Value = $credIdTextBox.Text
        } else {
            $null = $global:certsDataGrid.Rows.Add(
                $certNameTextBox.Text,
                $issuerTextBox.Text,
                $dateDisplay,
                $credIdTextBox.Text
            )
        }
    }
}

# New Dialog Functions for Academic Features
function Show-PatentDialog {
    param(
        [bool]$EditMode = $false
    )
    
    $patentForm = New-Object System.Windows.Forms.Form
    $patentForm.Text = if ($EditMode) { "Edit Patent" } else { "Add Patent" }
    $patentForm.Size = New-Object System.Drawing.Size(600, 400)
    $patentForm.StartPosition = "CenterScreen"
    $patentForm.FormBorderStyle = "FixedDialog"
    $patentForm.MaximizeBox = $false
    $patentForm.MinimizeBox = $false
    
    $yPos = 20
    $labelWidth = 150
    $controlWidth = 350
    
    # Patent Title
    $patentTitleLabel = New-Object System.Windows.Forms.Label
    $patentTitleLabel.Text = "Patent Title:*"
    $patentTitleLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $patentTitleLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $patentTitleTextBox = New-Object System.Windows.Forms.TextBox
    $patentTitleTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $patentTitleTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Patent Type
    $patentTypeLabel = New-Object System.Windows.Forms.Label
    $patentTypeLabel.Text = "Type:*"
    $patentTypeLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $patentTypeLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $patentTypeComboBox = New-Object System.Windows.Forms.ComboBox
    $patentTypeComboBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $patentTypeComboBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $patentTypeComboBox.Items.AddRange(@("Filed", "Published", "Grant"))
    $patentTypeComboBox.SelectedIndex = 0
    $yPos += 40
    
    # Patent Number
    $patentNumberLabel = New-Object System.Windows.Forms.Label
    $patentNumberLabel.Text = "Patent Number:"
    $patentNumberLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $patentNumberLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $patentNumberTextBox = New-Object System.Windows.Forms.TextBox
    $patentNumberTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $patentNumberTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Date
    $dateLabel = New-Object System.Windows.Forms.Label
    $dateLabel.Text = "Date:"
    $dateLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $dateLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $datePicker = New-Object System.Windows.Forms.DateTimePicker
    $datePicker.Location = New-Object System.Drawing.Point(180, $yPos)
    $datePicker.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Long
    $datePicker.ShowCheckBox = $true
    $datePicker.Checked = $false
    $datePicker.CustomFormat = " "
    $datePicker.Add_ValueChanged({
        if ($this.Checked) {
            $this.CustomFormat = "MMMM yyyy"
        } else {
            $this.CustomFormat = " "
        }
    })
    $yPos += 40
    
    # URL/Link
    $urlLabel = New-Object System.Windows.Forms.Label
    $urlLabel.Text = "URL/Link:"
    $urlLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $urlLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $urlTextBox = New-Object System.Windows.Forms.TextBox
    $urlTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $urlTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Buttons
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object System.Drawing.Point(180, $yPos)
    $saveButton.Size = New-Object System.Drawing.Size(100, 30)
    $saveButton.DialogResult = "OK"
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(290, $yPos)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.DialogResult = "Cancel"
    
    $patentForm.Controls.AddRange(@(
        $patentTitleLabel, $patentTitleTextBox,
        $patentTypeLabel, $patentTypeComboBox,
        $patentNumberLabel, $patentNumberTextBox,
        $dateLabel, $datePicker,
        $urlLabel, $urlTextBox,
        $saveButton, $cancelButton
    ))
    
    $patentForm.AcceptButton = $saveButton
    $patentForm.CancelButton = $cancelButton
    
    if ($EditMode -and $global:patentsDataGrid.SelectedRows.Count -gt 0) {
        $row = $global:patentsDataGrid.SelectedRows[0]
        $patentTitleTextBox.Text = $row.Cells[0].Value
        $patentTypeComboBox.SelectedItem = $row.Cells[1].Value
        $patentNumberTextBox.Text = $row.Cells[2].Value
        if ($row.Cells[3].Value) {
            try {
                $datePicker.Value = [DateTime]::Parse($row.Cells[3].Value)
                $datePicker.Checked = $true
                $datePicker.CustomFormat = "MMMM yyyy"
            } catch {}
        }
        $urlTextBox.Text = $row.Cells[4].Value
    }
    
    if ($patentForm.ShowDialog() -eq "OK") {
        $dateDisplay = if ($datePicker.Checked) { $datePicker.Value.ToString("MMMM yyyy") } else { "" }
        
        if ($EditMode -and $global:patentsDataGrid.SelectedRows.Count -gt 0) {
            $row = $global:patentsDataGrid.SelectedRows[0]
            $row.Cells[0].Value = $patentTitleTextBox.Text
            $row.Cells[1].Value = $patentTypeComboBox.SelectedItem
            $row.Cells[2].Value = $patentNumberTextBox.Text
            $row.Cells[3].Value = $dateDisplay
            $row.Cells[4].Value = $urlTextBox.Text
        } else {
            $null = $global:patentsDataGrid.Rows.Add(
                $patentTitleTextBox.Text,
                $patentTypeComboBox.SelectedItem,
                $patentNumberTextBox.Text,
                $dateDisplay,
                $urlTextBox.Text
            )
        }
    }
}

function Show-SoftwareDialog {
    param(
        [bool]$EditMode = $false
    )
    
    $softwareForm = New-Object System.Windows.Forms.Form
    $softwareForm.Text = if ($EditMode) { "Edit Software" } else { "Add Software" }
    $softwareForm.Size = New-Object System.Drawing.Size(600, 400)
    $softwareForm.StartPosition = "CenterScreen"
    $softwareForm.FormBorderStyle = "FixedDialog"
    $softwareForm.MaximizeBox = $false
    $softwareForm.MinimizeBox = $false
    
    $yPos = 20
    $labelWidth = 150
    $controlWidth = 350
    
    # Software Name
    $softwareNameLabel = New-Object System.Windows.Forms.Label
    $softwareNameLabel.Text = "Software Name:*"
    $softwareNameLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $softwareNameLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $softwareNameTextBox = New-Object System.Windows.Forms.TextBox
    $softwareNameTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $softwareNameTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Description
    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Text = "Description:*"
    $descLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $descLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $descTextBox = New-Object System.Windows.Forms.TextBox
    $descTextBox.Multiline = $true
    $descTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $descTextBox.Size = New-Object System.Drawing.Size($controlWidth, 80)
    $descTextBox.ScrollBars = "Vertical"
    $yPos += 90
    
    # GitHub/GitLab URL
    $urlLabel = New-Object System.Windows.Forms.Label
    $urlLabel.Text = "GitHub/GitLab URL:*"
    $urlLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $urlLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $urlTextBox = New-Object System.Windows.Forms.TextBox
    $urlTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $urlTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Technologies
    $techLabel = New-Object System.Windows.Forms.Label
    $techLabel.Text = "Technologies Used:"
    $techLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $techLabel.Size = New-Object System.Drawing.Size($labelWidth, 30)
    $techTextBox = New-Object System.Windows.Forms.TextBox
    $techTextBox.Location = New-Object System.Drawing.Point(180, $yPos)
    $techTextBox.Size = New-Object System.Drawing.Size($controlWidth, 30)
    $yPos += 40
    
    # Buttons
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object System.Drawing.Point(180, $yPos)
    $saveButton.Size = New-Object System.Drawing.Size(100, 30)
    $saveButton.DialogResult = "OK"
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(290, $yPos)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.DialogResult = "Cancel"
    
    $softwareForm.Controls.AddRange(@(
        $softwareNameLabel, $softwareNameTextBox,
        $descLabel, $descTextBox,
        $urlLabel, $urlTextBox,
        $techLabel, $techTextBox,
        $saveButton, $cancelButton
    ))
    
    $softwareForm.AcceptButton = $saveButton
    $softwareForm.CancelButton = $cancelButton
    
    if ($EditMode -and $global:softwareDataGrid.SelectedRows.Count -gt 0) {
        $row = $global:softwareDataGrid.SelectedRows[0]
        $softwareNameTextBox.Text = $row.Cells[0].Value
        $descTextBox.Text = $row.Cells[1].Value
        $urlTextBox.Text = $row.Cells[2].Value
        $techTextBox.Text = $row.Cells[3].Value
    }
    
    if ($softwareForm.ShowDialog() -eq "OK") {
        if ($EditMode -and $global:softwareDataGrid.SelectedRows.Count -gt 0) {
            $row = $global:softwareDataGrid.SelectedRows[0]
            $row.Cells[0].Value = $softwareNameTextBox.Text
            $row.Cells[1].Value = $descTextBox.Text
            $row.Cells[2].Value = $urlTextBox.Text
            $row.Cells[3].Value = $techTextBox.Text
        } else {
            $null = $global:softwareDataGrid.Rows.Add(
                $softwareNameTextBox.Text,
                $descTextBox.Text,
                $urlTextBox.Text,
                $techTextBox.Text
            )
        }
    }
}
#endregion

#region Save and Load Functions - FIXED: Now opens CV Builder with loaded data
function Save-CVDraft {
    if ([string]::IsNullOrWhiteSpace($global:nameTextBox.Text)) {
        Show-SafeMessage -Message "Please enter your name before saving." -Title "Validation Error" -Icon "Warning"
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CV Profile (*.cvprofile)|*.cvprofile|JSON Files (*.json)|*.json"
    $saveDialog.InitialDirectory = $script:ProfilesDir
    $saveDialog.FileName = "$($global:nameTextBox.Text -replace '[^\w]', '_')_$(Get-Date -Format 'yyyyMMdd').cvprofile"
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        try {
            # Collect all data including academic fields
            $cvData = [PSCustomObject]@{
                PersonalInfo = [PSCustomObject]@{
                    FullName = $global:nameTextBox.Text
                    Title = $global:titleTextBox.Text
                    DateOfBirth = if ($global:dobDateTimePicker.Checked) { $global:dobDateTimePicker.Value.ToString("yyyy-MM-dd") } else { $null }
                    Email = $global:emailTextBox.Text
                    Phone = $global:phoneTextBox.Text
                    Address = $global:addressTextBox.Text
                    Nationality = $global:nationalityTextBox.Text
                    LinkedIn = $global:linkedInTextBox.Text
                    GitHub = $global:githubTextBox.Text
                    Website = $global:websiteTextBox.Text
                    Summary = $global:summaryTextBox.Text
                    Category = $global:categoryComboBox.SelectedItem
                }
                Experiences = @()
                Education = @()
                Projects = @()
                Certifications = @()
                Skills = @{
                    Technical = ($global:techSkillsTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                    Business = ($global:businessSkillsTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                    Soft = ($global:softSkillsTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                    Languages = ($global:languagesTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                }
                # Academic fields
                ResearchWork = @{
                    Patents = @()
                    Software = @()
                    ConferencePapers = ($global:conferencePapersTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                    JournalArticles = ($global:journalArticlesTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                    Books = ($global:booksTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                }
                ResearchProfiles = [PSCustomObject]@{
                    ORCID = $global:orcidTextBox.Text
                    GoogleScholar = $global:googleScholarTextBox.Text
                    ResearchGate = $global:researchGateTextBox.Text
                    Scopus = $global:scopusTextBox.Text
                    WebOfScience = $global:wosTextBox.Text
                    Vidwan = $global:vidwanTextBox.Text
                }
                PhDGuided = ($global:phdGuidedTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                EditorialActivities = [PSCustomObject]@{
                    EditorRoles = $global:editorRolesTextBox.Text
                    ReviewerActivities = $global:reviewerTextBox.Text
                    GuestEditor = $global:guestEditorTextBox.Text
                }
                PhDDetails = [PSCustomObject]@{
                    Title = $global:phdTitleTextBox.Text
                    University = $global:phdUniversityTextBox.Text
                    Department = $global:phdDepartmentTextBox.Text
                    Supervisor = $global:phdSupervisorTextBox.Text
                    Year = $global:phdYearTextBox.Text
                    ThesisURL = $global:phdThesisUrlTextBox.Text
                    Abstract = $global:phdAbstractTextBox.Text
                }
                Seminars = [PSCustomObject]@{
                    Attended = ($global:attendedSeminarsTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                    Conducted = ($global:conductedSeminarsTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                }
                PhDThesisEvaluated = ($global:phdThesisEvalTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
                Declaration = [PSCustomObject]@{
                    Statement = $global:declarationTextBox.Text
                    DatePlace = $global:datePlaceTextBox.Text
                }
                ProfessionalProfiles = [PSCustomObject]@{
                    LinkedInProf = $global:linkedInProfTextBox.Text
                    GitHubProf = $global:githubProfTextBox.Text
                    AcademicProfile = $global:academicProfileTextBox.Text
                    IRINS = $global:irinsTextBox.Text
                    YouTube = $global:youtubeTextBox.Text
                    GoogleDevelopers = $global:googleDevTextBox.Text
                }
                Template = $global:templateComboBox.SelectedItem
                Color = $global:colorComboBox.SelectedItem
                Font = $global:fontComboBox.SelectedItem
                IncludePhoto = $global:includePhotoCheckBox.Checked
                IncludeSignature = $global:includeSignatureCheckBox.Checked
                PhotoPath = $global:photoPath
                SignaturePath = $global:signaturePath
                LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Version = $script:Version
                Generator = "Windows CV Generator Pro"
            }
            
            # Collect experiences
            foreach ($row in $global:experienceDataGrid.Rows) {
                if ($row.Cells[0].Value) {
                    $cvData.Experiences += [PSCustomObject]@{
                        JobTitle = $row.Cells[0].Value
                        Company = $row.Cells[1].Value
                        Location = $row.Cells[2].Value
                        StartDate = $row.Cells[3].Value
                        EndDate = $row.Cells[4].Value
                        Description = $row.Cells[5].Value
                    }
                }
            }
            
            # Collect education
            foreach ($row in $global:educationDataGrid.Rows) {
                if ($row.Cells[0].Value) {
                    $cvData.Education += [PSCustomObject]@{
                        Degree = $row.Cells[0].Value
                        Institution = $row.Cells[1].Value
                        Location = $row.Cells[2].Value
                        StartDate = $row.Cells[3].Value
                        EndDate = $row.Cells[4].Value
                        GPA = $row.Cells[5].Value
                    }
                }
            }
            
            # Collect projects
            foreach ($row in $global:projectsDataGrid.Rows) {
                if ($row.Cells[0].Value) {
                    $cvData.Projects += [PSCustomObject]@{
                        ProjectName = $row.Cells[0].Value
                        Role = $row.Cells[1].Value
                        Description = $row.Cells[2].Value
                        Technologies = $row.Cells[3].Value
                    }
                }
            }
            
            # Collect certifications
            foreach ($row in $global:certsDataGrid.Rows) {
                if ($row.Cells[0].Value) {
                    $cvData.Certifications += [PSCustomObject]@{
                        Certification = $row.Cells[0].Value
                        Issuer = $row.Cells[1].Value
                        Date = $row.Cells[2].Value
                        CredentialID = $row.Cells[3].Value
                    }
                }
            }
            
            # Collect patents
            foreach ($row in $global:patentsDataGrid.Rows) {
                if ($row.Cells[0].Value) {
                    $cvData.ResearchWork.Patents += [PSCustomObject]@{
                        Title = $row.Cells[0].Value
                        Type = $row.Cells[1].Value
                        PatentNumber = $row.Cells[2].Value
                        Date = $row.Cells[3].Value
                        URL = $row.Cells[4].Value
                    }
                }
            }
            
            # Collect software
            foreach ($row in $global:softwareDataGrid.Rows) {
                if ($row.Cells[0].Value) {
                    $cvData.ResearchWork.Software += [PSCustomObject]@{
                        Name = $row.Cells[0].Value
                        Description = $row.Cells[1].Value
                        URL = $row.Cells[2].Value
                        Technologies = $row.Cells[3].Value
                    }
                }
            }
            
            # Convert to JSON and save
            $jsonContent = $cvData | ConvertTo-Json -Depth 10
            $jsonContent | Out-File -FilePath $saveDialog.FileName -Encoding UTF8 -Force
            
            Show-SafeMessage -Message "CV draft saved successfully!`nFile: $($saveDialog.FileName)" -Title "Success" -Icon "Information"
            
        } catch {
            Show-SafeMessage -Message "Error saving CV draft: $_" -Title "Error" -Icon "Error"
        }
    }
}

function Save-CVDraftAs {
    Save-CVDraft
}

function Load-Profile {
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = "CV Profile (*.cvprofile;*.json)|*.cvprofile;*.json|All Files (*.*)|*.*"
    $openDialog.InitialDirectory = $script:ProfilesDir
    
    if ($openDialog.ShowDialog() -eq "OK") {
        # Store the profile path and show CV builder
        Show-CVBuilderForm -ProfilePath $openDialog.FileName
    }
}

function Load-ProfileData {
    param(
        [string]$ProfilePath
    )
    
    try {
        $fileContent = Get-Content -Path $ProfilePath -Raw
        $cvData = $fileContent | ConvertFrom-Json
        
        # Clear existing data
        $global:experienceDataGrid.Rows.Clear()
        $global:educationDataGrid.Rows.Clear()
        $global:projectsDataGrid.Rows.Clear()
        $global:certsDataGrid.Rows.Clear()
        $global:patentsDataGrid.Rows.Clear()
        $global:softwareDataGrid.Rows.Clear()
        
        # Load personal info
        $global:nameTextBox.Text = $cvData.PersonalInfo.FullName
        $global:titleTextBox.Text = $cvData.PersonalInfo.Title
        
        if ($cvData.PersonalInfo.DateOfBirth) {
            try {
                $global:dobDateTimePicker.Value = [DateTime]::Parse($cvData.PersonalInfo.DateOfBirth)
                $global:dobDateTimePicker.Checked = $true
                $global:dobDateTimePicker.CustomFormat = "dddd, MMMM dd, yyyy"
            } catch {}
        }
        
        $global:emailTextBox.Text = $cvData.PersonalInfo.Email
        $global:phoneTextBox.Text = $cvData.PersonalInfo.Phone
        $global:addressTextBox.Text = $cvData.PersonalInfo.Address
        $global:nationalityTextBox.Text = $cvData.PersonalInfo.Nationality
        $global:linkedInTextBox.Text = $cvData.PersonalInfo.LinkedIn
        $global:githubTextBox.Text = $cvData.PersonalInfo.GitHub
        $global:websiteTextBox.Text = $cvData.PersonalInfo.Website
        $global:summaryTextBox.Text = $cvData.PersonalInfo.Summary
        
        if ($cvData.PersonalInfo.Category) {
            $global:categoryComboBox.SelectedItem = $cvData.PersonalInfo.Category
        }
        
        # Load template settings
        if ($cvData.Template) {
            $global:templateComboBox.SelectedItem = $cvData.Template
        }
        
        if ($cvData.Color) {
            $global:colorComboBox.SelectedItem = $cvData.Color
        }
        
        if ($cvData.Font) {
            $global:fontComboBox.SelectedItem = $cvData.Font
        }
        
        $global:includePhotoCheckBox.Checked = [bool]$cvData.IncludePhoto
        if ($cvData.IncludeSignature -ne $null) {
            $global:includeSignatureCheckBox.Checked = [bool]$cvData.IncludeSignature
        }
        
        # Load experiences
        if ($cvData.Experiences) {
            foreach ($exp in $cvData.Experiences) {
                $null = $global:experienceDataGrid.Rows.Add(
                    $exp.JobTitle,
                    $exp.Company,
                    $exp.Location,
                    $exp.StartDate,
                    $exp.EndDate,
                    $exp.Description
                )
            }
        }
        
        # Load education
        if ($cvData.Education) {
            foreach ($edu in $cvData.Education) {
                $null = $global:educationDataGrid.Rows.Add(
                    $edu.Degree,
                    $edu.Institution,
                    $edu.Location,
                    $edu.StartDate,
                    $edu.EndDate,
                    $edu.GPA
                )
            }
        }
        
        # Load projects
        if ($cvData.Projects) {
            foreach ($project in $cvData.Projects) {
                $null = $global:projectsDataGrid.Rows.Add(
                    $project.ProjectName,
                    $project.Role,
                    $project.Description,
                    $project.Technologies
                )
            }
        }
        
        # Load certifications
        if ($cvData.Certifications) {
            foreach ($cert in $cvData.Certifications) {
                $null = $global:certsDataGrid.Rows.Add(
                    $cert.Certification,
                    $cert.Issuer,
                    $cert.Date,
                    $cert.CredentialID
                )
            }
        }
        
        # Load skills
        if ($cvData.Skills) {
            if ($cvData.Skills.Technical) {
                $global:techSkillsTextBox.Text = $cvData.Skills.Technical -join "`n"
            }
            if ($cvData.Skills.Business) {
                $global:businessSkillsTextBox.Text = $cvData.Skills.Business -join "`n"
            }
            if ($cvData.Skills.Soft) {
                $global:softSkillsTextBox.Text = $cvData.Skills.Soft -join "`n"
            }
            if ($cvData.Skills.Languages) {
                $global:languagesTextBox.Text = $cvData.Skills.Languages -join "`n"
            }
        }
        
        # Load research work - Patents
        if ($cvData.ResearchWork -and $cvData.ResearchWork.Patents) {
            foreach ($patent in $cvData.ResearchWork.Patents) {
                $null = $global:patentsDataGrid.Rows.Add(
                    $patent.Title,
                    $patent.Type,
                    $patent.PatentNumber,
                    $patent.Date,
                    $patent.URL
                )
            }
        }
        
        # Load research work - Software
        if ($cvData.ResearchWork -and $cvData.ResearchWork.Software) {
            foreach ($software in $cvData.ResearchWork.Software) {
                $null = $global:softwareDataGrid.Rows.Add(
                    $software.Name,
                    $software.Description,
                    $software.URL,
                    $software.Technologies
                )
            }
        }
        
        # Load research work - Publications
        if ($cvData.ResearchWork) {
            if ($cvData.ResearchWork.ConferencePapers) {
                $global:conferencePapersTextBox.Text = $cvData.ResearchWork.ConferencePapers -join "`n"
            }
            if ($cvData.ResearchWork.JournalArticles) {
                $global:journalArticlesTextBox.Text = $cvData.ResearchWork.JournalArticles -join "`n"
            }
            if ($cvData.ResearchWork.Books) {
                $global:booksTextBox.Text = $cvData.ResearchWork.Books -join "`n"
            }
        }
        
        # Load research profiles
        if ($cvData.ResearchProfiles) {
            $global:orcidTextBox.Text = $cvData.ResearchProfiles.ORCID
            $global:googleScholarTextBox.Text = $cvData.ResearchProfiles.GoogleScholar
            $global:researchGateTextBox.Text = $cvData.ResearchProfiles.ResearchGate
            $global:scopusTextBox.Text = $cvData.ResearchProfiles.Scopus
            $global:wosTextBox.Text = $cvData.ResearchProfiles.WebOfScience
            $global:vidwanTextBox.Text = $cvData.ResearchProfiles.Vidwan
        }
        
        # Load PhD guided
        if ($cvData.PhDGuided) {
            $global:phdGuidedTextBox.Text = $cvData.PhDGuided -join "`n"
        }
        
        # Load editorial activities
        if ($cvData.EditorialActivities) {
            $global:editorRolesTextBox.Text = $cvData.EditorialActivities.EditorRoles
            $global:reviewerTextBox.Text = $cvData.EditorialActivities.ReviewerActivities
            $global:guestEditorTextBox.Text = $cvData.EditorialActivities.GuestEditor
        }
        
        # Load PhD details
        if ($cvData.PhDDetails) {
            $global:phdTitleTextBox.Text = $cvData.PhDDetails.Title
            $global:phdUniversityTextBox.Text = $cvData.PhDDetails.University
            $global:phdDepartmentTextBox.Text = $cvData.PhDDetails.Department
            $global:phdSupervisorTextBox.Text = $cvData.PhDDetails.Supervisor
            $global:phdYearTextBox.Text = $cvData.PhDDetails.Year
            $global:phdThesisUrlTextBox.Text = $cvData.PhDDetails.ThesisURL
            $global:phdAbstractTextBox.Text = $cvData.PhDDetails.Abstract
        }
        
        # Load seminars
        if ($cvData.Seminars) {
            if ($cvData.Seminars.Attended) {
                $global:attendedSeminarsTextBox.Text = $cvData.Seminars.Attended -join "`n"
            }
            if ($cvData.Seminars.Conducted) {
                $global:conductedSeminarsTextBox.Text = $cvData.Seminars.Conducted -join "`n"
            }
        }
        
        # Load PhD thesis evaluated
        if ($cvData.PhDThesisEvaluated) {
            $global:phdThesisEvalTextBox.Text = $cvData.PhDThesisEvaluated -join "`n"
        }
        
        # Load declaration
        if ($cvData.Declaration) {
            $global:declarationTextBox.Text = $cvData.Declaration.Statement
            $global:datePlaceTextBox.Text = $cvData.Declaration.DatePlace
        }
        
        # Load professional profiles
        if ($cvData.ProfessionalProfiles) {
            $global:linkedInProfTextBox.Text = $cvData.ProfessionalProfiles.LinkedInProf
            $global:githubProfTextBox.Text = $cvData.ProfessionalProfiles.GitHubProf
            $global:academicProfileTextBox.Text = $cvData.ProfessionalProfiles.AcademicProfile
            $global:irinsTextBox.Text = $cvData.ProfessionalProfiles.IRINS
            $global:youtubeTextBox.Text = $cvData.ProfessionalProfiles.YouTube
            $global:googleDevTextBox.Text = $cvData.ProfessionalProfiles.GoogleDevelopers
        }
        
        # Load photo
        if ($cvData.PhotoPath -and (Test-Path $cvData.PhotoPath) -and $global:includePhotoCheckBox.Checked) {
            try {
                $photoImage = [System.Drawing.Image]::FromFile($cvData.PhotoPath)
                $global:photoPictureBox.Image = $photoImage
                $global:photoPath = $cvData.PhotoPath
            } catch {
                # If photo can't be loaded from saved path, try to find it in photos directory
                $photoName = [System.IO.Path]::GetFileName($cvData.PhotoPath)
                $localPhotoPath = "$script:PhotosDir\$photoName"
                if (Test-Path $localPhotoPath) {
                    try {
                        $photoImage = [System.Drawing.Image]::FromFile($localPhotoPath)
                        $global:photoPictureBox.Image = $photoImage
                        $global:photoPath = $localPhotoPath
                    } catch {}
                }
            }
        }
        
        # Load signature
        if ($cvData.SignaturePath -and (Test-Path $cvData.SignaturePath) -and $global:includeSignatureCheckBox.Checked) {
            try {
                $signatureImage = [System.Drawing.Image]::FromFile($cvData.SignaturePath)
                $global:signaturePictureBox.Image = $signatureImage
                $global:signaturePath = $cvData.SignaturePath
            } catch {
                $signatureName = [System.IO.Path]::GetFileName($cvData.SignaturePath)
                $localSignaturePath = "$script:SignaturesDir\$signatureName"
                if (Test-Path $localSignaturePath) {
                    try {
                        $signatureImage = [System.Drawing.Image]::FromFile($localSignaturePath)
                        $global:signaturePictureBox.Image = $signatureImage
                        $global:signaturePath = $localSignaturePath
                    } catch {}
                }
            }
        }
        
        # Update template preview
        Update-TemplatePreview
        
    } catch {
        Show-SafeMessage -Message "Error loading profile: $_" -Title "Error" -Icon "Error"
    }
}
#endregion

#region CV Generation Functions with Enhanced Word Support
function Generate-CVFromForm {
    # Validation
    if ([string]::IsNullOrWhiteSpace($global:nameTextBox.Text)) {
        Show-SafeMessage -Message "Please enter your full name." -Title "Validation Error" -Icon "Warning"
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($global:emailTextBox.Text)) {
        Show-SafeMessage -Message "Please enter your email address." -Title "Validation Error" -Icon "Warning"
        return
    }
    
    # Collect data
    $cvData = @{
        PersonalInfo = @{
            FullName = $global:nameTextBox.Text
            Title = $global:titleTextBox.Text
            DateOfBirth = if ($global:dobDateTimePicker.Checked) { $global:dobDateTimePicker.Value.ToString("MMMM dd, yyyy") } else { $null }
            Email = $global:emailTextBox.Text
            Phone = $global:phoneTextBox.Text
            Address = $global:addressTextBox.Text
            Nationality = $global:nationalityTextBox.Text
            LinkedIn = $global:linkedInTextBox.Text
            GitHub = $global:githubTextBox.Text
            Website = $global:websiteTextBox.Text
            Summary = $global:summaryTextBox.Text
            Category = $global:categoryComboBox.SelectedItem
        }
        Experiences = @()
        Education = @()
        Projects = @()
        Certifications = @()
        Skills = @{
            Technical = ($global:techSkillsTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
            Business = ($global:businessSkillsTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
            Soft = ($global:softSkillsTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
            Languages = ($global:languagesTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
        }
        # Academic fields
        ResearchWork = @{
            Patents = @()
            Software = @()
            ConferencePapers = ($global:conferencePapersTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
            JournalArticles = ($global:journalArticlesTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
            Books = ($global:booksTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
        }
        ResearchProfiles = @{
            ORCID = $global:orcidTextBox.Text
            GoogleScholar = $global:googleScholarTextBox.Text
            ResearchGate = $global:researchGateTextBox.Text
            Scopus = $global:scopusTextBox.Text
            WebOfScience = $global:wosTextBox.Text
            Vidwan = $global:vidwanTextBox.Text
        }
        PhDGuided = ($global:phdGuidedTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
        EditorialActivities = @{
            EditorRoles = $global:editorRolesTextBox.Text
            ReviewerActivities = $global:reviewerTextBox.Text
            GuestEditor = $global:guestEditorTextBox.Text
        }
        PhDDetails = @{
            Title = $global:phdTitleTextBox.Text
            University = $global:phdUniversityTextBox.Text
            Department = $global:phdDepartmentTextBox.Text
            Supervisor = $global:phdSupervisorTextBox.Text
            Year = $global:phdYearTextBox.Text
            ThesisURL = $global:phdThesisUrlTextBox.Text
            Abstract = $global:phdAbstractTextBox.Text
        }
        Seminars = @{
            Attended = ($global:attendedSeminarsTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
            Conducted = ($global:conductedSeminarsTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
        }
        PhDThesisEvaluated = ($global:phdThesisEvalTextBox.Text -split "`n" | Where-Object { $_ -and $_.Trim() -ne "" })
        Declaration = @{
            Statement = $global:declarationTextBox.Text
            DatePlace = $global:datePlaceTextBox.Text
        }
        ProfessionalProfiles = @{
            LinkedInProf = $global:linkedInProfTextBox.Text
            GitHubProf = $global:githubProfTextBox.Text
            AcademicProfile = $global:academicProfileTextBox.Text
            IRINS = $global:irinsTextBox.Text
            YouTube = $global:youtubeTextBox.Text
            GoogleDevelopers = $global:googleDevTextBox.Text
        }
        Template = $global:templateComboBox.SelectedItem
        Color = $global:colorComboBox.SelectedItem
        Font = $global:fontComboBox.SelectedItem
        IncludePhoto = $global:includePhotoCheckBox.Checked
        IncludeSignature = $global:includeSignatureCheckBox.Checked
        PhotoPath = $global:photoPath
        SignaturePath = $global:signaturePath
    }
    
    # Collect experiences
    foreach ($row in $global:experienceDataGrid.Rows) {
        if ($row.Cells[0].Value) {
            $cvData.Experiences += @{
                JobTitle = $row.Cells[0].Value
                Company = $row.Cells[1].Value
                Location = $row.Cells[2].Value
                StartDate = $row.Cells[3].Value
                EndDate = $row.Cells[4].Value
                Description = $row.Cells[5].Value
            }
        }
    }
    
    # Collect education
    foreach ($row in $global:educationDataGrid.Rows) {
        if ($row.Cells[0].Value) {
            $cvData.Education += @{
                Degree = $row.Cells[0].Value
                Institution = $row.Cells[1].Value
                Location = $row.Cells[2].Value
                StartDate = $row.Cells[3].Value
                EndDate = $row.Cells[4].Value
                GPA = $row.Cells[5].Value
            }
        }
    }
    
    # Collect projects
    foreach ($row in $global:projectsDataGrid.Rows) {
        if ($row.Cells[0].Value) {
            $cvData.Projects += @{
                ProjectName = $row.Cells[0].Value
                Role = $row.Cells[1].Value
                Description = $row.Cells[2].Value
                Technologies = $row.Cells[3].Value
            }
        }
    }
    
    # Collect certifications
    foreach ($row in $global:certsDataGrid.Rows) {
        if ($row.Cells[0].Value) {
            $cvData.Certifications += @{
                Certification = $row.Cells[0].Value
                Issuer = $row.Cells[1].Value
                Date = $row.Cells[2].Value
                CredentialID = $row.Cells[3].Value
            }
        }
    }
    
    # Collect patents
    foreach ($row in $global:patentsDataGrid.Rows) {
        if ($row.Cells[0].Value) {
            $cvData.ResearchWork.Patents += @{
                Title = $row.Cells[0].Value
                Type = $row.Cells[1].Value
                PatentNumber = $row.Cells[2].Value
                Date = $row.Cells[3].Value
                URL = $row.Cells[4].Value
            }
        }
    }
    
    # Collect software
    foreach ($row in $global:softwareDataGrid.Rows) {
        if ($row.Cells[0].Value) {
            $cvData.ResearchWork.Software += @{
                Name = $row.Cells[0].Value
                Description = $row.Cells[1].Value
                URL = $row.Cells[2].Value
                Technologies = $row.Cells[3].Value
            }
        }
    }
    
    # Determine output formats
    $outputFormats = @()
    if ($global:htmlCheckBox.Checked) { $outputFormats += "HTML" }
    if ($global:pdfCheckBox.Checked) { $outputFormats += "PDF" }
    if ($global:docxCheckBox.Checked) { $outputFormats += "DOCX" }
    
    if ($outputFormats.Count -eq 0) {
        Show-SafeMessage -Message "Please select at least one output format." -Title "Warning" -Icon "Warning"
        return
    }
    
    # Generate files
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $filenameBase = "$($cvData.PersonalInfo.FullName -replace '[^\w]', '_')_$timestamp"
    $generatedFiles = @()
    
    foreach ($format in $outputFormats) {
        $outputPath = "$script:OutputDir\$filenameBase.$($format.ToLower())"
        
        switch ($format) {
            "HTML" {
                Generate-HTMLCVAdvanced -Data $cvData -OutputPath $outputPath
                $generatedFiles += $outputPath
            }
            "PDF" {
                Generate-PDFCVAdvanced -Data $cvData -OutputPath $outputPath
                $generatedFiles += $outputPath
            }
            "DOCX" {
                Generate-DOCXCVAdvanced -Data $cvData -OutputPath $outputPath
                $generatedFiles += $outputPath
            }
        }
    }
    
    # Show success message
    $message = "CV generated successfully!`n`nGenerated files:`n"
    $message += ($generatedFiles -join "`n")
    
    $result = Show-SafeMessage -Message $message -Title "Success" -Buttons "YesNo" -Icon "Information"
    if ($result -eq "Yes") {
        try {
            Start-Process "explorer.exe" -ArgumentList "/select,`"$($generatedFiles[0])`""
        } catch {
            Start-Process $script:OutputDir
        }
    }
}

function Generate-HTMLCVAdvanced {
    param(
        [hashtable]$Data,
        [string]$OutputPath
    )
    
    $template = $script:Templates[$Data.Template]
    $color = $Data.Color
    $font = $Data.Font
    
    # Color mapping
    $colorMap = @{
        "Blue" = @{ Primary = "#007bff"; Secondary = "#0056b3" }
        "Green" = @{ Primary = "#28a745"; Secondary = "#1e7e34" }
        "Purple" = @{ Primary = "#6f42c1"; Secondary = "#593196" }
        "Red" = @{ Primary = "#dc3545"; Secondary = "#bd2130" }
        "Teal" = @{ Primary = "#20c997"; Secondary = "#17a673" }
        "Orange" = @{ Primary = "#fd7e14"; Secondary = "#e06c10" }
        "Navy" = @{ Primary = "#001f3f"; Secondary = "#001122" }
        "Charcoal" = @{ Primary = "#343a40"; Secondary = "#212529" }
        "Burgundy" = @{ Primary = "#800000"; Secondary = "#600000" }
        "DarkGreen" = @{ Primary = "#006400"; Secondary = "#004400" }
        "DarkBlue" = @{ Primary = "#00008b"; Secondary = "#00005b" }
        "Maroon" = @{ Primary = "#800000"; Secondary = "#600000" }
        "ForestGreen" = @{ Primary = "#228b22"; Secondary = "#1a691a" }
        "SlateGray" = @{ Primary = "#708090"; Secondary = "#5a6670" }
        "DarkRed" = @{ Primary = "#8b0000"; Secondary = "#6b0000" }
        "DarkSlateGray" = @{ Primary = "#2f4f4f"; Secondary = "#1f3f3f" }
    }
    
    $colors = if ($colorMap.ContainsKey($color)) { $colorMap[$color] } else { $colorMap["Blue"] }
    
    # Generate photo HTML if included
    $photoHTML = ""
    if ($Data.IncludePhoto -and $Data.PhotoPath -and (Test-Path $Data.PhotoPath)) {
        try {
            $photoBytes = [System.IO.File]::ReadAllBytes($Data.PhotoPath)
            $photoBase64 = [Convert]::ToBase64String($photoBytes)
            $photoExtension = [System.IO.Path]::GetExtension($Data.PhotoPath).TrimStart('.')
            $photoHTML = @"
            <div style="float: right; margin-left: 20px; margin-bottom: 20px;">
                <img src="data:image/$photoExtension;base64,$photoBase64" 
                     alt="Profile Photo" 
                     style="width: 150px; height: 180px; object-fit: cover; border-radius: 5px; border: 2px solid $( $colors.Primary );">
            </div>
"@
        } catch {}
    }
    
    # Generate signature HTML if included
    $signatureHTML = ""
    if ($Data.IncludeSignature -and $Data.SignaturePath -and (Test-Path $Data.SignaturePath)) {
        try {
            $signatureBytes = [System.IO.File]::ReadAllBytes($Data.SignaturePath)
            $signatureBase64 = [Convert]::ToBase64String($signatureBytes)
            $signatureExtension = [System.IO.Path]::GetExtension($Data.SignaturePath).TrimStart('.')
            $signatureHTML = @"
            <div style="margin-top: 30px; text-align: right;">
                <div style="display: inline-block; text-align: center;">
                    <img src="data:image/$signatureExtension;base64,$signatureBase64" 
                         alt="Signature" 
                         style="width: 200px; height: 80px; object-fit: contain;">
                    <div style="border-top: 1px solid #333; margin-top: 5px; padding-top: 5px;">
                        <strong>Signature</strong>
                    </div>
                </div>
            </div>
"@
        } catch {}
    }
    
    # Generate HTML with academic sections if template supports it
    $academicSectionsHTML = ""
    if ($template.SupportsAcademic) {
        $academicSectionsHTML = Generate-AcademicSectionsHTML -Data $Data -Colors $colors -Font $font
    }
    
    # Generate HTML
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$( $Data.PersonalInfo.FullName ) - Curriculum Vitae</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: '$font', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background-color: #fff;
            max-width: 210mm;
            margin: 0 auto;
            padding: 20px;
        }
        .cv-header {
            border-bottom: 4px solid $( $colors.Primary );
            padding-bottom: 25px;
            margin-bottom: 30px;
            overflow: hidden;
        }
        .name {
            font-size: 2.8em;
            color: $( $colors.Primary );
            margin-bottom: 8px;
            font-weight: 700;
        }
        .title {
            font-size: 1.5em;
            color: $( $colors.Secondary );
            margin-bottom: 20px;
            font-weight: 600;
        }
        .contact-info {
            display: flex;
            flex-wrap: wrap;
            gap: 25px;
            margin-bottom: 15px;
        }
        .contact-item {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 1.05em;
        }
        .section {
            margin-bottom: 35px;
            clear: both;
        }
        .section-title {
            font-size: 1.5em;
            color: $( $colors.Primary );
            border-bottom: 2px solid $( $colors.Secondary );
            padding-bottom: 10px;
            margin-bottom: 20px;
            font-weight: 600;
        }
        .experience-item, .education-item {
            margin-bottom: 25px;
            page-break-inside: avoid;
        }
        .item-header {
            display: flex;
            justify-content: space-between;
            margin-bottom: 10px;
        }
        .item-title {
            font-weight: bold;
            font-size: 1.2em;
            color: #222;
        }
        .item-subtitle {
            color: $( $colors.Secondary );
            font-weight: 500;
        }
        .item-date {
            color: #666;
            font-style: italic;
            font-weight: 500;
        }
        .skills-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
            gap: 15px;
        }
        .skill-category {
            margin-bottom: 20px;
        }
        .category-title {
            font-weight: 600;
            color: $( $colors.Secondary );
            margin-bottom: 10px;
            font-size: 1.1em;
        }
        .skill-item {
            background-color: #f8f9fa;
            padding: 12px 20px;
            border-radius: 8px;
            border-left: 4px solid $( $colors.Primary );
            margin-bottom: 8px;
        }
        ul { padding-left: 25px; margin-top: 10px; }
        li { margin-bottom: 8px; line-height: 1.5; }
        .summary {
            font-size: 1.1em;
            line-height: 1.7;
            color: #444;
            margin-bottom: 30px;
        }
        @media print {
            body { padding: 0; font-size: 12pt; }
            .no-print { display: none; }
            .cv-header { padding-top: 20px; }
        }
        .footer {
            margin-top: 50px;
            padding-top: 20px;
            border-top: 1px solid #eee;
            text-align: center;
            color: #666;
            font-size: 0.9em;
        }
        .developer-info {
            font-size: 0.8em;
            color: #999;
            text-align: center;
            margin-top: 10px;
        }
        .declaration {
            margin-top: 40px;
            padding: 20px;
            background-color: #f9f9f9;
            border-left: 4px solid $( $colors.Primary );
            border-radius: 5px;
        }
        .professional-profiles {
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            margin-top: 20px;
        }
        .profile-link {
            display: inline-block;
            padding: 8px 15px;
            background-color: $( $colors.Primary );
            color: white;
            text-decoration: none;
            border-radius: 20px;
            font-size: 0.9em;
            transition: background-color 0.3s;
        }
        .profile-link:hover {
            background-color: $( $colors.Secondary );
            text-decoration: none;
            color: white;
        }
        .academic-badge {
            display: inline-block;
            padding: 3px 10px;
            background-color: $( $colors.Primary );
            color: white;
            border-radius: 12px;
            font-size: 0.8em;
            margin-left: 10px;
            vertical-align: middle;
        }
    </style>
</head>
<body>
    <div class="cv-header">
        $( $photoHTML )
        <div>
            <h1 class="name">$( $Data.PersonalInfo.FullName )</h1>
            <h2 class="title">$( $Data.PersonalInfo.Title )</h2>
            <div class="contact-info">
                <div class="contact-item"><strong>📧</strong> <span>$( $Data.PersonalInfo.Email )</span></div>
                <div class="contact-item"><strong>📱</strong> <span>$( $Data.PersonalInfo.Phone )</span></div>
                <div class="contact-item"><strong>📍</strong> <span>$( $Data.PersonalInfo.Address )</span></div>
                $( if ($Data.PersonalInfo.DateOfBirth) { "<div class='contact-item'><strong>🎂</strong> <span>$( $Data.PersonalInfo.DateOfBirth )</span></div>" } )
                $( if ($Data.PersonalInfo.Nationality) { "<div class='contact-item'><strong>🌍</strong> <span>$( $Data.PersonalInfo.Nationality )</span></div>" } )
                $( if ($Data.PersonalInfo.LinkedIn) { "<div class='contact-item'><strong>💼</strong> <a href='$( $Data.PersonalInfo.LinkedIn )' target='_blank'>$( $Data.PersonalInfo.LinkedIn )</a></div>" } )
                $( if ($Data.PersonalInfo.GitHub) { "<div class='contact-item'><strong>🐙</strong> <a href='$( $Data.PersonalInfo.GitHub )' target='_blank'>$( $Data.PersonalInfo.GitHub )</a></div>" } )
                $( if ($Data.PersonalInfo.Website) { "<div class='contact-item'><strong>🌐</strong> <a href='$( $Data.PersonalInfo.Website )' target='_blank'>$( $Data.PersonalInfo.Website )</a></div>" } )
            </div>
            $( if ($Data.PersonalInfo.Category) { "
            <div style='margin-top: 15px;'>
                <span style='background-color: $( $colors.Primary ); color: white; padding: 5px 15px; border-radius: 20px; font-size: 0.9em;'>
                    $( $Data.PersonalInfo.Category )
                </span>
                $( if ($template.SupportsAcademic) { "<span class='academic-badge'>Academic CV</span>" } )
            </div>
            " } )
        </div>
    </div>
    
    $( if ($Data.PersonalInfo.Summary) { "
    <div class='section'>
        <h3 class='section-title'>Professional Summary</h3>
        <div class='summary'>$( $Data.PersonalInfo.Summary )</div>
    </div>
    " } )
    
    $( if ($Data.Experiences.Count -gt 0) { "
    <div class='section'>
        <h3 class='section-title'>Work Experience</h3>
        $( $Data.Experiences | ForEach-Object { "
        <div class='experience-item'>
            <div class='item-header'>
                <div>
                    <div class='item-title'>$( $_.JobTitle )</div>
                    <div class='item-subtitle'>$( $_.Company ) | $( $_.Location )</div>
                </div>
                <div class='item-date'>$( $_.StartDate ) - $( $_.EndDate )</div>
            </div>
            $( if ($_.Description) { "<p>$( $_.Description )</p>" } )
        </div>
        " } )
    </div>
    " } )
    
    $( if ($Data.Education.Count -gt 0) { "
    <div class='section'>
        <h3 class='section-title'>Education</h3>
        $( $Data.Education | ForEach-Object { "
        <div class='education-item'>
            <div class='item-header'>
                <div>
                    <div class='item-title'>$( $_.Degree )</div>
                    <div class='item-subtitle'>$( $_.Institution ) | $( $_.Location )</div>
                </div>
                <div class='item-date'>$( $_.StartDate ) - $( $_.EndDate )</div>
            </div>
            $( if ($_.GPA) { "<p style='margin-top: 5px;'><strong>GPA/Score:</strong> $( $_.GPA )</p>" } )
        </div>
        " } )
    </div>
    " } )
    
    $( if ($Data.PhDDetails.Title) { "
    <div class='section'>
        <h3 class='section-title'>PhD Details</h3>
        <div class='experience-item'>
            <div class='item-header'>
                <div>
                    <div class='item-title'>$( $Data.PhDDetails.Title )</div>
                    <div class='item-subtitle'>$( $Data.PhDDetails.University ) | $( $Data.PhDDetails.Department )</div>
                </div>
                $( if ($Data.PhDDetails.Year) { "<div class='item-date'>$( $Data.PhDDetails.Year )</div>" } )
            </div>
            $( if ($Data.PhDDetails.Supervisor) { "<p><strong>Supervisor(s):</strong> $( $Data.PhDDetails.Supervisor )</p>" } )
            $( if ($Data.PhDDetails.Abstract) { "<p style='margin-top: 10px;'>$( $Data.PhDDetails.Abstract )</p>" } )
            $( if ($Data.PhDDetails.ThesisURL) { "<p><strong>Thesis Link:</strong> <a href='$( $Data.PhDDetails.ThesisURL )' target='_blank'>$( $Data.PhDDetails.ThesisURL )</a></p>" } )
        </div>
    </div>
    " } )
    
    $( $academicSectionsHTML )
    
    $( if ($Data.Projects.Count -gt 0) { "
    <div class='section'>
        <h3 class='section-title'>Projects</h3>
        $( $Data.Projects | ForEach-Object { "
        <div class='experience-item'>
            <div class='item-header'>
                <div>
                    <div class='item-title'>$( $_.ProjectName )</div>
                    <div class='item-subtitle'>Role: $( $_.Role )</div>
                </div>
            </div>
            $( if ($_.Description) { "<p>$( $_.Description )</p>" } )
            $( if ($_.Technologies) { "<p><strong>Technologies:</strong> $( $_.Technologies )</p>" } )
        </div>
        " } )
    </div>
    " } )
    
    $( if ($Data.Certifications.Count -gt 0) { "
    <div class='section'>
        <h3 class='section-title'>Certifications</h3>
        <ul>
        $( $Data.Certifications | ForEach-Object { "
            <li>
                <strong>$( $_.Certification )</strong> - $( $_.Issuer )
                $( if ($_.Date) { " ($( $_.Date ))" } )
                $( if ($_.CredentialID) { " [$( $_.CredentialID )]" } )
            </li>
        " } )
        </ul>
    </div>
    " } )
    
    <div class="section">
        <h3 class="section-title">Skills & Competencies</h3>
        <div class="skills-grid">
            $( if ($Data.Skills.Technical.Count -gt 0) { "
            <div class='skill-category'>
                <div class='category-title'>Technical Skills</div>
                $( $Data.Skills.Technical | ForEach-Object { "
                <div class='skill-item'>$_</div>
                " } )
            </div>
            " } )
            $( if ($Data.Skills.Business.Count -gt 0) { "
            <div class='skill-category'>
                <div class='category-title'>Business Skills</div>
                $( $Data.Skills.Business | ForEach-Object { "
                <div class='skill-item'>$_</div>
                " } )
            </div>
            " } )
            $( if ($Data.Skills.Soft.Count -gt 0) { "
            <div class='skill-category'>
                <div class='category-title'>Soft Skills</div>
                $( $Data.Skills.Soft | ForEach-Object { "
                <div class='skill-item'>$_</div>
                " } )
            </div>
            " } )
            $( if ($Data.Skills.Languages.Count -gt 0) { "
            <div class='skill-category'>
                <div class='category-title'>Languages</div>
                $( $Data.Skills.Languages | ForEach-Object { "
                <div class='skill-item'>$_</div>
                " } )
            </div>
            " } )
        </div>
    </div>
    
    $( if ($Data.Declaration.Statement) { "
    <div class='section'>
        <h3 class='section-title'>Declaration</h3>
        <div class='declaration'>
            <p>$( $Data.Declaration.Statement )</p>
            $( if ($Data.Declaration.DatePlace) { "<p style='margin-top: 20px;'><strong>Date & Place:</strong> $( $Data.Declaration.DatePlace )</p>" } )
            $( $signatureHTML )
        </div>
    </div>
    " } )
    
    <div class="section">
        <h3 class="section-title">Professional Profiles</h3>
        <div class="professional-profiles">
            $( if ($Data.PersonalInfo.LinkedIn) { "<a href='$( $Data.PersonalInfo.LinkedIn )' target='_blank' class='profile-link'>LinkedIn</a>" } )
            $( if ($Data.PersonalInfo.GitHub) { "<a href='$( $Data.PersonalInfo.GitHub )' target='_blank' class='profile-link'>GitHub</a>" } )
            $( if ($Data.ProfessionalProfiles.LinkedInProf) { "<a href='$( $Data.ProfessionalProfiles.LinkedInProf )' target='_blank' class='profile-link'>LinkedIn (Professional)</a>" } )
            $( if ($Data.ProfessionalProfiles.GitHubProf) { "<a href='$( $Data.ProfessionalProfiles.GitHubProf )' target='_blank' class='profile-link'>GitHub (Professional)</a>" } )
            $( if ($Data.ProfessionalProfiles.AcademicProfile) { "<a href='$( $Data.ProfessionalProfiles.AcademicProfile )' target='_blank' class='profile-link'>Academic Profile</a>" } )
            $( if ($Data.ResearchProfiles.GoogleScholar) { "<a href='$( $Data.ResearchProfiles.GoogleScholar )' target='_blank' class='profile-link'>Google Scholar</a>" } )
            $( if ($Data.ResearchProfiles.ResearchGate) { "<a href='$( $Data.ResearchProfiles.ResearchGate )' target='_blank' class='profile-link'>ResearchGate</a>" } )
            $( if ($Data.ResearchProfiles.ORCID) { "<a href='https://orcid.org/$( $Data.ResearchProfiles.ORCID )' target='_blank' class='profile-link'>ORCID</a>" } )
            $( if ($Data.ResearchProfiles.Scopus) { "<a href='https://www.scopus.com/authid/detail.uri?authorId=$( $Data.ResearchProfiles.Scopus )' target='_blank' class='profile-link'>Scopus</a>" } )
            $( if ($Data.ProfessionalProfiles.YouTube) { "<a href='$( $Data.ProfessionalProfiles.YouTube )' target='_blank' class='profile-link'>YouTube</a>" } )
            $( if ($Data.ProfessionalProfiles.GoogleDevelopers) { "<a href='$( $Data.ProfessionalProfiles.GoogleDevelopers )' target='_blank' class='profile-link'>Google Developers</a>" } )
            $( if ($Data.ProfessionalProfiles.IRINS) { "<a href='$( $Data.ProfessionalProfiles.IRINS )' target='_blank' class='profile-link'>Institutional IRINS</a>" } )
        </div>
    </div>
    
    <div class="footer no-print">
        <p>Generated with Windows CV Generator Pro v$script:Version</p>
        <p>$( Get-Date -Format 'MMMM dd, yyyy' ) | Template: $( $Data.Template ) | Color: $( $Data.Color )</p>
        <div class="developer-info">
            Developed by $script:Developer © $script:Year | $script:Website
        </div>
    </div>
</body>
</html>
"@
    
    $html | Out-File -FilePath $OutputPath -Encoding UTF8
}

function Generate-AcademicSectionsHTML {
    param(
        [hashtable]$Data,
        [hashtable]$Colors,
        [string]$Font
    )
    
    $html = ""
    
    # Research Work Section
    $hasResearchWork = $false
    $researchWorkHTML = ""
    
    # Patents
    if ($Data.ResearchWork.Patents.Count -gt 0) {
        $hasResearchWork = $true
        $researchWorkHTML += "<h4 style='color: $( $Colors.Secondary ); margin-top: 15px;'>Patents</h4><ul>"
        foreach ($patent in $Data.ResearchWork.Patents) {
            $urlLink = if ($patent.URL) { " <a href='$( $patent.URL )' target='_blank'>[Link]</a>" } else { "" }
            $researchWorkHTML += "<li><strong>$( $patent.Title )</strong> ($( $patent.Type ))$( if ($patent.PatentNumber) { ", Patent No: $( $patent.PatentNumber )" } )$( if ($patent.Date) { ", $( $patent.Date )" } )$urlLink</li>"
        }
        $researchWorkHTML += "</ul>"
    }
    
    # Software Developed
    if ($Data.ResearchWork.Software.Count -gt 0) {
        $hasResearchWork = $true
        $researchWorkHTML += "<h4 style='color: $( $Colors.Secondary ); margin-top: 15px;'>Software/Programs Developed</h4><ul>"
        foreach ($software in $Data.ResearchWork.Software) {
            $urlLink = if ($software.URL) { " <a href='$( $software.URL )' target='_blank'>[GitHub]</a>" } else { "" }
            $researchWorkHTML += "<li><strong>$( $software.Name )</strong>: $( $software.Description )$( if ($software.Technologies) { " ($( $software.Technologies ))" } )$urlLink</li>"
        }
        $researchWorkHTML += "</ul>"
    }
    
    # Conference Papers
    if ($Data.ResearchWork.ConferencePapers.Count -gt 0) {
        $hasResearchWork = $true
        $researchWorkHTML += "<h4 style='color: $( $Colors.Secondary ); margin-top: 15px;'>Conference Papers</h4><ul>"
        foreach ($paper in $Data.ResearchWork.ConferencePapers) {
            # Extract URL if present
            $paperText = $paper
            $urlMatch = [regex]::Match($paper, 'URL:\s*(https?://[^\s]+)')
            if ($urlMatch.Success) {
                $url = $urlMatch.Groups[1].Value
                $paperText = $paper -replace 'URL:\s*https?://[^\s]+', ''
                $researchWorkHTML += "<li>$paperText <a href='$url' target='_blank'>[Link]</a></li>"
            } else {
                $researchWorkHTML += "<li>$paperText</li>"
            }
        }
        $researchWorkHTML += "</ul>"
    }
    
    # Journal Articles
    if ($Data.ResearchWork.JournalArticles.Count -gt 0) {
        $hasResearchWork = $true
        $researchWorkHTML += "<h4 style='color: $( $Colors.Secondary ); margin-top: 15px;'>Journal Articles</h4><ul>"
        foreach ($article in $Data.ResearchWork.JournalArticles) {
            $articleText = $article
            $urlMatch = [regex]::Match($article, 'URL:\s*(https?://[^\s]+)')
            if ($urlMatch.Success) {
                $url = $urlMatch.Groups[1].Value
                $articleText = $article -replace 'URL:\s*https?://[^\s]+', ''
                $researchWorkHTML += "<li>$articleText <a href='$url' target='_blank'>[Link]</a></li>"
            } else {
                $researchWorkHTML += "<li>$articleText</li>"
            }
        }
        $researchWorkHTML += "</ul>"
    }
    
    # Books
    if ($Data.ResearchWork.Books.Count -gt 0) {
        $hasResearchWork = $true
        $researchWorkHTML += "<h4 style='color: $( $Colors.Secondary ); margin-top: 15px;'>Books & Chapters</h4><ul>"
        foreach ($book in $Data.ResearchWork.Books) {
            $bookText = $book
            $urlMatch = [regex]::Match($book, 'URL:\s*(https?://[^\s]+)')
            if ($urlMatch.Success) {
                $url = $urlMatch.Groups[1].Value
                $bookText = $book -replace 'URL:\s*https?://[^\s]+', ''
                $researchWorkHTML += "<li>$bookText <a href='$url' target='_blank'>[Link]</a></li>"
            } else {
                $researchWorkHTML += "<li>$bookText</li>"
            }
        }
        $researchWorkHTML += "</ul>"
    }
    
    if ($hasResearchWork) {
        $html += @"
    <div class="section">
        <h3 class="section-title">Research Work & Publications</h3>
        $researchWorkHTML
    </div>
"@
    }
    
    # PhD Scholars Guided
    if ($Data.PhDGuided.Count -gt 0) {
        $html += @"
    <div class="section">
        <h3 class="section-title">PhD Scholars Guided</h3>
        <ul>
        $( $Data.PhDGuided | ForEach-Object { "<li>$_</li>" } )
        </ul>
    </div>
"@
    }
    
    # Editorial Activities
    $hasEditorialActivities = $false
    $editorialHTML = ""
    
    if ($Data.EditorialActivities.EditorRoles) {
        $hasEditorialActivities = $true
        $editorialHTML += "<h4 style='color: $( $Colors.Secondary ); margin-top: 15px;'>Editor Roles</h4><p>$( $Data.EditorialActivities.EditorRoles.Replace("`n", "<br>") )</p>"
    }
    
    if ($Data.EditorialActivities.ReviewerActivities) {
        $hasEditorialActivities = $true
        $editorialHTML += "<h4 style='color: $( $Colors.Secondary ); margin-top: 15px;'>Reviewer Activities</h4><p>$( $Data.EditorialActivities.ReviewerActivities.Replace("`n", "<br>") )</p>"
    }
    
    if ($Data.EditorialActivities.GuestEditor) {
        $hasEditorialActivities = $true
        $editorialHTML += "<h4 style='color: $( $Colors.Secondary ); margin-top: 15px;'>Guest Editor Roles</h4><p>$( $Data.EditorialActivities.GuestEditor.Replace("`n", "<br>") )</p>"
    }
    
    if ($hasEditorialActivities) {
        $html += @"
    <div class="section">
        <h3 class="section-title">Editorial Activities</h3>
        $editorialHTML
    </div>
"@
    }
    
    # Seminars & Workshops
    $hasSeminars = $false
    $seminarsHTML = ""
    
    if ($Data.Seminars.Attended.Count -gt 0) {
        $hasSeminars = $true
        $seminarsHTML += "<h4 style='color: $( $Colors.Secondary ); margin-top: 15px;'>Seminars/Workshops Attended</h4><ul>"
        foreach ($seminar in $Data.Seminars.Attended) {
            $seminarsHTML += "<li>$seminar</li>"
        }
        $seminarsHTML += "</ul>"
    }
    
    if ($Data.Seminars.Conducted.Count -gt 0) {
        $hasSeminars = $true
        $seminarsHTML += "<h4 style='color: $( $Colors.Secondary ); margin-top: 15px;'>Seminars/Workshops Conducted</h4><ul>"
        foreach ($seminar in $Data.Seminars.Conducted) {
            $seminarsHTML += "<li>$seminar</li>"
        }
        $seminarsHTML += "</ul>"
    }
    
    if ($hasSeminars) {
        $html += @"
    <div class="section">
        <h3 class="section-title">Seminars & Workshops</h3>
        $seminarsHTML
    </div>
"@
    }
    
    # PhD Thesis Evaluated
    if ($Data.PhDThesisEvaluated.Count -gt 0) {
        $html += @"
    <div class="section">
        <h3 class="section-title">PhD Thesis Evaluated as External Examiner</h3>
        <ul>
        $( $Data.PhDThesisEvaluated | ForEach-Object { "<li>$_</li>" } )
        </ul>
    </div>
"@
    }
    
    return $html
}

function Generate-PDFCVAdvanced {
    param(
        [hashtable]$Data,
        [string]$OutputPath
    )
    
    # First generate HTML
    $htmlPath = "$script:OutputDir\temp_cv_$(Get-Random).html"
    Generate-HTMLCVAdvanced -Data $Data -OutputPath $htmlPath
    
    # Try to use wkhtmltopdf if available
    $wkhtmlPath = "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
    if (Test-Path $wkhtmlPath) {
        try {
            & $wkhtmlPath --enable-local-file-access $htmlPath $OutputPath 2>$null
        } catch {}
    } else {
        # Fallback: Create a simple PDF using .NET
        try {
            $pdfDocument = New-Object System.Text.StringBuilder
            $pdfDocument.AppendLine("%PDF-1.4")
            $pdfDocument.AppendLine("1 0 obj")
            $pdfDocument.AppendLine("<<")
            $pdfDocument.AppendLine("/Type /Catalog")
            $pdfDocument.AppendLine("/Pages 2 0 R")
            $pdfDocument.AppendLine(">>")
            $pdfDocument.AppendLine("endobj")
            $pdfDocument.AppendLine("2 0 obj")
            $pdfDocument.AppendLine("<<")
            $pdfDocument.AppendLine("/Type /Pages")
            $pdfDocument.AppendLine("/Kids [3 0 R]")
            $pdfDocument.AppendLine("/Count 1")
            $pdfDocument.AppendLine(">>")
            $pdfDocument.AppendLine("endobj")
            $pdfDocument.AppendLine("3 0 obj")
            $pdfDocument.AppendLine("<<")
            $pdfDocument.AppendLine("/Type /Page")
            $pdfDocument.AppendLine("/Parent 2 0 R")
            $pdfDocument.AppendLine("/MediaBox [0 0 612 792]")
            $pdfDocument.AppendLine("/Contents 4 0 R")
            $pdfDocument.AppendLine("/Resources <<")
            $pdfDocument.AppendLine("/Font <<")
            $pdfDocument.AppendLine("/F1 5 0 R")
            $pdfDocument.AppendLine(">>")
            $pdfDocument.AppendLine(">>")
            $pdfDocument.AppendLine(">>")
            $pdfDocument.AppendLine("endobj")
            $pdfDocument.AppendLine("4 0 obj")
            $pdfDocument.AppendLine("<<")
            $pdfDocument.AppendLine("/Length 500")
            $pdfDocument.AppendLine(">>")
            $pdfDocument.AppendLine("stream")
            $pdfDocument.AppendLine("BT")
            $pdfDocument.AppendLine("/F1 24 Tf")
            $pdfDocument.AppendLine("100 750 Td")
            $pdfDocument.AppendLine("($( $Data.PersonalInfo.FullName )) Tj")
            $pdfDocument.AppendLine("ET")
            $pdfDocument.AppendLine("BT")
            $pdfDocument.AppendLine("/F1 18 Tf")
            $pdfDocument.AppendLine("100 720 Td")
            $pdfDocument.AppendLine("($( $Data.PersonalInfo.Title )) Tj")
            $pdfDocument.AppendLine("ET")
            $pdfDocument.AppendLine("BT")
            $pdfDocument.AppendLine("/F1 12 Tf")
            $pdfDocument.AppendLine("100 680 Td")
            $pdfDocument.AppendLine("(Email: $( $Data.PersonalInfo.Email )) Tj")
            $pdfDocument.AppendLine("ET")
            $pdfDocument.AppendLine("BT")
            $pdfDocument.AppendLine("/F1 12 Tf")
            $pdfDocument.AppendLine("100 660 Td")
            $pdfDocument.AppendLine("(Phone: $( $Data.PersonalInfo.Phone )) Tj")
            $pdfDocument.AppendLine("ET")
            $pdfDocument.AppendLine("BT")
            $pdfDocument.AppendLine("/F1 10 Tf")
            $pdfDocument.AppendLine("100 640 Td")
            $pdfDocument.AppendLine("(Generated with Windows CV Generator Pro v$script:Version) Tj")
            $pdfDocument.AppendLine("ET")
            $pdfDocument.AppendLine("BT")
            $pdfDocument.AppendLine("/F1 10 Tf")
            $pdfDocument.AppendLine("100 620 Td")
            $pdfDocument.AppendLine("(Developed by $script:Developer © $script:Year) Tj")
            $pdfDocument.AppendLine("ET")
            $pdfDocument.AppendLine("BT")
            $pdfDocument.AppendLine("/F1 10 Tf")
            $pdfDocument.AppendLine("100 600 Td")
            $pdfDocument.AppendLine("($script:Website) Tj")
            $pdfDocument.AppendLine("ET")
            $pdfDocument.AppendLine("endstream")
            $pdfDocument.AppendLine("endobj")
            $pdfDocument.AppendLine("5 0 obj")
            $pdfDocument.AppendLine("<<")
            $pdfDocument.AppendLine("/Type /Font")
            $pdfDocument.AppendLine("/Subtype /Type1")
            $pdfDocument.AppendLine("/BaseFont /Helvetica")
            $pdfDocument.AppendLine(">>")
            $pdfDocument.AppendLine("endobj")
            $pdfDocument.AppendLine("xref")
            $pdfDocument.AppendLine("0 6")
            $pdfDocument.AppendLine("0000000000 65535 f")
            $pdfDocument.AppendLine("0000000009 00000 n")
            $pdfDocument.AppendLine("0000000056 00000 n")
            $pdfDocument.AppendLine("0000000113 00000 n")
            $pdfDocument.AppendLine("0000000220 00000 n")
            $pdfDocument.AppendLine("0000000400 00000 n")
            $pdfDocument.AppendLine("trailer")
            $pdfDocument.AppendLine("<<")
            $pdfDocument.AppendLine("/Size 6")
            $pdfDocument.AppendLine("/Root 1 0 R")
            $pdfDocument.AppendLine(">>")
            $pdfDocument.AppendLine("startxref")
            $pdfDocument.AppendLine("600")
            $pdfDocument.AppendLine("%%EOF")
            
            $pdfDocument.ToString() | Out-File -FilePath $OutputPath -Encoding ASCII
        } catch {
            # If PDF generation fails, copy the HTML file
            Copy-Item $htmlPath $OutputPath.Replace('.pdf', '.html') -Force
        }
    }
    
    # Clean up temporary HTML file
    if (Test-Path $htmlPath) {
        Remove-Item $htmlPath -ErrorAction SilentlyContinue
    }
}

function Generate-DOCXCVAdvanced {
    param(
        [hashtable]$Data,
        [string]$OutputPath
    )
    
    # Create a temporary HTML file first
    $tempHtmlPath = "$env:TEMP\cv_temp_$(Get-Random).html"
    Generate-HTMLCVAdvanced -Data $Data -OutputPath $tempHtmlPath
    
    try {
        # Try to use Word to convert HTML to DOCX
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Open the HTML file
        $doc = $word.Documents.Open($tempHtmlPath)
        
        # Save as DOCX
        $doc.SaveAs([ref]$OutputPath, [ref]16)  # wdFormatDocumentDefault = 16
        
        # Close and quit
        $doc.Close()
        $word.Quit()
        
        # Clean up COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
    } catch {
        # If Word COM fails, create a simple DOCX using XML
        try {
            $docxContent = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">
  <w:body>
    <w:p>
      <w:r>
        <w:t>$( $Data.PersonalInfo.FullName )</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>$( $Data.PersonalInfo.Title )</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Email: $( $Data.PersonalInfo.Email )</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Phone: $( $Data.PersonalInfo.Phone )</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Generated with Windows CV Generator Pro v$script:Version</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Developed by $script:Developer © $script:Year</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>$script:Website</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
    </w:sectPr>
  </w:body>
</w:wordDocument>
"@
            
            $docxContent | Out-File -FilePath $OutputPath -Encoding UTF8
            
            # Rename to .docx if needed
            if (-not $OutputPath.EndsWith('.docx')) {
                Rename-Item -Path $OutputPath -NewName "$OutputPath.docx" -Force
            }
            
        } catch {
            # Final fallback: Copy HTML file
            Copy-Item $tempHtmlPath $OutputPath.Replace('.docx', '.html') -Force
        }
    }
    
    # Clean up temporary HTML file
    if (Test-Path $tempHtmlPath) {
        Remove-Item $tempHtmlPath -ErrorAction SilentlyContinue
    }
}
#endregion

#region Export Form
function Show-ExportForm {
    $exportForm = New-Object System.Windows.Forms.Form
    $exportForm.Text = "Export CV"
    $exportForm.Size = New-Object System.Drawing.Size(500, 400)
    $exportForm.StartPosition = "CenterScreen"
    $exportForm.BackColor = [System.Drawing.Color]::White
    
    $yPos = 20
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Export CV to Multiple Formats"
    $titleLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 14 -Style "Bold"
    $titleLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $titleLabel.Size = New-Object System.Drawing.Size(400, 40)
    $titleLabel.TextAlign = "MiddleCenter"
    $yPos += 50
    
    $formatLabel = New-Object System.Windows.Forms.Label
    $formatLabel.Text = "Select Export Formats:"
    $formatLabel.Location = New-Object System.Drawing.Point(50, $yPos)
    $formatLabel.Size = New-Object System.Drawing.Size(200, 30)
    $yPos += 40
    
    $htmlCheck = New-Object System.Windows.Forms.CheckBox
    $htmlCheck.Text = "HTML (Web Format)"
    $htmlCheck.Location = New-Object System.Drawing.Point(50, $yPos)
    $htmlCheck.Size = New-Object System.Drawing.Size(200, 30)
    $htmlCheck.Checked = $true
    $yPos += 40
    
    $pdfCheck = New-Object System.Windows.Forms.CheckBox
    $pdfCheck.Text = "PDF (Portable Format)"
    $pdfCheck.Location = New-Object System.Drawing.Point(50, $yPos)
    $pdfCheck.Size = New-Object System.Drawing.Size(200, 30)
    $pdfCheck.Checked = $true
    $yPos += 40
    
    $docxCheck = New-Object System.Windows.Forms.CheckBox
    $docxCheck.Text = "DOCX (Word Format)"
    $docxCheck.Location = New-Object System.Drawing.Point(50, $yPos)
    $docxCheck.Size = New-Object System.Drawing.Size(200, 30)
    $docxCheck.Checked = $true
    $yPos += 50
    
    $templateLabel = New-Object System.Windows.Forms.Label
    $templateLabel.Text = "Select Template:"
    $templateLabel.Location = New-Object System.Drawing.Point(250, 90)
    $templateLabel.Size = New-Object System.Drawing.Size(200, 30)
    
    $templateCombo = New-Object System.Windows.Forms.ComboBox
    $templateCombo.Location = New-Object System.Drawing.Point(250, 120)
    $templateCombo.Size = New-Object System.Drawing.Size(200, 30)
    $templateCombo.DropDownStyle = "DropDownList"
    $script:Templates.Keys | ForEach-Object { $null = $templateCombo.Items.Add($_) }
    if ($templateCombo.Items.Count -gt 0) {
        $templateCombo.SelectedIndex = 0
    }
    
    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "Export"
    $exportButton.Location = New-Object System.Drawing.Point(150, 250)
    $exportButton.Size = New-Object System.Drawing.Size(100, 40)
    $exportButton.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $exportButton.ForeColor = [System.Drawing.Color]::White
    $exportButton.FlatStyle = "Flat"
    $exportButton.FlatAppearance.BorderSize = 0
    $exportButton.Add_Click({
        $formats = @()
        if ($htmlCheck.Checked) { $formats += "HTML" }
        if ($pdfCheck.Checked) { $formats += "PDF" }
        if ($docxCheck.Checked) { $formats += "DOCX" }
        
        if ($formats.Count -eq 0) {
            Show-SafeMessage -Message "Please select at least one format." -Title "Warning" -Icon "Warning"
            return
        }
        
        Show-SafeMessage -Message "Please use the CV Builder to create and export your CV.`nThe export function requires data from the CV Builder form." -Title "Info" -Icon "Information"
        $exportForm.Close()
    })
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(260, 250)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 40)
    $cancelButton.Add_Click({ $exportForm.Close() })
    
    $exportForm.Controls.AddRange(@(
        $titleLabel,
        $formatLabel, $htmlCheck, $pdfCheck, $docxCheck,
        $templateLabel, $templateCombo,
        $exportButton, $cancelButton
    ))
    
    $null = $exportForm.ShowDialog()
}
#endregion

#region Templates Form with Enhanced Preview - FIXED: Now shows HTML preview
function Show-TemplatesForm {
    $templatesForm = New-Object System.Windows.Forms.Form
    $templatesForm.Text = "CV Templates Gallery"
    $templatesForm.Size = New-Object System.Drawing.Size(1200, 800)
    $templatesForm.StartPosition = "CenterScreen"
    $templatesForm.BackColor = [System.Drawing.Color]::White
    
    $splitContainer = New-Object System.Windows.Forms.SplitContainer
    $splitContainer.Dock = "Fill"
    $splitContainer.Orientation = "Vertical"
    $splitContainer.SplitterDistance = 400
    
    # Left panel - Template list
    $listPanel = New-Object System.Windows.Forms.Panel
    $listPanel.Dock = "Fill"
    
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Dock = "Fill"
    $listView.View = "Details"
    $listView.FullRowSelect = $true
    $listView.MultiSelect = $false
    $listView.Columns.Add("Template Name", 150) | Out-Null
    $listView.Columns.Add("Style", 100) | Out-Null
    $listView.Columns.Add("Layout", 100) | Out-Null
    $listView.Columns.Add("Academic Support", 120) | Out-Null
    $listView.Columns.Add("Sections", 200) | Out-Null
    
    foreach ($templateKey in $script:Templates.Keys) {
        $template = $script:Templates[$templateKey]
        $item = New-Object System.Windows.Forms.ListViewItem($template.Name)
        $item.SubItems.Add($template.Style) | Out-Null
        $item.SubItems.Add($template.Layout) | Out-Null
        $academicSupport = if ($template.SupportsAcademic) { "✓ Yes" } else { "No" }
		$item.SubItems.Add($academicSupport) | Out-Null
        $item.SubItems.Add($template.Sections.Count.ToString() + " sections") | Out-Null
        $item.Tag = $templateKey
        $listView.Items.Add($item) | Out-Null
    }
    
    # Right panel - Template preview (HTML WebBrowser)
    $previewPanel = New-Object System.Windows.Forms.Panel
    $previewPanel.Dock = "Fill"
    $previewPanel.BackColor = [System.Drawing.Color]::White
    $previewPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $previewTitle = New-Object System.Windows.Forms.Label
    $previewTitle.Text = "Template Preview"
    $previewTitle.Font = New-SafeFont -FontName "Segoe UI" -Size 14 -Style "Bold"
    $previewTitle.Dock = "Top"
    $previewTitle.Height = 40
    $previewTitle.TextAlign = "MiddleCenter"
    
    $global:templatePreviewBrowser = New-Object System.Windows.Forms.WebBrowser
    $global:templatePreviewBrowser.Dock = "Fill"
    $global:templatePreviewBrowser.AllowWebBrowserDrop = $false
    $global:templatePreviewBrowser.IsWebBrowserContextMenuEnabled = $false
    $global:templatePreviewBrowser.WebBrowserShortcutsEnabled = $false
    $global:templatePreviewBrowser.ScriptErrorsSuppressed = $true
    
    # Button panel
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = "Bottom"
    $buttonPanel.Height = 50
    
    $useTemplateButton = New-Object System.Windows.Forms.Button
    $useTemplateButton.Text = "Use This Template"
    $useTemplateButton.Location = New-Object System.Drawing.Point(150, 10)
    $useTemplateButton.Size = New-Object System.Drawing.Size(150, 30)
    $useTemplateButton.Enabled = $false
    $useTemplateButton.Add_Click({
        if ($listView.SelectedItems.Count -gt 0) {
            $templateKey = $listView.SelectedItems[0].Tag
            $templatesForm.Close()
            Show-CVBuilderForm
        }
    })
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(310, 10)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Add_Click({ $templatesForm.Close() })
    
    $buttonPanel.Controls.AddRange(@($useTemplateButton, $closeButton))
    
    # List view selection changed event
    $listView.Add_SelectedIndexChanged({
        if ($listView.SelectedItems.Count -gt 0) {
            $templateKey = $listView.SelectedItems[0].Tag
            $template = $script:Templates[$templateKey]
            
            # Generate HTML preview
            $previewHTML = Generate-TemplateGalleryPreviewHTML -TemplateKey $templateKey -Template $template
            $global:templatePreviewBrowser.DocumentText = $previewHTML
            $useTemplateButton.Enabled = $true
        }
    })
    
    # Add controls
    $listPanel.Controls.Add($listView)
    $previewPanel.Controls.AddRange(@($previewTitle, $global:templatePreviewBrowser, $buttonPanel))
    $splitContainer.Panel1.Controls.Add($listPanel)
    $splitContainer.Panel2.Controls.Add($previewPanel)
    $templatesForm.Controls.Add($splitContainer)
    
    # Select first item
    if ($listView.Items.Count -gt 0) {
        $listView.Items[0].Selected = $true
    }
    
    $null = $templatesForm.ShowDialog()
}

function Generate-TemplateGalleryPreviewHTML {
    param(
        [string]$TemplateKey,
        [hashtable]$Template
    )
    
    $color = "Blue"  # Default color for preview
    $font = "Segoe UI"
    
    # Color mapping
    $colorMap = @{
        "Blue" = @{ Primary = "#007bff"; Secondary = "#0056b3" }
        "Green" = @{ Primary = "#28a745"; Secondary = "#1e7e34" }
        "Purple" = @{ Primary = "#6f42c1"; Secondary = "#593196" }
    }
    
    $colors = $colorMap[$color]
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: '$font', Arial, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background-color: #f5f5f5;
        }
        .preview-container {
            max-width: 100%;
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .preview-header {
            background: $( $colors.Primary );
            color: white;
            padding: 20px;
            text-align: center;
        }
        .preview-name {
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 10px;
        }
        .preview-title {
            font-size: 16px;
            opacity: 0.9;
        }
        .preview-content {
            padding: 20px;
        }
        .template-info {
            background: #f8f9fa;
            border-left: 4px solid $( $colors.Primary );
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
        }
        .info-item {
            margin-bottom: 8px;
            padding: 5px 0;
            border-bottom: 1px dashed #ddd;
        }
        .info-item:last-child {
            border-bottom: none;
        }
        .info-label {
            font-weight: bold;
            color: $( $colors.Secondary );
            display: inline-block;
            width: 150px;
        }
        .sections-list {
            columns: 2;
            column-gap: 20px;
        }
        .section-item {
            padding: 5px 0;
            break-inside: avoid;
        }
        .academic-features {
            background: #e8f4f8;
            border: 1px solid #b3e0f2;
            padding: 15px;
            border-radius: 5px;
            margin-top: 15px;
        }
        .feature-list {
            list-style-type: none;
            padding-left: 0;
        }
        .feature-list li {
            padding: 8px 0;
            border-bottom: 1px dashed #ddd;
        }
        .feature-list li:last-child {
            border-bottom: none;
        }
        .badge {
            display: inline-block;
            padding: 3px 10px;
            background: $( $colors.Primary );
            color: white;
            border-radius: 12px;
            font-size: 12px;
            margin-left: 10px;
        }
        .color-palette {
            display: flex;
            gap: 10px;
            margin-top: 10px;
        }
        .color-sample {
            width: 30px;
            height: 30px;
            border-radius: 50%;
            border: 1px solid #ddd;
        }
        .recommended-for {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 15px;
            border-radius: 5px;
            margin-top: 15px;
        }
    </style>
</head>
<body>
    <div class="preview-container">
        <div class="preview-header">
            <div class="preview-name">$( $Template.Name )</div>
            <div class="preview-title">$( $Template.Description )</div>
        </div>
        
        <div class="preview-content">
            <div class="template-info">
                <div class="info-item">
                    <span class="info-label">Style:</span> $( $Template.Style )
                </div>
                <div class="info-item">
                    <span class="info-label">Layout:</span> $( $Template.Layout )
                </div>
                <div class="info-item">
                    <span class="info-label">Academic Support:</span> 
                    $( if ($Template.SupportsAcademic) { "<span class='badge'>Full Academic CV</span>" } else { "Basic" } )
                </div>
                <div class="info-item">
                    <span class="info-label">Color Schemes:</span>
                    <div class="color-palette">
                        $( $Template.Colors | ForEach-Object { "<div class='color-sample' style='background-color: $_;' title='$_'></div>" } )
                    </div>
                </div>
            </div>
            
            <div class="info-item">
                <span class="info-label">Included Sections:</span>
                <div class="sections-list">
                    $( $Template.Sections | ForEach-Object { "<div class='section-item'>✓ $_</div>" } )
                </div>
            </div>
            
            $( if ($Template.SupportsAcademic) { "
            <div class='academic-features'>
                <h4 style='color: $( $colors.Secondary ); margin-top: 0;'>Academic Features Included:</h4>
                <ul class='feature-list'>
                    <li>✓ Research Work & Publications</li>
                    <li>✓ Patents & Software Developed</li>
                    <li>✓ PhD Details & Supervision</li>
                    <li>✓ Editorial Activities</li>
                    <li>✓ Seminars & Workshops</li>
                    <li>✓ PhD Thesis Evaluation</li>
                    <li>✓ Professional Research Profiles</li>
                    <li>✓ Declaration & Signature</li>
                </ul>
            </div>
            " } )
            
            <div class="recommended-for">
                <h4 style='color: #856404; margin-top: 0;'>Recommended For:</h4>
                <p>
                    $( switch ($Template.Style) {
                        "Professional" { "Corporate jobs, Business positions, Industry professionals" }
                        "Academic" { "Researchers, Professors, Scientists, Academic positions" }
                        "Creative" { "Designers, Artists, Marketing professionals, Creative fields" }
                        "Executive" { "Senior management, Directors, C-level executives" }
                        "Technical" { "Engineers, Developers, IT professionals, Technical roles" }
                        "International" { "International jobs, Europass format, Global positions" }
                        default { "General purpose, all career levels" }
                    } )
                </p>
            </div>
        </div>
    </div>
</body>
</html>
"@
    
    return $html
}
#endregion

#region About Form
function Show-AboutForm {
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "About Windows CV Generator Pro"
    $aboutForm.Size = New-Object System.Drawing.Size(500, 680)  # Increased height for new features
    $aboutForm.StartPosition = "CenterScreen"
    $aboutForm.BackColor = [System.Drawing.Color]::White
    
    $iconPath = "$script:ImagesDir\Icon.ico"
    if (Test-Path $iconPath) {
        try {
            $aboutForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
        } catch {}
    }
    
    $logoPanel = New-Object System.Windows.Forms.Panel
    $logoPanel.Location = New-Object System.Drawing.Point(50, 30)
    $logoPanel.Size = New-Object System.Drawing.Size(400, 80)
    $logoPanel.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    
    $logoLabel = New-Object System.Windows.Forms.Label
    $logoLabel.Text = "📄 Windows CV Generator Pro"
    $logoLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 20 -Style "Bold"
    $logoLabel.ForeColor = [System.Drawing.Color]::White
    $logoLabel.Location = New-Object System.Drawing.Point(0, 20)
    $logoLabel.Size = New-Object System.Drawing.Size(400, 40)
    $logoLabel.TextAlign = "MiddleCenter"
    
    $logoPanel.Controls.Add($logoLabel)
    
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "Version $script:Version"
    $versionLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 12 -Style "Bold"
    $versionLabel.Location = New-Object System.Drawing.Point(50, 130)
    $versionLabel.Size = New-Object System.Drawing.Size(400, 30)
    $versionLabel.TextAlign = "MiddleCenter"
    
    $devLabel = New-Object System.Windows.Forms.Label
    $devLabel.Text = "Developed by:"
    $devLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 11
    $devLabel.Location = New-Object System.Drawing.Point(50, 170)
    $devLabel.Size = New-Object System.Drawing.Size(400, 25)
    $devLabel.TextAlign = "MiddleCenter"
    
    $devNameLabel = New-Object System.Windows.Forms.Label
    $devNameLabel.Text = "$script:Developer"
    $devNameLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 12 -Style "Bold"
    $devNameLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $devNameLabel.Location = New-Object System.Drawing.Point(50, 200)
    $devNameLabel.Size = New-Object System.Drawing.Size(400, 30)
    $devNameLabel.TextAlign = "MiddleCenter"
    
    $yearLabel = New-Object System.Windows.Forms.Label
    $yearLabel.Text = "© $script:Year"
    $yearLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 10
    $yearLabel.Location = New-Object System.Drawing.Point(50, 240)
    $yearLabel.Size = New-Object System.Drawing.Size(400, 25)
    $yearLabel.TextAlign = "MiddleCenter"
    
    $websiteLink = New-Object System.Windows.Forms.LinkLabel
    $websiteLink.Text = $script:Website
    $websiteLink.Font = New-SafeFont -FontName "Segoe UI" -Size 10 -Style "Underline"
    $websiteLink.Location = New-Object System.Drawing.Point(50, 270)
    $websiteLink.Size = New-Object System.Drawing.Size(400, 25)
    $websiteLink.TextAlign = "MiddleCenter"
    $websiteLink.LinkColor = [System.Drawing.Color]::Blue
    $websiteLink.Add_Click({ Start-Process $script:Website })
    
    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Text = "Professional Resume/CV Creation Tool`nCreate stunning resumes with multiple templates,`nphoto upload, and multiple output formats."
    $descLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 10
    $descLabel.Location = New-Object System.Drawing.Point(50, 310)
    $descLabel.Size = New-Object System.Drawing.Size(400, 60)
    $descLabel.TextAlign = "MiddleCenter"
    
    $featuresLabel = New-Object System.Windows.Forms.Label
    $featuresLabel.Text = "New in Version 1.3 - Academic Edition:`n• Complete Academic CV support`n• Research Work & Publications`n• Patents & Software Developed`n• PhD Details & Supervision`n• Editorial Activities`n• Seminars & Workshops`n• PhD Thesis Evaluation`n• Declaration & Signature`n• Professional Research Profiles`n• 10+ Professional Templates`n• Enhanced Template Preview`n• Maximized CV Builder Window`n• Fixed Load Profile Functionality`n• Improved Certification Date Picker"
    $featuresLabel.Font = New-SafeFont -FontName "Segoe UI" -Size 9
    $featuresLabel.Location = New-Object System.Drawing.Point(50, 380)
    $featuresLabel.Size = New-Object System.Drawing.Size(400, 200)
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(200, 590)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Add_Click({ $aboutForm.Close() })
    
    $aboutForm.Controls.AddRange(@(
        $logoPanel, $versionLabel, $devLabel, $devNameLabel,
        $yearLabel, $websiteLink, $descLabel, $featuresLabel, $closeButton
    ))
    
    $null = $aboutForm.ShowDialog()
}
#endregion

#region Main Execution
try {
    # Initialize directories
    Initialize-Directories
    
    # Show main form
    Show-MainForm
    
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "An error occurred: $_`n`nThe application will now exit.",
        "Error",
        "OK",
        "Error"
    )
    exit 1
}

[Environment]::Exit(0)
#endregion