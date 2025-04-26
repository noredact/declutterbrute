#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, StdOut

; Main script for Desktop Cleanup Utility
class DesktopCleanup {
    static Version := "1.0"
    static AppName := "Desktop Cleanup Utility"
    
    static DesktopPath := A_Desktop
    static ArchiveRoot := A_Desktop "\Archive"
    static SortFolder := A_Desktop "\Sorted Files"
    
    ; File type categories (extend as needed)
    static FileCategories := Map(
        "Documents", ["doc", "docx", "pdf", "txt", "rtf", "odt"],
        "Spreadsheets", ["xls", "xlsx", "csv", "ods"],
        "Presentations", ["ppt", "pptx", "odp"],
        "Images", ["jpg", "jpeg", "png", "gif", "bmp", "svg", "webp"],
        "Audio", ["mp3", "wav", "ogg", "flac", "m4a"],
        "Video", ["mp4", "mov", "avi", "mkv", "wmv"],
        "Archives", ["zip", "rar", "7z", "tar", "gz"],
        "Executables", ["exe", "msi", "ahk", "bat", "cmd"],
        "Code", ["ahk", "py", "js", "html", "css", "json", "xml", "ini", "config"]
    )
    
    ; Initialize and run the application
    static Run() {
        this.CreateNeededFolders()
        files := this.GetDesktopFiles()
        
        if !files.Length {
            MsgBox "No files found on desktop to organize.", this.AppName, "Iconi"
            return
        }
        
        this.ShowFileSummary(files)
    }
    
    ; Create all needed directory structure
    static CreateNeededFolders() {
        ; Create root folders if they don't exist
        for folder in [this.ArchiveRoot, this.SortFolder] {
            if !DirExist(folder)
                DirCreate(folder)
        }
        
        ; Create subfolders for file categories
        for category in this.FileCategories {
            path := this.SortFolder "\" category
            if !DirExist(path)
                DirCreate(path)
        }
    }
    
    ; Get all files on desktop (ignoring directories)
    static GetDesktopFiles() {
        files := []
        loop files, this.DesktopPath "\*", "F" {
            ; Skip our own archive and sorted files folders
            if InStr(A_LoopFilePath, this.ArchiveRoot) || InStr(A_LoopFilePath, this.SortFolder)
                continue
            files.Push(A_LoopFilePath)
        }
        return files
    }
    
    ; Show GUI with file summary and options
    static ShowFileSummary(files) {
        ; Analyze files
        fileStats := this.AnalyzeFiles(files)
        totalFiles := files.Length
        
        ; Create GUI
        myGui := Gui(, this.AppName " v" this.Version)
        myGui.OnEvent("Close", (*) => ExitApp())
        myGui.SetFont("s10", "Segoe UI")
        
        ; Header
        myGui.Add("Text", "w600 Center", "Desktop Cleanup Utility will organize " totalFiles " files")
        myGui.Add("Text", "w600 Center", "The original files will be moved to: " this.SortFolder)
        myGui.Add("Text", "w600 Center", "Shortcuts will be created in: " this.ArchiveRoot)
        myGui.Add("Text", "w600 Center", "`nFile types to be organized:")
        
        ; File type statistics
        statsText := ""
        for category, count in fileStats {
            statsText .= "`n" category ": " count " files"
        }
        myGui.Add("Text", "w600 Center", statsText)
        
        ; Action buttons
        btnContinue := myGui.Add("Button", "w120", "Continue")
        ; btnContinue := btnRow.Add("Button", "Default w120", "&Continue")
        btnContinue.OnEvent("Click", (*) => myGui.Destroy())
        btnCancel := myGui.Add("Button", "xp+130 w120", "&Cancel")
        btnCancel.OnEvent("Click", (*) => ExitApp())
        
        myGui.Show()
        
        ; Wait for user to continue
        WinWaitClose(myGui)
        
        ; Proceed with file organization
        this.OrganizeFiles(files)
    }
    
    ; Helper method to check if value exists in array
    static HasValue(arr, value) {
        for item in arr {
            if (item = value)
                return true
        }
        return false
    }
    
    ; Analyze files and return statistics
    static AnalyzeFiles(files) {
        stats := Map()
        
        ; Initialize categories
        for category in this.FileCategories {
            stats[category] := 0
        }
        stats["Other"] := 0
        
        ; Count files by category
        for file in files {
            ext := LTrim(SubStr(file, InStr(file, ".", , -1)), ".") ; Get extension
            ext := StrLower(ext) ; Convert to lowercase for case-insensitive comparison
            
            ; Skip files with no extension
            if (ext = "") {
                stats["Other"]++
                continue
            }
            
            categorized := false
            
            ; Check each category for matching extensions
            for category, exts in this.FileCategories {
                if this.HasValue(exts, ext) {
                    stats[category]++
                    categorized := true
                    break
                }
            }
            
            if !categorized
                stats["Other"]++
        }
        
        return stats
    }
    
    ; Main file organization logic
    static OrganizeFiles(files) {
        ; Progress GUI
        progressGui := Gui(, "Organizing Files...")
        progressGui.SetFont("s10", "Segoe UI")
        progressGui.Add("Text", "w400", "Moving and organizing files...")
        progressBar := progressGui.Add("Progress", "w400 h20 Range0-" files.Length)
        statusText := progressGui.Add("Text", "w400", "Preparing...")
        progressGui.Show()
        
        movedFiles := 0
        errors := 0
        
        ; Process each file
        for i, file in files {
            try {
                statusText.Text := "Processing: " SubStr(file, InStr(file, "\", , -1) + 1)
                progressBar.Value := i
                
                ; Get file info
                SplitPath(file, &name, &dir, &ext, &nameNoExt)
                ext := LTrim(ext, ".")
                
                ; Skip files with no extension
                if (ext = "") {
                    FileAppend "Skipping file with no extension: " file "`n", A_Temp "\DesktopCleanupErrors.log"
                    continue
                }
                
                ; Determine category
                category := "Other"
                for cat, exts in this.FileCategories {
                    if this.HasValue(exts, StrLower(ext)) {
                        category := cat
                        break
                    }
                }
                
                ; Move file to categorized folder
                destDir := this.SortFolder "\" category
                destPath := destDir "\" name
                
                if !DirExist(destDir)
                    DirCreate(destDir)
                
                FileMove(file, destPath, 1) ; 1 = overwrite
                
                ; Create various shortcut types
                this.CreateShortcuts(destPath, file)
                
                movedFiles++
            } catch as e {
                errors++
                FileAppend "Error processing " file ": " e.Message "`n", A_Temp "\DesktopCleanupErrors.log"
            }
        }
        
        progressGui.Destroy()
        
        ; Show completion message
        msg := "Operation complete!`n`n"
        msg .= "Files moved: " movedFiles "`n"
        msg .= "Errors encountered: " errors "`n"
        if errors
            msg .= "See " A_Temp "\DesktopCleanupErrors.log for details"
        
        MsgBox(msg, this.AppName, "Iconi")
    }
    
    ; Create all shortcut types for a file
    static CreateShortcuts(newPath, originalPath) {
        ; Get file info
        SplitPath(newPath, &name, &dir, &ext, &nameNoExt)
        ext := LTrim(ext, ".")
        
        ; Get timestamps
        modifiedTime := FileGetTime(newPath, "M")
        createdTime := FileGetTime(newPath, "C")
        movedTime := FormatTime(, "yyyyMMdd")
        
        modifiedYear := SubStr(modifiedTime, 1, 4)
        modifiedMonth := SubStr(modifiedTime, 5, 2)
        createdYear := SubStr(createdTime, 1, 4)
        createdMonth := SubStr(createdTime, 5, 2)
        
        ; Create Modified time-based shortcut
        modPath := this.ArchiveRoot "\Modified\" modifiedYear "\" modifiedMonth
        if !DirExist(modPath)
            DirCreate(modPath)
        this.CreateShortcut(newPath, modPath "\" name ".lnk")
        
        ; Create Created time-based shortcut
        createdPath := this.ArchiveRoot "\Created\" createdYear "\" createdMonth
        if !DirExist(createdPath)
            DirCreate(createdPath)
        this.CreateShortcut(newPath, createdPath "\" name ".lnk")
        
        ; Create Moved time-based shortcut
        movedPath := this.ArchiveRoot "\Moved\" SubStr(movedTime, 1, 4) "\" SubStr(movedTime, 5, 2)
        if !DirExist(movedPath)
            DirCreate(movedPath)
        this.CreateShortcut(newPath, movedPath "\" name ".lnk")
        
        ; Create Alphabetical shortcut
        firstChar := SubStr(nameNoExt, 1, 1)
        if firstChar ~= "[0-9]"
            alphaCat := "0-9"
        else if firstChar ~= "[A-Fa-f]"
            alphaCat := "A-F"
        else if firstChar ~= "[G-Mg-m]"
            alphaCat := "G-M"
        else if firstChar ~= "[N-Tn-t]"
            alphaCat := "N-T"
        else
            alphaCat := "U-Z"
        
        alphaPath := this.ArchiveRoot "\Alphabetical\" alphaCat
        if !DirExist(alphaPath)
            DirCreate(alphaPath)
        this.CreateShortcut(newPath, alphaPath "\" name ".lnk")
    }
    
    ; Helper to create a shortcut
    static CreateShortcut(target, linkPath) {
        try {
            FileCreateShortcut(target, linkPath)
        } catch as e {
            FileAppend "Error creating shortcut " linkPath ": " e.Message "`n", A_Temp "\DesktopCleanupErrors.log"
        }
    }
}

; Start the application
DesktopCleanup.Run()