Add-Type -AssemblyName system.io.compression.FileSystem
$SourceFiles = "C:\Users\Microsoft\Documents\sortfiles"
$ZipOutput = "C:\Output"
$Dirs = Get-ChildItem -Path "C:\Users\Microsoft\Documents\caseId\" -Directory -Exclude done, .ipynb_checkpoints

foreach ($Dir in $Dirs) {

    #Copy files in
    Copy-Item ($SourceFiles + '\createArrowHTMLPages.py') -Destination $Dir.FullName

    # Kick off the html script
    Set-Location $Dir.FullName
    Python createArrowHTMLPages.py

    # Now Copy the ToC
    $cpSource = [string]::concat($sourceFiles,'\generateToC.py')
    Copy-Item $cpSource -Destination ( $Dir.FullName + '\html\extractedText' )
    Copy-Item $cpSource -Destination ( $Dir.FullName + '\html\regularFiles' )

    # Then Execute
    Set-Location ($Dir.FullName + '\html\regularFiles')
    Python generateToC.py

    Set-Location ($Dir.FullName + '\html\extractedText')
    Python generateToC.py

    #Remove Files
    Remove-Item ([string]::Concat( $Dir.FullName,'\html\extractedText\generateToC.py'))
    Remove-Item ([string]::concat($Dir.FullName,'\html\regularFiles\generateToC.py'))

    #zip the File and place in zip output
    $Dest = [string]::concat($ZipOutput,'\',$Dir.Name,'.zip')
    $Source = [string]::concat($Dir.FullName,'\html')
    [System.IO.Compression.ZipFile]::CreateFromDirectory($Source,$Dest)

    #Move the Directory after being done
    Set-Location $Dir.Parent.FullName
    Move-Item -Path $Dir.FullName -Destination ([string]::Concat($Dir.Parent.FullName,'\done') ) 
}