Add-Type -Path 'D:\itti\mtc\ICB_Coding_Workspace\src\csharp\plugins\NavisworksClashReporter\packages\HtmlAgilityPack.1.11.71\lib\Net45\HtmlAgilityPack.dll'
[System.Reflection.Assembly]::LoadWithPartialName('System.Xml.Linq') | Out-Null

 = 'D:\itti\mtc\pluginresultcompare\Clash Detection for 100 BD_260518_XML.xml'
 = 'D:\itti\mtc\pluginresultcompare\html\Clash Detection for 100 BD_260518.html'

 = [System.Xml.Linq.XDocument]::Load()
 = @()
foreach ( in .Descendants('view')) {
     += .Attribute('name').Value.Trim()
}

 = New-Object HtmlAgilityPack.HtmlDocument
.Load() # Default encoding, just like C# code currently does
 = .DocumentNode.SelectNodes('//div[contains(@class, "viewpoint")]')

 = 0
foreach ( in ) {
     = .SelectSingleNode('.//h2')
    if () {
         = .InnerText.Trim()
         = [HtmlAgilityPack.HtmlEntity]::DeEntitize()
        
         =  -replace '[^\w]', ''
        
         = False
        foreach ( in ) {
             =  -replace '[^\w]', ''
            if ( -eq ) {
                 = True
                break
            }
        }
        if (-not ) {
            Write-Host "MISMATCH: "
            Write-Host "Loose HTML: "
            ++
            if ( -ge 5) { break }
        }
    }
}
Write-Host "Total Failures: "
