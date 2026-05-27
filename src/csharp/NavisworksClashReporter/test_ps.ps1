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
.Load(, [System.Text.Encoding]::UTF8)
 = .DocumentNode.SelectNodes('//div[contains(@class, "viewpoint")]')

foreach ( in ) {
     = .SelectSingleNode('.//h2')
    if () {
         = .InnerText.Trim()
         = [HtmlAgilityPack.HtmlEntity]::DeEntitize()
        
         =  -replace '[^\p{L}\p{N}]', ''
        
         = False
        foreach ( in ) {
             =  -replace '[^\p{L}\p{N}]', ''
            if ( -eq ) {
                 = True
                break
            }
        }
        if (-not ) {
            Write-Host "MISMATCH: "
            break
        }
    }
}
Write-Host "Done checking"
