# Helper for VSTO deploy, post action
# CWD is project folder of the addin
#
# Accepts fully qualified postAction class as parameter
# Optionaly accepts version number (defaults to 1.0.0.0)
# Adds a postActions Node to the current version of the application manifest
# Uses pfx file in the current project
#  signs the application manifest
#  updates the deployment manifest trust
# copy the .vsto file to the current version Application Files folder

param(
    [string] $postActionClass,
    [string] $postActionVer = "1.0.0.0"
)

$CWD = $PWD

if($postActionClass.IndexOf(".") -gt -1){
    $postActionClassName = $postActionClass.Split(".")[1]
}else{
    $postActionClassName = $postActionClass
}

# Get the publish path from the .csproj file

$configExtension = ".csproj"
$certExtension = ".pfx"

$ProjectFolder = Get-ChildItem $PWD
$ConfigFile = $ProjectFolder | Where-Object {$_.Extension -eq $configExtension}
[xml]$Config = Get-Content $ConfigFile.Name

[string]$publishPath = ($Config.Project.PropertyGroup | Where-Object {$_.PublishUrl -ne $null}).PublishUrl
Write-Debug $publishPath

# Add a postActions Node to the current version of the application manifest

# Get the deployment manifest
$publishFolder = Get-ChildItem -Path $publishPath
$deploymentManifestFile = $publishFolder | Where-Object {$_.Extension -eq ".vsto"}

[xml]$deploymentManifest = Get-Content $deploymentManifestFile.FullName

# find the current application version and get the application manifest
$CurrentApplicationManifestPath = Join-Path $publishPath $deploymentManifest.assembly.dependency.dependentAssembly.codebase
[xml]$CurrentApplicationManifest = Get-Content -Path $CurrentApplicationManifestPath

# build the postActions node and add it to the application m/fest after the </vstav3:update> element
$doc = $CurrentApplicationManifest

[xml]$doc = @"
<xml xmlns:vstav3="urn:schemas-microsoft-com:vsta.v3">
    <vstav3:postActions>
      <vstav3:postAction>
        <vstav3:entryPoint
          class="FileCopyPDA.FileCopyPDA">
          <assemblyIdentity
            name="FileCopyPDA"
            version="1.0.0.0"
            language="neutral"
            processorArchitecture="msil" />
        </vstav3:entryPoint>
        <vstav3:postActionData>
        </vstav3:postActionData>
      </vstav3:postAction>
    </vstav3:postActions>
</xml>
"@
[Xml.XmlNode]$postActionsNode = $CurrentApplicationManifest.ImportNode($doc.xml.postActions, $true)
[Xml.XmlNode]$entryPoint = $postActionsNode.postAction.entryPoint
[Xml.XmlNode]$assemblyIdentity = $entryPoint.assemblyIdentity

$entryPoint.class = $postActionClass
$assemblyIdentity.name = $postActionClassName
$assemblyIdentity.version = $postActionVer

$sw=New-Object System.Io.Stringwriter
$writer=New-Object System.Xml.XmlTextWriter($sw)
$writer.Formatting = [System.Xml.Formatting]::Indented
$doc.WriteContentTo($writer)
$sw.ToString()

# Replace any existing version
$addin = $CurrentApplicationManifest.assembly.addIn
$prevPANode = $addin.postActions
if($prevPANode -eq $null){
    $addin.InsertAfter($postActionsNode, $addin.update)
}else{
    $addin.ReplaceChild($postActionsNode, $prevPANode)
}

$CurrentApplicationManifest.Save($CurrentApplicationManifestPath)

# Use the pfx file in the current project

$Certificate = $ProjectFolder | Where-Object {$_.Extension -eq $certExtension}

#  sign the application manifest
mage -sign $CurrentApplicationManifestPath -certfile $Certificate.FullName

#  update the deployment manifest trust
mage -update $deploymentManifestFile.FullName  -appmanifest $CurrentApplicationManifestPath -certfile $Certificate.FullName