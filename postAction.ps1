# Helper for VSTO deploy, post action
# CWD is project folder of the addin
#
# Accepts fully qualified postAction class as parameter
# Optionaly accepts version number (defaults to 1.0.0.0)
# Adds a postActions Node to the current version of the application manifest
# Uses pfx file in the current project
# signs the application manifest
# updates the deployment manifest
# copy the .vsto file to the current version Application Files folder

param(
    [string] $postActionClass,
    [string] $postActionVer = "1.0.0.0"
)

function appendAttr([xml]$doc, [Xml.XmlNode]$node, [string]$name, [string]$value) {
    
    $Attr = $doc.CreateAttribute($name)
    $Attr.value = $value
    $node.Attributes.Append($Attr)
    return $node
}

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

# find the current application version
$CurrentApplicationManifestPath = Join-Path $publishPath $deploymentManifest.assembly.dependency.dependentAssembly.codebase
[xml]$CurrentApplicationManifest = Get-Content -Path $CurrentApplicationManifestPath

# build the postActions node and add it to the application m/fest after the </vstav3:update> element
$doc = $CurrentApplicationManifest
$vstav3NS = "urn:schemas-microsoft-com:vsta.v3"
$postActionsNode = $doc.CreateElement("postActions", $vstav3NS)
[Xml.XmlNode]$postAction = $postActionsNode.AppendChild($doc.CreateElement("postAction", $vstav3NS))
[Xml.XmlNode]$entryPoint = $postAction.AppendChild($doc.CreateElement("entryPoint", $vstav3NS))
[Xml.XmlNode]$assemblyIdentity = $entryPoint.AppendChild($doc.CreateElement("assemblyIdentity", $vstav3NS))

appendAttr $doc $assemblyIdentity "name" $postActionClassName
appendAttr $doc $assemblyIdentity "version" $postActionVer
appendAttr $doc $assemblyIdentity "language" "neutral"
appendAttr $doc $assemblyIdentity "processorArchitecture" "msil"

appendAttr $doc $entryPoint "class" $postActionClass

$postActionsNode.postAction.AppendChild($doc.CreateElement("postActionData", $vstav3NS))

Write-Debug $postActionsNode.OuterXml

# Replace any existing version
$addin = $doc.assembly.addIn
$prevPANode = $addin.GetElementsByTagName("postActions", $vstav3NS)[0]
if($prevPANode -eq $null){
    $addin.InsertAfter($postActionsNode, $doc.GetElementsByTagName("update", $vstav3NS))
}else{
    $addin.ReplaceChild($postActionsNode, $prevPANode)
}

$doc.Save($CurrentApplicationManifestPath)

$Certificate = $ProjectFolder | Where-Object {$_.Extension -eq $certExtension}

