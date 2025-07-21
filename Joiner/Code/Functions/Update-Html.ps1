# Description: Initalises a HTML object and also updates it with any data passed.
#              The data will be presented in a table.
# Accepts: [Hashtable]$Data - The data that needs to be added to the table.
#          [Switch]$Initialise - Initialises the object containing the HTML. This should only be done once.
# Returns: Null
Function Update-Html
{
    param(
        [Parameter(Mandatory = $True)][HashTable]$Data#,
        #[Parameter(Mandatory = $True, ParameterSetName = "Assyst")][String]$AssystRef
    )

    Function Update-HtmlRowColor
    {
        If ($Script:RowColour -eq "#b6b6ba")
        {
            $Script:RowColour = "#d8d8e0"
        }
        else
        {
            $Script:RowColour = "#b6b6ba"
        }
    }

    If ($Null -eq $Script:Html)
    {

        $Script:html += "<html><head><title>New Starter Provisioning</title>"
        $Script:html += "<style>"
        $Script:html += ".ABFTextPink{color:#ce003d;font-weight:bold;}"
        $Script:html += ".ABFTextGrey{color:#868689;font-weight:bold;}"
        $Script:html += "body{font-family:arial;}"
        $Script:html += "table{border-spacing:0px;padding:5px;margin-left:auto;margin-right:auto;}"
        $Script:html += ".subTitle{background-color:#ce003d;padding:5px;color:white;}"
        $Script:html += "</style></head>"
        $Script:Html += "<body><table width=1000px><tr><td height='100px'><div class='ABFTextPink'>New Starter Provisioning</div></td>"
        $Script:Html += "<td style=text-align:right;><div class='ABFTextPink' style=display:inline-block;vertical-align:middle;text-align:right;text-size:9pt;>Created on: </div><div class='ABFTextGrey' style=display:inline-block;vertical-align:middle;text-align:right;text-size:9pt;>" + (Get-Date -Format "dd/MM/yy HH:mm:ss") + "</div>"
        $Script:Html += "<tr><td Class='subTitle'>Attribute</td><td Class='subTitle'>Val ueSet</td></tr>"
    }

    Foreach ($Row in $Data.GetEnumerator())
    {
        Switch($Row.Name)
        {
            #Hashtables processed here
            (@("OtherAttributes") -like $_)
            {
                $Row = $Data[$Row.Name]
 
                Foreach($Item in $Row.GetEnumerator())
                {
                    Update-HtmlRowColor

                    $Script:Html += "<tr>"
                    $Script:Html += "<td style=background-color:$Script:RowColour>$($Item.Key)</td>"
                    $Script:Html += "<td style=background-color:$Script:RowColour>$($Item.Value)</td>"
                    $Script:Html += "</tr>"
                }
            }
 
            #Arrays processed here
            "Groups Added"
            {
                Update-HtmlRowColor
                Write-Host ("Array Entry is " + $Row.Name)
                $ArrayValue = ""
                Foreach($Item in $Data[$Row.name])
                {
                    Write-Host $Item
                    $ArrayValue += $Item + "<br/>"
                }
 
                $Script:Html += "<tr>"
                $Script:Html += "<td style=background-color:$Script:RowColour>$($Row.Name)</td>"
                $Script:Html += "<td style=background-color:$Script:RowColour>$ArrayValue</td>"
                $Script:Html += "</tr>"
            }
 
            Default
            {
                Update-HtmlRowColor
                Write-Host ("Entry is " + $Row.Name)
                $Script:HtmlObject =@{
                    LeftColumn = $Row.Name
                    RightColumn = $Row.Value
                }
 
                $Script:Html += "<tr>"
                $Script:Html += "<td style=background-color:$Script:RowColour>$($Row.Name)</td>"
                $Script:Html += "<td style=background-color:$Script:RowColour>$($Row.Value)</td>"
                $Script:Html += "</tr>"
            }
        }
    }

    $Script:Html += "</table></body></html>"
    $HtmlPath = "\\pxgbssc1aop001\d$\Automation\NewStarter\TestFiles\HtmlTemporaryStorage\NewStarter-$TranscriptingGuid.html"
    $Script:Html | Out-File $HtmlPath

    try
    {
        #Add-AssystAttachment -IncidentId $AssystRef -Path $HtmlPath -AssystEnvironment $Environment -Description "New Start Provisioning Log for attempt $ProvisionAttemptGuid"
        #Remove-Item $HtmlPath
    }
    catch
    {
        Write-Error -Message $_.Exception.Message
    }
}