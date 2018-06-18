<#
.REQUIREMENTS
Requires PnP-PowerShell version 2.27.1806.1 or later
https://github.com/OfficeDev/PnP-PowerShell/releases

.SYNOPSIS
Set a custom theme for a specific classic team site

.EXAMPLE
#Set a custom theme / composed look on for a classic experience
PS C:\> .\Set-SPTheme.ps1 -TargetSiteUrl "https://intranet.mydomain.com/sites/targetSite" 
#Set a modern theme, that must already exist on a classic or modern site
PS C:\> .\Set-SPTheme.ps1 -TargetSiteUrl "https://intranet.mydomain.com/sites/targetSite" -Theme "theme name"

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> .\Set-SPTheme.ps1 -TargetSiteUrl "https://intranet.mydomain.com/sites/targetSite" -ThemeName "Custom Theme Name" -MasterUrl "oslo.master" -Credentials $creds
PS C:\> .\Set-SPTheme.ps1 -TargetSiteUrl "https://intranet.mydomain.com/sites/targetSite" -Theme "Existing Modern Theme" -Credentials $creds
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the URL of the target asset location, i.e. site collection root web, e.g. 'https://intranet.mydomain.com/sites/targetWeb'")]
    [String]
    $targetSiteUrl,

	[Parameter(Mandatory = $false, HelpMessage="The composed look name to create and/or set. If not provided for a classic site, a composed look will not be created")]
    [String]
	$themeName,

    [Parameter(Mandatory = $false, HelpMessage="The theme master page url, relateive to Master Page Gallery of the target Web Url. Defaults to seattle.master")]
    [String]
	$masterUrl,
	
	[Parameter(Mandatory = $false, HelpMessage="The available modern theme name to apply to the site")]
    [String]
    $theme,

    [Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    [PSCredential]
    $Credentials
)

if($Credentials -eq $null)
{
	$Credentials = Get-Credential -Message "Enter Admin Credentials"
}
if ($masterUrl -eq "")
{
    $masterUrl = "seattle.master"
}

Write-Host -ForegroundColor White "--------------------------------------------------------"
Write-Host -ForegroundColor White "|                   Set Custom Theme                   |"
Write-Host -ForegroundColor White "--------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Yellow "Target site: $($targetSiteUrl)"
Write-Host ""

try
{
	#set up general variables
	$rootPath = $targetSiteUrl.Substring($targetSiteUrl.IndexOf('/',8))
	$relativeWebUrl = $rootPath

	#$themeName = "Custom Classic Theme"	
	$colorPaletteUrl = "$rootPath/_catalogs/theme/15/custom.theme.spcolor"
	$fontSchemeUrl = "$rootPath/_catalogs/theme/15/custom.theme.spfont"
	$bgImageUrl = "$rootPath/SiteAssets/custom.theme.bg.jpg"

	Connect-PnPOnline $targetSiteUrl -Credentials $Credentials
	
	#get the web template of this current web "GROUP" ==  modern
	$web = Get-PnPWeb -includes webtemplate
	$webTemplate = $web.webtemplate
	$modern = if ($webTemplate -eq "GROUP" -or $webTemplate -eq "SITEPAGEPUBLISHING") { $true } else { $false }


	#if a modern site, 
	if ($modern -eq $true)
	{
		Write-Host -ForegroundColor Yellow "Configuring theming for modern experiences to $($targetSiteUrl)"

		if ($theme -ne "")
		{
			Write-Host -ForegroundColor White "Setting modern theme: $($theme)"
			Set-PnPWebTheme -Theme $theme
			Write-Host -ForegroundColor Green "Theme set"
		}
		else {
			Write-Host -ForegroundColor Yellow "Modern theme name required to set theme"
		}
	}
	else {
		Write-Host -ForegroundColor Yellow "Configuring theming for classic experiences to $($targetSiteUrl)"

		if ($theme -ne "") {
			Write-Host -ForegroundColor White "Applying a modern theme to $($targetSiteUrl)"
			
			Set-PnPWebTheme -Theme $theme
		}
		else {
			Write-Host -ForegroundColor White "Applying a composed look to $($targetSiteUrl)"

			Write-Host -ForegroundColor White "Provisioning asset files to $($targetSiteUrl)"
			Apply-PnPProvisioningTemplate -Path .\Custom.SPTheme.Infrastructure.xml -Handlers Files

			if ($themeName -ne "") {
				#add componsed look only if not already added
				$composedLook = Get-PnPListItem -List "Composed Looks" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$themeName</Value></Eq></Where></Query></View>"
				if($composedLook -eq $null) {
					Write-Host -ForegroundColor White "Adding composed look $($targetSiteUrl)"
					Add-PnPListItem -List "Composed Looks" -ContentType "Item" -Values @{"Title"=$themeName; "Name"=$themeName; "MasterPageUrl"=$relativeWebUrl+"/_catalogs/masterpage/seattle.master, "+$relativeWebUrl+"/_catalogs/masterpage/seattle.master"; "ThemeUrl"=$colorPaletteUrl+", "+$colorPaletteUrl+""; "FontSchemeUrl"=$fontSchemeUrl+", "+$fontSchemeUrl+""; "ImageUrl"=$bgImageUrl+", "+$bgImageUrl+""; "DisplayOrder"="1"}
					Write-Host -ForegroundColor Green "Composed look added to: $($targetSiteUrl)"
				}
				else {
					Write-Host -ForegroundColor Yellow "Composed look already found. Ignored."
				}
			}

			Write-Host -ForegroundColor White "Setting theme for $($targetSiteUrl)"
		
			#https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/set-pnptheme?view=sharepoint-ps
			Set-PnPTheme -ColorPaletteUrl $colorPaletteUrl -FontSchemeUrl $fontSchemeUrl -BackgroundImageUrl $bgImageUrl

			if ($themeName -ne "") {
				#set current composed look
				$composedLookCurrent = Get-PnPListItem -List "Composed Looks" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Current</Value></Eq></Where></Query></View>"
				if($composedLookCurrent -ne $null) {
					Write-Host -ForegroundColor White "Setting current composed look"
					Set-PnPListItem -List "Composed Looks" -Identity $composedLookCurrent -Values @{"MasterPageUrl"=$relativeWebUrl+"/_catalogs/masterpage/seattle.master, "+$relativeWebUrl+"/_catalogs/masterpage/seattle.master"; "ThemeUrl"=$colorPaletteUrl+", "+$colorPaletteUrl+""; "FontSchemeUrl"=$fontSchemeUrl+", "+$fontSchemeUrl+""; "ImageUrl"=$bgImageUrl+", "+$bgImageUrl+""}
				}
				else {
					Write-Host -ForegroundColor Yellow "Unable to find current composed look"
				}
			}

			Write-Host -ForegroundColor Green "Theme set for $($targetSiteUrl)"

			#now set the master page
			$masterUrl = "$rootPath/_catalogs/masterpage/$masterUrl"
			Write-Host -ForegroundColor White "Setting master page to $($masterUrl)"

			Set-PnPWeb -MasterUrl $masterUrl
			
			Write-Host -ForegroundColor Green "Master page set for $($targetSiteUrl)"
		}

		
	}

	Write-Host ""
	Write-Host -ForegroundColor Green "Theming applied"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}