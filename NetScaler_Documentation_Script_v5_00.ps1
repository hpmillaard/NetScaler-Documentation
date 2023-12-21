#Requires -Version 3.0
#This File is in Unicode format. Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Creates a complete inventory of a NetScaler configuration using Microsoft Word.
.DESCRIPTION
	Creates a complete inventory of a NetScaler configuration using Microsoft Word and PowerShell.
	Creates a Word document named after the NetScaler Configuration.
	Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

	Script requires at least PowerShell version 3 but runs best in version 5.

.PARAMETER NSIP
	NetScaler IP address, could be NSIP or SNIP with management enabled
.PARAMETER Credential
	Accepts a PSCredential object with Username and Password to be used to Authenticate to the NetScaler Appliance
.PARAMETER NSUsername
	Accepts a Username to authenticate to the NetScaler Appliance
.PARAMETER NSPassword
	Accepts a Password to authenticate to the NetScaler Appliance.
	Note: It is recommended to create a PSCredential object for stored authentication as storing passwords in plaintext is inherrently insecure.
.PARAMETER UseNSSSL
	EXPERIMENTAL: Require SSL/TLS for communication with the NetScaler Nitro API. 
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
	If either registry key does not exist and this parameter is not specified, the report will
	not contain a Company Name on the cover page.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly works in 2010 but 
			Subtitle/Subject & Author fields need to be moved after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be manually resized or font 
			changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be manually resized or font 
			changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually resized or font 
			changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	Default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be ReportName_2020-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER ReportFileName
	Specifies the optional name for the output report.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	Default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	Default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER Offline
	ALIAS Export
	This disables the detection of MS Word and exports the API information to XML Files
	stored by default in an 'ADCDocsExport' subfolder in the directory in the script folder. 
	This location can be overridden with the OfflinePath or ExportPath Parameter
.PARAMETER OfflinePath
	This overrides the path that the offline XML files are exported to when run with the Offline parameter
.PARAMETER Import
	This generates the word output of the script using the XML files captured when running in Offline or Export mode.
	By default this will load content from an ADCDocsExport subfolder in the script working directory. This path 
	can be overridden using the ImportPath parameter.

	IMPORTANT: When running in Import mode, the NSIP of the appliance that the data was exported from MUST be provided 
	to successfully generate the report.
.PARAMETER ImportPath
	This overrides the location that XML content is loaded from when running in Import mode to generate a report using offline content.
	The import path should be a full path to the folder containing the export data to be used.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\NetScaler_Documentation_Script_v5_00.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\NetScaler_Documentation_Script_v5_00.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -AddDateTime
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be Script_Template_2020-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be Script_Template_2020-06-01_1800.PDF
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -Folder \\FileServer\ShareName
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -SmtpServer mail.domain.tld -From XDAdmin@domain.tld -To ITGroup@domain.tld
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, sending to ITGroup@domain.tld.
	Script will use the default SMPTP port 25 and will not use SSL.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 -UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will use the email server smtp.office365.com on port 587 using SSL, sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.EXAMPLE 
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -Export
	OR
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -Offline

	Will run without MS Word installed and create an export of API data to create a configuration report on another machine. API data will be exported to C:\PSScript\ADCDocsExport\.
.EXAMPLE 
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -Export -ExportPath "C:\ADCExport"
	OR
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -Offline -OfflinePath "C:\ADCExport"

	Will run without MS Word installed and create an export of API data to create a configuration report on another machine. API data will be exported to C:\ADCExport\.
.EXAMPLE 
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -Import

	Will create a configuration report using the API data stored in C:\PSScript\ADCDocsExport.
.EXAMPLE 
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -Import -ImportPath "C:\ADCExport"

	Will create a configuration report using the API data stored in C:\ADCExport.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -NSIP 172.16.20.10 -Credential $MyCredentials

	Will execute the script silently connecting to an ADC appliance on 172.16.20.10 using credentials stored in the PSCredential Object $Mycredentials
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 -NSIP 172.16.20.10 -NSUserName nsroot -NSPassword nsroot

	Will execute the script silently connecting to an ADC appliance on 172.16.20.10 using credentials nsroot/nsroot
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 
	-SmtpServer mail.domain.tld
	-From XDAdmin@domain.tld 
	-To ITGroup@domain.tld	

	The script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.

	The script will use the default SMTP port 25 and will not use SSL.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 
	-SmtpServer mailrelay.domain.tld
	-From Anonymous@domain.tld 
	-To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script will use the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld, sending to ITGroup@domain.tld.

	To send unauthenticated email using an email relay server requires the From email account 
	to use the name Anonymous.

	The script will use the default SMTP port 25 and will not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script will generate an anonymous secure password for the anonymous@domain.tld 
	account.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 
	-SmtpServer labaddomain-com.mail.protection.outlook.com
	-UseSSL
	-From SomeEmailAddress@labaddomain.com 
	-To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multiFunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script will use the email server labaddomain-com.mail.protection.outlook.com, 
	sending from SomeEmailAddress@labaddomain.com, sending to ITGroupDL@labaddomain.com.

	The script will use the default SMTP port 25 and will use SSL.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 
	-SmtpServer smtp.office365.com 
	-SmtpPort 587
	-UseSSL 
	-From Webster@CarlWebster.com 
	-To ITGroup@CarlWebster.com	

	The script will use the email server smtp.office365.com on port 587 using SSL, 
	sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Documentation_Script_v5_00.ps1 
	-SmtpServer smtp.gmail.com 
	-SmtpPort 587
	-UseSSL 
	-From Webster@CarlWebster.com 
	-To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	*** NOTE ***
	
	The script will use the email server smtp.gmail.com on port 587 using SSL, 
	sending from webster@gmail.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.

.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates a Word, PDF, Formatted Text or HTML document.
.NOTES
	NAME: NetScaler_Script_v5_00.ps1
	VERSION NetScaler Script: 5.00
	AUTHOR NetScaler script version 5 and up: Harm Peter Millaard
	AUTHOR NetScaler script until version 4: Barry Schiffer & Andy McCullough
	AUTHOR NetScaler script Functions: Iain Brighton
	AUTHOR Script template: Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters
	LASTEDIT: November 11, 2023
#>

#region changelog
<#
.COMMENT
	If you find issues with saving the final document or table layout is messed up please use the X86 version of Powershell!
	
.Release Notes version 5.00
#	Modes and features synced with 14.1
#	Renamed headings to names in 14.1 and rearranged the order of all paragraphs
#	Removed from final report what is not used or feature is disabled, hide non configured and default settings
#	Reformatted CS, LB, Services and Service Group properties
#	LB redirurl in LB vServer properties (no longer special heading)
#	Add TACACS
#	Fix CS Policy binding prios
#	Add CS Rewrite and Responder policies
#	Add ReportFileName parameter
#	Add Advanced Authentication Policies
#	LB HTTP monitors specified configuration
#	Cipher groups specified
#	CS and LB vserver SSL profile
#	Add SNMP View, Group, Users
#	Default Gateway Portal Themes are removed from the report, only non defaults will be documented.
#	Add system user memberof / system group members and other additional properties
#	Add AAA Groups and AAA Users
#	Add NetScaler Gateway Traffic Policies
#	Changed some naming (Citrix ADC no longer exists, so back to NetScaler)
#	Removed some actions under every item and moved them to an existing function
#	Set the defauls for tables in the function variables instead of parameter for every function call
#	Various bug fixes

.Release Notes version 4.52InvokevNetScalerNitroMethod
#	Add checking for a Word version of 0, which indicates the Office installation needs repairing
#	Change location of the -Dev, -Log, and -ScriptInfo output files from the script folder to the -Folder location (Thanks to Guy Leech for the "suggestion")
#	Remove code to check for $Null parameter values
#	Reformatted the terminating Write-Error messages to make them more visible and readable in the console
#	Remove the SMTP parameterset and manually verify the parameters
#	Update Function SendEmail to handle anonymous unauthenticated email
#	Update Help Text

.Release Notes version 4.51
#	Fix Swedish Table of Contents (Thanks to Johan Kallio)
#		From 
#			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
#		To
#			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
#	Updated help text

.Release Notes version 4.5
#	FIX: Issue connecting to NetScaler when using untrusted certificate on the management interface.
#	NEW: Pass PSCredential object to -Credential parameter to authenticate to NetScaler silently
#	NEW: -NSUserName and -NSPassword paramters allow authentication to NetScaler silently
#	FIX: Fixed issue where some users were prompted for missing parameter when running the script
#	FIX: Modes table had incorrect header values
#	FIX: Output issues for Certificates, ADC Servers, SAML Authentication, Location Database, Network Profiles, NSGW Session Profiles
#	FIX: NSGW Session Profiles using wrong value for SSO Domain
#	FIX: Formatting issue for HTTP Callouts 

.Release Notes version 4.4
#	FIX: Some bindings were not being correctly reported due to incorrect handling of null return values - Thanks to Aaron Kahn for reporting this.

.Release Notes version 4.3
#	Offline Usage - Added ability to export data on a workstation without Word installed and create report on another workstation 

.Release Notes version 4.2
#	FIX: Get-vNetScalerObjectCount always connects using non-SSL - thanks to Eglan Kurek for reporting
#	Added User Administration > Database Users, SMPP Users and Command Policies
#	Added Appflow Policies, Actions, Policy Labels and Analytics Profiles
#	Added Logout of API session on script completion to clean up old connections
#	Fixed issue where logon session to the NetScaler can time-out causing null values to be returned
#	Added SSL Certificate bindings for Load Balancing and Content Switching vServers and Gateway
#	Added TLS 1.3 to SSL Parameters

.Release Notes version 4.1
#	Name change from NetScaler to NetScaler (R.I.P NetScaler)
#	Official NetScaler 12.1 Support
#	Updated features and modes to 12.1 levels
#	NetScaler Gateway - Added RDP Client and Server Profiles
#	FIX: Service Group Monitors and Advanced Config missing - Thanks to Nico Stylemans
#	Added Unified Gateway SaaS Application Templates (System and User Defined)
#	Updated SSL Profiles with new options

.Release Notes version 4.0
#	Official NetScaler v12 support
#	Fixed NetScaler SSL connections
#	Added SAML Authentication policies
#	Updated GSLB Parameters to include late 11.1 build enhancements
#	Added Support for NetScaler Clustering
#	Added AppExpert
	- Pattern Sets
	- HTTP Callouts
	- Data Sets
#	Numerous bug fixes

.Release Notes version 3.6
	The script is now fully compatible with NetScaler 11.1 released in July 2016.
#	Added NetScaler Functionality
#	Added NetScaler Gateway reporting for Custom Themes
#	Added HTTPS redirect for Load Balancing
#	Added Policy Based Routing
#	Added several items to advanced configuration for Load Balancer and Services
#	Numerous bug fixes

.Release Notes version 3.5
	Most work on version 3.5 has been done by Andy McCullough!
	After the release of version 3.0 in May 2016, which was a major overhaul of the NetScaler documentation script we found a few issues which have been fixed in the update.

	The script is now fully compatible with NetScaler 11.1 released in July 2016.

#	Added NetScaler Functionality
#	Added NetScaler 11.1 Features, LSN / RDP Proxy / REP
#	Added Auditing Section
#	Added GSLB Section, vServer / Services / Sites
#	Added Locations Database section to support GSLB configuration using Static proximity.
#	Added additional DNS Records to the NetScaler DNS Section
#	Added RPC Nodes section
#	Added NetScaler SSL Chapter, moved existing Functionality and added detailed information
#	Added AppFW Profiles and Policies
#	Added AAA vServers

	Added NetScaler Gateway Functionality
#	Updated NSGW Global Settings Client Experience to include new parameters
#	Updated NSGW Global Settings Published Applications to include new parameters
#	Added Section NSGW "Global Settings AAA Parameters"
#	Added SSL Parameters section for NSGW Virtual Servers
#	Added Rewrite Policies section for each NSGW vServer
#	Updated CAG vServer basic configuration section to include new parameters
#	Updated NetScaler Gateway Session Action > Security to include new attributed
#	Added Section NetScaler Gateway Session Action > Client Experience
#	Added Section NetScaler Gateway Policies > NetScaler Gateway AlwaysON Policies
#	Added NSGW Bookmarks
#	Added NSGW Intranet IP's
#	Added NSGW Intranet Applications
#	Added NSGW SSL Ciphers

	Webster's Updates

#	Updated help text to match other documentation scripts
#	Removed all code related to TEXT and HTML output since Barry does not offer those
#	Added support for specifying an output folder to match other documentation scripts
#	Added support for the -Dev and -ScriptInfo parameters to match other documentation scripts
#	Added support for emailing the output file to match other documentation scripts
#	Removed unneeded Functions
#	Brought script code in line with the other documentation scripts
#	Temporarily disabled the use of the UseNSSSL parameter
	
.Release Notes version 3
	Overall
	The script has had a major overhaul and is now completely utilizing the Nitro API instead of the NS.Conf.
	The Nitro API offers a lot more information and most important end result is much more predictable. Adding NetScaler Functionality is also much easier.
	Added Functionality because of Nitro
#	Hardware and license information
#	Complete routing tables including default routes
#	Complete monitoring information including default monitors

.Release Notes version 2
	Overall
	Test group has grown from 5 to 20 people. A lot more testing on a lot more configs has been done.
        The result is that I've received a lot of nitty gritty bugs that are now solved. To many to list them all but this release is very very stable.
	New Script Functionality
	New table Function that now utilizes native word tables. Looks a lot better and is way faster
	Performance improvements; over 500% faster
	Better support for multi language Word versions. Will now always utilize cover page and TOC
	New NetScaler Functionality:
#	NetScaler Gateway
		Global Settings
		Virtual Servers settings and policies
		Policies Session/Traffic
		NetScaler administration users and groups
#	NetScaler Authentication
		Policies LDAP / Radius
		Actions Local / RADIUS
		Action LDAP more configuration reported and changed table layout
#	NetScaler Networking
		Channels
		ACL
#	NetScaler Cache redirection
	Bugfixes
#	Naming of items with spaces and quotes fixed
#	Expressions with spaces, quotes, dashes and slashed fixed
#	Grammatical corrections
#	Rechecked all settings like enabled/disabled or on/off and corrected when necessary
#	Time zone not show correctly when in GMT+....
#	A lot more small items
.Release Notes version 1
	Version 1.0 supports the following NetScaler Functionality:
#	NetScaler System Information
#	Version / NSIP / vLAN
#	NetScaler Global Settings
#	NetScaler Feature and mode state
#	NetScaler Networking
#	IP Address / vLAN / Routing Table / DNS
#	NetScaler Authentication
#	Local / LDAP
#	NetScaler Traffic Domain
#	Assigned Content Switch / Load Balancer / Service  / Server
#	NetScaler Monitoring
#	NetScaler Certificate
#	NetScaler Content Switches
#	Assigned Load Balancer / Service  / Server
#	NetScaler Load Balancer
#	Assigned Service  / Server
#	NetScaler Service
#	Assigned Server / monitor
#	NetScaler Service Group
#	Assigned Server / monitor
#	NetScaler Server
#	NetScaler Custom Monitor
#	NetScaler Policy
#	NetScaler Action
#	NetScaler Profile
#>
#endregion changelog

#region script template
#region input
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
    [parameter(ParameterSetName = "Word", Mandatory = $False)] 
    [Switch]$MSWord = $False,

    [parameter(ParameterSetName = "PDF", Mandatory = $False)] 
    [Switch]$PDF = $False,

    [parameter(Mandatory = $False )] 
    [Switch]$AddDateTime = $False,
	
    [parameter(Mandatory = $False )]
    [string] $NSIP,
    
    [parameter(Mandatory = $false ) ]
    #[PSCredential] $Credential = (Get-Credential -Message 'Enter NetScaler credentials'),
    [PSCredential] $Credential,

    [parameter(Mandatory = $false ) ]
    [String] $NSUserName,
    
    [parameter(Mandatory = $false ) ]
    [String] $NSPassword,
   
    ## EXPERIMENTAL: Require SSL/TLS, e.g. https://. This requires the client to trust to the NetScaler's certificate.
    [parameter(Mandatory = $false )]
    [System.Management.Automation.SwitchParameter] $UseNSSSL,
    
    [parameter(Mandatory = $False)] 
    [string]$Folder = "",
	
    [parameter(Mandatory = $False)]
    [string]$ReportFileName = "NetScaler Documentation",

    [parameter(ParameterSetName = "Word", Mandatory = $False)] 
    [parameter(ParameterSetName = "PDF", Mandatory = $False)] 
    [Alias("CN")]
    [ValidateNotNullOrEmpty()]
    [string]$CompanyName = "",
    
    [parameter(ParameterSetName = "Word", Mandatory = $False)] 
    [parameter(ParameterSetName = "PDF", Mandatory = $False)] 
    [Alias("CP")]
    [ValidateNotNullOrEmpty()]
    [string]$CoverPage = "Sideline", 

    [parameter(ParameterSetName = "Word", Mandatory = $False)] 
    [parameter(ParameterSetName = "PDF", Mandatory = $False)] 
    [Alias("UN")]
    [ValidateNotNullOrEmpty()]
    [string]$UserName = $env:username,

    [parameter(Mandatory = $False)] 
    [string]$SmtpServer = "",

    [parameter(Mandatory = $False)] 
    [int]$SmtpPort = 25,

    [parameter(Mandatory = $False)] 
    [switch]$UseSSL = $False,

    [parameter(Mandatory = $False)] 
    [string]$From = "",

    [parameter(Mandatory = $False)] 
    [string]$To = "",
	
    [parameter(Mandatory = $False)] 
    [Switch]$Dev = $False,
    
    [parameter(Mandatory = $False)] 
    [Switch]$Log = $False,

    [parameter(ParameterSetName = "Export", Mandatory = $False)]
    [parameter(ParameterSetName = "Word", Mandatory = $False)]
    [Alias("Export")] 
    [Switch]$Offline = $False,

    [parameter(ParameterSetName = "Export", Mandatory = $False)]
    [parameter(ParameterSetName = "Word", Mandatory = $False)]
    [Alias("ExportPath")] 
    [String]$OfflinePath = "$pwd\ADCDocsExport",

    [parameter(ParameterSetName = "Import", Mandatory = $False)] 
    [parameter(ParameterSetName = "Word", Mandatory = $False)]
    [Switch]$Import = $False,

    [parameter(ParameterSetName = "Import", Mandatory = $False)]
    [parameter(ParameterSetName = "Word", Mandatory = $False)] 
    [String]$ImportPath = "$pwd\ADCDocsExport",
	
    [parameter(Mandatory = $False)] 
    [Alias("SI")]
    [Switch]$ScriptInfo = $False
)

#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on June 1, 2016

Set-StrictMode -Version 2

#force -verbose on
#$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'
#recommended by webster
#$Error.Clear()

If (!(Test-Path Variable:NSIP) -or ("" -eq $NSIP)) {
    If (!$Import) {
        $NSIP = Read-Host "Please enter the Management IP address for NetScaler" 
    }
}
If ($Offline -and $Import) {
    #If both are specified then run normally as the admin wants to export and generate the word doc
    $Offline = $false
    $Import = $false
    Write-Host "$(Get-Date): Script Mode: Classic" -ForegroundColor Green
}
ElseIf ($Offline) {
    Write-Host "$(Get-Date): Script Mode: Export" -ForegroundColor Green
}
ElseIf ($Import) {
    Write-Host "$(Get-Date): Script Mode: Import" -ForegroundColor Green
}

If ($Null -eq $MSWord) {
    If ($PDF) {
        $MSWord = $False
    }
    Else {
        $MSWord = $True
    }
}

If ($MSWord -eq $False -and $PDF -eq $False) {
    $MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If ($MSWord) {
    Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf ($PDF) {
    Write-Verbose "$(Get-Date): PDF is set"
}
Else {
    $ErrorActionPreference = $SaveEAPreference
    Write-Verbose "$(Get-Date): Unable to determine output parameter"
    If ($Null -eq $MSWord) {
        Write-Verbose "$(Get-Date): MSWord is Null"
    }
    ElseIf ($Null -eq $PDF) {
        Write-Verbose "$(Get-Date): PDF is Null"
    }
    Else {
        Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
        Write-Verbose "$(Get-Date): PDF is $($PDF)"
    }
    Write-Error "
	`n`n
	`t`t
	Unable to determine output parameter.
	`n`n
	`t`t
	Script cannot continue.
	`n`n
	"
    Exit
}

If ($Folder -ne "") {
    Write-Verbose "$(Get-Date): Testing folder path"
    #does it exist
    If (Test-Path $Folder -EA 0) {
        #it exists, now check to see if it is a folder and not a file
        If (Test-Path $Folder -pathType Container -EA 0) {
            #it exists and it is a folder
            Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
        }
        Else {
            #it exists but it is a file not a folder
            Write-Error "
			`n`n
			`t`t
			Folder $Folder is a file, not a folder.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
            Exit
        }
    }
    Else {
        #does not exist
        Write-Error "
		`n`n
		`t`t
		Folder $Folder does not exist.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
        Exit
    }
}

If ($Folder -eq "") {
    $Script:pwdpath = $pwd.Path
}
Else {
    $Script:pwdpath = $Folder
}

If ($Script:pwdpath.EndsWith("\")) {
    #remove the trailing \
    $Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
}

If ($Dev) {
    $Error.Clear()
    $Script:DevErrorFile = "$Script:pwdpath\NetScalerDocumentationScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If ($Log) {
    $Error.Clear()
    $Script:LogFile = "$Script:pwdpath\NetScalerDocumentationLogFile_$(Get-Date -f yyyy-MM-dd_HHmm_ss).txt"
}

If (![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To)) {
    Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer but did not include a From or To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
    Exit
}
If (![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To)) {
    Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
    Exit
}
If (![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From)) {
    Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
    Exit
}
If (![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer)) {
    Write-Error "
	`n`n
	`t`t
	You specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
    Exit
}
If (![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer)) {
    Write-Error "
	`n`n
	`t`t
	You specified a From email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
    Exit
}
If (![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer)) {
    Write-Error "
	`n`n
	`t`t
	You specified a To email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
    Exit
}
#endregion input

#region initialize variables for word html and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If ($MSWord -or $PDF) {
    #try and fix the issue with the $CompanyName variable
    $Script:CoName = $CompanyName
    Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"
	
    #the following values were attained from 
    #http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
    #http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
    [int]$wdAlignPageNumberRight = 2
    [long]$wdColorGray15 = 14277081
    [long]$wdColorGray05 = 15987699 
    [int]$wdMove = 0
    [int]$wdSeekMainDocument = 0
    [int]$wdSeekPrimaryFooter = 4
    [int]$wdStory = 6
    [long]$wdColorRed = 255
    [int]$wdColorBlack = 0
    [int]$wdWord2007 = 12
    [int]$wdWord2010 = 14
    [int]$wdWord2013 = 15
    [int]$wdWord2016 = 16
    [int]$wdFormatDocumentDefault = 16
    [int]$wdFormatPDF = 17
    #http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
    #http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
    [int]$wdAlignParagraphLeft = 0
    [int]$wdAlignParagraphCenter = 1
    [int]$wdAlignParagraphRight = 2
    #http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
    [int]$wdCellAlignVerticalTop = 0
    [int]$wdCellAlignVerticalCenter = 1
    [int]$wdCellAlignVerticalBottom = 2
    #http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
    [int]$wdAutoFitFixed = 0
    [int]$wdAutoFitContent = 1
    [int]$wdAutoFitWindow = 2
    #http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
    [int]$wdAdjustNone = 0
    [int]$wdAdjustProportional = 1
    [int]$wdAdjustFirstColumn = 2
    [int]$wdAdjustSameWidth = 3

    [int]$PointsPerTabStop = 36
    [int]$Indent0TabStops = 0 * $PointsPerTabStop
    [int]$Indent1TabStops = 1 * $PointsPerTabStop
    [int]$Indent2TabStops = 2 * $PointsPerTabStop
    [int]$Indent3TabStops = 3 * $PointsPerTabStop
    [int]$Indent4TabStops = 4 * $PointsPerTabStop

    # http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
    [int]$wdStyleHeading1 = -2
    [int]$wdStyleHeading2 = -3
    [int]$wdStyleHeading3 = -4
    [int]$wdStyleHeading4 = -5
	[int]$wdStyleHeading5 = -6
    [int]$wdStyleNoSpacing = -158
    [int]$wdTableGrid = -155
	
    #http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
    [int]$wdLineStyleNone = 0
    [int]$wdLineStyleSingle = 1

    [int]$wdHeadingFormatTrue = -1
    [int]$wdHeadingFormatFalse = 0 
}
#endregion initialize variables for word html and text

#region email Function
Function SendEmail {
    Param([array]$Attachments)
    Write-Verbose "$(Get-Date): Prepare to email"

    $emailAttachment = $Attachments
    $emailSubject = $Script:Title
    $emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.
"@
    If ($Dev) {
        Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
    }

    $error.Clear()
	
    If ($From -Like "anonymous@*") {
        #https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
        $anonUsername = "anonymous"
        $anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
        $anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername, $anonPassword)

        If ($UseSSL) {
            Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
                -Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
                -UseSSL -credential $anonCredentials *>$Null 
        }
        Else {
            Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
                -Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
                -credential $anonCredentials *>$Null 
        }
		
        If ($?) {
            Write-Verbose "$(Get-Date): Email successfully sent using anonymous credentials"
        }
        ElseIf (!$?) {
            $e = $error[0]

            Write-Verbose "$(Get-Date): Email was not sent:"
            Write-Warning "$(Get-Date): Exception: $e.Exception" 
        }
    }
    Else {
        If ($UseSSL) {
            Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
            Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
                -Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
                -UseSSL *>$Null
        }
        Else {
            Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
            Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
                -Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
        }

        If (!$?) {
            $e = $error[0]
			
            #error 5.7.57 is O365 and error 5.7.0 is gmail
            If ($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7")) {
                #The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
                Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

                If ($Dev) {
                    Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
                }

                $error.Clear()

                $emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

                If ($UseSSL) {
                    Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
                        -Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
                        -UseSSL -credential $emailCredentials *>$Null 
                }
                Else {
                    Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
                        -Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
                        -credential $emailCredentials *>$Null 
                }

                If ($?) {
                    Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
                }
                ElseIf (!$?) {
                    $e = $error[0]

                    Write-Verbose "$(Get-Date): Email was not sent:"
                    Write-Warning "$(Get-Date): Exception: $e.Exception" 
                }
            }
            Else {
                Write-Verbose "$(Get-Date): Email was not sent:"
                Write-Warning "$(Get-Date): Exception: $e.Exception" 
            }
        }
    }
}
#endregion email Function

#region word specific Functions
Function SetWordHashTable {
    Param([string]$CultureCode)

    #optimized by Michael B. SMith
	
    # DE and FR translations for Word 2010 by Vladimir Radojevic
    # Vladimir.Radojevic@Commerzreal.com

    # DA translations for Word 2010 by Thomas Daugaard
    # Citrix Infrastructure Specialist at edgemo A/S

    # CA translations by Javier Sanchez 
    # CEO & Founder 101 Consulting

    #ca - Catalan
    #da - Danish
    #de - German
    #en - English
    #es - Spanish
    #fi - Finnish
    #fr - French
    #nb - Norwegian
    #nl - Dutch
    #pt - Portuguese
    #sv - Swedish
    #zh - Chinese
	
    [string]$toc = $(
        Switch ($CultureCode) {
            'ca-'	{ 'Taula automática 2'; Break }
            'da-'	{ 'Automatisk tabel 2'; Break }
            'de-'	{ 'Automatische Tabelle 2'; Break }
            'en-'	{ 'Automatic Table 2'; Break }
            'es-'	{ 'Tabla automática 2'; Break }
            'fi-'	{ 'Automaattinen taulukko 2'; Break }
            'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
            'nb-'	{ 'Automatisk tabell 2'; Break }
            'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
            'pt-'	{ 'Sumário Automático 2'; Break }
            # fix in 2.23 thanks to Johan Kallio 'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
            'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
            'zh-'	{ '自动目录 2'; Break }
        }
    )

    $Script:myHash = @{}
    $Script:myHash.Word_TableOfContents = $toc
    $Script:myHash.Word_NoSpacing = $wdStyleNoSpacing
    $Script:myHash.Word_Heading1 = $wdStyleheading1
    $Script:myHash.Word_Heading2 = $wdStyleheading2
    $Script:myHash.Word_Heading3 = $wdStyleheading3
    $Script:myHash.Word_Heading4 = $wdStyleheading4
	$Script:myHash.Word_Heading5 = $wdStyleheading5
    $Script:myHash.Word_TableGrid = $wdTableGrid
}

Function GetCulture {
    Param([int]$WordValue)
	
    #codes obtained from http://support.microsoft.com/kb/221435
    #http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
    $CatalanArray = 1027
    $ChineseArray = 2052, 3076, 5124, 4100
    $DanishArray = 1030
    $DutchArray = 2067, 1043
    $EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
    $FinnishArray = 1035
    $FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
    $GermanArray = 1031, 3079, 5127, 4103, 2055
    $NorwegianArray = 1044, 2068
    $PortugueseArray = 1046, 2070
    $SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
    $SwedishArray = 1053, 2077

    #ca - Catalan
    #da - Danish
    #de - German
    #en - English
    #es - Spanish
    #fi - Finnish
    #fr - French
    #nb - Norwegian
    #nl - Dutch
    #pt - Portuguese
    #sv - Swedish
    #zh - Chinese

    Switch ($WordValue) {
        { $CatalanArray -contains $_ } { $CultureCode = "ca-" }
        { $ChineseArray -contains $_ } { $CultureCode = "zh-" }
        { $DanishArray -contains $_ } { $CultureCode = "da-" }
        { $DutchArray -contains $_ } { $CultureCode = "nl-" }
        { $EnglishArray -contains $_ } { $CultureCode = "en-" }
        { $FinnishArray -contains $_ } { $CultureCode = "fi-" }
        { $FrenchArray -contains $_ } { $CultureCode = "fr-" }
        { $GermanArray -contains $_ } { $CultureCode = "de-" }
        { $NorwegianArray -contains $_ } { $CultureCode = "nb-" }
        { $PortugueseArray -contains $_ } { $CultureCode = "pt-" }
        { $SpanishArray -contains $_ } { $CultureCode = "es-" }
        { $SwedishArray -contains $_ } { $CultureCode = "sv-" }
        Default { $CultureCode = "en-" }
    }
    Return $CultureCode
}

Function ValidateCoverPage {
    Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
    $xArray = ""
	
    Switch ($CultureCode) {
        'ca-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
                    "Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
                    "Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
                    "Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
                    "Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
                    "Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
                    "Sector (fosc)", "Semàfor", "Visualització", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabet", "Anual", "Austin", "Conservador",
                    "Contrast", "Cubicles", "Diplomàtic", "Exposició",
                    "Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
                    "Perspectiva", "Piles", "Quadrícula", "Sobri",
                    "Transcendir", "Trencaclosques")
            }
        }

        'da-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
                    "Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
                    "Retro", "Semafor", "Sidelinje", "Stribet", 
                    "Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
                    "Retro", "Semafor", "Visningsmaster", "Integral",
                    "Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
                    "Udsnit (mørk)", "Ion (mørk)", "Austin")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
                    "Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
                    "Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
                    "Nålestribet", "Årlig", "Avispapir", "Tradionel")
            }
        }

        'de-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
                    "Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
                    "Pfiff", "Randlinie", "Raster", "Rückblick", 
                    "Segment (dunkel)", "Segment (hell)", "Semaphor", 
                    "ViewMaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
                    "Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
                    "ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
                    "Randlinie", "Austin", "Integral", "Facette")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
                    "Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
                    "Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
                    "Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
            }
        }

        'en-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
                    "Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
                    "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
                    "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
                    "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
                    "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
            }
        }

        'es-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
                    "Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
                    "Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
                    "Semáforo", "Slice (luz)", "Vista principal", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
                    "Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
                    "Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
                    "Ion (claro)", "Integral", "Con bandas")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
                    "Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
                    "Moderno", "Mosaicos", "Movimiento", "Papel periódico",
                    "Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
            }
        }

        'fi-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
                    "Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
                    "Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
                    "Kuiskaus", "Liike", "Ruudukko", "Sivussa")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
                    "Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
                    "Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
                    "Kiehkura", "Liike", "Ruudukko", "Sivussa")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
                    "Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
                    "Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
                    "Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
            }
        }

        'fr-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
                    "Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
                    "Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
                    "Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
                    "Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
                    "Exposition", "Guide", "Ligne latérale", "Moderne", 
                    "Mosaïques", "Mots croisés", "Papier journal", "Perspective",
                    "Quadrillage", "Rayures fines", "Transcendant")
            }
        }

        'nb-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
                    "Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
                    "Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
                    "ViewMaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
                    "BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
                    "Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
                    "Smale striper", "Stabler", "Transcenderende")
            }
        }

        'nl-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
                    "Integraal", "Ion (donker)", "Ion (licht)", "Raster",
                    "Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
                    "Terugblik", "Terzijde", "ViewMaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
                    "Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
                    "Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
                    "Puzzel", "Raster", "Stapels",
                    "Tegels", "Terzijde")
            }
        }

        'pt-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
                    "Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
                    "Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
                    "Retrospectiva", "Semáforo")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
                    "Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
                    "Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
                    "Quebra-cabeça", "Transcend")
            }
        }

        'sv-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
                    "Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
                    "Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
                    "Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
                    "RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
                    "Övergående")
            }
        }

        'zh-'	{
            If ($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
                    '离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
                    '切片(深色)', '丝状', '网格', '镶边', '信号灯',
                    '运动型')
            }
        }

        Default	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
                    "Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
                    "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
                    "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
                    "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
                    "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
            }
        }
    }
	
    If ($xArray -contains $xCP) {
        $xArray = $Null
        Return $True
    }
    Else {
        $xArray = $Null
        Return $False
    }
}

Function CheckWordPrereq {
    If ((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False) {
        $ErrorActionPreference = $SaveEAPreference
        Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
        Exit
    }

    #find out our session (usually "1" except on TS/RDC or Citrix)
    $SessionID = (Get-Process -PID $PID).SessionId
	
    #Find out if winword is running in our session
    [bool]$wordrunning = ((Get-Process 'WinWord' -ea 0) | ? { $_.SessionId -eq $SessionID }) -ne $Null
    If ($wordrunning) {
        $ErrorActionPreference = $SaveEAPreference
        Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
        Exit
    }
}

Function ValidateCompanyName {
    [bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
    If ($xResult) {
        Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
    }
    Else {
        $xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
        If ($xResult) {
            Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
        }
        Else {
            Return ""
        }
    }
}

Function Set-DocumentProperty {
    param ([String]$DocProperty, [string]$Value)
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $Script:Doc.BuiltInDocumentProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd {
    If (!$Offline) {
        #return focus to main document    
        $Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
        #move to the end of the current document
        $Script:Selection.EndKey($wdStory, $wdMove) | Out-Null
    }
}

Function SetupWord {
    Write-Verbose "$(Get-Date): Setting up Word"
    
    # Setup word for output
    Write-Verbose "$(Get-Date): Create Word comObject."
    $Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
    If (!$? -or $Null -eq $Script:Word) {
        Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "
		`n`n
		`t`t
		The Word object could not be created.  You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
        Exit
    }

    Write-Verbose "$(Get-Date): Determine Word language value"
    If ( ( validStateProp $Script:Word Language Value__ ) ) {
        [int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
    }
    Else {
        [int]$Script:WordLanguageValue = [int]$Script:Word.Language
    }

    If (!($Script:WordLanguageValue -gt -1)) {
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "
		`n`n
		`t`t
		Unable to determine the Word language value.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
        AbortScript
    }
    Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
    $Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
    SetWordHashTable $Script:WordCultureCode
	
    [int]$Script:WordVersion = [int]$Script:Word.Version
    If ($Script:WordVersion -eq $wdWord2016) {
        $Script:WordProduct = "Word 2016"
    }
    ElseIf ($Script:WordVersion -eq $wdWord2013) {
        $Script:WordProduct = "Word 2013"
    }
    ElseIf ($Script:WordVersion -eq $wdWord2010) {
        $Script:WordProduct = "Word 2010"
    }
    ElseIf ($Script:WordVersion -eq $wdWord2007) {
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "
		`n`n
		`t`t
		Microsoft Word 2007 is no longer supported.
		`n`n
		`t`t
		Script will end.
		`n`n
		"
        AbortScript
    }
    ElseIf ($Script:WordVersion -eq 0) {
        Write-Error "
		`n`n
		`t`t
		The Word Version is 0. You should run a full online repair of your Office installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
        Exit
    }
    Else {
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "
		`n`n
		`t`t
		You are running an untested or unsupported version of Microsoft Word.
		`n`n
		`t`t
		Script will end.
		`n`n
		`t`t
		Please send info on your version of Word to webster@carlwebster.com
		`n`n
		"
        AbortScript
    }

    #only validate CompanyName if the field is blank
    If ([String]::IsNullOrEmpty($Script:CoName)) {
        Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
        $TmpName = ValidateCompanyName
		
        If ([String]::IsNullOrEmpty($TmpName)) {
            Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
            Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
            Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
        }
        Else {
            $Script:CoName = $TmpName
            Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
        }
    }

    If ($Script:WordCultureCode -ne "en-") {
        Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
        [bool]$CPChanged = $False
        Switch ($Script:WordCultureCode) {
            'ca-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Línia lateral"
                    $CPChanged = $True
                }
            }

            'da-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Sidelinje"
                    $CPChanged = $True
                }
            }

            'de-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Randlinie"
                    $CPChanged = $True
                }
            }

            'es-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Línea lateral"
                    $CPChanged = $True
                }
            }

            'fi-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Sivussa"
                    $CPChanged = $True
                }
            }

            'fr-'	{
                If ($CoverPage -eq "Sideline") {
                    If ($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016) {
                        $CoverPage = "Lignes latérales"
                        $CPChanged = $True
                    }
                    Else {
                        $CoverPage = "Ligne latérale"
                        $CPChanged = $True
                    }
                }
            }

            'nb-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Sidelinje"
                    $CPChanged = $True
                }
            }

            'nl-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Terzijde"
                    $CPChanged = $True
                }
            }

            'pt-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Linha Lateral"
                    $CPChanged = $True
                }
            }

            'sv-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Sidlinje"
                    $CPChanged = $True
                }
            }

            'zh-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "边线型"
                    $CPChanged = $True
                }
            }
        }

        If ($CPChanged) {
            Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
        }
    }

    Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
    [bool]$ValidCP = $False
	
    $ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
    If (!$ValidCP) {
        $ErrorActionPreference = $SaveEAPreference
        Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
        Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
        Write-Error "
		`n`n
		`t`t
		For $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
        AbortScript
    }

    ShowScriptOptions

    $Script:Word.Visible = $False

    #http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
    #using Jeff's Demo-WordReport.ps1 file for examples
    Write-Verbose "$(Get-Date): Load Word Templates"

    [bool]$Script:CoverPagesExist = $False
    [bool]$BuildingBlocksExist = $False

    $Script:Word.Templates.LoadBuildingBlocks()
    #word 2010/2013/2016
    $BuildingBlocksCollection = $Script:Word.Templates | Where { $_.name -eq "Built-In Building Blocks.dotx" }

    Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
    $part = $Null

    $BuildingBlocksCollection | 
    ForEach {
        If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) {
            $BuildingBlocks = $_
        }
    }        

    If ($Null -ne $BuildingBlocks) {
        $BuildingBlocksExist = $True

        Try {
            $part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
        }

        Catch {
            $part = $Null
        }

        If ($Null -ne $part) {
            $Script:CoverPagesExist = $True
        }
    }

    If (!$Script:CoverPagesExist) {
        Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
        Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
        Write-Warning "This report will not have a Cover Page."
    }

    Write-Verbose "$(Get-Date): Create empty word doc"
    $Script:Doc = $Script:Word.Documents.Add()
    If ($Null -eq $Script:Doc) {
        Write-Verbose "$(Get-Date): "
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "
		`n`n
		`t`t
		An empty Word document could not be created.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
        AbortScript
    }

    $Script:Selection = $Script:Word.Selection
    If ($Null -eq $Script:Selection) {
        Write-Verbose "$(Get-Date): "
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "
		`n`n
		`t`t
		An unknown error happened selecting the entire Word document for default formatting options.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
        AbortScript
    }

    #set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
    #36 = .50"
    $Script:Word.ActiveDocument.DefaultTabStop = 36

    #Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
    Write-Verbose "$(Get-Date): Disable grammar and spell checking"
    #bug reported 1-Apr-2014 by Tim Mangan
    #save current options first before turning them off
    $Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
    $Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
    $Script:Word.Options.CheckGrammarAsYouType = $False
    $Script:Word.Options.CheckSpellingAsYouType = $False

    If ($BuildingBlocksExist) {
        #insert new page, getting ready for table of contents
        Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
        $part.Insert($Script:Selection.Range, $True) | Out-Null
        $Script:Selection.InsertNewPage()

        #table of contents
        Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
        $toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
        If ($Null -eq $toc) {
            Write-Verbose "$(Get-Date): "
            Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
            Write-Warning "This report will not have a Table of Contents."
        }
        Else {
            $toc.insert($Script:Selection.Range, $True) | Out-Null
        }
    }
    Else {
        Write-Verbose "$(Get-Date): Table of Contents are not installed."
        Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
    }

    #set the footer
    Write-Verbose "$(Get-Date): Set the footer"
    [string]$footertext = "Report created by $username"

    #get the footer
    Write-Verbose "$(Get-Date): Get the footer and format font"
    $Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
    #get the footer and format font
    $footers = $Script:Doc.Sections.Last.Footers
    ForEach ($footer in $footers) {
        If ($footer.exists) {
            $footer.range.Font.name = "Calibri"
            $footer.range.Font.size = 8
            $footer.range.Font.Italic = $True
            $footer.range.Font.Bold = $True
        }
    } #end ForEach
    Write-Verbose "$(Get-Date): Footer text"
    $Script:Selection.HeaderFooter.Range.Text = $footerText

    #add page numbering
    Write-Verbose "$(Get-Date): Add page numbering"
    $Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

    FindWordDocumentEnd
    Write-Verbose "$(Get-Date):"
    #end of Jeff Hicks 
}

Function UpdateDocumentProperties {
    Param([string]$AbstractTitle, [string]$SubjectTitle)
    #Update document properties
    If ($MSWORD -or $PDF) {
        If ($Script:CoverPagesExist) {
            Write-Verbose "$(Get-Date): Set Cover Page Properties"
			Set-DocumentProperty Author $UserName
			Set-DocumentProperty Company $Script:CoName
            Set-DocumentProperty Subject $SubjectTitle
            Set-DocumentProperty Title $Script:title

            #Get the Coverpage XML part
            $cp = $Script:Doc.CustomXMLParts | Where { $_.NamespaceURI -match "coverPageProps$" }

            #get the abstract XML part
            $ab = $cp.documentelement.ChildNodes | Where { $_.basename -eq "Abstract" }

            #set the text
            If ([String]::IsNullOrEmpty($Script:CoName)) {
                [string]$abstract = $AbstractTitle
            }
            Else {
                [string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
            }

            $ab.Text = $abstract

            $ab = $cp.documentelement.ChildNodes | Where { $_.basename -eq "PublishDate" }
            #set the text
            [string]$abstract = (Get-Date -Format d).ToString()
            $ab.Text = $abstract

            Write-Verbose "$(Get-Date): Update the Table of Contents"
            #update the Table of Contents
            $Script:Doc.TablesOfContents.item(1).Update()
            $cp = $Null
            $ab = $Null
            $abstract = $Null
        }
    }
}

Function New-BindingTable {
	[CmdletBinding()]
	param (
		# Name to query bindings for
		[Parameter(Mandatory)] [System.String] $Name,
		# Binding Type to Query
		[Parameter(Mandatory)] [System.String] $BindingType,
		# Friendly Name for Binding Type (used in Header)
		[Parameter(Mandatory)] [System.String[]] $BindingTypeName,
		# Array of Object Properties to output
		[Parameter(Mandatory)] [System.String[]] $Properties,
		# Retrieve Builk Bindings for an object
		[Parameter(Mandatory)] [System.String[]] $Headers,
		[Parameter(Mandatory)] [System.String] $Style
	)
	If ((Get-vNetScalerObjectCount -Type $BindingType -Name $Name).__count -ge 1) {
		$BindingObject = Get-vNetScalerObject -Type $BindingType -Name $Name
		WriteWordLine "$Style" 0 "$BindingTypeName"
		[System.Collections.Hashtable[]] $POLICIESH = @()
		$ArrProperties = $Properties.split(",")
		ForEach ($Binding in $BindingObject) {
			[System.Collections.HashTable] $TempHash = @{}
			foreach ($objProperty in $ArrProperties) {
				$objValue = $Binding."$objProperty"
				Try {
					$TempHash.Add($objProperty, $objValue)
				}
				Catch {
					Write-Log $_.exception
				}
			}
			$POLICIESH += $TempHash
		}
		$Params = $null
		$Params = @{
			Hashtable = $POLICIESH
			Columns   = $Properties.Split(",")
			Headers   = $Headers.Split(",")
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
}

Function New-SSLSettings{
	[CmdletBinding()]
	param (
		# vServer Name to query bindings for
		[Parameter(Mandatory)] [System.String] $Name,
		[Parameter(Mandatory)] [System.String] $Type,
		[Parameter(Mandatory)] [System.String] $Style
	)
	WriteWordLine "$Style" 0 "SSL Settings"
	$sslsettings = Get-vNetScalerObject -Type $Type -Name $Name
	If ($sslsettings.sslprofile -ne $null){
		[System.Collections.Hashtable[]] $SSLSETTINGSH = @(
			@{ Description = "Description"; Value = "Value"}
			@{ Description = "SSL Profile"; Value = $sslsettings.sslprofile}
		)
	} Else {
		WriteWordLine 0 0 "Only non-default values are reported, defaults are in [brackets]"
		$Ciphers = (Get-vNetScalerObject -Type "$($Type)_sslciphersuite_binding" -Name $Name).Ciphername -Join ", "
		[System.Collections.Hashtable[]] $SSLSETTINGSH = @(
			@{ Description = "Description"; Value = "Value"}
			@{ Description = "Ciphers"; Value = $Ciphers}
			If ($Type -eq "sslvserver"){
				@{ Description = "Enable DH Param [disabled]"; Value = $sslsettings.dh}
				If ($sslsettings.dh -ne "DISABLED"){
					@{ Description = "Diffe-Hellman Refresh Count"; Value = $sslsettings.dhcount}
					@{ Description = "Diffe-Hellman Key File"; Value = $sslsettings.dhfile}
					@{ Description = "Enable DH Key Expire Size Limit"; Value = $sslsettings.dhkeyexpsizelimit}
				}
				If ($sslsettings.ersa -ne "ENABLED"){@{ Description = "Enable Ephemeral RSA [enabled]"; Value = $sslsettings.ersa}}
				If ($sslsettings.ersa -eq "ENABLED" -and $sslsettings.ersacount -ne "0"){@{ Description = "Ephemeral RSA Refresh Count [0]"; Value = $sslsettings.ersacount}}
			}
            If ($sslsettings.sessreuse -ne "ENABLED"){@{ Description = "Enable Session Reuse [enabled]"; Value = $sslsettings.sessreuse}}
            If ($sslsettings.sessreuse -eq "ENABLED" -and $Type -eq "sslvserver"){@{ Description = "Session Time-out [120]"; Value = $sslsettings.sesstimeout}}
			If ($sslsettings.sessreuse -eq "ENABLED" -and $Type -ne "sslvserver"){@{ Description = "Session Time-out [300]"; Value = $sslsettings.sesstimeout}}
			If ($Type -ne "sslvserver"){
				@{ Description = "Enable Server Authentication"; Value = $sslsettings.serverauth}
				@{ Description = "Common Name"; Value = $sslsettings.commonname}
			}
			If ($Type -eq "sslvserver"){
				If ($sslsettings.cipherredirect -ne "DISABLED"){
					@{ Description = "Enable Cipher Redirect [disabled]"; Value = $sslsettings.cipherredirect}
					@{ Description = "Cipher Redirect URL"; Value = $sslsettings.cipherurl}
				}
				If ($sslsettings.sslv2redirect -eq "ENABLED"){
					@{ Description = "SSLv2 Redirect [disabled]"; Value = $sslsettings.sslv2redirect}
					If (![string]::IsNullOrWhiteSpace($sslsettings.sslv2url)){@{ Description = "SSLv2 Redirect URL"; Value = $sslsettings.sslv2url}}
				}
				If ($sslsettings.clientauth -ne "DISABLED"){
					@{ Description = "Client Authentication [disabled]"; Value = $sslsettings.clientauth}
					@{ Description = "Client Certificates"; Value = $sslsettings.clientcert}
				}
			}
			If ($sslsettings.ocspstapling -ne "DISABLED"){@{ Description = "OCSP Stapling [disabled]"; Value = $sslsettings.ocspstapling}}
			If ($Type -eq "sslvserver"){
				If ($sslsettings.sslredirect -ne "DISABLED"){
					@{ Description = "SSL Redirect [disabled]"; Value = $sslsettings.sslredirect}
					If ($sslsettings.redirectportrewrite -ne "DISABLED"){@{ Description = "SSL Redirect Port Rewrite [disabled]"; Value = $sslsettings.redirectportrewrite}}
				}
			}
            If ($sslsettings.snienable -ne "DISABLED"){@{ Description = "Server Name Indication (SNI) [disabled]"; Value = $sslsettings.snienable}}
			If ($sslsettings.sendclosenotify -ne "YES"){@{ Description = "Send Close-Notify [yes]"; Value = $sslsettings.sendclosenotify}}
			If ($Type -eq "sslvserver"){
				If ($sslsettings.cleartextport -ne "0"){@{ Description = "Clear Text Port [0]"; Value = $sslsettings.cleartextport}}
				If ($sslsettings.pushenctrigger -ne "Always"){@{ Description = "Push Encryption Trigger"; Value = $sslsettings.pushenctrigger}}
			}
			@{ Description = "Strict Signature Digest Check [disabled]"; Value = $sslsettings.strictsigdigestcheck}
			If ($Type -eq "sslvserver"){
				@{ Description = "Enable Stricy Transport Security (HSTS) [disabled]"; Value = $sslsettings.hsts}
				If ($sslsettings.maxage -ne "0"){@{ Description = "HSTS: Maximum Age [0]"; Value = $sslsettings.maxage}}
				If ($sslsettings.includesubdomains -ne "NO"){@{ Description = "HSTS: Include Subdomains [no]"; Value = $sslsettings.includesubdomains}}
				If ($sslsettings.preload -ne "NO"){@{ Description = "HSTS: Preload"; Value = $sslsettings.preload}}
			}
			@{ Description = "SSL 2 [disabled]"; Value = $sslsettings.ssl2}
            @{ Description = "SSL 3 [disabled]"; Value = $sslsettings.ssl3}
            @{ Description = "TLS 1 [enabled]"; Value = $sslsettings.tls1}
            @{ Description = "TLS 1.1 [enabled]"; Value = $sslsettings.tls11}
            @{ Description = "TLS 1.2 [enabled]"; Value = $sslsettings.tls12}
            @{ Description = "TLS 1.3 [disabled]"; Value = $sslsettings.tls13}
			@{ Description = "DTLS 1.0 [enabled]"; Value = $sslsettings.dtls1}
			@{ Description = "DTLS 1.2 [disabled]"; Value = $sslsettings.dtls12}
		)
	}
	$Params = $null
	$Params = @{
		Hashtable = $SSLSETTINGSH
		Columns   = "Description", "Value"
	}
	$Table = AddWordTable @Params -List
	FindWordDocumentEnd
}
#endregion word specific Functions

#region registry Functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name) {
    $key = Get-Item -LiteralPath $path -EA 0
    $key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name) {
    $key = Get-Item -LiteralPath $path -EA 0
    If ($key) {
        $key.GetValue($name, $Null)
    }
    Else {
        $Null
    }
}
#endregion registry Functions

#region word output Functions
Function WriteWordLine {
    #Function created by Ryan Revord
    #@rsrevord on Twitter
    #Function created to make output to Word easy in this script
    #updated 27-Mar-2014 to include font name, font size, italics and bold options
    Param([int]$style = 0, 
        [int]$tabs = 0, 
        [string]$name = '',
        [string]$value = '',
        [string]$fontName = "",
        [int]$fontSize = 0,
        [bool]$italics = $False,
        [bool]$boldface = $False,
        [Switch]$nonewline)
	
    #Build output style
    [string]$output = ""
    Switch ($style) {
        0 { $Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break }
        1 { $Script:Selection.Style = $Script:MyHash.Word_Heading1; Set-Progress $Name; Break }
        2 { $Script:Selection.Style = $Script:MyHash.Word_Heading2; Set-Progress $Name; Break }
        3 { $Script:Selection.Style = $Script:MyHash.Word_Heading3; Set-Progress $Name; Break }
        4 { $Script:Selection.Style = $Script:MyHash.Word_Heading4; Set-Progress $Name; Break }
		5 { $Script:Selection.Style = $Script:MyHash.Word_Heading5; Break }
        Default { $Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break }
    }
	
    #build # of tabs
    While ($tabs -gt 0) { 
        $output += "`t"; $tabs--
    }
 
    If (![String]::IsNullOrEmpty($fontName)) {
        $Script:Selection.Font.name = $fontName
    } 

    If ($fontSize -ne 0) {
        $Script:Selection.Font.size = $fontSize
    } 
 
    If ($italics -eq $True) {
        $Script:Selection.Font.Italic = $True
    } 
 
    If ($boldface -eq $True) {
        $Script:Selection.Font.Bold = $True
    } 

    #output the rest of the parameters.
    $output += $name + $value
    $Script:Selection.TypeText($output)
 
    #test for new WriteWordLine 0.
    If ($nonewline) {
        # Do nothing.
    } 
    Else { 
        If (!$Offline) {
            $Script:Selection.TypeParagraph()
        }
    }
}
#endregion word output Functions

#region Iain's Word table Functions
<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This Function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this Function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable {
    [CmdletBinding()]
    Param
    (
        # Array of Hashtable (including table headers)
        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True, ParameterSetName = 'Hashtable', Position = 0)]
        [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
        # Array of PSCustomObjects
        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True, ParameterSetName = 'CustomObject', Position = 0)]
        [ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
        # Array of Hashtable key names or PSCustomObject property names to include, in display order.
        # If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
        [Parameter(ValueFromPipelineByPropertyName = $True)] [AllowNull()] [string[]] $Columns = $Null,
        # Array of custom table header strings in display order.
        [Parameter(ValueFromPipelineByPropertyName = $True)] [AllowNull()] [string[]] $Headers = $Null,
        # AutoFit table behavior.
        [Parameter(ValueFromPipelineByPropertyName = $True)] [AllowNull()] [int] $AutoFit = 1,
        # List view (no headers)
        [Switch] $List,
        # Grid lines
        [Switch] $NoGridLines,
        [Switch] $NoInternalGridLines,
        # Built-in Word table formatting style constant
        # Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
        [Parameter(ValueFromPipelineByPropertyName = $True)] [int] $Format = -161
			<# https://www.thedoctools.com/downloads/Create-List-Of-BuiltIn-Styles_DocTools.docm
			-161 = Light Grid
			-175 = Light Grid Accent 1
			-193 = Light Grid Accent 2
			-207 = Light Grid Accent 3
			-221 = Light Grid Accent 4
			-235 = Light Grid Accent 5
			-249 = Light Grid Accent 6
			#>
    )

    Begin {
        Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName)
        ## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
        If (($Null -eq $Columns) -and ($Null -ne $Headers)) {
            Write-Warning "No columns specified and therefore, specified headers will be ignored."
            $Columns = $Null
        }
        ElseIf (($Null -ne $Columns) -and ($Null -ne $Headers)) {
            ## Check if number of specified -Columns matches number of specified -Headers
            If ($Columns.Length -ne $Headers.Length) {
                Write-Error "The specified number of columns does not match the specified number of headers."
            }
        } ## end ElseIf
    } ## end Begin

    Process {
        If (!$Offline) {
            ## Build the Word table data string to be converted to a range and then a table later.
            [System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder

            Switch ($PSCmdlet.ParameterSetName) {
                'CustomObject' {
                    If ($Null -eq $Columns) {
                        ## Build the available columns from all available PSCustomObject note properties
                        [string[]] $Columns = @()
                        ## Add each NoteProperty name to the array
                        ForEach ($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) { 
                            $Columns += $Property.Name 
                        }
                    }

                    ## Add the table headers from -Headers or -Columns (except when in -List(view)
                    If (-not $List) {
                        Write-Debug ("$(Get-Date): `t`tBuilding table headers")
                        If ($Null -ne $Headers) {
                            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers))
                        }
                        Else { 
                            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns))
                        }
                    }

                    ## Iterate through each PSCustomObject
                    Write-Debug ("$(Get-Date): `t`tBuilding table rows")
                    ForEach ($Object in $CustomObject) {
                        $OrderedValues = @()
                        ## Add each row item in the specified order
                        ForEach ($Column in $Columns) { 
                            $OrderedValues += $Object.$Column 
                        }
                        ## Use the ordered list to add each column in specified order
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues))
                    } ## end foreach
                    Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count))
                } ## end CustomObject

                Default {
                    ## Hashtable
                    If ($Null -eq $Columns) {
                        ## Build the available columns from all available hashtable keys. Hopefully
                        ## all Hashtables have the same keys (they should for a table).
                        $Columns = $Hashtable[0].Keys
                    }

                    ## Add the table headers from -Headers or -Columns (except when in -List(view)
                    If (-not $List) {
                        Write-Debug ("$(Get-Date): `t`tBuilding table headers")
                        If ($Null -ne $Headers) { 
                            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers))
                        }
                        Else {
                            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns))
                        }
                    }
                
                    ## Iterate through each Hashtable
                    Write-Debug ("$(Get-Date): `t`tBuilding table rows")
                    ForEach ($Hash in $Hashtable) {
                        $OrderedValues = @()
                        ## Add each row item in the specified order
                        ForEach ($Column in $Columns) { 
                            $OrderedValues += $Hash.$Column 
                        }
                        ## Use the ordered list to add each column in specified order
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues))
                    } ## end foreach

                    Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count)
                } ## end default
            } ## end switch

            ## Create a MS Word range and set its text to our tab-delimited, concatenated string
            Write-Debug ("$(Get-Date): `t`tBuilding table range")
            $WordRange = $Script:Doc.Application.Selection.Range
            $WordRange.Text = $WordRangeString.ToString()

            ## Create hash table of named arguments to pass to the ConvertToTable method
            $ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs}

            ## Negative built-in styles are not supported by the ConvertToTable method
            If ($Format -ge 0) {
                $ConvertToTableArguments.Add("Format", $Format)
                $ConvertToTableArguments.Add("ApplyBorders", $True)
                $ConvertToTableArguments.Add("ApplyShading", $True)
                $ConvertToTableArguments.Add("ApplyFont", $True)
                $ConvertToTableArguments.Add("ApplyColor", $True)
                If (!$List) { 
                    $ConvertToTableArguments.Add("ApplyHeadingRows", $True) 
                }
                $ConvertToTableArguments.Add("ApplyLastRow", $True)
                $ConvertToTableArguments.Add("ApplyFirstColumn", $True)
                $ConvertToTableArguments.Add("ApplyLastColumn", $True)
            }

            ## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
            ## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
            Write-Debug ("$(Get-Date): `t`tConverting range to table")
            ## Store the table reference just in case we need to set alternate row coloring
            $WordTable = $WordRange.GetType().InvokeMember(
                "ConvertToTable", # Method name
                [System.Reflection.BindingFlags]::InvokeMethod, # Flags
                $Null, # Binder
                $WordRange, # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)), ## Named argument values
                $Null, # Modifiers
                $Null, # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
            )

            ## Implement grid lines (will wipe out any existing formatting)
            If ($Format -lt 0) {
                Write-Debug ("$(Get-Date): `t`tSetting table format")
                $WordTable.Style = $Format
				$WordTable.ApplyStyleFirstColumn = $False
            }

            ## Set the table autofit behavior
            If ($AutoFit -ne -1) { 
                $WordTable.AutoFitBehavior($AutoFit)
            }

            If (!$List) {
                #the next line causes the heading row to flow across page breaks
                $WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue
            }

            If (!$NoGridLines) {
                $WordTable.Borders.InsideLineStyle = $wdLineStyleSingle
                $WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle
            }
            If ($NoGridLines) {
                $WordTable.Borders.InsideLineStyle = $wdLineStyleNone
                $WordTable.Borders.OutsideLineStyle = $wdLineStyleNone
            }
            If ($NoInternalGridLines) {
                $WordTable.Borders.InsideLineStyle = $wdLineStyleNone
                $WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle
            }

            Return $WordTable
        } # end If not offline
    } ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This Function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) returns a single Word COM cells object.
#>

Function SetWordCellFormat {
    [CmdletBinding(DefaultParameterSetName = 'Collection')]
    Param (
        # Word COM object cell collection reference
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Collection', Position = 0)] [ValidateNotNullOrEmpty()] $Collection,
        # Word COM object individual cell reference
        [Parameter(Mandatory = $true, ParameterSetName = 'Cell', Position = 0)] [ValidateNotNullOrEmpty()] $Cell,
        # Hashtable of cell co-ordinates
        [Parameter(Mandatory = $true, ParameterSetName = 'Hashtable', Position = 0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
        # Word COM object table reference
        [Parameter(Mandatory = $true, ParameterSetName = 'Hashtable', Position = 1)] [ValidateNotNullOrEmpty()] $Table,
        # Font name
        [Parameter()] [AllowNull()] [string] $Font = $Null,
        # Font color
        [Parameter()] [AllowNull()] $Color = $Null,
        # Font size
        [Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
        # Cell background color
        [Parameter()] [AllowNull()] $BackgroundColor = $Null,
        # Force solid background color
        [Switch] $Solid,
        [Switch] $Bold,
        [Switch] $Italic,
        [Switch] $Underline
    )

    Begin {
        Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName)
    }

    Process {
        Switch ($PSCmdlet.ParameterSetName) {
            'Collection' {
                ForEach ($Cell in $Collection) {
                    If ($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor}
                    If ($Bold) { $Cell.Range.Font.Bold = $true}
                    If ($Italic) { $Cell.Range.Font.Italic = $true}
                    If ($Underline) { $Cell.Range.Font.Underline = 1}
                    If ($Null -ne $Font) { $Cell.Range.Font.Name = $Font}
                    If ($Null -ne $Color) { $Cell.Range.Font.Color = $Color}
                    If ($Size -ne 0) { $Cell.Range.Font.Size = $Size}
                    If ($Solid) { $Cell.Shading.Texture = 0} ## wdTextureNone
                } # end foreach
            } # end Collection
            'Cell' {
                If ($Bold) { $Cell.Range.Font.Bold = $true}
                If ($Italic) { $Cell.Range.Font.Italic = $true}
                If ($Underline) { $Cell.Range.Font.Underline = 1}
                If ($Null -ne $Font) { $Cell.Range.Font.Name = $Font}
                If ($Null -ne $Color) { $Cell.Range.Font.Color = $Color}
                If ($Size -ne 0) { $Cell.Range.Font.Size = $Size}
                If ($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor}
                If ($Solid) { $Cell.Shading.Texture = 0} ## wdTextureNone
            } # end Cell
            'Hashtable' {
                ForEach ($Coordinate in $Coordinates) {
                    $Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column)
                    If ($Bold) { $Cell.Range.Font.Bold = $true}
                    If ($Italic) { $Cell.Range.Font.Italic = $true}
                    If ($Underline) { $Cell.Range.Font.Underline = 1}
                    If ($Null -ne $Font) { $Cell.Range.Font.Name = $Font}
                    If ($Null -ne $Color) { $Cell.Range.Font.Color = $Color}
                    If ($Size -ne 0) { $Cell.Range.Font.Size = $Size}
                    If ($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor}
                    If ($Solid) { $Cell.Shading.Texture = 0} ## wdTextureNone
                }
            } # end Hashtable
        } # end switch
    } # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This Function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This Function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this Function is called by the AddWordTable Function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor {
    [CmdletBinding()]
    Param (
        # Word COM object table reference
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)] [ValidateNotNullOrEmpty()] $Table,
        # Alternate row background color
        [Parameter(Mandatory = $true, Position = 1)] [ValidateNotNull()] [int] $BackgroundColor,
        # Alternate row starting seed
        [Parameter(ValueFromPipelineByPropertyName = $true, Position = 2)] [ValidateSet('First', 'Second')] [string] $Seed = 'First'
    )

    Process {
        $StartDateTime = Get-Date
        Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime)

        ## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
        If ($Seed.ToLower() -eq 'second') { 
            $StartRowIndex = 2 
        }
        Else { 
            $StartRowIndex = 1 
        }

        For ($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) { 
            $Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor
        }

        ## I've put verbose calls in here we can see how expensive this Functionality actually is.
        $EndDateTime = Get-Date
        $ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime
        Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds)
    }
}
#endregion

#region general script Functions
Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel ) {
    #Function created 8-jan-2014 by Michael B. Smith
    If ( $object ) {
        If ((gm -Name $topLevel -InputObject $object)) {
            If ((gm -Name $secondLevel -InputObject $object.$topLevel)) {
                Return $True
            }
        }
    }
    Return $False
}

Function AbortScript {
    If ($MSWord -or $PDF) {
        $Script:Word.quit()
        Write-Verbose "$(Get-Date): System Cleanup"
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
        If (Test-Path variable:global:word) {
            Remove-Variable -Name word -Scope Global
        }
    }
    [gc]::collect() 
    [gc]::WaitForPendingFinalizers()
    Write-Verbose "$(Get-Date): Script has been aborted"
    $ErrorActionPreference = $SaveEAPreference
    Exit
}

Function ShowScriptOptions {
    Write-Verbose "$(Get-Date): "
    Write-Verbose "$(Get-Date): "
    Write-Verbose "$(Get-Date): AddDateTime     : $($AddDateTime)"
    If ($MSWORD -or $PDF) {
        Write-Verbose "$(Get-Date): Company Name    : $($Script:CoName)"
    }
    If ($MSWORD -or $PDF) {
        Write-Verbose "$(Get-Date): Cover Page      : $($CoverPage)"
    }
    Write-Verbose "$(Get-Date): Dev             : $($Dev)"
    If ($Dev) {
        Write-Verbose "$(Get-Date): DevErrorFile    : $($Script:DevErrorFile)"
    }
    Write-Verbose "$(Get-Date): Filename1       : $($Script:filename1)"
    If ($PDF) {
        Write-Verbose "$(Get-Date): Filename2       : $($Script:filename2)"
    }
    Write-Verbose "$(Get-Date): Folder          : $($Folder)"
    Write-Verbose "$(Get-Date): From            : $($From)"
    Write-Verbose "$(Get-Date): NSIP            : $($NSIP)"
    Write-Verbose "$(Get-Date): Save As PDF     : $($PDF)"
    Write-Verbose "$(Get-Date): Save As WORD    : $($MSWORD)"
    Write-Verbose "$(Get-Date): ScriptInfo      : $($ScriptInfo)"
    Write-Verbose "$(Get-Date): Smtp Port       : $($SmtpPort)"
    Write-Verbose "$(Get-Date): Smtp Server     : $($SmtpServer)"
    Write-Verbose "$(Get-Date): Title           : $($Script:Title)"
    Write-Verbose "$(Get-Date): To              : $($To)"
    Write-Verbose "$(Get-Date): Use NS SSL      : $($UseNSSSL)"
    Write-Verbose "$(Get-Date): Use SSL         : $($UseSSL)"
    If ($MSWORD -or $PDF) {
        Write-Verbose "$(Get-Date): User Name       : $($UserName)"
    }
    Write-Verbose "$(Get-Date): "
    Write-Verbose "$(Get-Date): OS Detected     : $($Script:RunningOS)"
    Write-Verbose "$(Get-Date): PoSH version    : $($Host.Version)"
    Write-Verbose "$(Get-Date): PSCulture       : $($PSCulture)"
    Write-Verbose "$(Get-Date): PSUICulture     : $($PSUICulture)"
    If ($MSWORD -or $PDF) {
        Write-Verbose "$(Get-Date): Word language   : $($Script:WordLanguageValue)"
        Write-Verbose "$(Get-Date): Word version    : $($Script:WordProduct)"
    }
    Write-Verbose "$(Get-Date): "
    Write-Verbose "$(Get-Date): Script start    : $($Script:StartTime)"
    Write-Verbose "$(Get-Date): "
    Write-Verbose "$(Get-Date): "
}

Function SaveandCloseDocumentandShutdownWord {
    #bug fix 1-Apr-2014
    #reset Grammar and Spelling options back to their original settings
    $Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
    $Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

    Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
    If ($Script:WordVersion -eq $wdWord2010) {
        #the $saveFormat below passes StrictMode 2
        #I found this at the following two links
        #http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
        #http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
        If ($PDF) {
            Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
        }
        Else {
            Write-Verbose "$(Get-Date): Saving DOCX file"
        }
        If ($AddDateTime) {
            $Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
            If ($PDF) {
                $Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
            }
        }
        Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
        $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
        $Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
        If ($PDF) {
            Write-Verbose "$(Get-Date): Now saving as PDF"
            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
            $Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
        }
    }
    ElseIf ($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016) {
        If ($PDF) {
            Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
        }
        Else {
            Write-Verbose "$(Get-Date): Saving DOCX file"
        }
        If ($AddDateTime) {
            $Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
            If ($PDF) {
                $Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
            }
        }
        Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
        $Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
        If ($PDF) {
            Write-Verbose "$(Get-Date): Now saving as PDF"
            $Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
        }
    }

    Write-Verbose "$(Get-Date): Closing Word"
    $Script:Doc.Close()
    $Script:Word.Quit()
    If ($PDF) {
        [int]$cnt = 0
        While (Test-Path $Script:FileName1) {
            $cnt++
            If ($cnt -gt 1) {
                Write-Verbose "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $($cnt))"
                Start-Sleep -Seconds 10
                $Script:Word.Quit()
                If ($cnt -gt 2) {
                    #kill the winword process

                    #find out our session (usually "1" except on TS/RDC or Citrix)
                    $SessionID = (Get-Process -PID $PID).SessionId
					
                    #Find out if winword is running in our session
                    $wordprocess = ((Get-Process 'WinWord' -ea 0) | ? { $_.SessionId -eq $SessionID }).Id
                    If ($wordprocess -gt 0) {
                        Write-Verbose "$(Get-Date): Attempting to stop WinWord process # $($wordprocess)"
                        Stop-Process $wordprocess -EA 0
                    }
                }
            }
            Write-Verbose "$(Get-Date): Attempting to delete $($Script:FileName1) since only $($Script:FileName2) is needed (try # $($cnt))"
            Remove-Item $Script:FileName1 -EA 0 4>$Null
        }
    }
    Write-Verbose "$(Get-Date): System Cleanup"
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
    If (Test-Path variable:global:word) {
        Remove-Variable -Name word -Scope Global 4>$Null
    }
    $SaveFormat = $Null
    [gc]::collect() 
    [gc]::WaitForPendingFinalizers()
	
    #is the winword process still running? kill it

    #find out our session (usually "1" except on TS/RDC or Citrix)
    $SessionID = (Get-Process -PID $PID).SessionId

    #Find out if winword is running in our session
    $wordprocess = $Null
    $wordprocess = ((Get-Process 'WinWord' -ea 0) | ? { $_.SessionId -eq $SessionID }).Id
    If ($null -ne $wordprocess -and $wordprocess -gt 0) {
        Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
        Stop-Process $wordprocess -EA 0
    }
}

Function SetFileName1andFileName2 {
    Param([string]$OutputFileName)
    #set $filename1 and $filename2 with no file extension
    If ($AddDateTime) {
        [string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName)"
        If ($PDF) {
            [string]$Script:FileName2 = "$($Script:pwdpath)\$($OutputFileName)"
        }
    }
    If ($MSWord -or $PDF) {
        If (!$Offline) {
            CheckWordPreReq
        }
        If (!$AddDateTime) {
            [string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName).docx"
            If ($PDF) {
                [string]$Script:FileName2 = "$($Script:pwdpath)\$($OutputFileName).pdf"
            }
        }
        If (!$Offline) {
            SetupWord
        }
    }
}
#endregion

#region script end
Function ProcessScriptEnd {
    Write-Verbose "$(Get-Date): Script has completed"
    Write-Verbose "$(Get-Date): "

    #http://poshtips.com/measuring-elapsed-time-in-powershell/
    Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
    Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
    $runtime = $(Get-Date) - $Script:StartTime
    $Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
        $runtime.Days,
        $runtime.Hours,
        $runtime.Minutes,
        $runtime.Seconds,
        $runtime.Milliseconds)
    Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

    If ($Dev) {
        If ($SmtpServer -eq "") {
            Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
        }
        Else {
            Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
        }
    }

    If ($ScriptInfo) {
        $SIFile = "$Script:pwdpath\NSInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
        Out-File -FilePath $SIFile -InputObject "" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Add DateTime   : $($AddDateTime)" 4>$Null
        If ($MSWORD -or $PDF) {
            Out-File -FilePath $SIFile -Append -InputObject "Company Name   : $($Script:CoName)" 4>$Null		
        }
        If ($MSWORD -or $PDF) {
            Out-File -FilePath $SIFile -Append -InputObject "Cover Page     : $($CoverPage)" 4>$Null
        }
        Out-File -FilePath $SIFile -Append -InputObject "Dev            : $($Dev)" 4>$Null
        If ($Dev) {
            Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile   : $($Script:DevErrorFile)" 4>$Null
        }
        Out-File -FilePath $SIFile -Append -InputObject "Filename1      : $($Script:FileName1)" 4>$Null
        If ($PDF) {
            Out-File -FilePath $SIFile -Append -InputObject "Filename2      : $($Script:FileName2)" 4>$Null
        }
        Out-File -FilePath $SIFile -Append -InputObject "Folder         : $($Folder)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "From           : $($From)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "NSIP           : $($NSIP)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Save As PDF    : $($PDF)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Save As WORD   : $($MSWORD)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Script Info    : $($ScriptInfo)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Smtp Port      : $($SmtpPort)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Smtp Server    : $($SmtpServer)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Title          : $($Script:Title)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "To             : $($To)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Use NS SSL     : $($UseNSSSL)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Use SSL        : $($UseSSL)" 4>$Null
        If ($MSWORD -or $PDF) {
            Out-File -FilePath $SIFile -Append -InputObject "User Name      : $($UserName)" 4>$Null
        }
        Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "OS Detected    : $($Script:RunningOS)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "PoSH version   : $($Host.Version)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "PSCulture      : $($PSCulture)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "PSUICulture    : $($PSUICulture)" 4>$Null
        If ($MSWORD -or $PDF) {
            Out-File -FilePath $SIFile -Append -InputObject "Word language  : $($Script:WordLanguageValue)" 4>$Null
            Out-File -FilePath $SIFile -Append -InputObject "Word version   : $($Script:WordProduct)" 4>$Null
        }
        Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Script start   : $($Script:StartTime)" 4>$Null
        Out-File -FilePath $SIFile -Append -InputObject "Elapsed time   : $($Str)" 4>$Null
    }

    $ErrorActionPreference = $SaveEAPreference
    [gc]::collect()
}
#endregion

#region general script Functions
Function ProcessDocumentOutput {
    If ($MSWORD -or $PDF) {
        SaveandCloseDocumentandShutdownWord
    }

    $GotFile = $False

    If ($PDF) {
        If (Test-Path "$($Script:FileName2)") {
            Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
            $GotFile = $True
        }
        Else {
            Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
            Write-Error "Unable to save the output file, $($Script:FileName2)"
        }
    }
    Else {
        If (Test-Path "$($Script:FileName1)") {
            Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
            $GotFile = $True
        }
        Else {
            Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
            Write-Error "Unable to save the output file, $($Script:FileName1)"
        }
    }
	
    #email output file if requested
    If ($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer )) {
        If ($PDF) {
            $emailAttachment = $Script:FileName2
        }
        Else {
            $emailAttachment = $Script:FileName1
        }
        SendEmail $emailAttachment
    }
    [gc]::collect()
}
#endregion

#Script begins

$script:startTime = Get-Date
#endregion script template

#region file name and title name
#The Function SetFileName1andFileName2 needs your script output filename
#change title for your report
[string]$Script:Title = "NetScaler Documentation $($Script:CoName)"
SetFileName1andFileName2 $ReportFileName
#endregion file name and title name
#endregion script template

#region Documentation Script Complete
#region Functions
#Variables for Progress Bar
[int]$script:ProgressSteps = 300
[int]$script:Progress = 0

## Barry Schiffer Use Stopwatch class to time script execution
#$sw = [Diagnostics.Stopwatch]::StartNew()

##Disable Strict Mode to handle missing parameters
Set-StrictMode -Off
$selection.InsertNewPage()

#Check Paths for import/export
If ($Offline) {
    #Does the export path exists
    If (!(Test-Path "$OfflinePath\")) {
        #If not then try and create it
        New-Item -Path "$OfflinePath\" -ItemType directory | Out-Null
        If (!(Test-Path "$OfflinePath\")) {
            #If it still doesn't exist then something is wrong - so exit
            Write-Host "Unable to find or create the export path: $OfflinePath" -ForegroundColor Red
            Write-Host "Please try again with a different path using the -ExportPath parameter or confirm the location is accessible. Exiting." -ForegroundColor Red
            Exit
        }
    }
}

If ($Import) {
    #Does the import path exist
    If (!(Test-Path "$ImportPath\")) {
        #If it doesn't exist then something is wrong - so exit
        Write-Host "Unable to find the import path: $ImportPath" -ForegroundColor Red
        Write-Host "Please try again with a different path using the -ImportPath parameter or confirm the location is accessible. Exiting." -ForegroundColor Red
        Exit
    }
    Else {
        $OfflinePath = $ImportPath
    }
  
}

#region Nitro Functions
Function Get-vNetScalerObjectList {
    <#
        .SYNOPSIS
            Returns a list of objects available in a NetScaler Nitro API container.
    #>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API Container, i.e. nitro/v1/stat/ or nitro/v1/config/
        [Parameter(Mandatory)] [ValidateSet('Stat', 'Config')] [System.String] $Container = 'Config'
    )
    begin {
        $Container = $Container.ToLower()
    }
    process {
        If ($script:nsSession.UseNSSSL) { $protocol = 'https'}
        Else { $protocol = 'http'}
        $uri = '{0}://{1}/nitro/v1/{2}/' -f $protocol, $script:nsSession.Address, $Container
        $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container
        $methodResponse = '{0}objects' -f $Container.ToLower()
        Write-Output $restResponse.($methodResponse).objects
    }
} #end Function Get-vNetScalerObjectList

Function Get-vNetScalerObject {
    <#
        .SYNOPSIS
            Returns a NetScaler Nitro API object(s) via its REST API.
    #>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API resource type, e.g. /nitro/v1/config/LBVSERVER
        [Parameter(Mandatory)] [Alias('Object', 'Type')] [System.String] $ResourceType,
        # NetScaler Nitro API resource name, e.g. /nitro/v1/config/lbvserver/MYLBVSERVER
        [Parameter()] [Alias('Name')] [System.String] $ResourceName,
        # NetScaler Nitro API optional attributes, e.g. /nitro/v1/config/lbvserver/mylbvserver?ATTRS=<attrib1>,<attrib2>
        [Parameter()] [System.String[]] $Attribute,
        # NetScaler Nitro API Container, i.e. nitro/v1/stat/ or nitro/v1/config/
        [Parameter()] [ValidateSet('Stat', 'Config')] [System.String] $Container = 'config',
        # Retrieve Builk Bindings for an object
        [Parameter()] [Alias('Bulk')] [switch] $BulkBindings = $false
    )
    begin {
	# To Lower does not work with items where the items are with capitals
	$Container = $Container.ToLower()
	$ResourceType = $ResourceType.ToLower()
        #$ResourceName = $ResourceName.ToLower()
    }
    process {
        If ($script:nsSession.UseNSSSL) { $protocol = 'https'}
        Else { $protocol = 'http'}
        $uri = '{0}://{1}/nitro/v1/{2}/{3}' -f $protocol, $script:nsSession.Address, $Container, $ResourceType
        If ($ResourceName) { $uri = '{0}/{1}' -f $uri, $ResourceName}
        If ($Attribute) {
            $attrs = [System.String]::Join(',', $Attribute)
            $uri = '{0}?attrs={1}' -f $uri, $attrs
        }
        If ($BulkBindings) {
            $uri = '{0}?bulkbindings=yes' -f $uri
        }
        $uri = [System.Uri]::EscapeUriString($uri)
        Write-Log "Get-vNetScalerObject Request URL: $uri"
        If (!$Import) {
            $restResponse = InvokevNetScalerNitroMethod -Uri $uri -Container $Container
            Write-Log "REST Response: $restResponse"
        }
        If ($Offline) {
            Write-Log "Convert URI to ASCII: $uri"
            $FileNameBytes = [System.Text.Encoding]::ASCII.GetBytes($uri)
            Write-Log "ASCII Encoded: $FileNameBytes"
            $tmpFileName = Get-CleanBase64([System.Convert]::ToBase64String($FileNameBytes))
            Write-Log "Base64 Encoded File Name: $tmpFileName"
            $OfflineExportPath = Join-Path -Path $OfflinePath -ChildPath "$tmpFileName.xml"
            Write-Log "Export Path: $OfflineExportPath"
            $LongPath = $false
            #Check Path Length
            If ($OfflineExportPath.Length -ge 254) {
                Write-Log "Path exceeds path limit - converting to literal path"
                $LongPath = $true
                #Make path literal
                $BasePath = '\\?\'
                $PathInfo = [System.URI]$OfflineExportPath
                If ($PathInfo.IsUNC) {
                    $BasePath = '\\?\UNC\'
                }
                $LiteralPath = Join-Path -Path $BasePath -ChildPath $OfflineExportPath
                Write-Log "New Literal Path: $LiteralPath"
            }
            #Disable-Verbose
            Try {
                If (!$LongPath) {
                    $restResponse.($ResourceType) | Export-CliXML -Path $OfflineExportPath | Out-Null
                }
                Else {
                    Try {
                        $restResponse.($ResourceType) | Export-CliXML -LiteralPath $LiteralPath | Out-Null
                    }
                    Catch {
                        Write-Log $_.exception
                        Write-Warning $_.exception
                    }
                }
            }
            Catch {
                Write-Log "$($_.Exception)"
            }
            Write-Output $restResponse.($ResourceType)
            #Enable-Verbose
        }
        ElseIf ($Import) {
            Write-Log "Convert URI to ASCII: $uri"
            $FileNameBytes = [System.Text.Encoding]::ASCII.GetBytes($uri)
            Write-Log "ASCII Encoded: $FileNameBytes"
            $tmpFileName = Get-CleanBase64([System.Convert]::ToBase64String($FileNameBytes))
            Write-Log "Base64 Encoded File Name: $tmpFileName"
            $OfflineExportPath = Join-Path -Path $OfflinePath -ChildPath "$tmpFileName.xml"
            Write-Log "Import Path: $OfflineExportPath"
            $LongPath = $false
            #Check Path Length
            If ($OfflineExportPath.Length -ge 254) {
                Write-Log "Path exceeds path limit - converting to literal path"
                $LongPath = $true
                #Make path literal
                $BasePath = '\\?\'
                $PathInfo = [System.URI]$OfflineExportPath
                If ($PathInfo.IsUNC) {
                    $BasePath = '\\?\UNC\'
                }
                $LiteralPath = Join-Path -Path $BasePath -ChildPath $OfflineExportPath
                Write-Log "New Literal Path: $LiteralPath"
            }
            Try {
                If (!$LongPath) {
                    Import-Clixml -Path $OfflineExportPath | Write-Output
                }
                Else {
                    Try {
                        Import-Clixml -LiteralPath $LiteralPath | Write-Output
                    }
                    Catch {
                        Write-Log $_.exception
                        Write-Warning $_.exception
                    }
                }
            }
            Catch {
                Write-Log "$($_.Exception)"
            }
        }
        Else {
            If ($null -ne $restResponse.($ResourceType)) { Write-Output $restResponse.($ResourceType)}
            Else { Write-Output $restResponse }
        }
    }
} #end Function Get-vNetScalerObject

Function Get-vNetScalerFile {
    <#
        .SYNOPSIS
            Returns a NetScaler Nitro API SystemFile object(s) via its REST API.
    #>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API resource name, e.g. /nitro/v1/config/SystemFile?args=filename:Filename,filelocation:FileLocation
        [Parameter()] [Alias('Name')][System.String] $FileName,
        # NetScaler Nitro API optional attributes, e.g. /nitro/v1/config/lbvserver/mylbvserver?ATTRS=<attrib1>,<attrib2>
        [Parameter()] [Alias('Location')] [System.String] $FileLocation
    )
    begin {
        #Don't lower case these as they are case sensitive
        #$FileName = $FileName.ToLower();
        $FileLocation = $FileLocation.Replace("/", "%2F")
        $Container = "config"
    }
    process {
        If ($script:nsSession.UseNSSSL) { $protocol = 'https'}
        Else { $protocol = 'http'}
        $uri = '{0}://{1}/nitro/v1/config/systemfile/{2}?args=filelocation:{3}' -f $protocol, $script:nsSession.Address, $FileName, $FileLocation
        
        #Don't URI encode as we've already replaced / with %2F as required - URL encoding after this, encodes the % which breaks the request
        #$uri = [System.Uri]::EscapeUriString($uri);
        #Write-Output $uri;
        Write-Log "Get-vNetScalerFile Request: $uri"
        If (!$Import) {
            $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container
        }

        If ($Offline) {
            $FileNameBytes = [System.Text.Encoding]::ASCII.GetBytes($uri)
            $tmpFileName = Get-CleanBase64([System.Convert]::ToBase64String($FileNameBytes))
            $OfflineExportPath = Join-Path -Path $OfflinePath -ChildPath "$tmpFileName.xml"
            #Disable-Verbose
            Write-Log "Export Path: $OfflineExportPath"
            $LongPath = $false
            #Check Path Length
            If ($OfflineExportPath.Length -ge 254) {
                Write-Log "Path exceeds path limit - converting to literal path"
                $LongPath = $true
                #Make path literal
                $BasePath = '\\?\'
                $PathInfo = [System.URI]$OfflineExportPath
                If ($PathInfo.IsUNC) {
                    $BasePath = '\\?\UNC\'
                }
                $LiteralPath = Join-Path -Path $BasePath -ChildPath $OfflineExportPath
                Write-Log "New Literal Path: $LiteralPath"
            }
            Try {
                If (!$LongPath) {
                    $restResponse.systemfile | Export-CliXML -Path $OfflineExportPath | Out-Null
                }
                Else {
                    Try {
                        $restResponse.systemfile | Export-CliXML -LiteralPath $LiteralPath | Out-Null
                    }
                    Catch {
                        Write-Log $_.exception
                        Write-Warning $_.exception
                    }
                }
            }
            Catch {
                Write-Log "$($_.Exception)"
            }
            #Enable-Verbose
            Write-Output $restResponse.systemfile
        }
        ElseIf ($Import) {
            $FileNameBytes = [System.Text.Encoding]::ASCII.GetBytes($uri)
            $tmpFileName = Get-CleanBase64([System.Convert]::ToBase64String($FileNameBytes))
            $OfflineExportPath = Join-Path -Path $OfflinePath -ChildPath "$tmpFileName.xml"
            Write-Log "Import Path: $OfflineExportPath"
            $LongPath = $false
            #Check Path Length
            If ($OfflineExportPath.Length -ge 254) {
                Write-Log "Path exceeds path limit - converting to literal path"
                $LongPath = $true
                #Make path literal
                $BasePath = '\\?\'
                $PathInfo = [System.URI]$OfflineExportPath
                If ($PathInfo.IsUNC) {
                    $BasePath = '\\?\UNC\'
                }
                $LiteralPath = Join-Path -Path $BasePath -ChildPath $OfflineExportPath
                Write-Log "New Literal Path: $LiteralPath"
            }
            Try {
                If (!$LongPath) {
                    Import-Clixml -Path $OfflineExportPath | Write-Output
                }
                Else {
                    Import-Clixml -LiteralPath $LiteralPath | Write-Output 
                }
            }
            Catch {
                Write-Log "$($_.Exception)"
            }
        }
        Else {
            If ($null -ne $restResponse.systemfile) { Write-Output $restResponse.systemfile}
            Else { Write-Output $restResponse }
        }
    }
} #end Function Get-vNetScalerFile

Function Read-vNetScalerFile {
    [CmdletBinding()]
    param (
        # NetScaler Nitro API resource name, e.g. /nitro/v1/config/SystemFile?args=filename:Filename,filelocation:FileLocation
        [Parameter()] [Alias('Name')][System.String] $FileName,
        # NetScaler Nitro API optional attributes, e.g. /nitro/v1/config/lbvserver/mylbvserver?ATTRS=<attrib1>,<attrib2>
        [Parameter()] [Alias('Location')] [System.String] $FileLocation    
    )
    Write-Host "Running"

    #Don't lower case these as they are case sensitive
    #$FileName = $FileName.ToLower();
    $FileLocation = $FileLocation.Replace("/", "%2F")
    Write-Host $FileLocation
    $Container = "config"
    
    # process {
    If ($script:nsSession.UseSSL) { $protocol = 'https'} Else { $protocol = 'http'}
    $uri = '{0}://{1}/rapi/read_file?filter=path:{2}' -f $protocol, $script:nsSession.Address, $FileLocation
    write-host $uri
        
    #Don't URI encode as we've already replaced / with %2F as required - URL encoding after this, encodes the % which breaks the request
    #$uri = [System.Uri]::EscapeUriString($uri);
    #Write-Output $uri;
    $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container
    If ($null -ne $restResponse.systemfile) { Write-Output $restResponse}
    Else { Write-Output $restResponse }
    #}
} #end Function Get-vNetScalerFile

Function InvokevNetScalerNitroMethod {
    <#
        .SYNOPSIS
            Calls a fully qualified NetScaler Nitro API
        .NOTES
            This is an internal Function and shouldn't be called directly
    #>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API uniform resource identifier
        [Parameter(Mandatory)] [string] $Uri,
        # NetScaler Nitro API Container, i.e. nitro/v1/stat/ or nitro/v1/config/
        [Parameter(Mandatory)] [ValidateSet('Stat', 'Config')] [string] $Container
    )
    begin {
        If ($script:nsSession -eq $null) { throw 'No valid NetScaler session configuration.'}
        If ($script:nsSession.Session -eq $null -or $script:nsSession.Expiry -eq $null) { throw 'Invalid NetScaler session cookie.'}
        If ($script:nsSession.Expiry -lt (Get-Date)) { throw 'NetScaler session has expired.'}
    }
    process {
        $irmParameters = @{
            Uri         = $Uri
            Method      = 'Get'
            WebSession  = $script:nsSession.Session
            ErrorAction = 'Continue'
            Verbose     = ($PSBoundParameters['Debug'] -eq $true)
        }
        If (!$Import) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
            Try {
                $response = Invoke-RestMethod @irmParameters
            }
            Catch {
                Write-Log "Error: $_.Exception"
            }  
            Write-Output $response
            #Write-Output (Invoke-RestMethod @irmParameters);
        }
    }
} #end Function InvokevNetScalerNitroMethod

Function Connect-vNetScalerSession {
    <#
        .SYNOPSIS
            Authenticates to the NetScaler and stores a session cookie.
    #>
    [CmdletBinding(DefaultParameterSetName = 'HTTP')]
    [OutputType([Microsoft.PowerShell.Commands.WebRequestSession])]
    param (
        # NetScaler uniform resource identifier
        [Parameter(Mandatory, ParameterSetName = 'HTTP')]
        [Parameter(Mandatory, ParameterSetName = 'HTTPS')]
        [System.String] $ComputerName,
        # NetScaler session timeout (seconds)
        [Parameter(ParameterSetName = 'HTTP')]
        [Parameter(ParameterSetName = 'HTTPS')]
        [ValidateNotNull()]
        [System.Int32] $Timeout = 3600,
        # NetScaler authentication credentials
        [Parameter(ParameterSetName = 'HTTP')]
        [Parameter(ParameterSetName = 'HTTPS')]
        [System.Management.Automation.PSCredential] $Credential,
        ## EXPERIMENTAL: Require SSL/TLS, e.g. https://. This requires the client to trust to the NetScaler's certificate.
        [Parameter(ParameterSetName = 'HTTPS')] [System.Management.Automation.SwitchParameter] $UseNSSSL
    )
    process {
        Write-Log "Connecting to NetScaler"
        If (!$Credential) {
            Write-Log "No PSCredential object found."
            If (($NSUserName -eq "") -or ($NSPassword -eq "")) {
                write-log "Either username or password parameters have not been provided."
                If (!$Import) {
                    Write-Log "Prompt for credentials"
                    $Credential = $(Get-Credential -Message "Provide NetScaler credentials for '$ComputerName'"; )
                }
            }
            Else {
                Write-Log "Create PSCredential Object using provided credentials"
                $SecurePassword = Convertto-SecureString $NSPassword -AsPlainText -Force
                $Credential = New-Object System.Management.Automation.PSCredential($NSUserName, $SecurePassword)
            }
        }

        If ($UseNSSSL) { $protocol = 'https'}
        Else { $protocol = 'http'}
        $script:nsSession = @{ Address = $ComputerName; UseNSSSL = $UseNSSSL }
        $json = '{{ "login": {{ "username": "{0}", "password": "{1}", "timeout": {2} }} }}'
        $invokeRestMethodParams = @{
            Uri             = ('{0}://{1}/nitro/v1/config/login' -f $protocol, $ComputerName)
            Method          = 'Post'
            Body            = ($json -f $Credential.UserName, $Credential.GetNetworkCredential().Password, $Timeout)
            ContentType     = 'application/json'
            SessionVariable = 'nsSessionCookie'
            ErrorAction     = 'Stop'
            Verbose         = ($PSBoundParameters['Debug'] -eq $true)
        }
        If (!$Import) {
            Try {
                $restResponse = Invoke-RestMethod @invokeRestMethodParams
                Write-Log "Login Response: $restResponse"
            }
            Catch {
                SaveandCloseDocumentandShutdownWord
                Remove-Variable -Name nsSession -Scope Script
                Write-Log "Login Status: $($_.Exception)" 
                Write-Error "
			`n`n
			`t`t
			Failed to connect to NetScaler: $($_.Exception)
			`n`n
			"
                Exit
            }
        }
        ## Store the session cookie at the script scope
        $script:nsSession.Session = $nsSessionCookie
        ## Store the session expiry
        $script:nsSession.Expiry = (Get-Date).AddSeconds($Timeout)
        ## Return the Rest Method response
        Write-Output $restResponse
        
    }
} #end Function Connect-vNetScalerSession

Function Logout-vNetScalerSession {
    <#
        .SYNOPSIS
            Authenticates to the NetScaler and stores a session cookie.
    #>
    process {
        Write-Log "Logout NetScaler Session"
        If ($UseNSSSL) { $protocol = 'https'}
        Else { $protocol = 'http'}
        $json = '{{ "logout": {}}'
        $irmParameters = @{
            Uri         = ('{0}://{1}/nitro/v1/config/logout' -f $protocol, $script:nsSession.Address) 
            Method      = 'Post'
            Body        = $json
            ContentType = 'application/vnd.com.citrix.netscaler.logout+json'
            WebSession  = $script:nsSession.Session
            ErrorAction = 'Stop'
            Verbose     = ($PSBoundParameters['Debug'] -eq $true)
        }
        If (!$Import) {
            $restResponse = Invoke-RestMethod @irmParameters
        }
        #Remove the Session Variable

        Write-Output $restResponse
        Remove-Variable -Name nsSession -Scope Script
    }
} #end Function Logout-vNetScalerSession

Function Get-vNetScalerObjectCount {
    <#
        .Synopsis
            Returns an individual NetScaler Nitro API object.
    #>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API Object, e.g. /nitro/v1/config/NSVERSION
        [Parameter(Mandatory)] [Alias('Object', 'Type')] [string] $ResourceType,
        # NetScaler Nitro API resource name, e.g. /nitro/v1/config/lbvserver/MYLBVSERVER
        [Parameter()] [Alias('Name')] [System.String] $ResourceName,
        # NetScaler Nitro API Container, i.e. nitro/v1/stat/ or nitro/v1/config/
        [Parameter()] [ValidateSet('Stat', 'Config')] [string] $Container = 'config'
    )

    begin {
        ## Check session cookie
        If ($script:nsSession.Session -eq $null) { throw 'Invalid NetScaler session cookie.'}
        If ($script:nsSession.UseNSSSL) { $protocol = 'https'}
        Else { $protocol = 'http'}
    }

    process {
        If ($ResourceName) {
            $uri = '{0}://{1}/nitro/v1/{2}/{3}/{4}?count=yes' -f $protocol, $script:nsSession.Address, $Container.ToLower(), $ResourceType.ToLower(), $ResourceName
        }
        Else {
            $uri = '{0}://{1}/nitro/v1/{2}/{3}?count=yes' -f $protocol, $script:nsSession.Address, $Container.ToLower(), $ResourceType.ToLower()
        }
        write-log "Get-vNetScalerObjectCount Request URL: $uri"
        If (!$Import) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
            try {
                $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container
            }
            Catch {
                Write-Log "Error: $_.Exception"
            }
        }
        
        If ($Offline) {
            $FileNameBytes = [System.Text.Encoding]::ASCII.GetBytes($uri)
            $tmpFileName = Get-CleanBase64([System.Convert]::ToBase64String($FileNameBytes))
            $OfflineExportPath = Join-Path -Path $OfflinePath -ChildPath "$tmpFileName.xml"
            Write-Log "Export Path: $OfflineExportPath"
            $LongPath = $false
            #Check Path Length
            If ($OfflineExportPath.Length -ge 254) {
                Write-Log "Path exceeds path limit - converting to literal path"
                $LongPath = $true
                #Make path literal
                $BasePath = '\\?\'
                $PathInfo = [System.URI]$OfflineExportPath
                If ($PathInfo.IsUNC) {
                    $BasePath = '\\?\UNC\'
                }
                $LiteralPath = Join-Path -Path $BasePath -ChildPath $OfflineExportPath
                Write-Log "New Literal Path: $LiteralPath"
            }
            #Disable-Verbose
            try {
                If (!$LongPath) { 
                    $restResponse.($ResourceType.ToLower()) | Export-CliXML -Path $OfflineExportPath | Out-Null
                }
                Else {
                    Try {
                        $restResponse.($ResourceType.ToLower()) | Export-CliXML -LiteralPath $LiteralPath | Out-Null
                    }
                    Catch {
                        Write-Log $_.exception
                        Write-Warning $_.exception
                    }
                }
            }
            Catch {
                Write-Log $_.exception
            }
            #Enable-Verbose
            Write-Output $restResponse.($ResourceType.ToLower())
        }
        ElseIf ($Import) {
            $FileNameBytes = [System.Text.Encoding]::ASCII.GetBytes($uri)
            $tmpFileName = Get-CleanBase64([System.Convert]::ToBase64String($FileNameBytes))
            $OfflineExportPath = Join-Path -Path $OfflinePath -ChildPath "$tmpFileName.xml"
            Write-Log "Import Path: $OfflineExportPath"
            If ($OfflineExportPath.Length -ge 254) {
                Write-Log "Path exceeds path limit - converting to literal path"
                $LongPath = $true
                #Make path literal
                $BasePath = '\\?\'
                $PathInfo = [System.URI]$OfflineExportPath
                If ($PathInfo.IsUNC) {
                    $BasePath = '\\?\UNC\'
                }
                $LiteralPath = Join-Path -Path $BasePath -ChildPath $OfflineExportPath
                Write-Log "New Literal Path: $LiteralPath"
            }
            try { 
                If (!$LongPath) {
                    Import-Clixml -Path $OfflineExportPath | Write-Output
                }
                Else {
                    Try {
                        Import-Clixml -LiteralPath $LiteralPath | Write-Output
                    }
                    Catch {
                        Write-Log $_.exception
                        Write-Warning $_.exception
                    }
                }
            }
            Catch {
                Write-Log $_.exception
            }
        }
        Else {
            # $objectResponse = '{0}objects' -f $Container.ToLower();
            Write-Output $restResponse.($ResourceType.ToLower())
        }
    }
}
#endregion Nitro Functions

#region CSS Functions
Function Get-AttributeFromCSS {
    <#
        .Synopsis
            Searches for and returns a defined CSS Attribute from a custom portal theme
    #>

    [CmdletBinding()]
    param (
        # Search pattern to be used
        [Parameter(Mandatory)] [string] $SearchPattern,
        # Name of attribute we're looking for
        [Parameter(Mandatory)] [string] $Attribute,
        # NNumber of lines to search for attribute after pattern match
        [Parameter(Mandatory)] [int] $Lines,
        # Input string we're going to searching
        [Parameter(Mandatory)] [string] $Inputstring
    )

    $arrInput = $InputString.Split([Environment]::NewLine)

    $filteredInput = $arrInput | Select-String -Pattern "$SearchPattern" -Context  0, $Lines

    [String]$filteredOutput = $FilteredInput.Context.PostContext | Select-String -Pattern "$Attribute"
    If ($filteredOutput) {
        $arrOutput = $filteredOutput.Split(":")
        $cleanedOutput = $arrOutput[$arrOutput.Length - 1].Trim()
        $cleanedOutput = $cleanedOutput.Replace(";", "") #remove the trailing semi-colon
        Write-Output $cleanedOutput
    }
    Else { Write-Output "Undefined" }
}
#endregion CSS Functions

#region generic Functions
Function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}
Function Write-Log([String]$Message) {
    If ($Log) {
        Write-Output "$(Get-TimeStamp) $Message" | Out-file $Script:LogFile -append
    }
}
Function Enable-Verbose() {
    $VerbosePreference = "Continue"
    Return $VerbosePreference
}
Function Disable-Verbose() {
    $VerbosePreference = "SilentlyContinue"
    Return $VerbosePreference
}
Function IsNull($objectToCheck) {
    If ($objectToCheck -eq $null) {
        return $true
    }
    If ($objectToCheck -is [String] -and $objectToCheck -eq [String]::Empty) {
        return $true
    }
    If ($objectToCheck -is [DBNull] -or $objectToCheck -is [System.Management.Automation.Language.NullString]) {
        return $true
    }
    return $false
}
Function Get-NonEmptyString($String) {
    If (-not $String) {
        Return "N/A"
    }
    Else {
        Return "$String "
    }
}
Function Get-CleanBase64($String) {
    Write-Log "Unchanged BASE64 encoded string: $String"
    $String = $String.Replace("/", "_")
    $String = $String.Replace("=", "_")
    $String = $String.Replace("+", "_")
    Write-Log "Removing unsafe characters"
    Write-Output $String
}
Function Get-StringFromBase64 {
    [CmdletBinding()]
    param (
        # Base64 Encoded String
        [Parameter(Mandatory)] [string] $Object,
        # IS the file contents UTF8 or ASCII encoded
        [Parameter(Mandatory)] [ValidateSet('UTF8', 'ASCII')] [string] $Encoding
    )
    Switch ($Encoding) {
        "UTF8" { $output = [System.Text.Encoding]::UTF8.Getstring([System.convert]::FromBase64String($Object)) }
        "ASCII" { $output = [System.Text.Encoding]::ASCII.Getstring([System.convert]::FromBase64String($Object)) }
        Default { $output = [System.Text.Encoding]::UTF8.Getstring([System.convert]::FromBase64String($Object)) }
    }
    Return $output
}
Function Set-Progress($Status) {
    $script:Progress++
    Write-Progress -Id 1 -Activity "NetScaler Documentation Script" -Status "Processing: $Status" -PercentComplete (($script:Progress / $script:ProgressSteps) * 100)
	Write-Verbose "$(Get-Date): $Status"
}
Function Close-Progress() { Write-Progress -Id 1 -Activity "NetScaler Documentation Script" -Completed }
#endregion generic Functions

#region Connect
If ($UseNSSSL) {
    ##Allow connecting to untrusted SSL Certificates
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    $AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
    [System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback =
    [System.Linq.Expressions.Expression]::Lambda(
        [System.Net.Security.RemoteCertificateValidationCallback],
        [System.Linq.Expressions.Expression]::Constant($true),
        [System.Linq.Expressions.ParameterExpression[]](
            [System.Linq.Expressions.Expression]::Parameter(
                [object], 'sender'),
            [System.Linq.Expressions.Expression]::Parameter(
                [X509Certificate], 'certificate'),
            [System.Linq.Expressions.Expression]::Parameter(
                [System.Security.Cryptography.X509Certificates.X509Chain], 'chain'),
            [System.Linq.Expressions.Expression]::Parameter(
                [System.Net.Security.SslPolicyErrors], 'sslPolicyErrors'))).
    Compile()
}
Set-Progress "Connecting to NetScaler"
## Ensure we can connect to the NetScaler appliance before we spin up Word!
## Connect to the API if there is no session cookie
## Note: repeated logons will result in 'Connection limit to cfe exceeded' errors.
If ($Import) {
    $OfflineExportPath = Join-Path -Path $OfflinePath -ChildPath "nsSession.xml"
    If (Test-Path -Path $OfflineExportPath) {
        $script:nsSession = Import-Clixml -Path $OfflineExportPath
    }
}
If (-not (Get-Variable -Name nsSession -Scope Script -ErrorAction SilentlyContinue)) { 
    Write-Log "nsSession variable doesn't exist, so start a new connection"
    [ref] $null = Connect-vNetScalerSession -ComputerName $nsip -UseNSSSL:$UseNSSSL -Credential $Credential -ErrorAction Stop
}
### If we are running in offline/export mode, export the NSSession Variable so we can use this on import
### Export $Script:nsSession
If ($Offline) {
    $OfflineExportPath = Join-Path -Path $OfflinePath -ChildPath "nsSession.xml"
    $script:nsSession | Export-CliXML -Path $OfflineExportPath
}
#endregion Connect
#endregion Functions

#region feature state
##Getting Feature states for usage later on and performance enhancements by not running parts of the script when feature is disabled
$NSFeatures = Get-vNetScalerObject -Type nsfeature -Verbose
If ($NSFEATURES.WL -eq "True") { $FEATWL = "Enabled" } Else { $FEATWL = "Disabled" }
If ($NSFEATURES.SP -eq "True") { $FEATSP = "Enabled" } Else { $FEATSP = "Disabled" }
If ($NSFEATURES.LB -eq "True") { $FEATLB = "Enabled" } Else { $FEATLB = "Disabled" }
If ($NSFEATURES.CS -eq "True") { $FEATCS = "Enabled" } Else { $FEATCS = "Disabled" }
If ($NSFEATURES.CR -eq "True") { $FEATCR = "Enabled" } Else { $FEATCR = "Disabled" }
If ($NSFEATURES.SC -eq "True") { $FEATSC = "Enabled" } Else { $FEATSC = "Disabled" } # removed
If ($NSFEATURES.CMP -eq "True") { $FEATCMP = "Enabled" } Else { $FEATCMP = "Disabled" }
If ($NSFEATURES.PQ -eq "True") { $FEATPQ = "Enabled" } Else { $FEATPQ = "Disabled" } # removed
If ($NSFEATURES.SSL -eq "True") { $FEATSSL = "Enabled" } Else { $FEATSSL = "Disabled" }
If ($NSFEATURES.GSLB -eq "True") { $FEATGSLB = "Enabled" } Else { $FEATGSLB = "Disabled" }
If ($NSFEATURES.HDSOP -eq "True") { $FEATHDOSP = "Enabled" } Else { $FEATHDOSP = "Disabled" } #removed
If ($NSFEATURES.CF -eq "True") { $FEATCF = "Enabled" } Else { $FEATCF = "Disabled" }
If ($NSFEATURES.IC -eq "True") { $FEATIC = "Enabled" } Else { $FEATIC = "Disabled" }
If ($NSFEATURES.SSLVPN -eq "True") { $FEATSSLVPN = "Enabled" } Else { $FEATSSLVPN = "Disabled" }
If ($NSFEATURES.AAA -eq "True") { $FEATAAA = "Enabled" } Else { $FEATAAA = "Disabled" }
If ($NSFEATURES.OSPF -eq "True") { $FEATOSPF = "Enabled" } Else { $FEATOSPF = "Disabled" }
If ($NSFEATURES.RIP -eq "True") { $FEATRIP = "Enabled" } Else { $FEATRIP = "Disabled" }
If ($NSFEATURES.BGP -eq "True") { $FEATBGP = "Enabled" } Else { $FEATBGP = "Disabled" }
If ($NSFEATURES.REWRITE -eq "True") { $FEATREWRITE = "Enabled" } Else { $FEATREWRITE = "Disabled" }
If ($NSFEATURES.IPv6PT -eq "True") { $FEATIPv6PT = "Enabled" } Else { $FEATIPv6PT = "Disabled" }
If ($NSFEATURES.APPFW -eq "True") { $FEATAppFw = "Enabled" } Else { $FEATAppFw = "Disabled" }
If ($NSFEATURES.RESPONDER -eq "True") { $FEATRESPONDER = "Enabled" } Else { $FEATRESPONDER = "Disabled" }
If ($NSFEATURES.HTMLInjection -eq "True") { $FEATHTMLInjection = "Enabled" } Else { $FEATHTMLInjection = "Disabled" } #removed
If ($NSFEATURES.PUSH -eq "True") { $FEATpush = "Enabled" } Else { $FEATpush = "Disabled" }
If ($NSFEATURES.APPFLOW -eq "True") { $FEATAppFlow = "Enabled" } Else { $FEATAppFlow = "Disabled" }
If ($NSFEATURES.CloudBridge -eq "True") { $FEATCloudBridge = "Enabled" } Else { $FEATCloudBridge = "Disabled" }
If ($NSFEATURES.ISIS -eq "True") { $FEATISIS = "Enabled" } Else { $FEATISIS = "Disabled" }
If ($NSFEATURES.CH -eq "True") { $FEATCH = "Enabled" } Else { $FEATCH = "Disabled" }
If ($NSFEATURES.APPQoE -eq "True") { $FEATAppQoE = "Enabled" } Else { $FEATAppQoE = "Disabled" }
If ($NSFEATURES.contentaccelerator -eq "True") { $FEATcontentaccelerator = "Enabled" } Else { $FEATcontentaccelerator = "Disabled" }
If ($NSFEATURES.feo -eq "True") { $FEATfeo = "Enabled" } Else { $FEATfeo = "Disabled" }
If ($NSFEATURES.lsn -eq "True") { $FEATlsn = "Enabled" } Else { $FEATlsn = "Disabled" }
If ($NSFEATURES.rdpproxy -eq "True") { $FEATrdpproxy = "Enabled" } Else { $FEATrdpproxy = "Disabled" }
If ($NSFEATURES.rep -eq "True") { $FEATrep = "Enabled" } Else { $FEATrep = "Disabled" }
If ($NSFEATURES.urlfiltering -eq "True") { $FEATurl = "Enabled" } Else { $FEATurl = "Disabled" }
If ($NSFEATURES.videooptimization -eq "True") { $FEATvideo = "Enabled" } Else { $FEATvideo = "Disabled" }
If ($NSFEATURES.forwardproxy -eq "True") { $FEATfp = "Enabled" } Else { $FEATfp = "Disabled" }
If ($NSFEATURES.sslinterception -eq "True") { $FEATsslint = "Enabled" } Else { $FEATsslint = "Disabled" }
If ($NSFEATURES.adaptivetcp -eq "True") { $FEATadaptivetcp = "Enabled" } Else { $FEATadaptivetcp = "Disabled" }
If ($NSFEATURES.cqa -eq "True") { $FEATcqa = "Enabled" } Else { $FEATcqa = "Disabled" }
If ($NSFEATURES.ci -eq "True") { $FEATci = "Enabled" } Else { $FEATci = "Disabled" }
If ($NSFEATURES.bot -eq "True") { $FEATbot = "Enabled" } Else { $FEATbot = "Disabled" }
If ($NSFEATURES.apigateway -eq "True") { $FEATapigw = "Enabled" } Else { $FEATapigw = "Disabled" }
If ($NSFEATURES.vPath -eq "True") { $FEATVpath = "Enabled" } Else { $FEATVpath = "Disabled" }
#endregion feature state

#region Version
## Get version and build
$NSVersion = ((Get-vNetScalerObject -Type nsversion).version -replace 'NetScaler', '' -replace ',', '' -replace ':', '' -replace 'NS', '').trim().split()
$Version = $NSVersion[0]
$BuildDate = $($NSVersion[5] + " " + $NSVersion[4] + " " + $NSVersion[6])
$Build = $NSVersion[2]
## Set script test version
## WIP THIS WORKS ONLY WHEN REGIONAL SETTINGS DIGIT IS SET TO . :)
$ScriptVersion = 14.1
If ($Version -gt $ScriptVersion) {WriteWordLine 0 0 "Warning: You are using NetScaler version $Version, features added since version $ScriptVersion will not be shown."}
#endregion Version

#region System Information
WriteWordLine 1 0 "System Information"

#region Version and configuration
WriteWordLine 2 0 "Version and configuration"

$nsconfig = Get-vNetScalerObject -Type nsconfig
$nshostname = Get-vNetScalerObject -Type nshostname
$License = Get-vNetScalerObject -Type nslicense
If ($license.isstandardlic) {$LIC = "Standard"}
If ($license.isenterpriselic) {$LIC = "Enterprise"}
If ($license.isplatinumlic) {$LIC = "Platinum"}
If ($license.f_sslvpn_users -eq "4294967295") { $sslvpnlicenses = "Unlimited" } Else { $sslvpnlicenses = $license.f_sslvpn_users }
[System.Collections.Hashtable[]] $SYSTEMH = @(
    If (![string]::IsNullOrWhiteSpace($nshostname.hostname)) {@{ Description = "Hostname"; Value = $nshostname.hostname}}
	@{ Description = "IP"; Value = $nsconfig.ipaddress}
	@{ Description = "NetMask"; Value = $nsconfig.netmask}
    @{ Description = "Version"; Value = $Version}
	@{ Description = "Build"; Value = $Build}
	@{ Description = "Build Date"; Value = $BuildDate}
	@{ Description = "Last Configuration Saved Date"; Value = $nsconfig.lastconfigsavetime}
	@{ Description = "Edition"; Value = $LIC}
	@{ Description = "SSL VPN Licenses"; Value = $sslvpnlicenses}
	@{ Description = "System Type"; Value = $nsconfig.systemtype}
	@{ Description = "Timezone"; Value = $nsconfig.timezone}
)
$Params = $null
$Params = @{
    Hashtable = $SYSTEMH
    Columns   = "Description", "Value"
}
$Table = AddWordTable @Params -List
FindWordDocumentEnd
#endregion Version and configuration

#region Status
WriteWordLine 2 0 "Status"
$nsstatus = Get-vNetScalerObject -Container stat -Object ns
[System.Collections.Hashtable[]] $NSSTATH = @(
    @{ Description = "Last Startup Time"; Value = $nsstatus.starttime }
    @{ Description = "Current HA Status"; Value = $nsstatus.hacurmasterstate }
    @{ Description = "Last HA Status Change"; Value = $nsstatus.transtime }
    @{ Description = "Number of SSL Cards"; Value = $nsstatus.sslcards }
    @{ Description = "Number of CPUs"; Value = $nsstatus.numcpus }
    @{ Description = "/Flash Space Used (Percentage)"; Value = $nsstatus.disk0perusage }
    @{ Description = "/Flash Available Space"; Value = $nsstatus.disk0avail }
    @{ Description = "/Var Space Used (Percentage)"; Value = $nsstatus.disk1perusage }
    @{ Description = "/Var Available Space"; Value = $nsstatus.disk1avail }
)
$Params = $null
$Params = @{
    Hashtable = $NSSTATH
    Columns   = "Description", "Value"
}
$Table = AddWordTable @Params
FindWordDocumentEnd
#endregion Status

#region Hardware
WriteWordLine 2 0 "Hardware"
$nshardware = Get-vNetScalerObject -Type nshardware
$nsmgmtcpu = Get-vNetScalerObject -Type systemextramgmtcpu
$nscpucfg = Get-vNetScalerObject -Type nsvpxparam
[System.Collections.Hashtable[]] $NSHARDWARETable = @(
    @{ Description = "Hardware Description"; Value = $nshardware.hwdescription }
    @{ Description = "Model"; Value = $license.modelid }
    @{ Description = "Hardware System ID"; Value = $nshardware.sysid }
    @{ Description = "Host ID"; Value = $nshardware.hostid }
    @{ Description = "Host (MAC Address)"; Value = $nshardware.host }
    @{ Description = "Extra Management CPU Status"; Value = $nsmgmtcpu.effectivestate }
    @{ Description = "Serial Number"; Value = $nshardware.serialno }
    If (![string]::IsNullOrWhiteSpace($nscpucfg.cpuyield)){@{ Description = "Yield CPU Time (VPX Only)"; Value = $nscpucfg.cpuyield }}
)
$Params = $null
$Params = @{
    Hashtable = $NSHARDWARETable
    Columns   = "Description", "Value"
}
$Table = AddWordTable @Params
FindWordDocumentEnd
#endregion Hardware

#region Capacity
WriteWordLine 2 0 "Capacity"
$nscapacity = Get-vNetScalerObject -Type nscapacity
[System.Collections.Hashtable[]] $NSCAPACITYH = @(
    If (![string]::IsNullOrWhiteSpace($nscapacity.bandwidth)){@{ Description = "System Bandwidth Limit"; Value = $nscapacity.bandwidth }}
    If (![string]::IsNullOrWhiteSpace($nscapacity.unit)){@{ Description = "Bandwidth Limit Unit"; Value = $nscapacity.unit }}
    If (![string]::IsNullOrWhiteSpace($nscapacity.vcpu)){@{ Description = "System Using vCPU Licensing"; Value = $nscapacity.vcpu }}
    If (![string]::IsNullOrWhiteSpace($nscapacity.edition)){@{ Description = "Product Edition"; Value = $nscapacity.edition }}
    If (![string]::IsNullOrWhiteSpace($nscapacity.actualbandwidth)){@{ Description = "Actual Bandwidth (Mbps)"; Value = $nscapacity.actualbandwidth }}
    If (![string]::IsNullOrWhiteSpace($nscapacity.vcpucount)){@{ Description = "vCPU Count"; Value = $nscapacity.vcpucount }}
    If (![string]::IsNullOrWhiteSpace($nscapacity.maxvcpucount)){@{ Description = "Maximum vCPU Count"; Value = $nscapacity.maxvcpucount }}
    If (![string]::IsNullOrWhiteSpace($nscapacity.maxbandwidth)){@{ Description = "Maximum Bandwidth"; Value = $nscapacity.maxbandwidth }}
)
$Params = $null
$Params = @{
    Hashtable = $NSCAPACITYH
    Columns   = "Description", "Value"
}
$Table = AddWordTable @Params
FindWordDocumentEnd
#endregion Capacity
#endregion System Information

#region System
#region System Settings
$selection.InsertNewPage()
WriteWordLine 1 0 "System"
WriteWordLine 2 0 "Settings"

#region Modes
WriteWordLine 3 0 "Modes"
$nsmode = Get-vNetScalerObject -Type nsmode 
[System.Collections.Hashtable[]] $ADVModes = @(
    @{ Description = "Fast Ramp"; Value = $nsmode.fr }
    @{ Description = "Layer 2 mode"; Value = $nsmode.l2 }
    @{ Description = "Use Source IP"; Value = $nsmode.usip }
    @{ Description = "Client SideKeep-alive"; Value = $nsmode.cka }
    @{ Description = "TCP Buffering"; Value = $nsmode.TCPB }
    @{ Description = "MAC-based forwarding"; Value = $nsmode.MBF }
    @{ Description = "Edge configuration"; Value = $nsmode.Edge }
    @{ Description = "Use Subnet IP"; Value = $nsmode.USNIP }
    @{ Description = "Use Layer 3 Mode"; Value = $nsmode.L3 }
    @{ Description = "Path MTU Discovery"; Value = $nsmode.PMTUD }
    @{ Description = "Media Classification"; Value = $nsmode.mediaclassification }
    @{ Description = "Static Route Advertisement"; Value = $nsmode.SRADV }
    @{ Description = "Direct Route Advertisement"; Value = $nsmode.DRADV }
    @{ Description = "Intranet Route Advertisement"; Value = $nsmode.IRADV }
    @{ Description = "Ipv6 Static Route Advertisement"; Value = $nsmode.SRADV6 }
    @{ Description = "Ipv6 Direct Route Advertisement"; Value = $nsmode.DRADV6 }
    @{ Description = "Bridge BPDUs" ; Value = $nsmode.BridgeBPDUs }
	@{ Description = "Unified Logging Framework Mode for adding/removing ULF services." ; Value = $nsmode.ULFD }
    # (removed)@{ Description = "Rise APBR"; Value = $nsmode.rise_apbr }
    # (removed)@{ Description = "Rise RHI" ; Value = $nsmode.rise_rhi }
)
$Params = $null
$Params = @{
    Hashtable = $ADVModes
    Columns   = "Description", "Value"
    Headers   = "Mode", "Enabled"
}
$Table = AddWordTable @Params
FindWordDocumentEnd
#endregion Modes

#region Features
#region Basic Features
WriteWordLine 3 0 "Basic Features"
[System.Collections.Hashtable[]] $AdvancedConfiguration = @(
    @{ Description = "SSL Offloading"; Value = $FEATSSL }
    @{ Description = "Load Balancing"; Value = $FEATLB }
    @{ Description = "Integrated Caching"; Value = $FEATIC }
    @{ Description = "NetScaler Gateway"; Value = $FEATSSLVPN }
    @{ Description = "HTTP Compression"; Value = $FEATCMP }
    @{ Description = "Content Switching"; Value = $FEATCS }
    @{ Description = "Rewrite"; Value = $FEATRewrite }
    @{ Description = "Authentication, Authorization and Auditing"; Value = $FEATAAA }
)
$Params = $null
$Params = @{
    Hashtable	= $AdvancedConfiguration
    Columns		= "Description", "Value"
	Headers		= "Feature", "State"
}
$Table = AddWordTable @Params
FindWordDocumentEnd
#endregion Basic Features

#region Advanced Features
WriteWordLine 3 0 "Advanced Features"
[System.Collections.Hashtable[]] $AdvancedFeatures = @(
	@{ Description = "Surge Protection"; Value = $FEATSP }
	@{ Description = "Priority Queuing"; Value = $FEATPQ }
	@{ Description = "Cache Redirection"; Value = $FEATCR }
	@{ Description = "Web Logging"; Value = $FEATWL }
	@{ Description = "RIP Routing"; Value = $FEATRIP }
	@{ Description = "IPv6 protocol translation "; Value = $FEATIPv6PT }
	@{ Description = "Edgesight Monitoring HTML Injection"; Value = $FEATHTMLInjection }
	@{ Description = "AppFlow"; Value = $FEATAppFlow }
	@{ Description = "ISIS Routing"; Value = $FEATISIS }
	@{ Description = "AppQoE"; Value = $FEATAppQoE }
	@{ Description = "Video Optimization"; Value = $FEATvideo }
	@{ Description = "vPath"; Value = $FEATVpath }
	@{ Description = "Reputation"; Value = $FEATrep }
	@{ Description = "Forward Proxy"; Value = $FEATfp }
	@{ Description = "Adaptive TCP"; Value = $FEATadaptivetcp }
	@{ Description = "Content Inspection"; Value = $FEATci }
	@{ Description = "Citrix Bot Management"; Value = $FEATbot }
	# removed @{ Description = "Sure Connect"; Value = $FEATSC }
	# removed @{ Description = "Http DoS Protection"; Value = $FEATHDOSP }
	@{ Description = "Global Server Load Balancing"; Value = $FEATGSLB }
	@{ Description = "OSPF Routing"; Value = $FEATOSPF }
	@{ Description = "BGP Routing"; Value = $FEATBGP }
	@{ Description = "Responder"; Value = $FEATRESPONDER }
	@{ Description = "NetScaler ADC Push"; Value = $FEATPUSH }
	@{ Description = "CloudBridge"; Value = $FEATCloudBridge }
	@{ Description = "CallHome"; Value = $FEATCH }
	@{ Description = "Front End Optimization"; Value = $FEATfeo }
	@{ Description = "Large Scale NAT"; Value = $FEATlsn }
	@{ Description = "RDP Proxy"; Value = $FEATrdpproxy }
	@{ Description = "URL Filtering"; Value = $FEATurl }
	@{ Description = "SSL Interception"; Value = $FEATsslint }
	@{ Description = "Connection Quality Analytics"; Value = $FEATcqa }
	@{ Description = "Citrix Web App Firewall"; Value = $FEATAppFw }
	# removed @{ Description = "Content Filter"; Value = $FEATCF }
	# removed @{ Description = "Integrated Caching"; Value = $FEATIC }
)
$Params = $null
$Params = @{
	Hashtable	= $AdvancedFeatures
	Columns		= "Description", "Value"
	Headers		= "Feature", "State"
}
$Table = AddWordTable @Params
FindWordDocumentEnd      
#endregion Advanced Features
#endregion Features

#region Global System Settings
$nsspparams = Get-vNetScalerObject -Type nsspparams
$nsparam = Get-vNetScalerObject -Type nsparam
$nsratecontrol = Get-vNetScalerObject -Type nsratecontrol
$systemparameter = Get-vNetScalerObject -Type systemparameter
$nsconsoleloginprompt = Get-vNetScalerObject -Type nsconsoleloginprompt
$nsweblogparam = Get-vNetScalerObject -Type nsweblogparam
$rnatparam = Get-vNetScalerObject -Type rnatparam
[System.Collections.Hashtable[]] $NSPARAMH = @(
	If ($nsspparams.basethreshold -ne "200"){@{ Description = "Surge Protection Base Threshold [200]"; Value = $nsspparams.basethreshold }}
	If ($nsspparams.throttle -ne "Normal"){@{ Description = "Surge Protection Throttle [normal]"; Value = $nsspparams.throttle }}
	If ($nsparam.pmtumin -ne "576"){@{ Description = "Path MTU Discovery Minimum Path MTU (bytes) [576]"; Value = $nsparam.pmtumin }}
	If ($nsparam.pmtutimeout -ne "10"){@{ Description = "Path MTU Discovery Path MTU entry Time Out (mins) [10]"; Value = $nsparam.pmtutimeout }}
	If ($nsratecontrol.udpthreshold -ne "0"){@{ Description = "Rate Control (per 10ms) UDP Threshold [0]"; Value = $nsratecontrol.udpthreshold }}
	If ($nsratecontrol.tcpthreshold -ne "0"){@{ Description = "Rate Control (per 10ms) TCP Threshold [0]"; Value = $nsratecontrol.tcpthreshold }}
	If ($nsratecontrol.tcprstthreshold -ne "100"){@{ Description = "Rate Control (per 10ms) TCP Reset Threshold [100]"; Value = $nsratecontrol.tcprstthreshold }}
	If ($nsratecontrol.icmpthreshold -ne "100"){@{ Description = "Rate Control (per 10ms) ICMP Threshold [100]"; Value = $nsratecontrol.icmpthreshold }}
	If ($systemparameter.natpcbforceflushlimit -ne "2147483647"){@{ Description = "NATPCB Force flush NATPCB's above [2147483647]"; Value = $systemparameter.natpcbforceflushlimit }}
	If ($systemparameter.natpcbrstontimeout -ne "DISABLED"){@{ Description = "NATPCB Send RST for NATPCB timeout [disabled]"; Value = $systemparameter.natpcbrstontimeout }}
	If ($nsparam.grantquotaspillover -ne "10"){@{ Description = "Spill Over Grant Quota (%) [10]"; Value = $nsparam.grantquotaspillover }}
	If ($nsparam.exclusivequotaspillover -ne "80"){@{ Description = "Spill Over Exclusive Quota (%) [80]"; Value = $nsparam.exclusivequotaspillover }}
	If ($nsparam.grantquotamaxclient -ne "10"){@{ Description = "Max Client Grant Quota (%) [10]"; Value = $nsparam.grantquotamaxclient }}
	If ($nsparam.exclusivequotamaxclient -ne "80"){@{ Description = "Max Client Exclusive Quota (%) [80]"; Value = $nsparam.exclusivequotamaxclient }}
	If (![string]::IsNullOrWhiteSpace($nsparam.ftpportrange)){@{ Description = "FTP Port range"; Value = $nsparam.ftpportrange }}
	If ($nsparam.aftpallowrandomsourceport -ne "DISABLED"){@{ Description = "Enable Random source port selection for Active FTP [disabled]"; Value = $nsparam.aftpallowrandomsourceport }}
	If (![string]::IsNullOrWhiteSpace($nsparam.crportrange)){@{ Description = "Cache Redirection Port Range"; Value = $nsparam.crportrange }}
	If (![string]::IsNullOrWhiteSpace($systemparameter.promptstring)){@{ Description = "Command Line Interface (CLI) Prompt"; Value = $systemparameter.promptstring }}
	If ($systemparameter.restrictedtimeout -ne "DISABLED"){@{ Description = "Command Line Interface (CLI) Restricted Timeout [disabled]"; Value = $systemparameter.restrictedtimeout }}
	If ($systemparameter.rbaonresponse -ne "ENABLED"){@{ Description = "Command Line Interface (CLI) RBA on response [enabled]"; Value = $systemparameter.rbaonresponse }}
	If ($nsconsoleloginprompt.promptstring -notmatch "default:\"){@{ Description = "Command Line Interface (CLI) Login Prompt [default:\]"; Value = $nsconsoleloginprompt.promptstring }}
	If ($systemparameter.cliloglevel -ne "INFORMATIONAL"){@{ Description = "Command Line Interface (CLI) Log Levels [informational]"; Value = $systemparameter.cliloglevel }}
	If ($systemparameter.localauth -ne "ENABLED"){@{ Description = "Command Line Interface (CLI) Local Authentication [enabled]"; Value = $systemparameter.localauth }}
	If ($systemparameter.strongpassword -ne "DISABLED"){@{ Description = "Strong Password [disabled]"; Value = $systemparameter.strongpassword }}
	If ($systemparameter.minpasswordlen -ne "1"){@{ Description = "Min Password Length [1]"; Value = $systemparameter.minpasswordlen }}
	If ($systemparameter.forcepasswordchange -ne "DISABLED"){@{ Description = "Force Password Change (nsroot) [disabled]"; Value = $systemparameter.forcepasswordchange }}
	If ($systemparameter.basicauth -ne "ENABLED"){@{ Description = "Basic Auth [enabled]"; Value = $systemparameter.basicauth }}
	If ($nsweblogparam.buffersizemb -ne "16"){@{ Description = "Web Logging Buffer Size (in MBytes) [16]"; Value = $nsweblogparam.buffersizemb }}
	If (![string]::IsNullOrWhiteSpace($nsweblogparam.customreqhdrs)){@{ Description = "Web Logging Custom HTTP Request Header"; Value = $nsweblogparam.customreqhdrs }}
	If (![string]::IsNullOrWhiteSpace($nsweblogparam.customrsphdrs)){@{ Description = "Web Logging Custom HTTP Response Header"; Value = $nsweblogparam.customrsphdrs }}
	If ($systemparameter.timeout -ne "900"){@{ Description = "Idle Session Timeout (secs) [900]"; Value = $systemparameter.timeout }}
	If ($nsparam.secureicaports -ne "443"){@{ Description = "Secure ICA port(s) [443]"; Value = $nsparam.secureicaports }}
	If (![string]::IsNullOrWhiteSpace($nsparam.icaports)){@{ Description = "ICA port(s)"; Value = $nsparam.icaports }}
	If ($nsparam.mgmthttpport -ne "80"){@{ Description = "Management HTTP Port [80]"; Value = $nsparam.mgmthttpport }}
	If ($nsparam.mgmthttpsport -ne "443"){@{ Description = "Management HTTPS Port [443]"; Value = $nsparam.mgmthttpsport }}
	If ($nsparam.useproxyport -ne "ENABLED"){@{ Description = "Use Proxy Port [enabled]"; Value = $nsparam.useproxyport }}
	If ($nsparam.proxyprotocol -ne "DISABLED"){@{ Description = "Proxy Protocol [disabled]"; Value = $nsparam.proxyprotocol }}
	If ($rnatparam.tcpproxy -ne "ENABLED"){@{ Description = "Enable RNAT TCP Proxy [enabled]"; Value = $rnatparam.tcpproxy }}
	If ($nsparam.advancedanalyticsstats -ne "DISABLED"){@{ Description = "Advanced Analytics Stats [disabled]"; Value = $nsparam.advancedanalyticsstats }}
	If ($rnatparam.srcippersistency -ne "DISABLED"){@{ Description = "Enable RNAT Source IP Persistency [disabled]"; Value = $rnatparam.srcippersistency }}
	If ($nsparam.internaluserlogin -ne "ENABLED"){@{ Description = "Use in-built system user to communicate with other appliances [enabled]"; Value = $nsparam.internaluserlogin }}
	If ($nsparam.tcpcip -ne "DISABLED"){@{ Description = "Client TCP/IP header insertion in TCP payload [disabled]"; Value = $nsparam.tcpcip }}
	If ($systemparameter.fipsusermode -ne "DISABLED"){@{ Description = "Enable FIPS User Mode [disabled]"; Value = $systemparameter.fipsusermode }}
	If ($systemparameter.allowdefaultpartition -ne "NO"){@{ Description = "Allow Default Partition [no]"; Value = $systemparameter.allowdefaultpartition }}
	If ($systemparameter.reauthonauthparamchange -ne "DISABLED"){@{ Description = "Reauthentication On Authentication Parameter Change [disabled]"; Value = $systemparameter.reauthonauthparamchange }}
	If ($systemparameter.removesensitivefiles -ne "DISABLED"){@{ Description = "Remove Sensitive Files [disabled]"; Value = $systemparameter.removesensitivefiles }}
	If ($nsparam.ipttl -ne "255"){@{ Description = "IP Time to Live [255]"; Value = $nsparam.ipttl }}
)
If ($NSPARAMH.Length -gt 0) {
	WriteWordLine 3 0 "Global System Settings"
	WriteWordLine 0 0 "Only non-default values are reported, defaults are in [brackets]"
	$Params = $null
	$Params = @{
		Hashtable = $NSPARAMH
		Columns   = "Description", "Value"
		Headers   = "Description", "Value"
	}
	$Table = AddWordTable @Params
	FindWordDocumentEnd
}
#endregion Global System Settings

#region Global HTTP Parameters
$nshttpparam = Get-vNetScalerObject -Type nshttpparam
[System.Collections.Hashtable[]] $NSHTTPPARAMH = @(
	If ($nsconfig.maxconn -ne "0"){@{ Description = "Max Connections [0]"; Value = $nsconfig.maxconn }}
	If ($nsconfig.maxreq -ne "0"){@{ Description = "Max Request [0]"; Value = $nsconfig.maxreq }}
	If ($nsconfig.cip -ne "DISABLED"){@{ Description = "Client IP Insertion [disabled]"; Value = $nsconfig.cip }}
	If (![string]::IsNullOrWhiteSpace($nsconfig.cipheader)){@{ Description = "Client IP Header"; Value = $nsconfig.cipheader }}
	If ($nsconfig.cookieversion -ne "0"){@{ Description = "Cookie Version [0]"; Value = $nsconfig.cookieversion }}
	If ($nsconfig.securecookie -ne "ENABLED"){@{ Description = "Enable Persistence Secure Cookie [enabled]"; Value = $nsconfig.securecookie }}
	If ($nshttpparam.dropinvalreqs -ne "OFF"){@{ Description = "HTTP Drop Invalid Request [off]"; Value = $nshttpparam.dropinvalreqs }}
	If ($nshttpparam.markhttp09inval -ne "OFF"){@{ Description = "Mark HTTP/0.9 requests as invalid [off]"; Value = $nshttpparam.markhttp09inval }}
	If ($nshttpparam.markconnreqinval -ne "OFF"){@{ Description = "Mark CONNECT requests as invalid [off]"; Value = $nshttpparam.markconnreqinval }}
	If ($nshttpparam.logerrresp -ne "ON"){@{ Description = "Log HTTP error responses [on]"; Value = $nshttpparam.logerrresp }}
	If ($nshttpparam.http2serverside -ne "OFF"){@{ Description = "HTTP/2 on Server Side [off]"; Value = $nshttpparam.http2serverside }}
	If ($nshttpparam.conmultiplex -ne "ENABLED"){@{ Description = "Connection Multiplexing [enabled]"; Value = $nshttpparam.conmultiplex }}
	If ($nshttpparam.maxreusepool -ne "0"){@{ Description = "Max connections in Reuse Pool [0]"; Value = $nshttpparam.maxreusepool }}
	If ($nshttpparam.ignoreconnectcodingscheme -ne "DISABLED"){@{ Description = "Ignore Coding Scheme in CONNECT Request [disabled]"; Value = $nshttpparam.ignoreconnectcodingscheme }}
	If ($nshttpparam.insnssrvrhdr -ne "OFF"){
		@{ Description = "Server Header Insertion [off]"; Value = $nshttpparam.insnssrvrhdr }
		@{ Description = "Server Header"; Value = $nshttpparam.nssrvrhdr }
	}
)
If ($NSHTTPPARAMH.Length -gt 0) {
	WriteWordLine 3 0 "Global HTTP Parameters"
	WriteWordLine 0 0 "Only non-default values are reported, defaults are in [brackets]"
	$Params = $null
	$Params = @{
		Hashtable = $NSHTTPPARAMH
		Columns   = "Description", "Value"
		Headers   = "Description", "Value"
	}
	$Table = AddWordTable @Params
	FindWordDocumentEnd
}
#endregion Global HTTP Parameters

#region Global TCP Parameters
WriteWordLine 3 0 "Global TCP Parameters"
$nstcpbufparam = Get-vNetScalerObject -Type nstcpbufparam
$nstcpparam = Get-vNetScalerObject -Type nstcpparam
[System.Collections.Hashtable[]] $NSTCPH = @(
	If ($nstcpbufparam.size -ne "64"){@{ Description = "TCP Buffer size (KBytes) [64]"; Value = $nstcpbufparam.size}}
	If ($nstcpbufparam.memlimit -ne "64"){@{ Description = "TCP Memory usage limit (MBytes) [64]"; Value = $nstcpbufparam.memlimit}}
	If ($nstcpparam.ws -ne "ENABLED"){
		@{ Description = "Window Scaling [enabled]"; Value = $nstcpparam.ws}
		If ($nstcpparam.wsval -ne "8"){@{ Description = "Window scaling factor [8]"; Value = $nstcpparam.wsval}}
	}
	If ($nstcpparam.slowstartincr -ne "2"){@{ Description = "Slow start increment [2]"; Value = $nstcpparam.slowstartincr}}
	If ($nstcpparam.maxdynserverprobes -ne "7"){@{ Description = "Maximum Server probes in 10ms [7]"; Value = $nstcpparam.maxdynserverprobes}}
	If ($nstcpparam.synholdfastgiveup -ne "1024"){@{ Description = "Max server probes give up threshold [1024]"; Value = $nstcpparam.synholdfastgiveup}}
	If ($nstcpparam.maxsynholdperprobe -ne "128"){@{ Description = "Maximum SYN queued per PCB [128]"; Value = $nstcpparam.maxsynholdperprobe}}
	If ($nstcpparam.maxsynhold -ne "16384"){@{ Description = "Maximum SYN held [16384]"; Value = $nstcpparam.maxsynhold}}
	If ($nstcpparam.msslearninterval -ne "180"){@{ Description = "Virtual Server MSS learning interval (sec) [180]"; Value = $nstcpparam.msslearninterval}}
	If ($nstcpparam.msslearndelay -ne "3600"){@{ Description = "Virtual Server MSS learning delay (sec) [3600]"; Value = $nstcpparam.msslearndelay}}
	If ($nstcpparam.maxtimewaitconn -ne "7000"){@{ Description = "Max Connection limit FIN TIME WAIT [7000]"; Value = $nstcpparam.maxtimewaitconn}}
	If ($nstcpparam.connflushifnomem -ne "NONE "){@{ Description = "Connection flush on memory failure [NONE]"; Value = $nstcpparam.connflushifnomem}}
	If ($nstcpparam.connflushthres -ne "4294967295"){@{ Description = "Connection Flush Threshold [4294967295]"; Value = $nstcpparam.connflushthres}}
	If ($nstcpparam.tcpmaxretries -ne "7"){@{ Description = "Max no of Retransmission Timeouts [7]"; Value = $nstcpparam.tcpmaxretries}}
	If ($nstcpparam.maxburst -ne "6"){@{ Description = "Maximum Burst Limit [6]"; Value = $nstcpparam.maxburst}}
	If ($nstcpparam.initialcwnd -ne "10"){@{ Description = "Initial Congestion Window Size [10]"; Value = $nstcpparam.initialcwnd}}
	If ($nstcpparam.delayedack -ne "100"){@{ Description = "TCP Delayed ACK Time-out (msec) [100]"; Value = $nstcpparam.delayedack}}
	If ($nstcpparam.oooqsize -ne "300"){@{ Description = "Maximum ooo packet queue size [300]"; Value = $nstcpparam.oooqsize}}
	If ($nstcpparam.maxpktpermss -ne "0"){@{ Description = "Maximum Packets per MSS [0]"; Value = $nstcpparam.maxpktpermss}}
	If ($nstcpparam.pktperretx -ne "1"){@{ Description = "Maximum Packets Per Retransmission [1]"; Value = $nstcpparam.pktperretx}}
	If ($nstcpparam.minrto -ne "1000"){@{ Description = "Minimum RTO (in millisec) [1000]"; Value = $nstcpparam.minrto}}
	If ($nstcpparam.maxsynackretx -ne "100"){@{ Description = "Max limit for SYN+ACK retransmissions [100]"; Value = $nstcpparam.maxsynackretx}}
	If ($nstcpparam.tcpfastopencookietimeout -ne "0"){@{ Description = "TCP Fast Open Cookie Timeout [0]"; Value = $nstcpparam.tcpfastopencookietimeout}}
	If ($nstcpparam.autosyncookietimeout -ne "30"){@{ Description = "Auto Syn Cookie Timeout [30]"; Value = $nstcpparam.autosyncookietimeout}}
	If ($nstcpparam.tcpfintimeout -ne "40"){@{ Description = "TCP Finish Timeout [40]"; Value = $nstcpparam.tcpfintimeout}}
	If ($nstcpparam.rfc5961chlgacklimit -ne "0" -and $nstcpparam.rfc5961chlgacklimit -ne $null){@{ Description = "RFC5961 Chlg Ack Limit [0]"; Value = $nstcpparam.rfc5961chlgacklimit}}
	If ($nstcpparam.sack -ne "ENABLED"){@{ Description = "Selective Acknowledgement [enabled]"; Value = $nstcpparam.sack}}
	If ($nstcpparam.ackonpush -ne "ENABLED"){@{ Description = "Immediate ACK on receiving packet with PUSH [enabled]"; Value = $nstcpparam.ackonpush}}
	If ($nstcpparam.downstaterst -ne "DISABLED"){@{ Description = "Down service reset [disabled]"; Value = $nstcpparam.downstaterst}}
	If ($nstcpparam.learnvsvrmss -ne "DISABLED"){@{ Description = "Learn Virtual Server MSS [disabled]"; Value = $nstcpparam.learnvsvrmss7}}
	If ($nstcpparam.nagle -ne "DISABLED"){@{ Description = "Use Nagle's algorithm [disabled]"; Value = $nstcpparam.nagle}}
	If ($nstcpparam.synattackdetection -ne "ENABLED"){@{ Description = "SYN Attack Detection [enabled]"; Value = $nstcpparam.synattackdetection}}
	If ($nstcpparam.limitedpersist -ne "ENABLED"){@{ Description = "Limit Persist Probes [enabled]"; Value = $nstcpparam.limitedpersist}}
	If ($nstcpparam.mptcpsftimeout -ne "0"){@{ Description = "Idle Subflow Timeout (secs) [0]"; Value = $nstcpparam.mptcpsftimeout}}
	If ($nstcpparam.mptcpmaxsf -ne "4"){@{ Description = "Max Established Subflows per Session [4]"; Value = $nstcpparam.mptcpmaxsf}}
	If ($nstcpparam.mptcppendingjointhreshold -ne "0"){@{ Description = "Max Pending Join Subflows [0]"; Value = $nstcpparam.mptcppendingjointhreshold}}
	If ($nstcpparam.mptcpsfreplacetimeout -ne "10"){@{ Description = "Idle Subflow Replacement Timeout (secs) [10]"; Value = $nstcpparam.mptcpsfreplacetimeout}}
	If ($nstcpparam.mptcpmaxsf -ne "4"){@{ Description = "Max Pending Subflows per Session [4]"; Value = $nstcpparam.mptcpmaxsf}}
	If ($nstcpparam.mptcprtostoswitchsf -ne "2"){@{ Description = "Max no of RTO's to Switch Subflow [2]"; Value = $nstcpparam.mptcprtostoswitchsf}}
	If ($nstcpparam.mptcpfastcloseoption -ne "ACK"){@{ Description = "mptcp Fast Close Option [ACK]"; Value = $nstcpparam.mptcpfastcloseoption}}
	If ($nstcpparam.mptcpusebackupondss -ne "ENABLED"){@{ Description = "Use Backup Subflow on DSS [enabled])"; Value = $nstcpparam.buffersize}}
	If ($nstcpparam.mptcpchecksum -ne "ENABLED"){@{ Description = "DSS checksum [ENABLED]"; Value = $nstcpparam.mptcpchecksum}}
	If ($nstcpparam.compacttcpoptionnoop -ne "DISABLED"){@{ Description = "Compact TCP Option Noop [disabled]"; Value = $nstcpparam.compacttcpoptionnoop}}
	If ($nstcpparam.mptcpreliableaddaddr -ne "DISABLED"){@{ Description = "mptcp Reliable Add Address [disabled]"; Value = $nstcpparam.mptcpreliableaddaddr}}
	If ($nstcpparam.mptcpconcloseonpassivesf -ne "ENABLED"){@{ Description = "DATA_FIN/FAST_CLOSE on passive subflow [enabled]"; Value = $nstcpparam.mptcpconcloseonpassivesf}}
	If ($nstcpparam.mptcpimmediatesfcloseonfin -ne "DISABLED"){@{ Description = "Close subflows immediately on FIN [DISABLED]"; Value = $nstcpparam.mptcpimmediatesfcloseonfin}}
	If ($nstcpparam.mptcpsendsfresetoption -ne "DISABLED"){@{ Description = "mptcp Send SF Reset Option [DISABLED]"; Value = $nstcpparam.mptcpsendsfresetoption}}
	If ($nstcpparam.delinkclientserveronrst -ne "DISABLED"){@{ Description = "Delink Client Server on RST [DISABLED]"; Value = $nstcpparam.delinkclientserveronrst}}
)
If ($NSTCPH.Length -gt 0) {
	$Params = $null
	$Params = @{
		Hashtable	= $NSTCPH
		Columns		= "Description", "Value"
		Headers		= "Description", "Configuration"
	}
	$Table = AddWordTable @Params
	FindWordDocumentEnd
}
#endregion Global TCP Parameters

#region Global Diameter Parameters
WriteWordLine 3 0 "Global Diameter Parameters"
$nsdiameter = Get-vNetScalerObject -Type nsdiameter 
$Params = $null
$Params = @{
    Hashtable = @{
        HOST  = $nsdiameter.identity
        Realm = $nsdiameter.realm
        Close = $nsdiameter.serverclosepropagation
    }
    Columns   = "HOST", "Realm", "Close"
    Headers   = "Host Identity", "Realm", "Server Close Propagation"
}
$Table = AddWordTable @Params
FindWordDocumentEnd
#endregion Global Diameter Parameters
#endregion System Settings

#region High Availability
If ($nsconfig.systemtype -eq "HA"){
	WriteWordLine 2 0 "High Availability"
	$HANodes = Get-vNetScalerObject -Type hanode
	[System.Collections.Hashtable[]] $HAH = @()
	foreach ($HANODE in $HANodes) {
		#Name attribute will not be returned for secondary appliance
		If ([string]::IsNullOrWhiteSpace($HANODE.name)) {$HANODENAME = ""} Else {$HANODENAME = $HANODE.name}
		$HAH += @{
			HANAME   = $HANODENAME
			HAIP     = $HANODE.ipaddress
			HASTATUS = $HANODE.state
			HASYNC   = $HANODE.hasync        
		}
		$HANODEname = $null
	}
	If ($HAH.Length -gt 0) {
		$Params = $null
		$Params = @{
			Hashtable = $HAH
			Columns   = "HANAME", "HAIP", "HASTATUS", "HASYNC"
			Headers   = "NetScaler Name", "IP Address", "HA Status", "HA Synchronization"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
}
#endregion High Availability

#region NTP
If ((Get-vNetScalerObjectCount -Type ntpserver).__count -ge 1) {
	WriteWordLine 2 0 "NTP"
	$NTPs = Get-vNetScalerObject -Type ntpserver
	[System.Collections.Hashtable[]] $NTPH = @()
	foreach ($NTP in $NTPs) {
		$NTPH += @{
			NAME	= $NTP.servername
			minpoll	= $NTP.minpoll
			maxpoll	= $NTP.maxpoll
			autokey	= $NTP.autokey
			key		= $NTP.key
			preferredntpserver	= $NTP.preferredntpserver
		}
	}
	If ($NTPH.Length -gt 0) {
		$Params = $null
		$Params = @{
			Hashtable = $NTPH
			Columns   = "Name","minpoll","maxpoll","autokey","key","preferredntpserver"
			Headers   = "Name","Min Poll Interval","Max Poll Interval","Auto Key","Key","Set as preferred NTP server"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
}
#endregion NTP

#region Profiles
$selection.InsertNewPage()
WriteWordLine 2 0 "Profiles"

#region TCP Profiles
WriteWordLine 3 0 "TCP Profiles"
WriteWordLine 0 0 "Only non-default values are reported, defaults are in [brackets]"
$tcpprofiles = Get-vNetScalerObject -Type nstcpprofile
foreach ($tcpprofile in $tcpprofiles) {
	WriteWordLine 4 0 "$($tcpprofile.name)"
	$tcpprof = Get-vNetScalerObject -Type nstcpprofile -name $tcpprofile.name
	[System.Collections.Hashtable[]] $TCPPROFILESH = @(
		If ($tcpprof.ws -ne "DISABLED"){
			@{ Description = "Window Scaling [disabled]"; Value = $tcpprof.ws}
			If ($tcpprof.wsval -ne "4"){@{ Description = "Window scaling factor [4]"; Value = $tcpprof.wsval}}
		}
		If ($tcpprof.maxburst -ne "6"){@{ Description = "Maximum Burst Limit [6]"; Value = $tcpprof.maxburst}}
		If ($tcpprof.initialcwnd -ne "4"){@{ Description = "Initial Congestion Window Size [4]"; Value = $tcpprof.initialcwnd}}
		If ($tcpprof.delayedack -ne "100"){@{ Description = "TCP Delayed ACK Time-out (msec) [100]"; Value = $tcpprof.delayedack}}
		If ($tcpprof.oooqsize -ne "64"){@{ Description = "Maximum ooo packet queue size [64]"; Value = $tcpprof.oooqsize}}
		If ($tcpprof.mss -ne "0"){@{ Description = "MSS [0]"; Value = $tcpprof.mss}}
		If ($tcpprof.maxpktpermss -ne $null){@{ Description = "Maximum Packets per MSS"; Value = $tcpprof.maxpktpermss}}
		If ($tcpprof.pktperretx -ne "1"){@{ Description = "Maximum Packets Per Retransmission [1]"; Value = $tcpprof.pktperretx}}
		If ($tcpprof.minrto -ne "1000"){@{ Description = "Minimum RTO (in millisec) [1000]"; Value = $tcpprof.minrto}}
		If ($tcpprof.slowstartincr -ne "2"){@{ Description = "Slow start increment [2]"; Value = $tcpprof.slowstartincr}}
		If ($tcpprof.slowstartthreshold -ne "524288"){@{ Description = "TCP Slow Start Threshold [524288]"; Value = $tcpprof.slowstartthreshold}}
		If ($tcpprof.buffersize -ne "8190"){@{ Description = "TCP Buffer Size (bytes) [8190])"; Value = $tcpprof.buffersize}}
		If ($tcpprof.sendbuffsize -ne "8190"){@{ Description = "TCP Send Buffer Size (bytes) [8190]"; Value = $tcpprof.sendbuffsize}}
		If ($tcpprof.maxcwnd -ne "524288"){@{ Description = "TCP Maximum Congestion Window Size [524288]"; Value = $tcpprof.maxcwnd}}
		If ($tcpprof.dupackthresh -ne "3"){@{ Description = "TCP Dupack Threshold [3]"; Value = $tcpprof.dupackthresh}}
		If ($tcpprof.burstratecontrol -ne "DISABLED"){@{ Description = "TCP Burst Rate Control [disabled]"; Value = $tcpprof.burstratecontrol}}
		If ($tcpprof.tcprate -ne "0"){@{ Description = "TCP Connection Payload Send Rate (Kb/s) [0]"; Value = $tcpprof.tcprate}}
		If ($tcpprof.rateqmax -ne "0"){@{ Description = "Maximum Connection Queue Size (bytes) [0]"; Value = $tcpprof.rateqmax}}
		If ($tcpprof.flavor -ne "Default"){@{ Description = "TCP Flavor [default]"; Value = $tcpprof.flavor}}
		If ($tcpprof.establishclientconn -ne "AUTOMATIC"){@{ Description = "Establish Client Connection [automatic]"; Value = $tcpprof.establishclientconn}}
		If ($tcpprof.tcpsegoffload -ne "AUTOMATIC"){@{ Description = "TCP Segmentation Offload [automatic]"; Value = $tcpprof.tcpsegoffload}}
		If ($tcpprof.tcpmode -ne "TRANSPARENT"){@{ Description = "TCP Optimization Mode [transparent]"; Value = $tcpprof.tcpmode}}
		If ($tcpprof.clientiptcpoption -ne "DISABLED"){
			@{ Description = "Client IP TCP Option [disabled]"; Value = $tcpprof.clientiptcpoption}
			@{ Description = "Client IP TCP Option Number"; Value = $tcpprof.clientiptcpoptionnumber}
		}
		If ($tcpprof.ka -ne "DISABLED"){@{ Description = "Keep-alive probes [disabled]"; Value = $tcpprof.ka}}
		If ($tcpprof.kaconnidletime -ne "900"){@{ Description = "Connection idle time before sending probe (secs) [900]"; Value = $tcpprof.kaconnidletime}}
		If ($tcpprof.kaprobeinterval -ne "75"){@{ Description = "Keep-alive probe interval (secs) [75]"; Value = $tcpprof.kaprobeinterval}}
		If ($tcpprof.kamaxprobes -ne "3"){@{ Description = "Maximum Keep-alive probes [3]"; Value = $tcpprof.kamaxprobes}}
		If ($tcpprof.kaprobeupdatelastactivity -ne "ENABLED"){@{ Description = "Update last activity for KA probes [enabled]"; Value = $tcpprof.kaprobeupdatelastactivity}}
		If ($tcpprof.mptcp -ne "DISABLED"){@{ Description = "Multipath TCP [disabled]"; Value = $tcpprof.mptcp}}
		If ($tcpprof.mptcpdropdataonpreestsf -ne "DISABLED"){@{ Description = "Drop Data on Pre-Established subflow [disabled]"; Value = $tcpprof.mptcpdropdataonpreestsf}}
		If ($tcpprof.mpcapablecbit -ne "DISABLED"){@{ Description = "Multipath TCP C bit [disanbled]"; Value = $tcpprof.mpcapablecbit}}
		If ($tcpprof.sendclientportintcpoption -ne "DISABLED"){@{ Description = "Send Client Port in TCP Option [disabled]"; Value = $tcpprof.sendclientportintcpoption}}
		If ($tcpprof.tcpfastopencookiesize -ne "8"){@{ Description = "TCP Fast Open CookiesSize [8]"; Value = $tcpprof.tcpfastopencookiesize}}
		If ($tcpprof.mptcpsessiontimeout -ne "0"){@{ Description = "Session Timeout [0]"; Value = $tcpprof.mptcpsessiontimeout}}
		If ($tcpprof.sack -eq "DISABLED"){@{ Description = "Selective Acknowledgement [disabled]"; Value = $tcpprof.sack}}
		If ($tcpprof.ackonpush -ne "ENABLED"){@{ Description = "Immediate ACK on receiving packet with PUSH [enabled]"; Value = $tcpprof.ackonpush}}
		If ($tcpprof.rstwindowattenuate -ne "DISABLED"){@{ Description = "RST Window Attenuation [disabled]"; Value = $tcpprof.rstwindowattenuate}}
		If ($tcpprof.ecn -ne "DISABLED"){@{ Description = "Explicit Congestion Notification (ECN) [disabled]"; Value = $tcpprof.ecn}}
		If ($tcpprof.ackaggregation -ne "DISABLED"){@{ Description = "ACK Aggregation [disabled]"; Value = $tcpprof.ackaggregation}}
		If ($tcpprof.hystart -ne "DISABLED"){@{ Description = "CUBIC Hystart [disabled]"; Value = $tcpprof.hystart}}
		If ($tcpprof.applyadaptivetcp -ne "DISABLED"){@{ Description = "Apply Adaptive TCP [disabled]"; Value = $tcpprof.applyadaptivetcp}}
		If ($tcpprof.fack -ne "DISABLED"){@{ Description = "Forward Acknowledgement [disabled]"; Value = $tcpprof.fack}}
		If ($tcpprof.syncookie -ne "ENABLED"){@{ Description = "TCP SYN Cookie [enabled]"; Value = $tcpprof.syncookie}}
		If ($tcpprof.rstmaxack -ne "DISABLED"){@{ Description = "RST Acceptance [disabled]"; Value = $tcpprof.rstmaxack}}
		If ($tcpprof.timestamp -ne "DISABLED"){@{ Description = "TCP Timestamp [disabled]"; Value = $tcpprof.timestamp}}
		If ($tcpprof.frto -ne "DISABLED"){@{ Description = "Forward RTO-Recovery [disabled]"; Value = $tcpprof.frto}}
		If ($tcpprof.drophalfclosedconnontimeout -ne "DISABLED"){@{ Description = "Drop Half Closed Connections On Idle Timeout [disabled]"; Value = $tcpprof.drophalfclosedconnontimeout}}
		If ($tcpprof.taillossprobe -ne "DISABLED"){@{ Description = "Tail Loss Probe [disabled]"; Value = $tcpprof.taillossprobe}}
		If ($tcpprof.nagle -ne "DISABLED"){@{ Description = "Use Nagle's algorithm [disabled]"; Value = $tcpprof.nagle}}
		If ($tcpprof.dynamicreceivebuffering -ne "DISABLED"){@{ Description = "Dynamic Receive Buffering [disabled]"; Value = $tcpprof.dynamicreceivebuffering}}
		If ($tcpprof.spoofsyndrop -ne "ENABLED"){@{ Description = "SYN Spoof Protection [enabled]"; Value = $tcpprof.spoofsyndrop}}
		If ($tcpprof.dsack -ne "ENABLED"){@{ Description = "Duplicate SACK [enabled]"; Value = $tcpprof.dsack}}
		If ($tcpprof.tcpfastopen -ne "DISABLED"){@{ Description = "TCP Fast Open [disabled]"; Value = $tcpprof.tcpfastopen}}
		If ($tcpprof.dropestconnontimeout -ne "DISABLED"){@{ Description = "Drop TCP Established Connections On Idle Timeout [disabled]"; Value = $tcpprof.dropestconnontimeout}}
		If ($tcpprof.rfc5961compliance -eq "ENABLED"){@{ Description = "RFC5961 Compliance [disabled]"; Value = $tcpprof.rfc5961compliance}}
    )
	If ($TCPPROFILESH.Length -gt 0) {
		$Params = $null
		$Params = @{
			Hashtable	= $TCPPROFILESH
			Columns		= "Description", "Value"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
}
#endregion TCP Profiles

#region HTTP Profiles
$selection.InsertNewPage()
WriteWordLine 3 0 "HTTP Profiles"
WriteWordLine 0 0 "Only non-default values are reported, defaults are in [brackets]"
$httpprofiles = Get-vNetScalerObject -Type nshttpprofile
foreach ($httpprofile in $httpprofiles) {
	WriteWordLine 4 0 "$($httpprofile.name)"
	[System.Collections.Hashtable[]] $HTTPPROFILESH = @(
		If ($httpprofile.minreusepool -ne "0") {@{ Description = "Min connections in reuse pool [0]"; Value = $httpprofile.minreusepool}}
		If ($httpprofile.maxreusepool -ne "0") {@{ Description = "Max connections in reuse pool [0]"; Value = $httpprofile.maxreusepool}}
		If ($httpprofile.reusepooltimeout -ne "0") {@{ Description = "Reuse Pool Timeout [0]"; Value = $httpprofile.reusepooltimeout}}
		If ($httpprofile.incomphdrdelay -ne "7000") {@{ Description = "Incomplete header delay [7000]"; Value = $httpprofile.incomphdrdelay}}
		If ($httpprofile.reqtimeout -ne "0") {@{ Description = "Request time out [0]"; Value = $httpprofile.reqtimeout}}
		If ($httpprofile.maxreq -ne "0") {@{ Description = "Max requests per connection [0]"; Value = $httpprofile.maxreq}}
		If ($httpprofile.reqtimeoutaction -ne $null) {@{ Description = "Request timeout action"; Value = $httpprofile.reqtimeoutaction}}
		If ($httpprofile.maxheaderlen -ne "24820") {@{ Description = "Maximum Header Length [24820]"; Value = $httpprofile.maxheaderlen}}
		If ($httpprofile.maxheaderfieldlen -ne "24820") {@{ Description = "Maximum Header Field Length [24820]"; Value = $httpprofile.maxheaderfieldlen}}
		If ($httpprofile.clientiphdrexpr -ne $null) {@{ Description = "Client IP Header Expression"; Value = $httpprofile.clientiphdrexpr}}
		If ($httpprofile.grpcholdlimit -ne "131072") {@{ Description = "gRPC Hold Limit [131072]"; Value = $httpprofile.grpcholdlimit}}
		If ($httpprofile.grpcholdtimeout -ne "1000") {@{ Description = "gRPC Hold Timeout [1000]"; Value = $httpprofile.grpcholdtimeout}}
		If ($httpprofile.apdexcltresptimethreshold -ne "500") {@{ Description = "APDEX Client Response Time Threshold [500]"; Value = $httpprofile.apdexcltresptimethreshold}}
		If ($httpprofile.httppipelinebuffsize -ne "131072") {@{ Description = "HTTP Pipe Line Buffer Size [131072]"; Value = $httpprofile.httppipelinebuffsize}}
		If ($httpprofile.http2 -eq "ENABLED"){
			@{ Description = "HTTP/2 [disabled]"; Value = $httpprofile.http2}
			If ($httpprofile.http2direct -ne "DISABLED") {@{ Description = "Direct HTTP/2 [disabled]"; Value = $httpprofile.http2direct}}
			If ($httpprofile.http2altsvcframe -ne "DISABLED") {@{ Description = "Send HTTP/2 ALTSVC frame [disabled]"; Value = $httpprofile.http2altsvcframe}}
			If ($httpprofile.http2headertablesize -ne "4096") {@{ Description = "HTTP/2 Header Table Size [4096]"; Value = $httpprofile.http2headertablesize}}
			If ($httpprofile.http2initialconnwindowsize -ne "65535") {@{ Description = "HTTP/2 Initial Connection Window Size [65535]"; Value = $httpprofile.http2initialconnwindowsize}}
			If ($httpprofile.http2initialwindowsize -ne "65535") {@{ Description = "HTTP/2 Initial Window Size [65535]"; Value = $httpprofile.http2initialwindowsize}}
			If ($httpprofile.http2maxconcurrentstreams -ne "100") {@{ Description = "HTTP/2 Maximum Concurrent Streams [100]"; Value = $httpprofile.http2maxconcurrentstreams}}
			If ($httpprofile.http2maxframesize -ne "16384") {@{ Description = "HTTP/2 Maximum Frame Size [16384]"; Value = $httpprofile.http2maxframesize}}
			If ($httpprofile.http2minseverconn -ne "20") {@{ Description = "HTTP/2 Minimum Server Connections [20]"; Value = $httpprofile.http2minseverconn}}
			If ($httpprofile.http2maxheaderlistsize -ne "24576") {@{ Description = "HTTP/2 Maximum Header List Size [24576]"; Value = $httpprofile.http2maxheaderlistsize}}
			If ($httpprofile.http2maxpingframespermin -ne "60") {@{ Description = "HTTP/2 Maximum Ping Frames Per Minute [60]"; Value = $httpprofile.http2maxpingframespermin}}
			If ($httpprofile.http2maxresetframespermin -ne "90") {@{ Description = "HTTP/2 Maximum Reset Frames Per Minute [90]"; Value = $httpprofile.http2maxresetframespermin}}
			If ($httpprofile.http2maxemptyframespermin -ne "60") {@{ Description = "HTTP/2 Maximum Empty Frames Per Minute [60]"; Value = $httpprofile.http2maxemptyframespermin}}
			If ($httpprofile.http2maxsettingsframespermin -ne "15") {@{ Description = "HTTP/2 Maximum Settings Frames Per Minute [15]"; Value = $httpprofile.http2maxsettingsframespermin}}
		}
		If ($httpprofile.http3 -eq "ENABLED"){
			@{ Description = "HTTP/3 [disabled]"; Value = $httpprofile.http3}
			If ($httpprofile.http3maxheaderfieldsectionsize -ne "24576") {@{ Description = "HTTP/3 Maximum Header Field Section Size [24576]"; Value = $httpprofile.http3maxheaderfieldsectionsize}}
			If ($httpprofile.http3maxheadertablesize -ne "4096") {@{ Description = "HTTP/3 Maximum Header Table Size [4096]"; Value = $httpprofile.http3maxheadertablesize}}
			If ($httpprofile.http3maxheaderblockedstreams -ne "100") {@{ Description = "HTTP/3 Maximum Header Blocked Streams [100]"; Value = $httpprofile.http3maxheaderblockedstreams}}
		}
		If ($httpprofile.altsvc -ne "DISABLED") {@{ Description = "Alternative Service [disabled]"; Value = $httpprofile.altsvc}}
		If ($httpprofile.altsvcvalue -ne $null) {@{ Description = "Alternative Service Value"; Value = $httpprofile.altsvcvalue}}
		If ($httpprofile.conmultiplex -ne "ENABLED") {@{ Description = "Connection Multiplexing [enabled]"; Value = $httpprofile.conmultiplex}}
		If ($httpprofile.markconnreqinval -ne "DISABLED") {@{ Description = "Mark CONNECT Requests as Invalid [disabled]"; Value = $httpprofile.markconnreqinval}}
		If ($httpprofile.markhttpheaderextrawserror -ne "DISABLED") {@{ Description = "Mark HTTP Header with Extra White Space as Invalid [disabled]"; Value = $httpprofile.markhttpheaderextrawserror}}
		If ($httpprofile.websocket -ne "DISABLED") {@{ Description = "Enable WebSocket connections [disabled]"; Value = $httpprofile.websocket}}
		If ($httpprofile.weblog -ne "ENABLED") {@{ Description = "HTTP Weblogging [enabled]"; Value = $httpprofile.weblog}}
		If ($httpprofile.grpclengthdelimitation -ne "ENABLED") {@{ Description = "gRPC Length Delimitation [enabled]"; Value = $httpprofile.grpclengthdelimitation}}
		If ($httpprofile.dropinvalreqs -ne "DISABLED") {@{ Description = "Drop invalid HTTP requests [disabled]"; Value = $httpprofile.dropinvalreqs}}
		If ($httpprofile.marktracereqinval -ne "DISABLED") {@{ Description = "Mark TRACE Requests as Invalid [disabled]"; Value = $httpprofile.marktracereqinval}}
		If ($httpprofile.cmponpush -ne "DISABLED") {@{ Description = "Compression on PUSH packet [disabled]"; Value = $httpprofile.cmponpush}}
		If ($httpprofile.rtsptunnel -ne "DISABLED") {@{ Description = "Enable RTSP Tunnel [disabled]"; Value = $httpprofile.rtsptunnel}}
		If ($httpprofile.persistentetag -ne "DISABLED") {@{ Description = "Persistent ETag [disabled]"; Value = $httpprofile.persistentetag}}
		If ($httpprofile.allowonlywordcharactersandhyphen -ne "DISABLED") {@{ Description = "Allow only word characters and hyphen [disabled]"; Value = $httpprofile.allowonlywordcharactersandhyphen}}
		If ($httpprofile.markhttp09inval -ne "DISABLED") {@{ Description = "Mark HTTP/0.9 requests as invalid [disabled]"; Value = $httpprofile.markhttp09inval}}
		If ($httpprofile.markrfc7230noncompliantinval -ne "DISABLED") {@{ Description = "Mark RFC7230 Non-Compliant Transaction as Invalid [disabled]"; Value = $httpprofile.markrfc7230noncompliantinval}}
		If ($httpprofile.dropextracrlf -ne "ENABLED") {@{ Description = "Drop extra CRLF [enabled]"; Value = $httpprofile.dropextracrlf}}
		If ($httpprofile.dropextradata -ne "DISABLED") {@{ Description = "Drop extra data from server [disabled]"; Value = $httpprofile.dropextradata}}
		If ($httpprofile.adpttimeout -ne "DISABLED") {@{ Description = "Adaptive Timeout [disabled]"; Value = $httpprofile.adpttimeout}}
		If ($httpprofile.passprotocolupgrade -ne "ENABLED") {@{ Description = "Pass Protocol Upgrade [enabled]"; Value = $httpprofile.passprotocolupgrade}}
	)
	If ($HTTPPROFILESH.Length -gt 0) {
		$Params = $null
		$Params = @{
			Hashtable	= $HTTPPROFILESH
			Columns		= "Description", "Value"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
}
#endregion HTTP Profiles

#region SSL Profiles
If ((Get-vNetScalerObjectCount -Type sslprofile).__count -ge 1) {
	$selection.InsertNewPage()
	WriteWordLine 3 0 "SSL Profiles"
	WriteWordLine 0 0 "Only non-default values are reported, defaults are in [brackets]"
	$SSLProfiles = Get-vNetScalerObject -Type sslprofile
    Foreach ($SSLProfile in $SSLProfiles) {
        WriteWordLine 4 0 "$($SSLProfile.name)"
		$ProfCiphers = (Get-vNetScalerObject -Type sslprofile_sslcipher_binding -Name $SSLProfile.name).cipheraliasname -Join ", "
		$ProfCert = (Get-vNetScalerObject -Type sslprofile_sslcertkey_binding -Name $SSLProfile.name).sslicacertkey -Join ", "
        [System.Collections.Hashtable[]] $SSLPROFILEH = @(
            @{ Description = "SSL Profile Type"; Value = $SSLprofile.sslprofiletype}
			If ($SSLProfile.pushenctrigger -ne "Always"){@{ Description = "Push Encryption Trigger"; Value = $SSLProfile.pushenctrigger}}
			If ($SSLProfile.encrypttriggerpktcount -ne "45"){@{ Description = "Encryption trigger packet count [45]"; Value = $SSLProfile.encrypttriggerpktcount}}
			If ($SSLProfile.pushflag -ne "0"){@{ Description = "Push Flag [0=auto]"; Value = $SSLProfile.pushflag}}
			If ($SSLProfile.pushenctriggertimeout -ne "1"){@{ Description = "PUSH encryption trigger timeout (ms) [1]"; Value = $SSLProfile.pushenctriggertimeout}}
			If ($SSLProfile.ssltriggertimeout -ne "100"){@{ Description = "Encryption trigger timeout (10 ms ticks) [100]"; Value = $SSLProfile.ssltriggertimeout}}
			If ($SSLProfile.snihttphostmatch -ne "CERT"){@{ Description = "SNI HTTP Host Match [cert]"; Value = $SSLProfile.snihttphostmatch}}
			If ($SSLprofile.sslprofiletype -ne "BackEnd" -and $SSLProfile.insertionencoding -ne "Unicode"){@{ Description = "Encoding type [unicode]"; Value = $SSLProfile.insertionencoding}}
            If ($SSLProfile.denysslreneg -ne "ALL"){@{ Description = "Deny SSL Renegotiation [all]"; Value = $SSLProfile.denysslreneg}}
            IF ($SSLProfile.quantumsize -ne "8192"){@{ Description = "SSL quantum size (KBytes) [8192]"; Value = $SSLProfile.quantumsize}}
            If ($SSLprofile.sslprofiletype -ne "BackEnd"){
				If ($SSLProfile.cleartextport -ne "0"){@{ Description = "Clear Text Port [0]"; Value = $SSLProfile.cleartextport}}
				If ($SSLProfile.alpnprotocol -ne "NONE"){@{ Description = "ALPN Protocol"; Value = $SSLProfile.alpnprotocol}}
			}
			@{ Description = "Enable DH Param [disabled]"; Value = $SSLProfile.dh}
            If ($SSLProfile.dh -ne "DISABLED"){
				@{ Description = "Diffe-Hellman Refresh Count"; Value = $SSLProfile.dhcount}
				@{ Description = "Diffe-Hellman Key File"; Value = $SSLProfile.dhfile}
				@{ Description = "Enable DH Key Expire Size Limit"; Value = $SSLProfile.dhkeyexpsizelimit}
			}
            If ($SSLProfile.ersa -ne "ENABLED"){@{ Description = "Enable Ephemeral RSA [enabled]"; Value = $SSLProfile.ersa}}
			If ($SSLProfile.ersa -eq "ENABLED" -and $SSLProfile.ersacount -ne "0"){@{ Description = "Ephemeral RSA Refresh Count [0]"; Value = $SSLProfile.ersacount}}
            If ($SSLProfile.sessreuse -ne "ENABLED"){@{ Description = "Enable Session Reuse [enabled]"; Value = $SSLProfile.sessreuse}}
            If ($SSLProfile.sessreuse -eq "ENABLED" -and $SSLprofile.sslprofiletype -ne "FrontEnd" -and $SSLProfile.sesstimeout -ne "300"){@{ Description = "Session Time-out"; Value = $SSLProfile.sesstimeout}}
			If ($SSLProfile.sessreuse -eq "ENABLED" -and $SSLprofile.sslprofiletype -ne "BackEnd" -and $SSLProfile.sesstimeout -ne "120"){@{ Description = "Session Time-out"; Value = $SSLProfile.sesstimeout}}
            If ($SSLProfile.cipherredirect -ne "DISABLED"){
				@{ Description = "Enable Cipher Redirect [disabled]"; Value = $SSLProfile.cipherredirect}
				@{ Description = "Cipher Redirect URL"; Value = $SSLProfile.cipherurl}
			}
            If ($SSLprofile.sslprofiletype -ne "BackEnd" -and $SSLProfile.clientauth -ne "DISABLED"){
				@{ Description = "Client Authentication [disabled]"; Value = $SSLProfile.clientauth}
				@{ Description = "Client Certificates"; Value = $SSLProfile.clientcert}
			}
            If ($SSLProfile.skipclientcertpolicycheck -ne "DISABLED"){@{ Description = "Skip Client Certificate Policy Check [disabled]"; Value = $SSLProfile.skipclientcertpolicycheck}}
            If ($SSLprofile.sslprofiletype -ne "FrontEnd" -and $SSLProfile.serverauth -ne "DISABLED"){
				@{ Description = "Enable Server Authentication [disabled]"; Value = $SSLProfile.serverauth}
				@{ Description = "Common Name"; Value = $SSLProfile.commonname}
            }
			If ($SSLProfile.ocspstapling -ne "DISABLED"){@{ Description = "OCSP Stapling [disabled]"; Value = $SSLProfile.ocspstapling}}
            If ($SSLProfile.sslredirect -ne "DISABLED"){
				@{ Description = "SSL Redirect [disabled]"; Value = $SSLProfile.sslredirect}
				If ($SSLProfile.redirectportrewrite -ne "DISABLED"){@{ Description = "SSL Redirect Port Rewrite [disabled]"; Value = $SSLProfile.redirectportrewrite}}
			}
            If ($SSLProfile.snienable -ne "DISABLED"){
				@{ Description = "Server Name Indication (SNI) [disabled]"; Value = $SSLProfile.snienable}
				If ($SSLProfile.allowunknownsni -ne "DISABLED"){@{ Description = "Allow Unknown SNI [disabled]"; Value = $SSLProfile.allowunknownsni}}
			}
			If ($SSLProfile.sendclosenotify -ne "YES"){@{ Description = "Send Close-Notify [yes]"; Value = $SSLProfile.sendclosenotify}}
            If ($SSLProfile.nonfipsciphers -ne "DISABLED"){@{ Description = "Non-FIPS Ciphers [disabled]"; Value = $SSLProfile.nonfipsciphers}}
            If ($SSLProfile.strictcachecks -ne "NO"){@{ Description = "Strict CA Checks [no]"; Value = $SSLProfile.strictcachecks}}
            If ($SSLprofile.sslprofiletype -ne "BackEnd" -and $SSLProfile.dropreqwithnohostheader -ne "NO"){@{ Description = "Drop requests for SNI enabled SSL sessions if host header is absent [no]"; Value = $SSLProfile.dropreqwithnohostheader}}
			If ($SSLProfile.clientauthuseboundcachain -ne "DISABLED"){@{ Description = "Enable Client Authentication using bound CA Chain [disabled]"; Value = $SSLProfile.clientauthuseboundcachain}}
            If ($SSLprofile.sslprofiletype -ne "BackEnd"){
				If ($SSLProfile.ssllogprofile -ne $null){@{ Description = "SSL Log Profile"; Value = $SSLProfile.ssllogprofile}}
				If ($SSLProfile.sessionticket -ne "DISABLED"){@{ Description = "Session Ticket [disabled]"; Value = $SSLProfile.sessionticket}}
				If ($SSLProfile.sessionticketlifetime -ne "300"){@{ Description = "Session Ticket Lifetime (secs) [300]"; Value = $SSLProfile.sessionticketlifetime}}
				If ($SSLProfile.sessionticketkeydata -ne $null){@{ Description = "Session Key"; Value = $SSLProfile.sessionticketkeydata}}
				If ($SSLProfile.sessionticketkeyrefresh -ne "ENABLED"){@{ Description = "Session Key Auto Refresh [enabled]"; Value = $SSLProfile.sessionticketkeyrefresh}}
				If ($SSLProfile.sessionkeylifetime -ne "3000"){@{ Description = "Session Key Lifetime (secs) [3000]"; Value = $SSLProfile.sessionkeylifetime}}
				If ($SSLProfile.prevsessionkeylifetime -ne "0"){@{ Description = "Previous Session Key Lifetime (secs) [0]"; Value = $SSLProfile.prevsessionkeylifetime}}
				@{ Description = "Enable Stricy Transport Security (HSTS) [disabled]"; Value = $SSLProfile.hsts}
			}
            If ($SSLProfile.maxage -ne "0"){@{ Description = "HSTS: Maximum Age [0]"; Value = $SSLProfile.maxage}}
            If ($SSLProfile.includesubdomains -ne "NO"){@{ Description = "HSTS: Include Subdomains [no]"; Value = $SSLProfile.includesubdomains}}
            If ($SSLProfile.preload -ne "NO"){@{ Description = "HSTS: Preload"; Value = $SSLProfile.preload}}
            @{ Description = "SSL 3 [disabled]"; Value = $SSLProfile.ssl3}
            @{ Description = "TLS 1 [enabled]"; Value = $SSLProfile.tls1}
            @{ Description = "TLS 1.1 [enabled]"; Value = $SSLProfile.tls11}
            @{ Description = "TLS 1.2 [enabled]"; Value = $SSLProfile.tls12}
            @{ Description = "TLS 1.3 [disabled]"; Value = $SSLProfile.tls13}
			If ($SSLProfile.tls13 -match "ENABLED"){
				If ($SSLProfile.zerorttearlydata -ne "DISABLED"){@{ Description = "Zero RTT Early Data [disabled]"; Value = $SSLProfile.zerorttearlydata}}
				If ($SSLProfile.dhekeyexchangewithpsk -ne "NO"){@{ Description = "DHE Key Exchange with PSK [no]"; Value = $SSLProfile.dhekeyexchangewithpsk}}
			}
			If ($SSLProfile.allowextendedmastersecret -ne "NO"){@{ Description = "Allow Extended Master Secret [no]"; Value = $SSLProfile.allowextendedmastersecret}}
            If ($SSLProfile.sslinterception -ne "DISABLED"){@{ Description = "SSL Session Interception [disabled]"; Value = $SSLProfile.sslinterception}}
			If ($SSLProfile.ssliverifyservercertforreuse -ne "ENABLED"){@{ Description = "Verify Server Certificate For Reuse On SSL Interception [enabled]"; Value = $SSLProfile.ssliverifyservercertforreuse}}
			If (![string]::IsNullOrWhiteSpace($ProfCiphers)){@{ Description = "Ciphers"; Value = $ProfCiphers}}
			If (![string]::IsNullOrWhiteSpace($ProfCert)){ @{ Description = "CA Cert"; Value = $ProfCert}}
        )
		If ($SSLPROFILEH.Length -gt 0) {
			$Params = $null
			$Params = @{
				Hashtable = $SSLPROFILEH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
    }
}
#endregion SSL Profiles

#region DTLS Profiles
If ((Get-vNetScalerObjectCount -Type ssldtlsprofile).__count -ge 1) {
	$nonDefaultFound = $false
	$dtlsprofiles = Get-vNetScalerObject -Type ssldtlsprofile
	foreach ($dtlsprofile in $dtlsprofiles) {
		[System.Collections.Hashtable[]] $DTLSPROFILESH = @(
			If ($dtlsprofile.maxrecordsize -ne "1459") {@{ Description = "Max Record Size [1459]"; Value = $dtlsprofile.maxrecordsize}}
			If ($dtlsprofile.maxpacketsize -ne "120") {@{ Description = "Max Packet Size [120]"; Value = $dtlsprofile.maxpacketsize}}
			If ($dtlsprofile.maxholdqlen -ne "32") {@{ Description = "Max HoldQ Size [32]"; Value = $dtlsprofile.maxholdqlen}}
			If ($dtlsprofile.maxretrytime -ne "3") {@{ Description = "Max Retry Time [3]"; Value = $dtlsprofile.maxretrytime}}
			If ($dtlsprofile.maxbadmacignorecount -ne "100") {@{ Description = "Max Bad Mac Ignore Count [100]"; Value = $dtlsprofile.maxbadmacignorecount}}
			If ($dtlsprofile.pmtudiscovery -ne "DISABLED") {@{ Description = "PMTU Discovery [disabled]"; Value = $dtlsprofile.pmtudiscovery}}
			If ($dtlsprofile.terminatesession -ne "DISABLED") {@{ Description = "Terminate Session [disabled]"; Value = $dtlsprofile.terminatesession}}
			If ($dtlsprofile.helloverifyrequest -ne "ENABLED") {@{ Description = "Hello Verify Request [enabled]"; Value = $dtlsprofile.helloverifyrequest}}
        )
		If ($DTLSPROFILESH.Length -gt 0) {
			$nonDefaultFound = $true
            if (!$titleWritten) {
                $selection.InsertNewPage()
                WriteWordLine 3 0 "DTLS Profiles"
                WriteWordLine 0 0 "Only non-default values are reported, defaults are in [brackets]"
                $titleWritten = $true
            }
			WriteWordLine 4 0 "$($dtlsprofile.name)"
			$Params = $null
			$Params = @{
				Hashtable = $DTLSPROFILESH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
    }
}
#endregion DTLS Profiles
#endregion Profiles

#region User Administration
$selection.InsertNewPage()
WriteWordLine 2 0 "User Administration"

#region System Users
WriteWordLine 3 0 "System Users"
$nssystemusers = Get-vNetScalerObject -Type systemuser
[System.Collections.Hashtable[]] $AUTHLOCH = @()
foreach ($nssystemuser in $nssystemusers) {
	$nssystemusercmdpol = Get-vNetScalerObject -Type systemuser_systemcmdpolicy_binding -Name $nssystemuser.username
    $AUTHLOCH += @{
        LocalUser 		= $nssystemuser.username
		ExternalAuth	= $nssystemuser.externalauth
		timeout			= $nssystemuser.timeout
		Logging			= $nssystemuser.logging
		allowedmanagementinterface	= $nssystemuser.allowedmanagementinterface -join ", "
		CMDPOL			= $nssystemusercmdpol.policyname -join ", "
    }
}
If ($AUTHLOCH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $AUTHLOCH
        Columns   = "LocalUser","ExternalAuth","timeout","Logging","allowedmanagementinterface","CMDPOL"
        Headers   = "Local User","Enable External Authentication","Idle session Timeout","Enable Logging Privilege","Allowed Management Interface","Command Policy"
    }
    $Table = AddWordTable @Params
    FindWordDocumentEnd
}
#endregion System Users

#region Database Users
If ((Get-vNetScalerObjectCount -Type dbuser).__count -ge 1) {
	WriteWordLine 3 0 "Database Users"
    $nsdbusers = Get-vNetScalerObject -Type dbuser
    [System.Collections.Hashtable[]] $DBUserH = @()
    foreach ($dbuser in $nsdbusers) {
        $DBUserH += @{
            DBUser = $dbuser.username
        }
    }
    If ($DBUserH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $DBUserH
            Columns   = "DBUser"
            Headers   = "Database User"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Database Users

#region System Groups
If ((Get-vNetScalerObjectCount -Type systemgroup).__count -ge 1) {
	WriteWordLine 3 0 "System Groups"
	$nssystemgroups = Get-vNetScalerObject -Type systemgroup
    [System.Collections.Hashtable[]] $AUTHGRPH = @()
    foreach ($nssystemgroup in $nssystemgroups) {
		$nssystemgroupmembers = Get-vNetScalerObject -Type systemgroup_systemuser_binding -Name $nssystemgroup.groupname
        $AUTHGRPH += @{
            SystemGroup = $nssystemgroup.groupname
			Prompt		= $nssystemgroup.promptstring
			mgmtint		= $nssystemgroup.allowedmanagementinterface -join ", "
			members		= $nssystemgroupmembers.username -join ", "
        }
    }
    If ($AUTHGRPH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $AUTHGRPH
            Columns   = "SystemGroup", "Prompt", "mgmtint", "members"
            Headers   = "System Group", "Prompt string", "Allowed Management Interface", "Members"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion System Groups

#region SMPP Users
If ((Get-vNetScalerObjectCount -Type smppuser).__count -ge 1) {
	WriteWordLine 3 0 "SMPP Users"
    $nssmppusers = Get-vNetScalerObject -Type smppuser
    $SMPPUserH = $null
    [System.Collections.Hashtable[]] $SMPPUserH = @()
    foreach ($smppuser in $nssmppusers) {
        $SMPPUserH += @{
            SMPPUser = $smppuser.username
        }
    }
    If ($SMPPUserH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $SMPPUserH
            Columns   = "SMPPUser"
            Headers   = "SMPP User"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion SMPP Users

#region Command Policies
If ((Get-vNetScalerObjectCount -Type systemcmdpolicy).__count -ge 1) {
	WriteWordLine 3 0 "Command Policies"
    $nscmdpols = Get-vNetScalerObject -Type systemcmdpolicy
    [System.Collections.Hashtable[]] $CMDPOLH = @()
    foreach ($nscmdpol in $nscmdpols) {
        $CMDPOLH += @{
            NAME    = $nscmdpol.policyname
            ACTION  = $nscmdpol.action
            CMDSPEC = $nscmdpol.cmdspec
        }
    }
    If ($CMDPOLH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $CMDPOLH
            Columns   = "NAME", "ACTION", "CMDSPEC"
            Headers   = "Policy Name", "Action", "Command Policy"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Command Policies

#endregion User Administration

#region Authentication
$authpols = (Get-vNetScalerObjectCount -Type authenticationpolicy).__count
$authpolsldapcount = (Get-vNetScalerObjectCount -Type authenticationldappolicy).__count
$authactsldapcount = (Get-vNetScalerObjectCount -Type authenticationldapaction).__count
$authpolsradiuscount = (Get-vNetScalerObjectCount -Type authenticationradiuspolicy).__count
$authactsradiuscount = (Get-vNetScalerObjectCount -Type authenticationradiusaction).__count
$authpolsTACACScount = (Get-vNetScalerObjectCount -Type authenticationTACACSpolicy).__count
$authactsTACACScount = (Get-vNetScalerObjectCount -Type authenticationTACACSaction).__count
$authpolssamlcount = (Get-vNetScalerObjectCount -Type authenticationsamlpolicy).__count
$authactssamlcount = (Get-vNetScalerObjectCount -Type authenticationsamlaction).__count
If ($authpols -ge 1 -or `
$authpolsldapcount -ge 1 -or `
$authactsldapcount -ge 1 -or `
$authpolsradiuscount -ge 1 -or ` 
$authactsradiuscount -ge 1 -or `
$authpolsTACACScount -ge 1 -or `
$authactsTACACScount -ge 1 -or `
$authpolssamlcount -ge 1 -or `
$authactssamlcount -ge 1){
	$selection.InsertNewPage()
	WriteWordLine 2 0 "Authentication"

	#region Global Policy bindings
	$authpolglobbinds = Get-vNetScalerObject -Type authenticationpolicy_systemglobal_binding -Bulk
	If ($authpolglobbinds){
		WriteWordLine 3 0 "Global Policy bindings"
		[System.Collections.Hashtable[]] $AUTHPOLBINDH = @()
		foreach ($authpolglobbind in $authpolglobbinds) {
			$AUTHPOLBINDH += @{
				Priority	= $authpolglobbind.priority
				Policy		= $authpolglobbind.name
				GotoExp		= $authpolglobbind.gotopriorityexpression
				nextfactor	= $authpolglobbind.nextfactor
			}
		}
		If ($AUTHPOLBINDH.Length -gt 0) {
			$Params = $null
			$Params = @{
				Hashtable = $AUTHPOLBINDH
				Columns   = "Priority", "Policy", "GotoExp", "nextfactor"
				Headers   = "Priority", "Policy", "Goto Expression", "Next Factor"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion Global Policy bindings
}

#region Advanced Authentication Policies
If ($authpols -ge 1) {
	WriteWordLine 3 0 "Advanced Authentication Policies"
	$authpols = Get-vNetScalerObject -Type authenticationpolicy
    [System.Collections.Hashtable[]] $AUTHPOLH = @()
    foreach ($authpol in $authpols) {
        $AUTHPOLH += @{
            Policy     = $authpol.name
            Expression = $authpol.rule
            Action     = $authpol.action
        }
    }
    If ($AUTHPOLH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $AUTHPOLH
            Columns   = "Policy", "Expression", "Action"
            Headers   = "Policy", "Expression", "Action"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }	
}
#endregion Advanced Authentication Policies

#region Authentication LDAP Policies
If ($authpolsldapcount -ge 1) {
	WriteWordLine 3 0 "LDAP Policies"
	$authpolsldap = Get-vNetScalerObject -Type authenticationldappolicy
    [System.Collections.Hashtable[]] $AUTHLDAPPOLH = @()
    foreach ($authpolldap in $authpolsldap) {
        $AUTHLDAPPOLH += @{
            Policy     = $authpolldap.name
            Expression = $authpolldap.rule
            Action     = $authpolldap.reqaction
        }
    }
    If ($AUTHLDAPPOLH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $AUTHLDAPPOLH
            Columns   = "Policy", "Expression", "Action"
            Headers   = "LDAP Policy", "Expression", "LDAP Action"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Authentication LDAP Policies

#region Authentication LDAP Servers
If ($authactsldapcount -ge 1) {
	WriteWordLine 3 0 "LDAP Authentication Servers"
	$authactsldap = Get-vNetScalerObject -Type authenticationldapaction
    foreach ($authactldap in $authactsldap) {
        WriteWordLine 4 0 "$($authactldap.name)"
        [System.Collections.Hashtable[]] $LDAPCONFIG = @(
            @{ Description = "Description"; Value = "Configuration"}
            @{ Description = "LDAP Server IP"; Value = $authactldap.serverip}
            @{ Description = "LDAP Server Port"; Value = $authactldap.serverport}
            @{ Description = "LDAP Server Time-Out"; Value = $authactldap.authtimeout}
            @{ Description = "Validate Certificate"; Value = $authactldap.validateservercert}
            @{ Description = "LDAP Base OU"; Value = $authactldap.ldapbase}
            @{ Description = "LDAP Bind DN"; Value = $authactldap.ldapbinddn}
            @{ Description = "Login Name"; Value = $authactldap.ldaploginname}
            @{ Description = "Sub Attribute Name"; Value = $authactldap.ssonameattribute}
            @{ Description = "Security Type"; Value = $authactldap.sectype}   
            @{ Description = "Password Changes"; Value = $authactldap.passwdchange}
            @{ Description = "Group attribute name"; Value = $authactldap.groupattrname}
            @{ Description = "LDAP Single Sign On Attribute"; Value = $authactldap.ssonameattribute}
            @{ Description = "Authentication"; Value = $authactldap.authentication}
            @{ Description = "User Required"; Value = $authactldap.requireuser}
            @{ Description = "LDAP Referrals"; Value = $authactldap.maxldapreferrals}
            @{ Description = "Nested Group Extraction"; Value = $authactldap.nestedgroupextraction}
            @{ Description = "Maximum Nesting level"; Value = $authactldap.maxnestinglevel}
        )
        $Params = $null
        $Params = @{
            Hashtable = $LDAPCONFIG
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
	}
}
#endregion Authentication LDAP Servers

#region Authentication Radius Policies
If ($authpolsradiuscount -ge 1) {
	WriteWordLine 3 0 "Radius Policies"
	$authpolsradius = Get-vNetScalerObject -Type authenticationradiuspolicy
    [System.Collections.Hashtable[]] $AUTHRADIUSPOLH = @()
    foreach ($authpolradius in $authpolsradius) {
        $AUTHRADIUSPOLH += @{
            Policy     = $authpolradius.name
            Expression = $authpolradius.rule
            Action     = $authpolradius.reqaction
        }
    }
    If ($AUTHRADIUSPOLH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $AUTHRADIUSPOLH
            Columns   = "Policy", "Expression", "Action"
            Headers   = "RADIUS Policy", "Expression", "RADIUS Action"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
	}
}
#endregion Authentication Radius Policies

#region Authentication RADIUS Servers
If ($authactsradiuscount -ge 1) {
	WriteWordLine 3 0 "Radius Authentication Servers"
	$authactsradius = Get-vNetScalerObject -Type authenticationradiusaction
    foreach ($authactradius in $authactsradius) {
        WriteWordLine 4 0 "$($authactradius.name)"    
        [System.Collections.Hashtable[]] $RADUIUSCONFIG = @(
            @{ Description = "Description"; Value = "Configuration"}
            @{ Description = "RADIUS Server IP"; Value = $authactradius.serverip}
            @{ Description = "RADIUS Server Port"; Value = $authactradius.serverport}
            @{ Description = "RADIUS Server Time-Out"; Value = $authactradius.authtimeout}
            @{ Description = "Radius NAS IP"; Value = $authactradius.radnasip}
            @{ Description = "IP Vendor ID"; Value = $authactradius.ipvendorid}
            @{ Description = "Accounting"; Value = $authactradius.accounting}
            @{ Description = "Calling Station ID"; Value = $authactradius.callingstationid}
        )
        $Params = $null
        $Params = @{
            Hashtable = $RADUIUSCONFIG
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
    }
}
#endregion Authentication RADIUS

#region Authentication TACACS Policies
If ($authpolsTACACScount -ge 1) {
	WriteWordLine 3 0 "TACACS Policies"
	$authpolsTACACS = Get-vNetScalerObject -Type authenticationTACACSpolicy
    [System.Collections.Hashtable[]] $AUTHTACACSPOLH = @()
    foreach ($authpoltacacs in $authpolstacacs) {
        $AUTHTACACSPOLH += @{
            Policy     = $authpoltacacs.name
            Expression = $authpoltacacs.rule
            Action     = $authpoltacacs.reqaction
        }
    }
    If ($AUTHtacacsPOLH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $AUTHTACACSPOLH
            Columns   = "Policy", "Expression", "Action"
            Headers   = "TACACS Policy", "Expression", "TACACS Action"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Authentication TACACS Policies

#region Authentication TACACS Server
If ($authactsTACACScount -ge 1) {
	WriteWordLine 3 0 "TACACS authentication Servers"
	$authactsTACACS = Get-vNetScalerObject -Type authenticationTACACSaction
    foreach ($authactTACACS in $authactsTACACS) {
        WriteWordLine 4 0 "$($authactTACACS.name)"
        [System.Collections.Hashtable[]] $RADUIUSCONFIG = @(
            @{ Description = "Description"; Value = "Configuration"}
            @{ Description = "TACACS Server IP"; Value = $authactTACACS.serverip}
            @{ Description = "TACACS Server Port"; Value = $authactTACACS.serverport}
            @{ Description = "TACACS Server Time-Out"; Value = $authactTACACS.authtimeout}
            @{ Description = "TACACS Authorization"; Value = $authactTACACS.authorization}
            @{ Description = "TACACS Accounting"; Value = $authactTACACS.accounting}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.groupattrname)){@{ Description = "Group Attribute Name"; Value = $authactTACACS.groupattrname}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.defaultauthenticationgroup)){@{ Description = "Default Authentication Group"; Value = $authactTACACS.defaultauthenticationgroup}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attributes)){@{ Description = "Attributes"; Value = $authactTACACS.attributes}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute1)){@{ Description = "Attribute1"; Value = $authactTACACS.attribute1}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute2)){@{ Description = "Attribute2"; Value = $authactTACACS.attribute2}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute3)){@{ Description = "Attribute3"; Value = $authactTACACS.attribute3}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute4)){@{ Description = "Attribute4"; Value = $authactTACACS.attribute4}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute5)){@{ Description = "Attribute5"; Value = $authactTACACS.attribute5}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute6)){@{ Description = "Attribute6"; Value = $authactTACACS.attribute6}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute7)){@{ Description = "Attribute7"; Value = $authactTACACS.attribute7}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute8)){@{ Description = "Attribute8"; Value = $authactTACACS.attribute8}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute9)){@{ Description = "Attribute9"; Value = $authactTACACS.attribute9}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute10)){@{ Description = "Attribute10"; Value = $authactTACACS.attribute10}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute11)){@{ Description = "Attribute11"; Value = $authactTACACS.attribute11}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute12)){@{ Description = "Attribute12"; Value = $authactTACACS.attribute12}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute13)){@{ Description = "Attribute13"; Value = $authactTACACS.attribute13}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute14)){@{ Description = "Attribute14"; Value = $authactTACACS.attribute14}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute15)){@{ Description = "Attribute15"; Value = $authactTACACS.attribute15}}
            If (![string]::IsNullOrWhiteSpace($authactTACACS.attribute16)){@{ Description = "Attribute16"; Value = $authactTACACS.attribute16}}
        )
        $Params = $null
        $Params = @{
            Hashtable = $RADUIUSCONFIG
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
    }
}
#endregion Authentication TACACS Server

#region Authentication SAML Policies
If ($authpolssamlcount -ge 1) {
	WriteWordLine 3 0 "SAML Policies"
	$authpolssaml = Get-vNetScalerObject -Type authenticationsamlpolicy
    [System.Collections.Hashtable[]] $AUTHSAMLPOLH = @()
    foreach ($authpolsaml in $authpolssaml) {
        $AUTHSAMLPOLH += @{
            Policy     = $authpolsaml.name
            Expression = $authpolsaml.rule
            Action     = $authpolsaml.reqaction
        }
    }
    If ($AUTHSAMLPOLH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $AUTHSAMLPOLH
            Columns   = "Policy", "Expression", "Action"
            Headers   = "SAML Policy", "Expression", "SAML Action"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Authentication SAML Policies

#region Authentication SAML Servers
If ($authactssamlcount -ge 1) {
	WriteWordLine 3 0 "SAML authentication Servers"
	$authactssaml = Get-vNetScalerObject -Type authenticationsamlaction
	foreach ($authactsaml in $authactssaml) {
        WriteWordLine 4 0 "$($authactsaml.name)"
        [System.Collections.Hashtable[]] $SAMLCONFIG = @(
            @{ Description = "Description"; Value = "Configuration"}
            @{ Description = "IDP Certificate Name"; Value = $authactsaml.samlidpcertname}
            @{ Description = "Signing Certificate Name"; Value = $authactsaml.samlsigningcertname}
            @{ Description = "Redirect URL"; Value = $authactsaml.samlredirecturl}
            @{ Description = "Assertion Consumer Service Index"; Value = $authactsaml.samlacsindex}
            @{ Description = "User Field"; Value = $authactsaml.samluserfield}
            @{ Description = "Reject Unsigned Authentication"; Value = $authactsaml.samlrejectunsignedassertion}
            @{ Description = "Issuer Name"; Value = $authactsaml.samlissuername}
            @{ Description = "Two factor"; Value = $authactsaml.samltwofactor}
            @{ Description = "Signature Algorithm"; Value = $authactsaml.signaturealg}
            @{ Description = "Digest Method"; Value = $authactsaml.digestmethod}
            @{ Description = "Requested Authentication Context"; Value = $authactsaml.requestedauthncontext}
            @{ Description = "Binding"; Value = $authactsaml.samlbinding}
            @{ Description = "Attribute Consuming Service Index"; Value = $authactsaml.attributeconsumingserviceindex}
            @{ Description = "Send Thumb Print"; Value = $authactsaml.sendthumbprint}
            @{ Description = "Enforce User Name"; Value = $authactsaml.enforceusername}
            @{ Description = "Single Logout URL"; Value = $authactsaml.logouturl}
            @{ Description = "Skew Time"; Value = $authactsaml.skewtime}
            @{ Description = "Force Authentication"; Value = $authactsaml.forceauthn}
        )
        $Params = $null
        $Params = @{
            Hashtable = $SAMLCONFIG
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
    }
}
#endregion Authentication SAML Servers
#endregion Authentication

#region Auditing
$selection.InsertNewPage()
WriteWordLine 2 0 "Auditing"

#region Syslog Parameters
WriteWordLine 3 0 "Syslog Parameters"
$syslogparams = Get-vNetScalerObject -Type auditsyslogparams
[System.Collections.Hashtable[]] $SYSLOGPARAMH = @(
    @{ Description = "Server IP"; Value = $syslogparams.serverip}
    @{ Description = "Server Port"; Value = $syslogparams.serverport}
    @{ Description = "Date Format"; Value = $syslogparams.dateformat}
    @{ Description = "Log level"; Value = $syslogparams.loglevel -join ", "}
    @{ Description = "Log Facility"; Value = $syslogparams.logfacility}
    @{ Description = "Log TCP Messages"; Value = $syslogparams.tcp}
    @{ Description = "Log ACL Messages"; Value = $syslogparams.acl}
    @{ Description = "TimeZone"; Value = $syslogparams.timezone}
    @{ Description = "Log User Defined Messages"; Value = $syslogparams.userdefinedauditlog}
    @{ Description = "AppFlow Export"; Value = $syslogparams.appflowexport}
    @{ Description = "Log Large Scale NAT Messages"; Value = $syslogparams.lsn}
    @{ Description = "Log ALG Messages"; Value = $syslogparams.alg}
    @{ Description = "Log Subscriber Session Messages"; Value = $syslogparams.subscriberlog}
    @{ Description = "Log DNS Messages"; Value = $syslogparams.dns}   
)
$Params = $null
$Params = @{
    Hashtable = $SYSLOGPARAMH
    Columns   = "Description", "Value"
    Headers   = "Description", "Value"
}
$Table = AddWordTable @Params
FindWordDocumentEnd
#endregion Syslog Parameters

#region Syslog Policies
##BUG Does not report on no configured syslog policies
##Fixed AM 09/05/2017
If ((Get-vNetScalerObjectCount -Type auditsyslogpolicy).__count -ge 1) {
	WriteWordLine 3 0 "Syslog Policies"
    $syslogpolicies = Get-vNetScalerObject -Type auditsyslogpolicy
    [System.Collections.Hashtable[]] $SYSLOGPOLH = @()
    foreach ($syslogpolicy in $syslogpolicies) {
        $SYSLOGPOLH += @{
            NAME   = $syslogpolicy.name
            RULE   = $syslogpolicy.rule
            ACTION = $syslogpolicy.action
        }
    }
    $Params = $null
    $Params = @{
        Hashtable = $SYSLOGPOLH
        Columns   = "NAME", "RULE", "ACTION"
        Headers   = "Policy Name", "Rule", "Action"
    }
    $Table = AddWordTable @Params
    FindWordDocumentEnd
}
#endregion Syslog Policies

#region Syslog Actions
##BUG Does not report on no configured syslog policies
##Fixed AM 09/05/2017
If ((Get-vNetScalerObjectCount -Type auditsyslogaction).__count -ge 1) {
	WriteWordLine 3 0 "Syslog Servers"
    $syslogservers = Get-vNetScalerObject -Type auditsyslogaction 
    foreach ($syslogserver in $syslogservers) {
        WriteWordLine 4 0 "$($syslogserver.name)"
        [System.Collections.Hashtable[]] $SYSLOGSRVH = @(
            @{ Description = "Server IP"; Value = $syslogserver.serverip}
            If (![string]::IsNullOrWhiteSpace($syslogserver.serverdomainname)){@{ Description = "Server Domain Name"; Value = $syslogserver.serverdomainname}}
            If (![string]::IsNullOrWhiteSpace($syslogserver.domainresolveretry)){@{ Description = "DNS Resolution Retry"; Value = $syslogserver.domainresolveretry}}
            If (![string]::IsNullOrWhiteSpace($syslogserver.lbvservername)){@{ Description = "LB vServer Name"; Value = $syslogserver.lbvservername}}
            @{ Description = "Server Port"; Value = $syslogserver.serverport}
            @{ Description = "Log level"; Value = $syslogserver.loglevel -join ", "}
            @{ Description = "Date Format"; Value = $syslogserver.dateformat}
            @{ Description = "Log Facility"; Value = $syslogserver.logfacility}
            @{ Description = "Time Zone"; Value = $syslogserver.timezone}
            @{ Description = "TCP Logging"; Value = $syslogserver.tcp}
            @{ Description = "ACL Logging"; Value = $syslogserver.acl}
            @{ Description = "User Configurable Log Messages"; Value = $syslogserver.userdefinedauditlog}
            @{ Description = "AppFlow Logging"; Value = $syslogserver.appflowexport}
            @{ Description = "Large Scale NAT Logging"; Value = $syslogserver.lsn}
            @{ Description = "ALG messages Logging"; Value = $syslogserver.alg}
            @{ Description = "Subscriber Logging"; Value = $syslogserver.subscriberlog}
            @{ Description = "DNS Logging"; Value = $syslogserver.dns}
            @{ Description = "Transport Type"; Value = $syslogserver.transport}
            If (![string]::IsNullOrWhiteSpace($syslogserver.netprofile)){@{ Description = "Net Profile"; Value = $syslogserver.netprofile}}
            If (![string]::IsNullOrWhiteSpace($syslogserver.maxlogdatasizetohold)){@{ Description = "Max Log Data to hold"; Value = $syslogserver.maxlogdatasizetohold}}
        )
        $Params = $null
        $Params = @{
            Hashtable = $SYSLOGSRVH
            Columns   = "Description", "Value"
            Headers   = "Description", "Value"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    } #end foreach
}
#endregion Syslog Actions

#region message Actions
If ((Get-vNetScalerObjectCount -Type auditmessageaction).__count -ge 1) {
	WriteWordLine 3 0 "Auditing Message Actions"
    $auditmessageactions = Get-vNetScalerObject -Type auditmessageaction
    [System.Collections.Hashtable[]] $AUDITMESSAGEACTIONH = @()
    foreach ($auditmessageaction in $auditmessageactions) {
        $AUDITMESSAGEACTIONH += @{
            NAME		= $auditmessageaction.name
            LOGLEVEL	= $auditmessageaction.loglevel1
            Expression	= $auditmessageaction.stringbuilderexpr
			LOGNS		= $auditmessageaction.logtonewnslog
        }
	}
	$Params = $null
	$Params = @{
		Hashtable = $AUDITMESSAGEACTIONH
		Columns   = "NAME", "LOGLEVEL", "Expression", "LOGNS"
		Headers   = "Name", "Log Level", "Expression", "Log in newnslog"
	}
	$Table = AddWordTable @Params
	FindWordDocumentEnd
}
#endregion message Actions
#endregion Auditing

#region SNMP Monitoring
$selection.InsertNewPage()
WriteWordLine 2 0 "SNMP Monitoring"
If ((Get-vNetScalerObjectCount -Type snmpcommunity).__count -ge 1) {
	WriteWordLine 3 0 "SNMP Community"
	$snmpcoms = Get-vNetScalerObject -Type snmpcommunity
    [System.Collections.Hashtable[]] $SNMPCOMH = @()
    foreach ($snmpcom in $snmpcoms) {
        $SNMPCOMH += @{
            SNMPCommunity = $snmpcom.communityname
            Permissions   = $snmpcom.permissions
        }
    }
    If ($SNMPCOMH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $SNMPCOMH
            Columns   = "SNMPCommunity", "Permissions"
            Headers   = "SNMP Community", "Permissions"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}

If ((Get-vNetScalerObjectCount -Type snmptrap).__count -ge 1) {
	WriteWordLine 3 0 "SNMP Traps"
	$snmptraps = Get-vNetScalerObject -Type snmptrap 
    [System.Collections.Hashtable[]] $SNMPTRAPSH = @()
    foreach ($snmptrap in $snmptraps) {
        $SNMPTRAPSH += @{
            Type        = $snmptrap.trapclass
            Destination = $snmptrap.trapdestination
            Version     = $snmptrap.version
            Port        = $snmptrap.destport
            Name        = $snmptrap.communityname
        }
    }
    If ($SNMPTRAPSH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $SNMPTRAPSH
            Columns   = "Type", "Destination", "Version", "Port", "Name"
            Headers   = "Type", "Trap Destination", "Version", "Destination Port", "Community Name"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}

If ((Get-vNetScalerObjectCount -Type snmpmanager).__count -ge 1) {
	WriteWordLine 3 0 "SNMP Manager"
    $snmpmanagers = Get-vNetScalerObject -Type snmpmanager 
    [System.Collections.Hashtable[]] $SNMPMANSH = @()
    foreach ($snmpmanager in $snmpmanagers) {
        $SNMPMANSH += @{
            SNMPManager = $snmpmanager.ipaddress
            Netmask     = $snmpmanager.netmask
        }
    }
    If ($SNMPMANSH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $SNMPMANSH
            Columns   = "SNMPManager", "Netmask"
            Headers   = "SNMP Manager", "Netmask"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}

WriteWordLine 3 0 "SNMP Alarms"
[System.Collections.Hashtable[]] $SNMPALERTSH = @()
$snmpalarms = Get-vNetScalerObject -Type snmpalarm 
foreach ($snmpalarm in $snmpalarms) {
    $SNMPALERTSH += @{
        Alarm    = $snmpalarm.trapname
        State    = $snmpalarm.state
        Time     = $snmpalarm.time
        TimeOut  = $snmpalarm.timeout
        Severity = $snmpalarm.severity
        Logging  = $snmpalarm.logging
    }
}
If ($SNMPALERTSH.Length -gt 0) {
    $Params = @{
        Hashtable = $SNMPALERTSH
        Columns   = "Alarm", "State", "Time", "TimeOut", "Severity", "Logging"
        Headers   = "Alarm", "State", "Time", "Time-Out", "Severity", "Logging"
    }
    $Table = AddWordTable @Params
    FindWordDocumentEnd
}

If ((Get-vNetScalerObjectCount -Type snmpview).__count -ge 1) {
	WriteWordLine 3 0 "SNMP Views"
	$snmpviews = Get-vNetScalerObject -Type snmpview 
    [System.Collections.Hashtable[]] $SNMPVIEWSH = @()
    foreach ($snmpview in $snmpviews) {
        $SNMPVIEWSH += @{
            Name		= $snmpview.name
            subtree		= $snmpview.subtree
            type		= $snmpview.type
            storagetype	= $snmpview.storagetype
			Status		= $snmpview.Status
        }
    }
    If ($SNMPVIEWH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $SNMPVIEWSH
            Columns   = "Name", "subtree", "type", "storagetype", "status"
            Headers   = "Name", "Subtree", "Type", "Storage Type", "Status"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}

If ((Get-vNetScalerObjectCount -Type snmpgroup).__count -ge 1) {
	WriteWordLine 3 0 "SNMP Groups"
	$snmpgroups = Get-vNetScalerObject -Type snmpgroup 
    [System.Collections.Hashtable[]] $SNMPGROUPSH = @()
    foreach ($snmpgroup in $snmpgroups) {
        $SNMPGROUPSH += @{
            Name			= $snmpgroup.name
            securitylevel	= $snmpgroup.securitylevel
            readviewname	= $snmpgroup.readviewname
            storagetype		= $snmpgroup.storagetype
			status			= $snmpgroup.status
        }
    }
    If ($SNMPGROUPSH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $SNMPGROUPSH
            Columns   = "Name", "securitylevel", "readviewname", "storagetype", "status"
            Headers   = "Name", "Security Level", "Read View Name", "Storage Type", "Status"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}

If ((Get-vNetScalerObjectCount -Type snmpuser).__count -ge 1) {
	WriteWordLine 3 0 "SNMP Users"
	$snmpusers = Get-vNetScalerObject -Type snmpuser 
    [System.Collections.Hashtable[]] $SNMPUSERSH = @()
    foreach ($snmpuser in $snmpusers) {
        $SNMPUSERSH += @{
            Name		= $snmpuser.name
            group		= $snmpuser.group
            authtype	= $snmpuser.authtype
            privtype	= $snmpuser.privtype
        }
    }
    If ($SNMPUSERSH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $SNMPUSERSH
            Columns   = "Name", "Group", "authtype", "privtype"
            Headers   = "Name", "Group", "Authentication Type", "Privacy Type", "Community Name"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Monitoring

#region AppFlow
If ($FEATAppFlow -eq "Enabled"){
	$selection.InsertNewPage()
	WriteWordLine 2 0 "AppFlow"

	#region AppFlow Parameters
	WriteWordLine 3 0 "AppFlow Parameters"
	$afparams = Get-vNetScalerObject -Type appflowparam
	[System.Collections.Hashtable[]] $AFPARAMH = @(
		@{ Description = "HTTP URL"; Value = $afparams.httpurl}
		@{ Description = "HTTP Cookie"; Value = $afparams.httocookie}
		@{ Description = "HTTP Method"; Value = $afparams.httpmethod}
		@{ Description = "HTTP User-Agent"; Value = $afparams.httpuseragent}
		@{ Description = "HTTP Authorization"; Value = $afparams.httpauthorization}
		@{ Description = "HTTP Via"; Value = $afparams.httpvia}
		@{ Description = "HTTP Setcookie"; Value = $afparams.httpsetcookie}
		@{ Description = "HTTP Client Traffic Only"; Value = $afparams.clienttrafficonly}
		@{ Description = "HTTP Domain"; Value = $afparams.httpdomain}
		@{ Description = "Stream Identifier Name Logging"; Value = $afparams.identifiername}
		@{ Description = "Cache Insight"; Value = $afparams.cacheinsight}
		@{ Description = "Subscriber Awareness"; Value = $afparams.subscriberawareness}
		@{ Description = "Security Insight Traffic"; Value = $afparams.securityinsighttraffic}
		@{ Description = "URL Category"; Value = $afparams.urlcategory}
		@{ Description = "CQA Reporting"; Value = $afparams.cqareporting}
		@{ Description = "AAA Username"; Value = $afparams.aaausername}
		@{ Description = "HTTP Referrer"; Value = $afparams.httpreferer}
		@{ Description = "HTTP Host"; Value = $afparams.httphost}
		@{ Description = "HTTP Content-Type"; Value = $afparams.httpcontenttype}
		@{ Description = "HTTP X-Forwarded-For"; Value = $afparams.httpxforwardedfor}
		@{ Description = "HTTP Location"; Value = $afparams.httplocation}
		@{ Description = "HTTP Setcookie2"; Value = $afparams.httpsetcookie2}
		@{ Description = "Connection Chaining"; Value = $afparams.connectionchaining}
		@{ Description = "Skip Cache Redirection HTTP Transaction"; Value = $afparams.skipcacheredirectionhttptransaction}
		@{ Description = "Stream Identifier Session Name Logging"; Value = $afparams.identifiersessionname}
		@{ Description = "Video Insight"; Value = $afparams.videoinsight}
		@{ Description = "Subscriber ID Obfuscation"; Value = $afparams.subscriberidobfuscation}
		@{ Description = "HTTP Query Segment Along With the URL"; Value = $afparams.httpquerywithurl}
		@{ Description = "LSN Logging"; Value = $afparams.lsnlogging}
		@{ Description = "User Email-ID Logging"; Value = $afparams.emailaddress}
		@{ Description = "Observation Domain ID"; Value = $afparams.observationdomainid}
		@{ Description = "Observation Domain Name"; Value = $afparams.observationdomainname}
		@{ Description = "Template Refresh Interval"; Value = $afparams.templaterefresh}
		@{ Description = "Appname Refresh Interval"; Value = $afparams.appnamerefresh}
		@{ Description = "Flow Record Export Interval"; Value = $afparams.flowrecordinterval}
		@{ Description = "UDP Max Transmission Unit"; Value = $afparams.udppmtu}
		@{ Description = "Security Insight Record Interval"; Value = $afparams.securityinsightrecordinterval}   
	)
	$Params = $null
	$Params = @{
		Hashtable = $AFPARAMH
		Columns   = "Description", "Value"
		Headers   = "Description", "Value"
	}
	$Table = AddWordTable @Params
	FindWordDocumentEnd
	#endregion AppFlow Parameters

	#region AppFlow Collectors
	If ((Get-vNetScalerObjectCount -Type appflowcollector).__count -ge 1) {
		WriteWordLine 3 0 "AppFlow Collectors"
		$afcols = Get-vNetScalerObject -Type appflowcollector 
		[System.Collections.Hashtable[]] $AFCOLSH = @()
		foreach ($afcol in $afcols) {
			$AFCOLSH += @{
				Name       = $afcol.name
				IP         = $afcol.ipaddress
				Port       = $afcol.port
				NetProfile = (Get-NonEmptyString $afcol.netprofile)
				Transport  = (Get-NonEmptyString $afcol.transport)
			}
		}
		If ($AFCOLSH.Length -gt 0) {
			$Params = $null
			$Params = @{
				Hashtable = $AFCOLSH
				Columns   = "Name", "IP", "Port", "NetProfile", "Transport"
				Headers   = "Name", "IP Address", "Port", "Net Profile", "Transport"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion AppFlow Collectors

	#region AppFlow Policies
	If ((Get-vNetScalerObjectCount -Type appflowpolicy).__count -ge 1) {
		WriteWordLine 3 0 "AppFlow Policies"
		$afpols = Get-vNetScalerObject -Type appflowpolicy 
		[System.Collections.Hashtable[]] $AFPOLSH = @()
		foreach ($afpol in $afpols) {
			$AFPOLSH += @{
				Name        = $afpol.name
				Rule        = $afpol.rule
				Action      = $afpol.action
				UNDEFaction = Get-NonEmptyString $afpol.undefaction
				Comment     = Get-NonEmptyString $afpol.comment
			}
		}
		If ($AFPOLSH.Length -gt 0) {
			$Params = $null
			$Params = @{
				Hashtable = $AFPOLSH
				Columns   = "Name", "Rule", "Action", "UNDEFaction", "Comment"
				Headers   = "Name", "Rule", "Action", "UNDEF Action", "Comments"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion AppFlow Policies

	#region AppFlow Actions
	If ((Get-vNetScalerObjectCount -Type appflowaction).__count -ge 1) {
		WriteWordLine 3 0 "AppFlow Actions"
		$afacts = Get-vNetScalerObject -Type appflowaction 
		ForEach ($afact in $afacts) {
			WriteWordLine 4 0 "$($afact.name)"
			[System.Collections.Hashtable[]] $AFACTH = @(
				@{ Description = "Collectors"; Value = $afact.collectors -join ","}
				@{ Description = "Enable Client Side Measurement"; Value = $afact.clientsidemeasurements}
				@{ Description = "Page Tracking"; Value = $afact.pagetracking}
				@{ Description = "Web Insight"; Value = $afact.webinsight}
				@{ Description = "Security Insight"; Value = $afact.securityinsight}
				@{ Description = "Distribution Algorithm"; Value = $afact.distributionalgorithm}
				@{ Description = "Video Analytics"; Value = $afact.videoanalytics}
				@{ Description = "Comments"; Value = $afact.comment}
			)
			$Params = $null
			$Params = @{
				Hashtable = $AFACTH
				Columns   = "Description", "Value"
				Headers   = "Description", "Value"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion AppFlow Actions

	#region AppFlow Policy Labels
	If ((Get-vNetScalerObjectCount -Type appflowpolicylabel).__count -ge 1) {
		WriteWordLine 3 0 "AppFlow Policy Labels"
		$afpollbls = Get-vNetScalerObject -Type appflowpolicylabel 
		ForEach ($afpollbl in $afpollbls) {
			WriteWordLine 4 0 "$($afpollbl.labelname)"
			[System.Collections.Hashtable[]] $AFPOLLBLH = @(
				@{ Description = "Label Type"; Value = $afpollbl.policylabeltype}
				@{ Description = "Number of Bound Policies"; Value = $afpollbl.numpols}
			)
			$Params = $null
			$Params = @{
				Hashtable = $AFPOLLBLH
				Columns   = "Description", "Value"
				Headers   = "Description", "Value"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd

			If ((Get-vNetScalerObjectCount -Type appflowpolicylabel_appflowpolicy_binding -Name $afpollbl.labelname).__count -ge 1) {
				$afpolbinds = Get-vNetScalerObject -Type appflowpolicylabel_appflowpolicy_binding -Name $afpollbl.labelname 
				[System.Collections.Hashtable[]] $AFPOLBINDSH = @()
				foreach ($afpolbind in $afpolbinds) {
					$AFPOLBINDSH += @{
						Priority       = $afpolbind.priority
						Name           = $afpolbind.policyname
						GoToExpression = Get-NonEmptyString $afpolbind.gotopriorityexpression
						Invoke         = Get-NonEmptyString $afpolbind.invoke
					}
				}
				If ($AFPOLBINDSH.Length -gt 0) {
					$Params = $null
					$Params = @{
						Hashtable = $AFPOLBINDSH
						Columns   = "Priority", "Name", "GoToExpression", "Invoke"
						Headers   = "Priority", "Policy Name", "GoTo Expression", "Invoke"
					}
					$Table = AddWordTable @Params
					FindWordDocumentEnd
				}
			}
		}
	}
	#endregion AppFlow Policy Labels

	#region AppFlow Analytics Profiles
	If ((Get-vNetScalerObjectCount -Type analyticsprofile).__count -ge 1) {
		WriteWordLine 3 0 "AppFlow Analytics Profiles"
		$aprofs = Get-vNetScalerObject -Type analyticsprofile 
		ForEach ($aprof in $aprofs) {
			WriteWordLine 4 0 "$($aprof.name)"
			[System.Collections.Hashtable[]] $APROFH = @(
				@{ Description = "Collectors"; Value = $aprof.collectors -join ","}
				@{ Description = "Type"; Value = $aprof.type}
				@{ Description = "HTTP Client Side Measurement"; Value = $aprof.httpclientsidemeasurements}
				@{ Description = "HTTP Page Tracking"; Value = $aprof.httppagetracking}
				@{ Description = "HTTP URL"; Value = $aprof.httpurl}
				@{ Description = "HTTP Host"; Value = $aprof.httphost}
				@{ Description = "HTP Method"; Value = $aprof.httpmethod}
				@{ Description = "HTTP Referrer"; Value = $aprof.httpreferer}
				@{ Description = "HTTP User Agent"; Value = $aprof.httpuseragent}
				@{ Description = "HTTP Cookie"; Value = $aprof.httpcookie}
				@{ Description = "HTTP Location"; Value = $aprof.httplocation}
				@{ Description = "URL Category"; Value = $aprof.urlcategory}
				@{ Description = "HTTP Content Type"; Value = $aprof.httpcontenttype}
				@{ Description = "HTTP Authentication"; Value = $aprof.httpauthentication}
				@{ Description = "HTTP Via"; Value = $aprof.httpvia}
				@{ Description = "HTTP X Forwarded For Header"; Value = $aprof.httpxforwardedforheader}
				@{ Description = "HTTP Set Cookie"; Value = $aprof.httpsetcookie}
				@{ Description = "HTTP Set Cookie2"; Value = $aprof.httpsetcookie2}
				@{ Description = "HTTP Domain Name"; Value = $aprof.httpdomainname}
				@{ Description = "HTTP URL Query"; Value = $aprof.httpurlquery}
				@{ Description = "Integrated Cache"; Value = $aprof.integratedcache}
				@{ Description = "TCP Burst Reporting"; Value = $aprof.tcpburstreporting}
			)
			$Params = $null
			$Params = @{
				Hashtable = $APROFH
				Columns   = "Description", "Value"
				Headers   = "Description", "Value"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion AppFlow Analytics Profiles
}
#endregion AppFlow

#region Clustering
If ((Get-vNetScalerObjectCount -Type clusternode).__count -ge 1) {
	$selection.InsertNewPage()
	WriteWordLine 2 0 "Cluster"

    #region Cluster Instances
    WriteWordLine 3 0 "Cluster Instances"
    $nsclusterinstance = Get-vNetScalerObject -Type clusterinstance
    [System.Collections.Hashtable[]] $CLUSTINSTH = @(
        @{ Description = "Cluster ID"; Value = $nsclusterinstance.clid}
        @{ Description = "Dead interval (seconds)"; Value = $nsclusterinstance.deadinterval}
        @{ Description = "Hello interval (milliseconds)"; Value = $nsclusterinstance.hellointerval}
        @{ Description = "Quorum Type"; Value = $nsclusterinstance.quorumtype}
        @{ Description = "Preemption"; Value = $nsclusterinstance.preemption}
        @{ Description = "INC Mode"; Value = $nsclusterinstance.inc}
        @{ Description = "Process Local"; Value = $nsclusterinstance.processlocal}
        @{ Description = "Retain Connection on Cluster"; Value = $nsclusterinstance.retainconnectionsoncluster}
        @{ Description = "Admin State"; Value = $nsclusterinstance.adminstate}
        @{ Description = "Operational State"; Value = $nsclusterinstance.operationalstate}
        @{ Description = "Status"; Value = $nsclusterinstance.status}
        @{ Description = "Propagation"; Value = $nsclusterinstance.propstate}
    )
    $Params = $null
    $Params = @{
        Hashtable = $CLUSTINSTH
        Columns   = "Description", "Value"
        Headers   = "Description", "Value"
    }
    $Table = AddWordTable @Params
    FindWordDocumentEnd
    #endregion Cluster Instances

    #region Cluster Nodes
    WriteWordLine 3 0 "Cluster Nodes"
    $nsclusternodes = Get-vNetScalerObject -Type clusternode
    foreach ($nsclusternode in $nsclusternodes) {
        [System.Collections.Hashtable[]] $CLUSTNODESH = @(
            @{ Description = "Node IP Address"; Value = $nsclusternode.ipaddress}
            @{ Description = "Backplane Interface"; Value = $nsclusternode.backplane}
            @{ Description = "Health"; Value = $nsclusternode.clusterhealth}
            @{ Description = "Operational State"; Value = $nsclusternode.effectivestate}
            @{ Description = "Sync State"; Value = $nsclusternode.operationalsyncstate}
            @{ Description = "Priority"; Value = $nsclusternode.priority}
            @{ Description = "Admin State"; Value = $nsclusternode.state}
            @{ Description = "Is Configuration Coordinator"; Value = $nsclusternode.isconfigurationcoordinator}
            @{ Description = "Local Node"; Value = $nsclusternode.islocalnode}
        )
        $Params = $null
        $Params = @{
            Hashtable = $CLUSTNODESH
            Columns   = "Description", "Value"
            Headers   = "Description", "Value"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
    #endregion Cluster Nodes
}
#endregion Clustering

#region Networking
$selection.InsertNewPage()
WriteWordLine 2 0 "Networking"

$TDcounter = (Get-vNetScalerObjectCount -Type nstrafficdomain).__count

#region IP addresses
WriteWordLine 3 0 "IP addresses"
$IPs = Get-vNetScalerObject -Type nsip
[System.Collections.Hashtable[]] $IPADDRESSH = @()
foreach ($IP in $IPs) {
	If ($TDcounter -ge 1){
		$IPADDRESSH += @{
			IPAddress     = $IP.ipaddress
			SubnetMask    = $IP.netmask
			TrafficDomain = $IP.td
			Type          = $IP.type
			vServer       = $IP.vserver
			MGMT          = $IP.mgmtaccess
			SNMP          = $IP.snmp
		}
	} Else {
		$IPADDRESSH += @{
			IPAddress     = $IP.ipaddress
			SubnetMask    = $IP.netmask
			Type          = $IP.type
			vServer       = $IP.vserver
			MGMT          = $IP.mgmtaccess
			SNMP          = $IP.snmp
		}
	}
}
$Params = $null
If ($TDcounter -ge 1){
	$Params = @{
		Hashtable = $IPADDRESSH
		Columns   = "IPAddress", "SubnetMask", "TrafficDomain", "Type", "vServer", "MGMT", "SNMP"
		Headers   = "IP Address", "Subnet Mask", "Traffic Domain", "Type", "vServer", "Management", "SNMP"
	}
} Else {
	$Params = @{
		Hashtable = $IPADDRESSH
		Columns   = "IPAddress", "SubnetMask", "Type", "vServer", "MGMT", "SNMP"
		Headers   = "IP Address", "Subnet Mask", "Type", "vServer", "Management", "SNMP"
	}
}
$Table = AddWordTable @Params
FindWordDocumentEnd
#endregion IP addresses

#region Interfaces
WriteWordLine 3 0 "Interfaces"
$Interfaces = Get-vNetScalerObject -Type interface
foreach ($Interface in $Interfaces) {
	WriteWordLine 4 0 "Interface $($interface.devicename)"
	[System.Collections.Hashtable[]] $NSINTFH = @(
		@{ Description = "Description"; Value = "Value"}
		@{ Description = "Device Description"; Value = $interface.description}
		@{ Description = "Interface Type"; Value = $interface.intftype}
		@{ Description = "HA Monitoring"; Value = $interface.hamonitor}
		@{ Description = "State"; Value = $interface.state}
		@{ Description = "Auto Negotiate"; Value = $interface.autoneg}
		@{ Description = "HA Heartbeats"; Value = $interface.haheartbeat}
		@{ Description = "MAC Address"; Value = $interface.mac}
		@{ Description = "Tag All VLANs"; Value = $interface.tagall}
	)
	$Params = $null
	$Params = @{
		Hashtable = $NSINTFH
		Columns   = "Description", "Value"
	}
	$Table = AddWordTable @Params -List
	FindWordDocumentEnd
}
#endregion Interfaces

#region Channels
If ((Get-vNetScalerObjectCount -Type channel).__count -ge 1) { 
	WriteWordLine 3 0 "Channels"
	$Channels = Get-vNetScalerObject -Type channel
    [System.Collections.Hashtable[]] $CHANH = @()
    foreach ($Channel in $Channels) {
        $CHANH += @{
            CHANNEL = $channel.devicename
            Alias   = $channel.ifalias
            HA      = $channel.hamonitor
            State   = $channel.state
            Speed   = $channel.reqspeed
            Tagall  = $channel.tagall
            MTU     = $channel.mtu
        }
    }
    If ($CHANH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $CHANH
            Columns   = "CHANNEL", "Alias", "HA", "State", "Speed", "Tagall", "MTU"
            Headers   = "Channel", "Alias", "HA Monitoring", "State", "Speed", "Tag all vLAN", "MTU"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Channels

#region Traffic Domains
If ($TDcounter -ge 1) {
	WriteWordLine 3 0 "Traffic Domains"
	$TDs = Get-vNetScalerObject -Type nstrafficdomain
    [System.Collections.Hashtable[]] $TDSH = @()
    foreach ($TD in $TDs) {
        $TDSH += @{
            ID    = $TD.td
            Alias = $TD.aliasname
            vmac  = $TD.vmac
            State = $TD.state
        }
    }
    If ($TDSH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $TDSH
            Columns   = "ID", "Alias", "vmac", "State"
            Headers   = "Traffic Domain ID", "Traffic Domain Alias", "Traffic Domain vmac", "State"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Traffic Domains

#region vLAN
If ((Get-vNetScalerObjectCount -Type vlan).__count -ge 1) {
	WriteWordLine 3 0 "vLANs"
	$VLANS = Get-vNetScalerObject -Type vlan
    foreach ($VLAN in $VLANS) {
		WriteWordLine 4 0 "vLAN $($VLAN.id)"
        [System.Collections.Hashtable[]] $NSVLANH = @(
            @{ Description = "Description"; Value = "Value"}
            @{ Description = "VLAN ID"; Value = $VLAN.id}
            If ($VLAN.id -eq "1") {
                @{ Description = "VLAN Name"; Value = "Default"}
            }
            Else {
                @{ Description = "VLAN Name"; Value = $VLAN.aliasname}
            }
            @{ Description = "Bound Interfaces"; Value = $VLAN.ifaces}
            @{ Description = "Tagged Interfaces"; Value = $VLAN.tagged}
            @{ Description = "RNAT"; Value = $VLAN.rnat}
            @{ Description = "VXLAN"; Value = $VLAN.vxlan}
        )
        $Params = $null
        $Params = @{
            Hashtable = $NSVLANH
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
    }
}
#endregion vLAN

#region VXLAN
If ((Get-vNetScalerObjectCount -Type vxlan).__count -ge 1) {
	WriteWordLine 3 0 "VXLANs"
	$VXLANS = Get-vNetScalerObject -Type vxlan
    foreach ($VXLAN in $VXLANS) {
		WriteWordLine 4 0 "VXLAN $($VXLAN.id)"
        [System.Collections.Hashtable[]] $NSVXLANH = @(
            @{ Description = "Description"; Value = "Value"}
            @{ Description = "VXLAN ID"; Value = $VXLAN.id}
            @{ Description = "VXLAN Port"; Value = $VXLAN.port}
            @{ Description = "Dynamic Routing"; Value = $VXLAN.dynamicrouting}
            @{ Description = "IPv6 Dynamic Routing"; Value = $VLAN.ipv6dynamicrouting}
            @{ Description = "Inner VLAN Tagging"; Value = $VLAN.innervlantagging}
            @{ Description = "Encapsulation Type"; Value = $VLAN.type}
        )
        $Params = $null
        $Params = @{
            Hashtable = $NSVXLANH
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd       
    }
}
#endregion VXLAN

#region ACL
$nssimpleaclCounter = (Get-vNetScalerObjectCount -Type nssimpleacl).__count
$nsaclCounter = (Get-vNetScalerObjectCount -Type nsacl).__count
If ($nssimpleaclCounter -ge 1 -or $nsaclCounter -ge 1) {
	WriteWordLine 3 0 "ACL Configuration"
}

#region Simple ACL
If ($nssimpleaclCounter -ge 1) {
	WriteWordLine 4 0 "Simple ACL"
	$nssimpleacls = Get-vNetScalerObject -Type nssimpleacl
    [System.Collections.Hashtable[]] $nssimpleaclH = @()
    foreach ($nssimpleacl in $nssimpleacls) {        
        $nssimpleaclH += @{
            ACLNAME  = $nssimpleacl.aclname
            ACTION   = $nssimpleacl.aclaction
            SOURCEIP = $nssimpleacl.srcip
            DESTPORT = $nssimpleacl.destport
            PROT     = $nssimpleacl.protocol
            TD       = $nssimpleacl.td
        }
    }
    If ($nssimpleaclH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $nssimpleaclH
            Columns   = "ACLNAME", "ACTION", "SOURCEIP", "DESTPORT", "PROT", "TD"
            Headers   = "ACL Name", "Action", "Source IP", "Destination Port", "Protocol", "Traffic Domain"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Simple ACL IPv4

#region Extended ACL
If ($nsaclCounter -ge 1) {
	WriteWordLine 4 0 "Extended ACL"    
	$nsacls = Get-vNetScalerObject -Type nsacl
    [System.Collections.Hashtable[]] $nsaclH = @()
    foreach ($nsacl in $nsacls) {        
        $nsaclH += @{
            ACLNAME  = $nsacl.aclname
            ACTION   = $nsacl.aclaction
            SOURCEIP = $nsacl.srcipval
            TD       = $nsacl.td
        }
    }
    If ($nsaclH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $nsaclH
            Columns   = "ACLNAME", "ACTION", "SOURCEIP", "TD"
            Headers   = "ACL Name", "Action", "Source IP", "Traffic Domain"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Extended ACL IPv4
#endregion ACL

#region PBRs
$NSPBRCounter = (Get-vNetScalerObjectCount -Type nspbr).__count
$NSPBR6Counter = (Get-vNetScalerObjectCount -Type nspbr6).__count
If ($NSPBRCounter -ge 1 -or $NSPBR6Counter -ge 1){WriteWordLine 2 0 "Policy Based Routes"}
If ($NSPBRCounter -ge 1) {
	WriteWordLine 3 0 "Policy Based Routes"
	$NSPBRS = Get-vNetScalerObject -Type nspbr
    foreach ($NSPBR in $NSPBRS) {
        [System.Collections.Hashtable[]] $NSPBRH = @(
            @{ Description = "Description"; Value = "Value"}
            @{ Description = "Name"; Value = $NSPBR.name}
            @{ Description = "Action"; Value = $NSPBR.action}
            @{ Description = "State"; Value = $NSPBR.state}
            @{ Description = "Traffic Domain"; Value = $NSPBR.td}
            @{ Description = "Source MAC Address"; Value = $NSPBR.srcmac}
            @{ Description = "Source MAC Address Mask"; Value = $NSPBR.srcmacmask}
            @{ Description = "Protocol"; Value = $NSPBR.protocol}
            @{ Description = "Protocol Number"; Value = $NSPBR.protocolnumber}
            @{ Description = "Source Port"; Value = $NSPBR.srcportval}
            @{ Description = "Source Port Operator"; Value = $NSPBR.srcportop}
            @{ Description = "Destination Port"; Value = $NSPBR.destportval}
            @{ Description = "Destination Port Operator"; Value = $NSPBR.destportop}
            @{ Description = "Source IP"; Value = $NSPBR.srcipval}
            @{ Description = "Source IP Operator"; Value = $NSPBR.srcipop}
            @{ Description = "Destination IP"; Value = $NSPBR.destipval}
            @{ Description = "Destination IP Operator"; Value = $NSPBR.destipop}
            @{ Description = "VLAN"; Value = $NSPBR.vlan}
            @{ Description = "VXLAN"; Value = $NSPBR.vxlan}
            @{ Description = "Interface"; Value = $NSPBR.interface}
            @{ Description = "Priority"; Value = $NSPBR.priority}
            @{ Description = "Next Hop"; Value = $NSPBR.nexthopval}
            @{ Description = "IP Tunnel Name"; Value = $NSPBR.iptunnelname}
            @{ Description = "VXLAN"; Value = $NSPBR.vxlanvlanmap}
            @{ Description = "Monitor"; Value = $NSPBR.monitor}
        )
        $Params = $null
        $Params = @{
            Hashtable = $NSPBRH
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
    }
}

If ($NSPBR6Counter -ge 1) {
	WriteWordLine 3 0 "Policy Based V6 Routes"
	$NSPBRS6 = Get-vNetScalerObject -Type nspbr6
    foreach ($NSPBR6 in $NSPBRS6) {
        [System.Collections.Hashtable[]] $NSPBR6H = @(
            @{ Description = "Description"; Value = "Value"}
            @{ Description = "Name"; Value = $NSPBR6.name}
            @{ Description = "Action"; Value = $NSPBR6.action}
            @{ Description = "State"; Value = $NSPBR6.state}
            @{ Description = "Traffic Domain"; Value = $NSPBR6.td}
            @{ Description = "Source MAC Address"; Value = $NSPBR6.srcmac}
            @{ Description = "Source MAC Address Mask"; Value = $NSPBR6.srcmacmask}
            @{ Description = "Protocol"; Value = $NSPBR6.protocol}
            @{ Description = "Protocol Number"; Value = $NSPBR6.protocolnumber}
            @{ Description = "Source Port"; Value = $NSPBR6.srcportval}
            @{ Description = "Source Port Operator"; Value = $NSPBR6.srcportop}
            @{ Description = "Destination Port"; Value = $NSPBR6.destportval}
            @{ Description = "Destination Port Operator"; Value = $NSPBR6.destportop}
            @{ Description = "Source IP"; Value = $NSPBR6.srcipv6val}
            @{ Description = "Source IP Operator"; Value = $NSPBR6.srcipop}
            @{ Description = "Destination IP"; Value = $NSPBR6.destipv6val}
            @{ Description = "Destination IP Operator"; Value = $NSPBR6.destipop}
            @{ Description = "VLAN"; Value = $NSPBR6.vlan}
            @{ Description = "VXLAN"; Value = $NSPBR6.vxlan}
            @{ Description = "Interface"; Value = $NSPBR6.interface}
            @{ Description = "Priority"; Value = $NSPBR6.priority}
            @{ Description = "Next Hop"; Value = $NSPBR6.nexthopval}
            @{ Description = "IP Tunnel Name"; Value = $NSPBR6.iptunnelname}
            @{ Description = "VXLAN"; Value = $NSPBR6.vxlanvlanmap}
            @{ Description = "Monitor"; Value = $NSPBR6.monitor}
        )
        $Params = $null
        $Params = @{
            Hashtable = $NSPBR6H
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
    }       
}
#endregion PBRs

#region RPC Nodes
If ((Get-vNetScalerObjectCount -Type nsrpcnode).__count -ge 1) {
	WriteWordLine 3 0 "RPC Nodes"
	$rpcnodes = Get-vNetScalerObject -Type nsrpcnode
    [System.Collections.Hashtable[]] $RPCCONFIGH = @()
    foreach ($rpcnode in $rpcnodes) {
        $RPCCONFIGH += @{
            IPADDR = $rpcnode.ipaddress
            SOURCE = $rpcnode.srcip
            SECURE = $rpcnode.secure
        }
    }
    If ($RPCCONFIGH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $RPCCONFIGH
            Columns   = "IPADDR", "SOURCE", "SECURE"
            Headers   = "IP Address", "Source IP Address", "Secure"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion RPC Nodes

#region routing table
WriteWordLine 3 0 "Routing Table"
$nsroute = Get-vNetScalerObject -Type route
[System.Collections.Hashtable[]] $ROUTESH = @()
foreach ($ROUTE in $nsroute) {
	If ($TDcounter -ge 1){
		$ROUTESH += @{
			Network  = $ROUTE.network
			Subnet   = $ROUTE.netmask
			Gateway  = $ROUTE.gateway
			Distance = $ROUTE.distance
			Weight   = $ROUTE.weight
			Cost     = $ROUTE.cost
			TD       = $ROUTE.td
		}
	} Else {
		$ROUTESH += @{
			Network  = $ROUTE.network
			Subnet   = $ROUTE.netmask
			Gateway  = $ROUTE.gateway
			Distance = $ROUTE.distance
			Weight   = $ROUTE.weight
			Cost     = $ROUTE.cost
		}
	}
}
If ($ROUTESH.Length -gt 0) {
    $Params = $null
    If ($TDcounter -ge 1){
		$Params = @{
			Hashtable = $ROUTESH
			Columns   = "Network", "Subnet", "Gateway", "Distance", "Weight", "Cost", "TD"
			Headers   = "Network", "Subnet", "Gateway", "Distance", "Weight", "Cost", "Traffic Domain"
		}
    } Else {
		$Params = @{
			Hashtable = $ROUTESH
			Columns   = "Network", "Subnet", "Gateway", "Distance", "Weight", "Cost"
			Headers   = "Network", "Subnet", "Gateway", "Distance", "Weight", "Cost"
		}
	}
    $Table = AddWordTable @Params
    FindWordDocumentEnd
}
#endregion routing table

#region IP Sets
If ((Get-vNetScalerObjectCount -Type netprofile).__count -ge 1){
	WriteWordLine 3 0 "IP Sets"
	$ipsets = Get-vNetScalerObject -Type ipset
	[System.Collections.Hashtable[]] $IPSETH = @()
	foreach ($ipset in $ipsets) {
		$ipsetbindings = Get-vNetScalerObject -Type ipset_nsip_binding -Name $ipset.name
		$ipsetbindings6 = Get-vNetScalerObject -Type ipset_nsip6_binding -Name $ipset.name
		$IPSETH += @{
			NAME    = $ipset.name
			TD      = $ipset.td
			IP4		= $ipsetbindings.ipaddress -join ", "
			IP6		= $ipsetbindings6.ipaddress -join ", "
		}
	}
	If ($IPSETH.Length -gt 0) {
		$Params = $null
		$Params = @{
			Hashtable = $IPSETH
			Columns   = "NAME", "TD", "IP4", "IP6"
			Headers   = "IP Set", "Traffic Domain", "IP 4", "IP 6"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
}
#endregion IP Sets

#region Network Profiles
If ((Get-vNetScalerObjectCount -Type netprofile).__count -ge 1){
	WriteWordLine 3 0 "Network Profiles"
	$netprofiles = Get-vNetScalerObject -Type netprofile
	[System.Collections.Hashtable[]] $NETPROFILESH = @()
	foreach ($netprofile in $netprofiles) {
		$NETPROFILESH += @{
			NAME    = $netprofile.name
			TD      = $netprofile.td
			SRCIP   = $netprofile.srcip
			PERSIST = $netprofile.srcippersistency
			LSN     = $netprofile.overridelsn
		}
	}
	If ($NETPROFILESH.Length -gt 0) {
		$Params = $null
		$Params = @{
			Hashtable = $NETPROFILESH
			Columns   = "NAME", "TD", "SRCIP", "PERSIST", "LSN"
			Headers   = "Net Profile", "Traffic Domain", "Source IP", "IP Persistency", "Override LSN"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
}
#endregion Network Profiles

#region Linksets
If ((Get-vNetScalerObjectCount -Type linkset).__count -ge 1) {
	WriteWordLine 3 0 "LinkSets"
    [System.Collections.Hashtable[]] $LSH = @()
    $NSLSS = Get-vNetScalerObject -Type linkset
    foreach ($NSLS in $NSLSS) {  
        $URLLSID = $NSLS.id -replace "/", "%2F"
        $LSIFBinds = Get-vNetScalerObject -Type linkset_interface_binding -Name $URLLSID
        $LSH += @{
            LSID  = $NSLS.id
            IFNUM = $LSIFBinds.ifnum -Join ", "
        }
    }
    If ($LSH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $LSH
            Columns   = "LSID", "IFNUM"
            Headers   = "Linkset ID", "Interfaces"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Linksets

#endregion Networking
#endregion System

#region AppExpert
$selection.InsertNewPage()
WriteWordLine 1 0 "AppExpert"

#region HTTP Callouts
If ((Get-vNetScalerObjectCount -Type policyhttpcallout).__count -ge 1) {
	WriteWordLine 2 0 "HTTP Callouts"
	$callouts = Get-vNetScalerObject -Type policyhttpcallout
    foreach ($callout in $callouts) {
        $calloutname = $callout.name
        WriteWordLine 3 0 "$calloutname"
        [System.Collections.Hashtable[]] $CALLOUTH = @(
            @{ Description = "Description"; Value = "Value"}
            If (![string]::IsNullOrWhiteSpace($callout.comment)){@{ Description = "Comment"; Value = $callout.comment}}
            If (![string]::IsNullOrWhiteSpace($callout.ip)){@{ Description = "IP Address"; Value = $callout.ip}}
            If (![string]::IsNullOrWhiteSpace($callout.port)){@{ Description = "Port"; Value = $callout.port}}
            If (![string]::IsNullOrWhiteSpace($callout.vserver)){@{ Description = "Virtual Server"; Value = $callout.vserver}}
            If (![string]::IsNullOrWhiteSpace($callout.httpmethod)){@{ Description = "Request Method"; Value = $callout.httpmethod}}
            If (![string]::IsNullOrWhiteSpace($callout.hostexpr)){@{ Description = "Host Expression"; Value = $callout.hostexpr}}
            If (![string]::IsNullOrWhiteSpace($callout.urlstemexpr)){@{ Description = "URL Stem Expression"; Value = $callout.urlstemexpr}}
            If (![string]::IsNullOrWhiteSpace($callout.bodyexpr)){@{ Description = "Body Expression"; Value = $callout.bodyexpr}}
            If (![string]::IsNullOrWhiteSpace($callout.headers)){@{ Description = "Headers"; Value = $callout.headers}}
            If (![string]::IsNullOrWhiteSpace($callout.parameters)){@{ Description = "Parameters"; Value = $callout.parameters}}
            If (![string]::IsNullOrWhiteSpace($callout.scheme)){@{ Description = "Scheme"; Value = $callout.scheme}}
            If (![string]::IsNullOrWhiteSpace($callout.returntype)){@{ Description = "Return Type"; Value = $callout.returntype}}
            If (![string]::IsNullOrWhiteSpace($callout.resultexpr)){@{ Description = "Expression to extract data from response"; Value = $callout.resultexpr}}
            If (![string]::IsNullOrWhiteSpace($callout.cacheforsecs)){@{ Description = "Cache Expiration Time (seconds)"; Value = $callout.cacheforsecs}}
        )
        $Params = $null
        $Params = @{
            Hashtable = $CALLOUTH
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
	}
}
#endregion HTTP Callouts

#region Pattern Sets
WriteWordLine 2 0 "Pattern Sets"
$pattsetpolicies = Get-vNetScalerObject -Type policypatset
foreach ($patternsetpolicy in $pattsetpolicies) {
    $patternset = Get-vNetScalerObject -Type policypatset_binding -Name $patternsetpolicy.name
    WriteWordLine 3 0 "$($patternsetpolicy.name)"
    [System.Collections.Hashtable[]] $PATSETS = @()
    foreach ($patternsetentry in $patternset.policypatset_pattern_binding) {
        $PATSETS += @{
            STRING  = "$($patternsetentry.charset)"
            CHARSET = "$($patternsetentry.string)"
            INDEX   = "$($patternsetentry.index)"
        }
    } #end foreach
    $Params = $null
    $Params = @{
        Hashtable = $PATSETS
        Columns   = "STRING", "CHARSET", "INDEX"
        Headers   = "Pattern", "Charset", "Index"
    }
	$Table = AddWordTable @Params
    FindWordDocumentEnd
}
#endregion Pattern Sets

#region Data Sets
If ((Get-vNetScalerObjectCount -Type policydataset).__count -ge 1) {
	$selection.InsertNewPage()
	WriteWordLine 2 0 "Data Sets"
	$datatsetpolicies = Get-vNetScalerObject -Type policydataset
	foreach ($datasetpolicy in $datatsetpolicies) {
		$dataset = Get-vNetScalerObject -Type policydataset_binding -Name $datasetpolicy.name
		$datasetname = $datasetpolicy.name
		WriteWordLine 3 0 "$datasetname"
		[System.Collections.Hashtable[]] $DATASETS = @()
		foreach ($datasetentry in $dataset.policydataset_pattern_binding) {
			$DATASETS += @{
				VALUE = $datasetentry.string 
				INDEX = $datasetentry.index
			}
		}
		$Params = $null
		$Params = @{
			Hashtable = $DATASETS
			Columns   = "VALUE", "INDEX"
			Headers   = "Value", "Index"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
}
#endregion Data Sets

#region URL Sets

#API not working in 12.0 Beta

#endregion URL Sets

#region String Maps
If ((Get-vNetScalerObjectCount -Type policystringmap).__count -ge 1) {
	$selection.InsertNewPage()
	WriteWordLine 2 0 "String Maps"
	$stringmaps = Get-vNetScalerObject -Type policystringmap
    foreach ($stringmap in $stringmaps) {
        $stringmappatterns = Get-vNetScalerObject -Type policystringmap_pattern_binding -Name $stringmap.name
        WriteWordLine 3 0 "$($stringmap.name)"
		If (![string]::IsNullOrWhiteSpace($stringmap.comment)){WriteWordLine 0 0 "$($stringmap.comment)"}
		If ((Get-vNetScalerObjectCount -Type policystringmap_pattern_binding -Name $stringmap.name).__count -le 0){
			WriteWordLine 0 0 "This stringmap does not contain any entries"
		} Else {
			[System.Collections.Hashtable[]] $SMSETS = @()
			foreach ($stringmappattern in $stringmappatterns) {
				$SMSETS += @{
					KEY   = $stringmappattern.key
					VALUE = $stringmappattern.value -replace "\r\n","\r\n"
				}
			}
			$Params = $null
			$Params = @{
				Hashtable = $SMSETS
				Columns   = "KEY", "VALUE"
				Headers   = "Key", "Value"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
    }
}
#endregion String Maps

#region XML NameSpaces
If ((Get-vNetScalerObjectCount -Type nsxmlnamespace).__count -ge 1) {
	$selection.InsertNewPage()
	WriteWordLine 2 0 "XML Namespaces"
	$xmlnscol = Get-vNetScalerObject -Type nsxmlnamespace
    [System.Collections.Hashtable[]] $XMLNS = @()
    $tempNS = " "
    $tempDesc = " "
    foreach ($xmlnsitem in $xmlnscol) {
        If (IsNull($xmlnsitem.Namespace)) { $tempNS = " " }Else { $tempNS = $xmlnsitem.Namespace }
        If (IsNull($xmlnsitem.description)) { $tempDesc = " " }Else { $tempDesc = $xmlnsitem.description }
        $XMLNS += @{
            PREFIX = $xmlnsitem.prefix 
            NS     = $tempNS
            DESC   = $tempDesc
        }
    }
    $Params = $null
    $Params = @{
        Hashtable = $XMLNS
        Columns   = "PREFIX", "NS", "DESC"
        Headers   = "Prefix", "Namespace", "Description"
    }
    $Table = AddWordTable @Params
    FindWordDocumentEnd
}
#endregion XML NameSpaces

#region Location
$nslocsCount = (Get-vNetScalerObjectCount -Type location).__count
$nslocdbsCount = (Get-vNetScalerObjectCount -Type locationfile).__count
If ($nslocsCount -ge 1 -and $nslocdbsCount -ge 2){$selection.InsertNewPage();WriteWordLine 2 0 "Location"}

#region Custom Location Entries
If ($nslocsCount -ge 1){
	WriteWordLine 3 0 "Custom Location Entries"
	$nslocs = Get-vNetScalerObject -Type location
    $LOCSH = $null    
    [System.Collections.Hashtable[]] $LOCSH = @()
    foreach ($nsloc in $nslocs) {
        If (IsNull($nsloc.preferredlocation)) { $locpreferredlocation = "" } Else { $locpreferredlocation = $nsloc.prefferedlocation }
        If (IsNull($nsloc.longitude)) { $loclongitude = "" } Else { $loclongitude = $nsloc.longitude }
        If (IsNull($nsloc.latitude)) { $loclatitude = "" } Else { $loclatitude = $nsloc.latitude }
        $LOCSH += @{
            IPFrom    = $nsloc.ipfrom
            IPTo      = $nsloc.ipto
            Preferred = $locpreferredlocation
            Longitude = $loclongitude
            Latitude  = $loclatitude
        }
    }
    If ($LOCSH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $LOCSH
            Columns   = "IPFrom", "IPTo", "Preferred", "Longitude", "Latitude"
            Headers   = "From IP Address", "To IP Address", "Location Name", "Longitude", "Latitude"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion Custom Location Entries

#region Location Database
If ($nslocdbsCount -ge 2) {
	WriteWordLine 3 0 "Location Database"
	$nslocdbs = Get-vNetScalerObject -Type locationfile
    [System.Collections.Hashtable[]] $LOCDBSH = @()
    foreach ($nslocdb in $nslocdbs) {
		If (![string]::IsNullOrWhiteSpace($nslocdb.Locationfile)){
			$LOCDBSH += @{
				LocationFile = $nslocdb.Locationfile
				Format       = $nslocdb.format
			}
		}
    }
    If ($LOCDBSH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $LOCDBSH
            Columns   = "Locationfile", "Format"
            Headers   = "Location File", "Format"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd   
    }
}
#endregion Location Database
#endregion Location

#region NS Variables
If ((Get-vNetScalerObjectCount -Type nsvariable).__count -ge 1) {
	WriteWordLine 2 0 "NS Variables"
	$nsvars = Get-vNetScalerObject -Type nsvariable
    foreach ($nsvar in $nsvars) {
        WriteWordLine 2 0 "$($nsvar.name)"
        [System.Collections.Hashtable[]] $NSVARH = @(
            @{ Description = "Description"; Value = "Value"}
            @{ Description = "Type"; Value = $nsvar.type}
            @{ Description = "Scope"; Value = $nsvar.scope}
            @{ Description = "If Full"; Value = $nsvar.iffull}
            @{ Description = "If Value is too big"; Value = $nsvar.ifvaluetoobig}
            @{ Description = "If No Value"; Value = $nsvar.ifnovalue}
            @{ Description = "Expires"; Value = $nsvar.expires}
            @{ Description = "Init Value"; Value = $nsvar.init}
            @{ Description = "Comment"; Value = $nsvar.comment}
        )
        $Params = $null
        $Params = @{
            Hashtable = $NSVARH
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
    }
}
#endregion NS Variables

#region NS Assignments
If ((Get-vNetScalerObjectCount -Type nsassignment).__count -ge 1) {
	WriteWordLine 2 0 "NS Assignments"
	$nsasses = Get-vNetScalerObject -Type nsassignment
    foreach ($nsass in $nsasses) {
        $nsassname = $nsass.name
        WriteWordLine 2 0 "$nsassname"       
        [System.Collections.Hashtable[]] $NSASSH = @(
            @{ Description = "Description"; Value = "Value"}
            @{ Description = "Variable"; Value = $nsass.variable}
            @{ Description = "Set"; Value = $nsass.set}
            @{ Description = "Add"; Value = $nsass.add}
            @{ Description = "Subtract"; Value = $nsass.sub}
            @{ Description = "Append"; Value = $nsass.append}
            @{ Description = "Clear"; Value = $nsass.clear}
            @{ Description = "Hits"; Value = $nsass.hits}
            @{ Description = "Comment"; Value = $nsass.comment}
        )
        $Params = $null
        $Params = @{
            Hashtable = $NSASSH
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
	}
}
#endregion NS Assignments

#region Policy Extensions

#placeholder

#endregion Policy Extensions

#region Expressions
$selection.InsertNewPage()
WriteWordLine 2 0 "Expressions"

#region Classic Expressions
WriteWordLine 3 0 "Classic Expressions"
$policyexpressions = Get-vNetScalerObject -Type policyexpression
[System.Collections.Hashtable[]] $CLASSICEXPH = @()
$tempNS = " "
$tempCSM = " "
foreach ($policyexpression in $policyexpressions) {
    $tempCSM = " "
    $tempVal = " "
    $tempComment = " "
	If ($policyexpression.isdefault -eq "False" -and $policyexpression.type1 -eq "CLASSIC") {
		If (IsNull($policyexpression.value)) { $tempVal = " " }Else { $tempVal = $policyexpression.value }
		If (IsNull($policyexpression.clientsecuritymessage)) { $tempCSM = " " }Else { $tempCSM = $policyexpression.clientsecuritymessage }
		If (IsNull($policyexpression.comment)) { $tempComment = " " }Else { $tempComment = $policyexpression.comment }
		$CLASSICEXPH += @{
			NAME    = $policyexpression.name 
			VALUE   = $tempVal
			CSM     = $tempCSM
			COMMENT = $tempComment
		}
    }
}
$Params = $null
$Params = @{
    Hashtable = $CLASSICEXPH
    Columns   = "NAME", "VALUE", "CSM", "COMMENT"
    Headers   = "Name", "Value", "Client Security Message", "Comment"
}
$Table = AddWordTable @Params
FindWordDocumentEnd
#endregion Classic Expressions

#region Advanced Expressions
WriteWordLine 3 0 "Advanced Expressions"
$policyexpressions = Get-vNetScalerObject -Type policyexpression
[System.Collections.Hashtable[]] $ADVEXPH = @()
foreach ($policyexpression in $policyexpressions) {
    $tempVal = " "
    $tempComment = " "
	If ($policyexpression.isdefault -eq "False" -and $policyexpression.type1 -eq "ADVANCED") {
		If ([string]::IsNullOrWhiteSpace($policyexpression.value)) { $tempVal = " " }Else { $tempVal = $policyexpression.value }
		If ([string]::IsNullOrWhiteSpace($policyexpression.comment)) { $tempComment = " " }Else { $tempComment = $policyexpression.comment }
		$ADVEXPH += @{
			NAME    = $policyexpression.name 
			VALUE   = $tempVal
			COMMENT = $tempComment
		}
	}
}
$Params = $null
$Params = @{
    Hashtable = $ADVEXPH
    Columns   = "NAME", "VALUE", "COMMENT"
    Headers   = "Name", "Value", "Comment"
}
$Table = AddWordTable @Params
FindWordDocumentEnd
#endregion Classic Expressions
#endregion Expressions

#region Rate Limiting
$selection.InsertNewPage()
WriteWordLine 2 0 "Rate Limiting"

#region selectors
If ((Get-vNetScalerObjectCount -Type streamselector).__count -ge 1) {
	WriteWordLine 3 0 "Selectors"
	$streamselectors = Get-vNetScalerObject -Type streamselector
	[System.Collections.Hashtable[]] $STREAMSELH = @()
	foreach ($streamselector in $streamselectors) {
		$STREAMSELH += @{ 
			NAME = $streamselector.name 
			RULE = $streamselector.rule -join ", "
		}
	} 
	$Params = $null
	$Params = @{
		Hashtable = $STREAMSELH
		Columns   = "NAME", "RULE"
		Headers   = "Name", "Expressions"
	}
	$Table = AddWordTable @Params
	FindWordDocumentEnd
}
#endregion selectors

#region rate limit identifiers
If ((Get-vNetScalerObjectCount -Type nslimitidentifier).__count -ge 1) { 
	WriteWordLine 3 0 "Limit Identifiers"
	$limitidentifiers = Get-vNetScalerObject -Type nslimitidentifier
	foreach ($limitidentifier in $limitidentifiers) {
		$limitname = $limitidentifier.limitidentifier
		WriteWordLine 4 0 "$limitname"
		[System.Collections.Hashtable[]] $NSLIMITH = @(
			@{ Description = "Description"; Value = "Value"}
			@{ Description = "Threshold"; Value = $limitidentifier.threshold}
			@{ Description = "Timeslice"; Value = $limitidentifier.timeslice}
			@{ Description = "Mode"; Value = $limitidentifier.mode}
			@{ Description = "Limit Type"; Value = $limitidentifier.limittype}
			@{ Description = "Selector"; Value = $limitidentifier.selectorname}
			@{ Description = "Bandwidth (Kbps)"; Value = $limitidentifier.bandwidth}
			@{ Description = "Traps"; Value = $limitidentifier.trapsintimeslice}
		)
		$Params = $null
		$Params = @{
			Hashtable = $NSLIMITH
			Columns   = "Description", "Value"
		}
		$Table = AddWordTable @Params -List
		FindWordDocumentEnd
	}
}
#endregion rate limit identifiers
#endregion rate limiting

#region Action Analytics
$selection.InsertNewPage()
WriteWordLine 2 0 "Action Analytics"

#region selectors
If ((Get-vNetScalerObjectCount -Type streamselector).__count -ge 1) {
	WriteWordLine 3 0 "Selectors"
	$streamselectors = Get-vNetScalerObject -Type streamselector
    [System.Collections.Hashtable[]] $STREAMSELH = @()
    foreach ($streamselector in $streamselectors) {
        $STREAMSELH += @{ 
            NAME = $streamselector.name 
            RULE = $streamselector.rule -join ", "
        }
    } 
    $Params = $null
    $Params = @{
        Hashtable = $STREAMSELH
        Columns   = "NAME", "RULE"
        Headers   = "Name", "Expressions"
    }
    $Table = AddWordTable @Params
    FindWordDocumentEnd
}
#endregion selectors

#region stream identifiers
If ((Get-vNetScalerObjectCount -Type streamidentifier).__count -ge 1) {
	WriteWordLine 3 0 "Stream Identifiers"
	$streamidentifiers = Get-vNetScalerObject -Type streamidentifier
    foreach ($streamidentifier in $streamidentifiers) {
        $streamname = $streamidentifier.name
        WriteWordLine 4 0 "$streamname"
        [System.Collections.Hashtable[]] $STREAMIDH = @(
            @{ Description = "Description"; Value = "Value"}
            @{ Description = "Selector Name"; Value = $streamidentifier.selectorname}
            @{ Description = "Interval"; Value = $streamidentifier.interval}
            @{ Description = "Samples"; Value = $streamidentifier.samplecount}
            @{ Description = "Sort"; Value = $streamidentifier.sort}
            @{ Description = "SNMP Traps"; Value = $streamidentifier.snmptrap}
            @{ Description = "AppFlow Logging"; Value = $streamidentifier.appflowlog}
            @{ Description = "Track Transactions"; Value = $streamidentifier.tracktransactions}
            @{ Description = "Maximum Transactions Threshold"; Value = $streamidentifier.maxtransactionthreshold}
            @{ Description = "Minimum Transactions Threshold"; Value = $streamidentifier.mintransactionthreshold}
            @{ Description = "Acceptance Threshold"; Value = $streamidentifier.acceptancethreshold}
            @{ Description = "Breach Threshold"; Value = $streamidentifier.breachthreshold}
        )
        $Params = $null
        $Params = @{
            Hashtable = $STREAMIDH
            Columns   = "Description", "Value"
        }
        $Table = AddWordTable @Params -List
        FindWordDocumentEnd
	}
}
#endregion stream identifiers
#endregion Action Analytics

#region AppQoE
If ($FEATAppQoE -eq "Enabled"){
	$selection.InsertNewPage()
	WriteWordLine 2 0 "AppQoE"
	
	#region AppQoE Parameters
	WriteWordLine 3 0 "AppQoE Paramters"
	$appqoeparams = Get-vNetScalerObject -Type appqoeparameter
	[System.Collections.Hashtable[]] $APPQOEPARAMH = @(
		@{ Description = "Description"; Value = "Value"}
		@{ Description = "Session Life (Secs)"; Value = $appqoeparams.sessionlife}
		@{ Description = "Average Waiting Client"; Value = $appqoeparams.avgwaitingclient}
		@{ Description = "Alternate Response Bandwidth Limit (Mbps)"; Value = $appqoeparams.maxaltrespbandwidth}
		@{ Description = "DOS Attack Threshold"; Value = $appqoeparams.dosattackthresh}
	)
	$Params = $null
	$Params = @{
		Hashtable = $APPQOEPARAMH
		Columns   = "Description", "Value"
	}
	$Table = AddWordTable @Params -List
	FindWordDocumentEnd
	#endregion AppQoE Parameters

	#region AppQoE Policies
	If ((Get-vNetScalerObjectCount -Type appqoepolicy).__count -ge 1) {
		WriteWordLine 3 0 "AppQoE Policies"
		$appqoepols = Get-vNetScalerObject -Type appqoepolicy
		[System.Collections.Hashtable[]] $APPQOEPOLH = @()
		foreach ($appqoepol in $appqoepols) {
			$APPQOEPOLH += @{ 
				NAME   = $appqoepol.name 
				RULE   = $appqoepol.rule
				ACTION = $appqoepol.action
			}
		} 
		$Params = $null
		$Params = @{
			Hashtable = $APPQOEPOLH
			Columns   = "NAME", "RULE", "ACTION"
			Headers   = "Name", "Expression", "Action"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
	#endregion AppQoE Policies

	#region AppQoE Actions
	If ((Get-vNetScalerObjectCount -Type appqoeaction).__count -ge 1) { 
		WriteWordLine 3 0 "AppQoE Actions"
		$appqoeactions = Get-vNetScalerObject -Type appqoeaction
		foreach ($appqoeaction in $appqoeactions) {
			$actionname = $appqoeaction.name
			WriteWordLine 4 0 "$actionname"
			[System.Collections.Hashtable[]] $QOEACTH = @(
				@{ Description = "Description"; Value = "Value"}
				@{ Description = "Priority"; Value = $appqoeaction.priority}
				@{ Description = "Action Type"; Value = $appqoeaction.respondwith}
				@{ Description = "Custom File"; Value = $appqoeaction.customfile}
				@{ Description = "Priority"; Value = $appqoeaction.priority}
				@{ Description = "Alternative Content Server Name"; Value = $appqoeaction.altcontentsvcname}
				@{ Description = "Alternate Content Path"; Value = $appqoeaction.altcontentpath}
				@{ Description = "Policy Queue Depth"; Value = $appqoeaction.polqdepth}
				@{ Description = "Queue Depth"; Value = $appqoeaction.priqdepth}
				@{ Description = "Maximum Connections"; Value = $appqoeaction.maxconn}
				@{ Description = "Delay (microseconds)"; Value = $appqoeaction.priority}
				@{ Description = "DOS Expression"; Value = $appqoeaction.dostrigexpression}
				@{ Description = "DOS Action"; Value = $appqoeaction.dosaction}
				@{ Description = "TCP Profile"; Value = $appqoeaction.tcpprofile}
			)
			$Params = $null
			$Params = @{
				Hashtable = $QOEACTH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
		} #end foreach
	}
	#endregion AppQoE Actions
}
#endregion AppQoE

#region Rewrite
If ($FEATREWRITE -eq "Enabled"){
	$selection.InsertNewPage()
	WriteWordLine 2 0 "Rewrite"
	
	#region Rewrite Policy Labels
	If ((Get-vNetScalerObjectCount -Type rewritepolicylabel).__count -ge 1) {
		WriteWordLine 3 0 "Rewrite Policy Labels"
		$rwpollbls = Get-vNetScalerObject -Type rewritepolicylabel
		foreach ($rwpollbl in $rwpollbls) {
			If ($rwpollbl.isdefault -ne "True" -and $rwpollbl.labelname -notmatch "ns_cvpn_v2_"){
				[System.Collections.Hashtable[]] $RWPLBLH = @()
				WriteWordLine 4 0 "$($rwpollbl.labelname) ($($rwpollbl.transform))"
				$rwpollblbindings = Get-vNetScalerObject -Type rewritepolicylabel_rewritepolicy_binding -Name $rwpollbl.labelname
				foreach ($rwplbl in $rwpollblbindings){
					$RWPLBLH += @{
						priority	= $rwplbl.priority
						policy		= $rwplbl.policyname
						gotoexp		= $rwplbl.gotopriorityexpression
						invoke		= $rwplbl.invoke
					}
				}
				If ($RWPLBLH.Length -gt 0) {
					$Params = $null
					$Params = @{
						Hashtable = $RWPLBLH
						Columns   = "priority", "policy", "gotoexp", "invoke"
						Headers   = "Priority", "Policy", "Goto Expression", "Invoke"
					}
					$Table = AddWordTable @Params
					FindWordDocumentEnd
				}
			}
		}
	}
	#endregion Rewrite Policy Labels
	
	#region Rewrite Policies
	If ((Get-vNetScalerObjectCount -Type rewritepolicy).__count -ge 1) {
		WriteWordLine 3 0 "Rewrite Policies"
		$rewritepolicies = Get-vNetScalerObject -Type rewritepolicy
		[System.Collections.Hashtable[]] $RWPPOLH = @()
		foreach ($rewritepolicy in $rewritepolicies) {
			If ($rewritepolicy.isdefault -ne "True"){
				$RWPPOLH += @{
					RWPOLNAME = $rewritepolicy.name
					RULE      = $rewritepolicy.rule
					ACTION    = $rewritepolicy.action
				}
			}
		}
		If ($RWPPOLH.Length -gt 0) {
			$Params = $null
			$Params = @{
				Hashtable = $RWPPOLH
				Columns   = "RWPOLNAME", "RULE", "ACTION"
				Headers   = "Rewrite Policy", "Rule", "Action"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion Rewrite Policies

	#region Rewrite Actions
	If ((Get-vNetScalerObjectCount -Type rewriteaction).__count -ge 1) {
		WriteWordLine 3 0 "Rewrite Actions"
		$rewriteactions = Get-vNetScalerObject -Type rewriteaction
		[System.Collections.Hashtable[]] $RWACTH = @()
		foreach ($rewriteaction in $rewriteactions) {
			If ($rewriteaction.isdefault -ne "True"){
				$RWACTH += @{ 
					REWRITE = $rewriteaction.name 
					Type    = $rewriteaction.type
					Target  = $rewriteaction.target -replace "\r\n","\r\n" -replace "`n",""
					STRING  = $rewriteaction.stringbuilderexpr
				}
			}
		}
		If ($RWACTH.Length -gt 0) {
			$Params = $null
			$Params = @{
				Hashtable = $RWACTH
				Columns   = "REWRITE", "Type", "Target", "STRING"
				Headers   = "Rewrite Policy", "Type", "Target", "String"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion Rewrite Actions
}
#endregion Rewrite

#region Responder
If ($FEATRESPONDER -eq "Enabled"){
	$selection.InsertNewPage()
	WriteWordLine 2 0 "Responder"
	
	#region Responder Policy Labels
	If ((Get-vNetScalerObjectCount -Type responderpolicylabel).__count -ge 1) {
		WriteWordLine 3 0 "Responder Policy Labels"
		$respollbls = Get-vNetScalerObject -Type responderpolicylabel
		foreach ($respollbl in $respollbls) {
			WriteWordLine 4 0 "$($respollbl.labelname) ($($respollbl.transform))"
			[System.Collections.Hashtable[]] $RSPPOLLBLH = @()
			$respollblbindings = Get-vNetScalerObject -Type responderpolicylabel_responderpolicy_binding -Name $respollbl.labelname
			foreach ($respollblbinding in $respollblbindings){
				$RSPPOLLBLH += @{
					priority	= $respollblbinding.priority
					POLNAME		= $respollblbinding.policyname
					GotoExp		= $respollblbinding.gotopriorityexpression
					invoke		= $respollblbinding.invoke
				}
			}
			If ($RSPPOLLBLH.Length -gt 0) {
				$Params = $null
				$Params = @{
					Hashtable = $RSPPOLLBLH
					Columns   = "priority", "POLNAME", "GoToExp", "invoke"
					Headers   = "Priority", "Policy", "Goto Expression", "Invoke"
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
		}
	}
	#endregion Rewrite Policy Labels

	#region Responder Policies
	If ((Get-vNetScalerObjectCount -Type responderpolicy).__count -ge 1) {
		WriteWordLine 3 0 "Responder Policies"
		$responderpolicies = Get-vNetScalerObject -Type responderpolicy
		[System.Collections.Hashtable[]] $RESPPOL = @()
		foreach ($responderpolicy in $responderpolicies) {
			If ($responderpolicy.builtin -notmatch "IMMUTABLE"){
				$RESPPOL += @{
					RESPOLNAME = $responderpolicy.name
					RULE       = $responderpolicy.rule
					ACTION     = $responderpolicy.action
				}
			}
		}
		If ($RESPPOL.Length -gt 0) {
			$Params = $null
			$Params = @{
				Hashtable = $RESPPOL
				Columns   = "RESPOLNAME", "RULE", "ACTION"
				Headers   = "Responder Policy", "Rule", "Action"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion Responder Policies

	#region Responder Actions
	If ((Get-vNetScalerObjectCount -Type responderaction).__count -ge 1) {
		WriteWordLine 3 0 "Responder Action"
		$resacts = Get-vNetScalerObject -Type responderaction
		[System.Collections.Hashtable[]] $RESACTH = @()
		foreach ($resact in $resacts) {
			if ($resact.isdefault -ne "True"){
				$RESACTH += @{ 
					Responder = $resact.name 
					Type      = $resact.type
					Target    = $resact.target -replace "\r\n","\r\n" -replace "`n",""
					RESPST    = $resact.responsestatuscode
				}
			}
		}
		If ($RESACTH.Length -gt 0) {
			$Params = $null
			$Params = @{
				Hashtable = $RESACTH
				Columns   = "Responder", "Type", "Target", "RESPST"
				Headers   = "Responder Policy", "Type", "Target", "Response Status Code"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion Responder Actions

	#region Responder HTML Page Imports
	$rescontent = (Get-StringFromBase64 -Object (Get-vNetScalerFile -FileName mapping-responder -FileLocation "/var/download" | Select -ExpandProperty filecontent) -Encoding UTF8) -split "`n"
	#If ((Get-vNetScalerObjectCount -Type responderhtmlpage).__count -ge 1) {
	If ($rescontent){
		WriteWordLine 3 0 "Responder HTML Page Imports"
		foreach ($respage in $rescontent) {
			$pagename = $respage.split(",")[0]
			If (![string]::IsNullOrWhiteSpace($pagename)){
				WriteWordLine 4 0 $pagename
				$pagecontent = (Get-StringFromBase64 -Object (Get-vNetScalerFile -FileName $pagename -FileLocation "/var/download/responder" | Select -ExpandProperty filecontent) -Encoding UTF8).trim() -replace "`n", "CHR(10)"
				$Params = $null
				$Params = @{
					Hashtable = @{
						Content = $pagecontent
					}
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
		}
	}
	#endregion Responder HTML Page Imports
}
#endregion Responder
#endregion AppExpert

#region Traffic Management
$selection.InsertNewPage()
WriteWordLine 1 0 "Traffic Management"

#region Load Balancers
If ($FEATLB -eq "Enabled"){
	WriteWordLine 2 0 "Load Balancing"
	
	#region LB vServers
	If ((Get-vNetScalerObjectCount -Type lbvserver).__count -ge 1) {
		WriteWordLine 3 0 "Virtual Servers"
		$lbvservers = Get-vNetScalerObject -Type lbvserver
		foreach ($LB in $lbvservers) {
			WriteWordLine 4 0 "$($LB.name)"
			[System.Collections.Hashtable[]] $LBSRVH = @(
				@{ Description = "Description"; Value = "Configuration" }
				@{ Description = "Current State"; Value = $LB.curstate}
				@{ Description = "Current Effective State"; Value = $LB.effectivestate}

				# Basic Settings
				@{ Description = "Protocol"; Value = $LB.servicetype}
				If ($LB.port -eq "0"){
					@{ Description = "IP Address Type"; Value = "Non Addressable"}
				}Else{
					@{ Description = "Port"; Value = $LB.port}
					If ([string]::IsNullOrWhiteSpace($LB.ipv46)){
						@{ Description = "IP Pattern"; Value = $LB.ippattern}
						@{ Description = "IP Mask"; Value = $LB.ipmask}
					}Else{
						@{ Description = "IP"; Value = $LB.ipv46}
					}
				}
				If ($LB.td -ne "0"){@{ Description = "Traffic Domain"; Value = $LB.td}}
				If (![string]::IsNullOrWhiteSpace($LB.ipset)){@{ Description = "IP Set"; Value = $LB.ipset}}
				If ($LB.m -ne "IP"){@{ Description = "Redirection Mode [IP]"; Value = $LB.m}}
				If ($LB.m -eq "TOS"){@{ Description = "TOS ID"; Value = $LB.tosid}}
				If (![string]::IsNullOrWhiteSpace($LB.listenpriority)){@{ Description = "Listen Priority"; Value = $LB.listenpriority}}
				If ($LB.listenpolicy -ne "NONE"){@{ Description = "Listen Policy Expression [none]"; Value = $LB.listenpolicy}}
				If (![string]::IsNullOrWhiteSpace($LB.probeprotocol)){@{ Description = "Probe Protocol"; Value = $LB.probeprotocol}}
				If (![string]::IsNullOrWhiteSpace($LB.probesuccessresponsecode)){@{ Description = "Probe Success Response Code"; Value = $LB.probesuccessresponsecode}}
				If (![string]::IsNullOrWhiteSpace($LB.probeport)){@{ Description = "Probe Port"; Value = $LB.probeport}}
				If ($LB.toggleorder -ne "ASCENDING"){@{ Description = "Toggle Order [ascending]"; Value = $LB.toggleorder}}
				If ($LB.orderthreshold -ne "0"){@{ Description = "Order Threshold [0]"; Value = $LB.orderthreshold}}
				If (![string]::IsNullOrWhiteSpace($LB.comment)){@{ Description = "Comment"; Value = $LB.comment}}
				If ($LB.rhistate -eq "ACTIVE"){@{ Description = "RHI State [passive]"; Value = $LB.rhistate}}
				If ($LB.appflowlog -eq "DISABLED"){@{ Description = "AppFlow logging [enabled]"; Value = $LB.appflowlog}}
				If ($LB.servicetype -eq "SSL"){
					If (![string]::IsNullOrWhiteSpace($LB.redirectfromport)){@{ Description = "Redirect From Port"; Value = $LB.redirectfromport}}
					If (![string]::IsNullOrWhiteSpace($LB.httpsredirecturl)){@{ Description = "HTTPS Redirect URL"; Value = $LB.httpsredirecturl}}
				}
				If ($LB.servicetype -eq "FTP" -or $LB.servicetype -eq "HTTP" -or $LB.servicetype -eq "SSL" -or $LB.servicetype -eq "TCP"){
					If ($LB.retainconnectionsoncluster -ne "NO"){@{ Description = "Retain Connections on Cluster [no]"; Value = $LB.retainconnectionsoncluster}}
				}
				If ($LB.servicetype -match "DNS"){
					If ($LB.dns64 -ne "DISABLED"){@{ Description = "DNS64 [disabled]"; Value = $LB.dns64}}
					If ($LB.recursionavailable -ne "NO"){@{ Description = "Recursion Available [no]"; Value = $LB.recursionavailable}}
					If ($LB.bypassaaaa -ne "NO"){@{ Description = "Bypass AAAA Requests [no]"; Value = $LB.bypassaaaa}}
				}

				# Method
				@{ Description = "Load Balancing Method"; Value = $LB.lbmethod}
				If (![string]::IsNullOrWhiteSpace($LB.newservicerequest) -and $LB.newservicerequest -ne "0"){@{ Description = "New Service Startup Request Rate [0]"; Value = $LB.newservicerequest}}
				If (![string]::IsNullOrWhiteSpace($LB.backuplbmethod) -and $LB.backuplbmethod -ne "ROUNDROBIN"){@{ Description = "Backup Load Balancing Method [roundrobin]"; Value = $LB.backuplbmethod}}
				If ($LB.newservicerequestunit -ne "PER_SECOND"){@{ Description = "New Service Request unit [per_second]"; Value = $LB.newservicerequestunit}}
				If (![string]::IsNullOrWhiteSpace($LB.newservicerequestincrementinterval) -and $LB.newservicerequestincrementinterval -ne "0"){@{ Description = "Increment Interval [0]"; Value = $LB.newservicerequestincrementinterval}}
				If (![string]::IsNullOrWhiteSpace($LB.netmask) -and $LB.netmask -ne "255.255.255.255"){@{ Description = "Method IPv4 Netmask [255.255.255.255]"; Value = $LB.netmask}}
				If (![string]::IsNullOrWhiteSpace($LB.v6netmasklen) -and $LB.v6netmasklen -ne "128"){@{ Description = "Method IPv6 Mask Length [128]"; Value = $LB.v6netmasklen}}
				If (![string]::IsNullOrWhiteSpace($LB.hashlength) -and $LB.hashlength -ne "80"){@{ Description = "Hash Length [80]"; Value = $LB.hashlength}}

				# Persistence
				@{ Description = "Persistence Type"; Value = $LB.persistencetype}
				If ($LB.persistencetype -ne "NONE"){
					@{ Description = "Persistence Time-out (mins)"; Value = $LB.timeout}
					If (![string]::IsNullOrWhiteSpace($LB.cookiename)){@{ Description = "Persistence Cookiename"; Value = $LB.cookiename}}
					If (![string]::IsNullOrWhiteSpace($LB.persistmask)){@{ Description = "Persistence IPv4 Netmask"; Value = $LB.persistmask}}
					If (![string]::IsNullOrWhiteSpace($LB.v6persistmasklen)){@{ Description = "Persistence IPv6 Mask Length [128]"; Value = $LB.v6persistmasklen}}
					If (![string]::IsNullOrWhiteSpace($LB.rule)){@{ Description = "Persistence Expression"; Value = $LB.rule}}
					If (![string]::IsNullOrWhiteSpace($LB.resrule)){@{ Description = "Persistence Response Expression"; Value = $LB.resrule}}
					If ($LB.persistencebackup -eq "SOURCEIP"){
						@{ Description = "Backup persistence [none]"; Value = $LB.persistencebackup}
						@{ Description = "Backup Persistence Time-out (mins) [2]"; Value = $LB.backuppersistencetimeout}
					}
				}
				
				# Traffic Settings
				If ($LB.healththreshold -ne "0"){@{ Description = "Health Threshold [0]"; Value = $LB.healththreshold}}
				@{ Description = "Client Time-out"; Value = $LB.clttimeout}
				If ($LB.minautoscalemembers -ne "0"){@{ Description = "Minimum Autoscale Members [0]"; Value = $LB.minautoscalemembers}}
				If ($LB.maxautoscalemembers -ne "0"){@{ Description = "Maximum Autoscale Members [0]"; Value = $LB.maxautoscalemembers}}
				If ($LB.insertvserveripport -eq "ON"){
					@{ Description = "Virtual Server IP Port Insertion [off]"; Value = $LB.insertvserveripport}
					If (![string]::IsNullOrWhiteSpace($LB.vipheader)){@{ Description = "Virtual Server IP Port Header [vip-header]"; Value = $LB.vipheader}}
				}
				If ($LB.icmpvsrresponse -ne "PASSIVE"){@{ Description = "ICMP Virtual Server Response [passive]"; Value = $LB.icmpvsrresponse}}
				If ($LB.cacheable -eq "YES"){@{ Description = "Route cacheable requests to a cache redirection server"; Value = $LB.cacheable}}
				If ($LB.sessionless -eq "ENABLED"){@{ Description = "Sessionless load balancing [disabled]"; Value = $LB.sessionless}}
				If ($LB.downstateflush -eq "DISABLED"){@{ Description = "Down State Flush [enabled]"; Value = $LB.downstateflush}}
				If ($LB.rtspnat -eq "ON"){@{ Description = "Use network address translation [off]"; Value = $LB.rtspnat}}
				If ($LB.redirectportrewrite -eq "ENABLED"){@{ Description = "Redirect Port Rewrite [disabled]"; Value = $LB.redirectportrewrite}}
				If ($LB.l2conn -eq "ON"){@{ Description = "Layer 2 Parameters [off]"; Value = $LB.l2conn}}
				If ($LB.skippersistency -ne "NONE"){@{ Description = "Skip Persistency [none]"; Value = $LB.skippersistency}}
				If ($LB.macmoderetainvlan -eq "ENABLED"){@{ Description = "Retain VLAN ID [disabled]"; Value = $LB.macmoderetainvlan}}
				If ($LB.trofspersistence -eq "DISABLED"){@{ Description = "Trofs Persistence [enabled]"; Value = $LB.trofspersistence}}
				
				# Protection
				If (![string]::IsNullOrWhiteSpace($LB.redirurl)){@{ Description = "Redirect URL"; Value = $LB.redirurl}}
				If (![string]::IsNullOrWhiteSpace($LB.backupvserver)){@{ Description = "Backup vServer"; Value = $LB.backupvserver}}
				If ($LB.disableprimaryondown -eq "ENABLED"){@{ Description = "Disable Primary When Down [disabled]"; Value = $LB.disableprimaryondown}}
				If ($LB.somethod -ne "NONE"){@{ Description = "Spillover Method [none]"; Value = $LB.somethod}}
				If (![string]::IsNullOrWhiteSpace($LB.sobackupaction)){@{ Description = "Spillover Backup Action"; Value = $LB.sobackupaction}}
				If ($LB.sopersistencetimeout -ne "2"){@{ Description = "Spillover Persistence Timeout (mins) [2]"; Value = $LB.sopersistencetimeout}}
				If ($LB.sopersistence -eq "ENABLED"){@{ Description = "Spillover Persistence [disabled]"; Value = $LB.sopersistence}}
				
				# Authentication
				If ($LB.authentication -eq "ON"){
					If ($LB.authn401 -eq "OFF"){
						@{ Description = "Form Based Authentication"; Value = "ON"}
						If (![string]::IsNullOrWhiteSpace($LB.authenticationhost)){@{ Description = "Authentication FQDN"; Value = $LB.authenticationhost}}
					}Else{
						@{ Description = "401 Based Authentication"; Value = "ON"}
					}
					If (![string]::IsNullOrWhiteSpace($LB.authnvsname)){@{ Description = "Authentication virtual server name"; Value = $LB.authnvsname}}
					If (![string]::IsNullOrWhiteSpace($LoadBalance.authnprofile)){@{ Description = "Authentication profile"; Value = $LoadBalance.authnprofile}}
				}
				
				# Databases
				If ($LB.servicetype -match "MSSQL"){@{ Description = "MSSQL Server Version"; Value = $LB.mssqlserverversion}}
				If ($LB.servicetype -match "ORACLE"){@{ Description = "Oracle Server Version"; Value = $LB.oracleserverversion}}
				If ($LB.servicetype -match "MYSQL"){
					@{ Description = "MySQL Protocol Version"; Value = $LB.mysqlprotocolversion}
					@{ Description = "MySQL Server Version"; Value = $LB.mysqlserverversion}
					@{ Description = "MySQL Character Set"; Value = $LB.mysqlcharacterset}
					@{ Description = "MySQL Server Capabilities"; Value = $LB.mysqlservercapabilities}
				}
				If ($LB.servicetype -match "SQL" -and $LB.dbslb -eq "ENABLED"){@{ Description = "Enable Database specific Load Balancing [disabled]"; Value = $LB.dbslb}}
				
				# Profiles
				If (![string]::IsNullOrWhiteSpace($LB.quicbridgeprofilename)){@{ Description = "QUIC BRIDGE Profile Name"; Value = $LB.quicbridgeprofilename}}
				If (![string]::IsNullOrWhiteSpace($LB.netprofile)){@{ Description = "Net Profile"; Value = $LB.netprofile}}
				If (![string]::IsNullOrWhiteSpace($LB.tcpprofilename)){@{ Description = "TCP Profile"; Value = $LB.tcpprofilename}}
				If (![string]::IsNullOrWhiteSpace($LB.lbprofilename)){@{ Description = "LB Profile"; Value = $LB.lbprofilename}}
				If (![string]::IsNullOrWhiteSpace($LB.quicprofilename)){@{ Description = "QUIC Profile Name"; Value = $LB.quicprofilename}}
				If (![string]::IsNullOrWhiteSpace($LB.httpprofilename)){@{ Description = "HTTP Profile"; Value = $LB.httpprofilename}}
				If (![string]::IsNullOrWhiteSpace($LB.dbprofilename)){@{ Description = "DB Profile"; Value = $LB.dbprofilename}}
				If (![string]::IsNullOrWhiteSpace($LB.dnsprofilename)){@{ Description = "DNS Profile Name"; Value = $LB.dnsprofilename}}
				If (![string]::IsNullOrWhiteSpace($LB.adfsproxyprofile)){@{ Description = "adfsProxy Profile Name"; Value = $LB.adfsproxyprofile}}
				
				# Push
				If ($LB.push -eq "ENABLED"){
					@{ Description = "Push"; Value = $LB.push}
					If (![string]::IsNullOrWhiteSpace($LB.pushmulticlients)){@{ Description = "Push Multiple Clients"; Value = $LB.pushmulticlients}}
					If (![string]::IsNullOrWhiteSpace($LB.pushvserver)){@{ Description = "Push Virtual Server"; Value = $LB.pushvserver}}
					If (![string]::IsNullOrWhiteSpace($LB.pushlabel)){@{ Description = "Push Label Expression"; Value = $LB.pushlabel}}
				}
			)
			$Params = $null
			$Params = @{
				Hashtable = $LBSRVH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			
			#region services and groups
			New-BindingTable -Name $LB.Name -BindingType "lbvserver_service_binding" -BindingTypeName "Services" -Properties "servicename" -Headers "Service Name" -Style 5
			New-BindingTable -Name $LB.Name -BindingType "lbvserver_servicegroup_binding" -BindingTypeName "Service Groups" -Properties "servicename" -Headers "Service Group Name" -Style 5
			#endregion services and groups

			#region SSL Settings
			#Don't process SSL unless we are using an SSL based Service Type
			If ($LB.ServiceType -match "SSL" -or $LB.ServiceType -match "TLS" -or $LB.ServiceType -match "HTTP_QUIC" -and $LB.ServiceType -notmatch "BRIDGE") {
				New-BindingTable -Name $LB.Name -BindingType "sslvserver_sslcertkey_binding" -BindingTypeName "Certificates" -Properties "certkeyname,ca,crlcheck,snicert,ocspcheck,cleartextport" -Headers "Certificate Name,CA Certificate,CRL Checks Enabled,SNI Enabled,OCSP Enabled,Clear Text Port" -Style 5
				New-SSLSettings $LB.Name sslvserver 5
			}
			#endregion SSL Settings

			#region LB policies
			New-BindingTable -Name $LB.name -BindingType "lbvserver_responderpolicy_binding" -BindingTypeName "Responder Policies" -Properties "priority,policyname,gotopriorityexpression" -Headers "Priority,Policy Name,Go To Expression" -Style 5
			New-BindingTable -Name $LB.name -BindingType "lbvserver_rewritepolicy_binding" -BindingTypeName "Rewrite Policies" -Properties "priority,policyname,bindpoint,gotopriorityexpression" -Headers "Priority,Policy Name,Bindpoint,Go To Expression" -Style 5
			New-BindingTable -Name $LB.name -BindingType "lbvserver_tmtrafficpolicy_binding" -BindingTypeName "Traffic Policies" -Properties "priority,policyname" -Headers "Priority,Policy Name" -Style 5
			New-BindingTable -Name $LB.name -BindingType "lbvserver_analyticsprofile_binding" -BindingTypeName "Analytics Policies" -Properties "analyticsprofile" -Headers "Analytics Profile" -Style 5
			New-BindingTable -Name $LB.name -BindingType "lbvserver_cachepolicy_binding" -BindingTypeName "Cache Policies" -Properties "priority,policyname,bindpoint,gotopriorityexpression" -Headers "Priority,Policy Name,BindPoint,Go To Expression" -Style 5
			#endregion LB policies

			$selection.InsertNewPage()
		}
	}
	#endregion LB vServers
	
	#region Services
	If ((Get-vNetScalerObjectCount -Type service).__count -ge 1) {
		WriteWordLine 3 0 "Services"
		$services = Get-vNetScalerObject -Type service
		foreach ($svc in $Services) {
			WriteWordLine 4 0 "$($svc.name)"
			[System.Collections.Hashtable[]] $AdvancedConfiguration = @(
				@{ Description = "Description"; Value = "Configuration" }
				@{ Description = "Current state of the service"; Value = $svc.srvstate }
				
				# Basic Settings
				@{ Description = "Server"; Value = $svc.ipaddress}
				@{ Description = "Protocol"; Value = $svc.servicetype}
				@{ Description = "Port"; Value = $svc.port}
				If ($svc.td -ne "0"){@{ Description = "Traffic Domain"; Value = $svc.td}}
				If (![string]::IsNullOrWhiteSpace($svc.hashid)){@{ Description = "Hash ID"; Value = $svc.hashid}}
				If ($svc.customserverid -ne "NONE"){@{ Description = "Server ID [none]"; Value = $svc.customserverid}}
				If (![string]::IsNullOrWhiteSpace($svc.cleartextport)){@{ Description = "Clear Text Port"; Value = $svc.cleartextport}}
				If (![string]::IsNullOrWhiteSpace($svc.cachetype)){@{ Description = "Cache Type"; Value = $svc.cachetype}}
				If ($svc.cacheable -eq "YES"){@{ Description = "Cacheable [no]"; Value = $svc.cacheable}}
				If ($svc.state -eq "DISABLED"){@{ Description = "State [enabled]"; Value = $svc.state }}
				If ($svc.healthmonitor -eq "NO"){@{ Description = "Health Monitoring [yes]"; Value = $svc.healthmonitor }}
				If ($svc.appflowlog -eq "DISABLED"){@{ Description = "AppFlow Logging [enabled]"; Value = $svc.appflowlog }}
				If (![string]::IsNullOrWhiteSpace($svc.comment)){@{ Description = "Comments"; Value = $svc.comment}}
				If ($svc.monconnectionclose -ne "NONE"){@{ Description = "Monitoring Connection Close Bit [none]"; Value = $svc.monconnectionclose }}
				
				# Settings
				If ($svc.sp -eq "ON"){@{ Description = "Surge protection [off]"; Value = $svc.sp }}
				If ($svc.useproxyport -eq "OFF"){@{ Description = "Use the proxy port [on]"; Value = $svc.useproxyport }}
				If ($svc.downstateflush -eq "DISABLED"){@{ Description = "Down State Flush [enabled]"; Value = $svc.downstateflush }}
				If ($svc.accessdown -eq "YES"){@{ Description = "Access Down [no]"; Value = $svc.accessdown }}
				If ($svc.usip -eq "YES"){@{ Description = "Use Source IP [no]"; Value = $svc.usip }}
				If ($svc.cka -eq "YES"){@{ Description = "Client keep-alive [no]"; Value = $svc.cka }}
				If ($svc.tcpb -eq "YES"){@{ Description = "TCP Buffering [no]"; Value = $svc.tcpb }}
				If ($svc.cmp -eq "YES"){@{ Description = "Compression [no]"; Value = $svc.cmp }}
				@{ Description = "Insert Client IP Address [enabled]"; Value = $svc.cip }
				If ($svc.cip -eq "ENABLED"){@{ Description = "Name for the HTTP Header"; Value = $svc.cipheader }}
				
				# Threshold & timeouts
				If ($svc.maxbandwidth -ne "0"){@{ Description = "Maximum Bandwidth (Kbps)"; Value = $svc.maxbandwidth}}
				If ($svc.monthreshold -ne "0"){@{ Description = "Monitor Threshold"; Value = $svc.monthreshold }}
				If ($svc.maxreq -ne "0"){@{ Description = "Maximum Requests"; Value = $svc.maxreq }}
				If ($svc.maxclient -ne "0"){@{ Description = "Maximum Clients"; Value = $svc.maxclient }}
				@{ Description = "Client Time-Out"; Value = $svc.clttimeout }
				@{ Description = "Server Time-Out"; Value = $svc.svrtimeout }
				
				# Profiles
				If (![string]::IsNullOrWhiteSpace($svc.netprofile)){@{ Description = "Network Profile Name"; Value = $svc.netprofile }}
				If (![string]::IsNullOrWhiteSpace($svc.tcpprofilename)){@{ Description = "TCP Profile Name"; Value = $svc.tcpprofilename }}
				If (![string]::IsNullOrWhiteSpace($svc.httpprofilename)){@{ Description = "HTTP Profile Name"; Value = $svc.httpprofilename }}
				If (![string]::IsNullOrWhiteSpace($svc.dnsprofilename)){@{ Description = "DNS Profile Name"; Value = $svc.dnsprofilename }}
				If (![string]::IsNullOrWhiteSpace($svc.contentinspectionprofilename)){@{ Description = "Content Inspection Profile Name"; Value = $svc.contentinspectionprofilename }}

				# Databases
				If ($svc.servicetype -match "ORACLE"){@{ Description = "Oracle Server Version"; Value = $svc.oracleserverversion}}
				
				If ($svc.gslb -ne "NONE"){@{ Description = "GSLB"; Value = $svc.gslb}}
				If ($svc.pathmonitor -eq "YES"){
					@{ Description = "Path Monitoring"; Value = $svc.pathmonitor }
					@{ Description = "Individual Path monitoring"; Value = $svc.pathmonitorindv }
				}
				If (![string]::IsNullOrWhiteSpace($svc.sc)){@{ Description = "SureConnect"; Value = $svc.sc }}
				@{ Description = "RTSP session ID mapping"; Value = $svc.rtspsessionidremap }
			)
			$Params = $null
			$Params = @{
				Hashtable = $AdvancedConfiguration
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			
			New-BindingTable -Name $svc.Name -BindingType "service_lbmonitor_binding" -BindingTypeName "Monitor" -Properties "monitor_name,weight" -Headers "Monitor Name, Weight" -Style 5
			
			#region SSL Settings
			#Don't process SSL unless we are using an SSL based Service Type
			If ($svc.ServiceType -match "SSL" -or $svc.ServiceType -match "TLS" -or $svc.ServiceType -match "HTTP_QUIC" -and $svc.ServiceType -notmatch "BRIDGE") {
				New-BindingTable -Name $svc.Name -BindingType "sslservice_sslcertkey_binding" -BindingTypeName "Certificates" -Properties "certkeyname,ca,crlcheck,snicert,ocspcheck,cleartextport" -Headers "Certificate Name,CA Certificate,CRL Checks Enabled,SNI Enabled,OCSP Enabled,Clear Text Port" -Style 5
				New-SSLSettings $svc.Name sslservice 5
			}
			#endregion SSL Settings
			
			$selection.InsertNewPage()
		}
	}
	#endregion Services

	#region Service Groups
	If ((Get-vNetScalerObjectCount -Type servicegroup).__count -ge 1) {
		WriteWordLine 3 0 "Service Groups"
		$ServiceGroups = Get-vNetScalerObject -Type servicegroup
		foreach ($svcgrp in $ServiceGroups) {
			WriteWordLine 4 0 "$($svcgrp.servicegroupname)"
			[System.Collections.Hashtable[]] $AdvancedConfiguration = @(
				@{ Description = "Description"; Value = "Configuration" }
				@{ Description = "Current state of the service group"; Value = $svcgrp.ServiceGroupeffectivestate }
				@{ Description = "Protocol"; Value = $svcgrp.servicetype }
				
				# Basic Settings
				If ($svcgrp.td -ne "0"){@{ Description = "Traffic Domain [0]"; Value = $svcgrp.td }}
				If ($svcgrp.cachetype -ne "SERVER"){@{ Description = "Cache Type [server]"; Value = $svcgrp.cachetype }}
				If ($svcgrp.autodisablegraceful -eq "YES"){@{ Description = "Auto Disable Graceful [no]"; Value = $svcgrp.autodisablegraceful }}
				If ($svcgrp.autoscale -ne "DISABLED"){@{ Description = "Auto Scale Mode"; Value = $svcgrp.autoscale }}
				If (![string]::IsNullOrWhiteSpace($svcgrp.autodelayedtrofs)){@{ Description = "Auto Delayed Trofs"; Value = $svcgrp.autodelayedtrofs }}
				If (![string]::IsNullOrWhiteSpace($svcgrp.autodisabledelay)){@{ Description = "Auto Disable Delay"; Value = $svcgrp.autodisabledelay }}
				If ($svcgrp.cacheable -eq "YES"){@{ Description = "Cacheable [no]"; Value = $svcgrp.cacheable }}
				If ($svcgrp.state -eq "DISABLED"){@{ Description = "State [enabled]"; Value = $svcgrp.state }}
				If ($svcgrp.healthmonitor -eq "NO"){@{ Description = "Health Monitoring [yes]"; Value = $svcgrp.healthmonitor }}
				If ($svcgrp.appflowlog -eq "DISABLED"){@{ Description = "AppFlow Logging [enabled]"; Value = $svcgrp.appflowlog }}
				If ($svcgrp.monconnectionclose -ne "NONE"){@{ Description = "Monitoring Connection Close Bit [none]"; Value = $svcgrp.monconnectionclose }}
				If (![string]::IsNullOrWhiteSpace($svcgrp.comment)){@{ Description = "Comment"; Value = $svcgrp.comment }}
				
				# Settings
				If ($svcgrp.sp -eq "ON"){@{ Description = "Surge protection [off]"; Value = $svcgrp.sp }}
				If ($svcgrp.useproxyport -eq "OFF"){@{ Description = "Use the proxy port [on]"; Value = $svcgrp.useproxyport }}
				If ($svcgrp.downstateflush -eq "DISABLED"){@{ Description = "Down State Flush [enabled]"; Value = $svcgrp.downstateflush }}
				If ($svcgrp.usip -eq "YES"){@{ Description = "Use Source IP [no]"; Value = $svcgrp.usip }}
				If ($svcgrp.cka -eq "YES"){@{ Description = "Client keep-alive [no]"; Value = $svcgrp.cka }}
				If ($svcgrp.tcpb -eq "YES"){@{ Description = "TCP Buffering [no]"; Value = $svcgrp.tcpb }}
				If ($svcgrp.cmp -eq "YES"){@{ Description = "Compression [no]"; Value = $svcgrp.cmp }}
				@{ Description = "Insert Client IP Address [enabled]"; Value = $svcgrp.cip }
				If ($svcgrp.cip -eq "ENABLED"){@{ Description = "Name for the HTTP Header"; Value = $svcgrp.cipheader }}

				# Threshold & timeouts
				If ($svcgrp.maxbandwidth -ne "0"){@{ Description = "Maximum Bandwidth (Kbps)"; Value = $svcgrp.maxbandwidth}}
				If ($svcgrp.monthreshold -ne "0"){@{ Description = "Monitor Threshold"; Value = $svcgrp.monthreshold }}
				If ($svcgrp.maxreq -ne "0"){@{ Description = "Max Requests"; Value = $svcgrp.maxreq }}
				If ($svcgrp.maxclient -ne "0"){@{ Description = "Max Clients"; Value = $svcgrp.maxclient }}
				@{ Description = "Client Idle Time-out"; Value = $svcgrp.clttimeout }
				@{ Description = "Server Idle Time-out"; Value = $svcgrp.svrtimeout}
								
				# Profiles
				If (![string]::IsNullOrWhiteSpace($svcgrp.netprofilename)){@{ Description = "Network Profile Name"; Value = $svcgrp.netprofilename }}
				If (![string]::IsNullOrWhiteSpace($svcgrp.tcpprofilename)){@{ Description = "TCP Profile Name"; Value = $svcgrp.tcpprofilename }}
				If (![string]::IsNullOrWhiteSpace($svcgrp.httpprofilename)){@{ Description = "HTTP Profile Name"; Value = $svcgrp.httpprofilename }}

				If ($svcgrp.pathmonitor -eq "YES"){
					@{ Description = "Path Monitoring [no]"; Value = $svcgrp.pathmonitor }
					@{ Description = "Individual Path monitoring [no]"; Value = $svcgrp.pathmonitorindv }
				}
				If (![string]::IsNullOrWhiteSpace($svcgrp.sc)){@{ Description = "SureConnect"; Value = $svcgrp.sc }}
				If ($svcgrp.rtspsessionidremap -eq "ON"){@{ Description = "RTSP session ID mapping [off]"; Value = $svcgrp.rtspsessionidremap }}
				If (![string]::IsNullOrWhiteSpace($svcgrp.customserverid)){@{ Description = "Unique identifier for the service"; Value = $svcgrp.customserverid}}
			)
			$Params = $null
			$Params = @{
				Hashtable = $AdvancedConfiguration
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			New-BindingTable -Name $svcgrp.servicegroupname -BindingType "servicegroup_servicegroupmember_binding" -BindingTypeName "Servers" -Properties "servername,port,weight,serverid,hashid,state" -Headers "Server,Port,Weight,ServerID,HashID,State" -Style 5
			New-BindingTable -Name $svcgrp.servicegroupname -BindingType "servicegroup_lbmonitor_binding" -BindingTypeName "Monitor" -Properties "monitor_name,weight" -Headers "Monitor Name, Weight" -Style 5
			
			#region SSL Settings
			#Don't process SSL unless we are using an SSL based Service Type
			If ($svcgrp.ServiceType -match "SSL" -or $svcgrp.ServiceType -match "TLS" -and $svcgrp.ServiceType -notmatch "BRIDGE") {
				New-BindingTable -Name $svcgrp.servicegroupname -BindingType "sslservicegroup_sslcertkey_binding" -BindingTypeName "Certificates" -Properties "certkeyname,ca,crlcheck,snicert,ocspcheck,cleartextport" -Headers "Certificate Name,CA Certificate,CRL Checks Enabled,SNI Enabled,OCSP Enabled,Clear Text Port" -Style 5
				New-SSLSettings $svcgrp.servicegroupname sslservicegroup 5
			}
			#endregion SSL Settings
			
			$selection.InsertNewPage()
		}
	}
	#endregion Service Groups

	#region Monitors
	WriteWordLine 3 0 "Monitors"
	$monitors = Get-vNetScalerObject -Type lbmonitor
	[System.Collections.Hashtable[]] $MONITORSH = @()
	foreach ($monitor in $monitors) {
		$MONITORSH += @{
			NAME            = $monitor.monitorname
			Type            = $monitor.type
			DestinationPort = $monitor.destport
			Interval        = $monitor.interval
			TimeOut         = $monitor.resptimeout
		}
	}
	If ($MONITORSH.Length -gt 0) {
		$Params = $null
		$Params = @{
			Hashtable = $MONITORSH
			Columns   = "NAME", "Type", "DestinationPort", "Interval", "TimeOut"
			Headers   = "Monitor Name", "Type", "Destination Port", "Interval", "Time-Out"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
	foreach ($monitor in $monitors) {
		If ($monitor.Type -eq "HTTP" -and $Monitor.monitorname -ne "http" -and $Monitor.monitorname -ne "https"){
			WriteWordLine 4 0 "$($Monitor.monitorname)"
			[System.Collections.Hashtable[]] $MONITORSHTTPH = @(
				If (![string]::IsNullOrWhiteSpace($monitor.respcode)){@{ Description = "Response Code"; Value = $monitor.respcode}}
				If (![string]::IsNullOrWhiteSpace($monitor.customheaders)){@{ Description = "Custom Header"; Value = $monitor.customheaders -replace "\r\n","\r\n"}}
				If (![string]::IsNullOrWhiteSpace($monitor.httprequest)){@{ Description = "HTTP Request"; Value = $monitor.httprequest}}
				@{ Description = "Secure"; Value = $monitor.secure}
				If ($monitor.downtime -ne "30"){@{ Description = "Down Time"; Value = $monitor.downtime}}
				If ($monitor.retries -ne "3"){@{ Description = "Rertries"; Value = $monitor.retries}}
				If ($monitor.successretries -ne "1"){@{ Description = "Success Retries"; Value = $monitor.successretries}}
				If ($monitor.failureretries -ne "0"){@{ Description = "Failure Retries"; Value = $monitor.failureretries}}
				If (![string]::IsNullOrWhiteSpace($monitor.netprofile)){@{ Description = "Net Profile"; Value = $monitor.netprofile}}
			)
			$Params = $null
			$Params = @{
				Hashtable = $MONITORSHTTPH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion Monitors

	#region Servers
	If ((Get-vNetScalerObjectCount -Type server).__count -ge 1) {
		$selection.InsertNewPage()
		WriteWordLine 3 0 "Servers"
		$servers = Get-vNetScalerObject -Type server
		[System.Collections.Hashtable[]] $SERVERSH = @()
		foreach ($Server in $servers) {
			If ($TDcounter -ge 1){
				$SERVERSH += @{
					Server = $server.name
					IP     = $Server.ipaddress
					TD     = $server.td
					STATE  = $server.state
				}
			} Else {
				$SERVERSH += @{
					Server = $server.name
					IP     = $Server.ipaddress
					STATE  = $server.state
				}
			}
		}
		If ($SERVERSH.Length -gt 0) {
			$Params = $null
			If ($TDcounter -ge 1){
				$Params = @{
					Hashtable = $SERVERSH
					Columns   = "Server", "IP", "TD", "STATE"
					Headers   = "Server", "IP Address", "Traffic Domain", "State"
				}
			} Else {
				$Params = @{
					Hashtable = $SERVERSH
					Columns   = "Server", "IP", "STATE"
					Headers   = "Server", "IP Address", "State"
				}
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion Servers
	
	#region Persistency groups
	If ((Get-vNetScalerObjectCount -Type lbgroup).__count -ge 1) {
		$selection.InsertNewPage()
		WriteWordLine 3 0 "Persistency Groups"
		$lbgroups = Get-vNetScalerObject -Type lbgroup
		foreach ($lbgroup in $lbgroups) {
			$lbgroupmembers = (Get-vNetScalerObject -Type lbgroup_lbvserver_binding -Name $lbgroup.name).vservername -join ", "
			WriteWordLine 4 0 "$($lbgroup.name)"
			[System.Collections.Hashtable[]] $LBGROUPH = @(
				@{ Description = "Persistence"; Value = $lbgroup.persistencetype}
				If (![string]::IsNullOrWhiteSpace($lbgroup.cookiename)){@{ Description = "Cookie Name"; Value = $lbgroup.cookiename}}
				If (![string]::IsNullOrWhiteSpace($lbgroup.cookiedomain)){@{ Description = "Cookie Domain"; Value = $lbgroup.cookiedomain}}
				If (![string]::IsNullOrWhiteSpace($lbgroup.rule)){@{ Description = "Expression"; Value = $lbgroup.rule}}
				If (![string]::IsNullOrWhiteSpace($lbgroup.persistmask)){@{ Description = "IPv4 Netmask"; Value = $lbgroup.persistmask}}
				If (![string]::IsNullOrWhiteSpace($lbgroup.v6persistmasklen)){@{ Description = "IPv6 Mask Length"; Value = $lbgroup.v6persistmasklen}}
				If (![string]::IsNullOrWhiteSpace($lbgroup.timeout)){@{ Description = "Time-out"; Value = $lbgroup.timeout}}
				If (![string]::IsNullOrWhiteSpace($lbgroup.persistencebackup)){@{ Description = "Backup Persistence"; Value = $lbgroup.persistencebackup}}
				If (![string]::IsNullOrWhiteSpace($lbgroup.usevserverpersistency)){@{ Description = "Use vServer Persistence"; Value = $lbgroup.usevserverpersistency}}
				If (![string]::IsNullOrWhiteSpace($lbgroupmembers)){@{ Description = "Virtual Server Name"; Value = $lbgroupmembers}}
			)
			If ($LBGROUPH.Length -gt 0) {
				$Params = $null
				$Params = @{
					Hashtable = $LBGROUPH
					Columns   = "Description", "Value"
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
		}
	}
	#endregion Persistency groups
}
#endregion Load Balancers

#region Content Switches
If ($FEATCS -eq "Enabled"){
	If ((Get-vNetScalerObjectCount -Type csvserver).__count -ge 1) {
		$selection.InsertNewPage()
		WriteWordLine 2 0 "Content Switching"
		
		#region Content Switching Policy Labels
		If ((Get-vNetScalerObjectCount -Type cspolicylabel).__count -ge 1) {
			WriteWordLine 3 0 "Content Switching Policy Labels"
			$cspollbls = Get-vNetScalerObject -Type cspolicylabel
			foreach ($cspollbl in $cspollbls) {
				[System.Collections.Hashtable[]] $CSPLBLH = @()
				WriteWordLine 4 0 "$($cspollbl.labelname) ($($cspollbl.transform))"
				$cspollblbindings = Get-vNetScalerObject -Type cspolicylabel_cspolicy_binding -Name $cspollbl.labelname
				foreach ($csplbl in $cspollblbindings){
					$CSPLBLH += @{
						priority	= $csplbl.priority
						policy		= $csplbl.policyname
						gotoexp		= $csplbl.gotopriorityexpression
						invoke		= $csplbl.invoke
					}
				}
				If ($CSPLBLH.Length -gt 0) {
					$Params = $null
					$Params = @{
						Hashtable = $CSPLBLH
						Columns   = "priority", "policy", "gotoexp", "invoke"
						Headers   = "Priority", "Policy", "Goto Expression", "Invoke"
					}
					$Table = AddWordTable @Params
					FindWordDocumentEnd
				}
			}
		}
		#endregion Content Switching Policy Labels
		
		#region Content Switching Policies
		If ((Get-vNetScalerObjectCount -Type cspolicy).__count -ge 1) {
			WriteWordLine 3 0 "Content Switch Policies"
			$CSPols = Get-vNetScalerObject -Type cspolicy
			[System.Collections.Hashtable[]] $CSPH = @()
			foreach ($CSPol in $CSPols) {
`				$CSPH += @{
					CSPOLNAME = $CSPol.policyname
					RULE      = $CSPol.rule
					ACTION    = $CSPol.action
					LOGACTION = $CSPol.logaction
				}
			}
			If ($CSPH.Length -gt 0) {
				$Params = $null
				$Params = @{
					Hashtable = $CSPH
					Columns   = "CSPOLNAME", "RULE", "ACTION", "LOGACTION"
					Headers   = "CS Policy", "Rule", "Action", "Log Action"
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
		}
		#endregion Content Switching Policies

		#region Content Switching Actions
		If ((Get-vNetScalerObjectCount -Type csaction).__count -ge 1) {
			WriteWordLine 3 0 "Content Switch Actions"
			$CSActions = Get-vNetScalerObject -Type csaction
			[System.Collections.Hashtable[]] $CSAH = @()
			foreach ($CSAction in $CSActions) {
				If (![string]::IsNullOrWhiteSpace($CSAction.comment)){WriteWordLine 0 0 "$CSAction.comment"}
				If (![string]::IsNullOrWhiteSpace($CSAction.targetlbvserver)){$CSAtarget = $CSAction.targetlbvserver}
				If (![string]::IsNullOrWhiteSpace($CSAction.targetvserver)){$CSAtarget = $CSAction.targetvserver}
				If (![string]::IsNullOrWhiteSpace($CSAction.targetvserverexpr)){$CSAtarget = $CSAction.targetvserverexpr}
				$CSAH += @{
					Name	= $CSAction.name
					target	= $CSAtarget
				}
			}
			If ($CSAH.Length -gt 0) {
				$Params = $null
				$Params = @{
					Hashtable = $CSAH
					Columns   = "Name", "Target"
					Headers   = "Name", "Target"
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
		}
		#endregion Content Switching Actions
		
		#region Content Switching vServers
		$csvservers = Get-vNetScalerObject -Type csvserver
		foreach ($CS in $csvservers) {
			$selection.InsertNewPage()
			WriteWordLine 3 0 "$($CS.name)"
			[System.Collections.Hashtable[]] $CSVSERVERSH = @(
				@{ Description = "Description"; Value = "Configuration" }
				@{ Description = "Status"; Value = $CS.status }
				
				# Basic Settings
				@{ Description = "Protocol"; Value = $CS.servicetype }
				If ($CS.targettype -eq "GSLB"){@{ Description = "Target Type [none]"; Value = $CS.targettype }}
				If ($CS.persistencetype -ne "NONE"){@{ Description = "Persistence Type"; Value = $CS.persistencetype }}
				If ($CS.persistmask -ne "255.255.255.255"){@{ Description = "Persist Mask [255.255.255.255]"; Value = $CS.persistmask }}
				If ($CS.v6persistmasklen -ne "128"){@{ Description = "IPv6 Persist Mask Length [128]"; Value = $CS.v6persistmasklen }}
				If ($CS.timeout -ne "2"){@{ Description = "Timeout [2]"; Value = $CS.timeout }}
				If ($CS.persistencetype -eq "COOKIE_INSERT"){
					@{ Description = "Cookie Name"; Value = $CS.cookiename }
					@{ Description = "Persistence Backup"; Value = $CS.persistencebackup }
				}
				If ($CS.backuppersistencetimeout -ne "2"){@{ Description = "Backup Persistence Timeout [2]"; Value = $CS.backuppersistencetimeout }}
				If ($CS.targettype -eq "GSLB"){
					@{ Description = "DNS record type"; Value = $CS.dnsrecordtype}
					@{ Description = "Persistence ID"; Value = $CS.persistenceid}
				}Else{
					If (![string]::IsNullOrWhiteSpace($CS.ipv46)){@{ Description = "IP"; Value = $CS.ipv46}}
					If (![string]::IsNullOrWhiteSpace($CS.ippattern)){@{ Description = "IP Pattern"; Value = $CS.ippattern}}
					If (![string]::IsNullOrWhiteSpace($CS.ipmask)){@{ Description = "IP Mask"; Value = $CS.ipmask}}
				}
				@{ Description = "Port"; Value = $CS.port}
				If ($CS.td -ne "0"){@{ Description = "Traffic Domain [0]"; Value = $CS.td}}
				If (![string]::IsNullOrWhiteSpace($CS.ipset)){@{ Description = "IPSet"; Value = $CS.ipset}}
				If (![string]::IsNullOrWhiteSpace($CS.range)){@{ Description = "Range"; Value = $CS.range}}
				If (![string]::IsNullOrWhiteSpace($CS.listenpriority)){@{ Description = "Listen Priority"; Value = $CS.listenpriority }}
				If (![string]::IsNullOrWhiteSpace($CS.listenpolicy)){@{ Description = "Listen Policy Expression"; Value = $CS.listenpolicy }}
				If (![string]::IsNullOrWhiteSpace($CS.comment)){@{ Description = "Comments"; Value = $CS.comment }}
				If ($CS.curstate -eq "DISABLED"){@{ Description = "State [enabled]"; Value = $CS.curstate }}
				If ($CS.rhistate -eq "ACTIVE"){@{ Description = "RHI State [passive]"; Value = $CS.rhistate }}
				If ($CS.dtls -eq "ON"){@{ Description = "DTLS [off]"; Value = $CS.dtls }}
				If ($CS.appflowlog -eq "DISABLED"){@{ Description = "Apply AppFlow logging [enabled]"; Value = $CS.appflowlog }}
				If (![string]::IsNullOrWhiteSpace($CS.probeprotocol)){@{ Description = "Probe Protocol"; Value = $CS.probeprotocol }}
				If (![string]::IsNullOrWhiteSpace($CS.probesuccessresponsecode)){@{ Description = "Probe Success Response Code"; Value = $CS.probesuccessresponsecode }}
				If (![string]::IsNullOrWhiteSpace($CS.probeport)){@{ Description = "Probe Port"; Value = $CS.probeport }}
				If (![string]::IsNullOrWhiteSpace($CS.redirectfromport)){@{ Description = "Redirect from port"; Value = $CS.redirectfromport }}
				If (![string]::IsNullOrWhiteSpace($CS.httpsredirecturl)){@{ Description = "HTTPS Redirect URL"; Value = $CS.httpsredirecturl }}

				If (![string]::IsNullOrWhiteSpace($CS.lbvserver)){@{ Description = "Default vServer"; Value = $CS.lbvserver }}
				
				# Traffic Settings
				@{ Description = "Client Idle Time-out"; Value = $CS.clttimeout }
				If ($CS.insertvserveripport -ne "OFF"){@{ Description = "Virtual Server IP Port Insertion [off]"; Value = $CS.insertvserveripport }}
				If (![string]::IsNullOrWhiteSpace($CS.vipheader)){@{ Description = "Virtual Server IP Port Insertion Value"; Value = $CS.vipheader }}
				If ($CS.icmpvsrresponse -ne "PASSIVE"){@{ Description = "ICMP Virtual Server Response [passive]"; Value = $CS.icmpvsrresponse }}
				If ($CS.cacheable -eq "ENABLED"){@{ Description = "Cacheable [disabled]"; Value = $CS.cacheable }}
				If ($CS.downstateflush -eq "DISABLED"){@{ Description = "Down State Flush [enabled]"; Value = $CS.downstateflush }}
				If ($CS.casesensitive -eq "ON"){@{ Description = "Case Sensitive [off]"; Value = $CS.casesensitive }}
				If ($CS.l2conn -eq "ON"){@{ Description = "Layer 2 Parameters [off]"; Value = $CS.l2conn }}
				If ($CS.stateupdate -eq "ENABLED"){@{ Description = "State Update [disabled]"; Value = $CS.stateupdate }}
				If ($CS.precedence -eq "RULE"){@{ Description = "Precedence [rule]"; Value = $CS.precedence }}
				
				# Authentication
				If ($CS.authentication -eq "ON"){
					If ($CS.authn401 -eq "OFF"){
						@{ Description = "Form Based Authentication"; Value = "ON"}
						If (![string]::IsNullOrWhiteSpace($CS.authenticationhost)){@{ Description = "Authentication FQDN"; Value = $CS.authenticationhost}}
					}Else{
						@{ Description = "401 Based Authentication"; Value = "ON"}
					}
					If (![string]::IsNullOrWhiteSpace($CS.authnvsname)){@{ Description = "Authentication virtual server name"; Value = $CS.authnvsname}}
					If (![string]::IsNullOrWhiteSpace($LoadBalance.authnprofile)){@{ Description = "Authentication profile"; Value = $LoadBalance.authnprofile}}
				}
				
				# Protection
				If (![string]::IsNullOrWhiteSpace($CS.redirecturl)){@{ Description = "Redirect URL"; Value = $CS.redirecturl }}
				If (![string]::IsNullOrWhiteSpace($CS.backupvserver)){@{ Description = "Backup vServer"; Value = $CS.backupvserver }}
				If ($CS.disableprimaryondown -eq "ENABLED"){@{ Description = "Disable Primary When Down [enabled]"; Value = $CS.disableprimaryondown }}

				# Spillover
				If ($CS.somethod -ne "NONE"){@{ Description = "Spillover Method [none]"; Value = $CS.somethod }}
				If (![string]::IsNullOrWhiteSpace($CS.sothreshold)){@{ Description = "Spillover Threshold"; Value = $CS.sothreshold }}
				If (![string]::IsNullOrWhiteSpace($CS.sobackupaction)){@{ Description = "Spillover Backup Action"; Value = $CS.sobackupaction }}
				If ($CS.sopersistencetimeout -ne "2"){@{ Description = "Spillover Persistence Timeout (mins) [2]"; Value = $CS.sopersistencetimeout }}
				If ($CS.sopersistence -eq "ENABLED"){@{ Description = "Spillover Persistence [disabled]"; Value = $CS.sopersistence }}
				
				# Profiles
				If (![string]::IsNullOrWhiteSpace($CS.netprofile)){@{ Description = "Network Profile"; Value = $CS.netprofile }}
				If (![string]::IsNullOrWhiteSpace($CS.tcpprofilename)){@{ Description = "TCP Profile"; Value = $CS.tcpprofilename }}
				If (![string]::IsNullOrWhiteSpace($CS.httpprofilename)){@{ Description = "HTTP Profile"; Value = $CS.httpprofilename }}
				If (![string]::IsNullOrWhiteSpace($CS.dbprofilename)){@{ Description = "DB Profile"; Value = $CS.dbprofilename }}
				If (![string]::IsNullOrWhiteSpace($CS.dnsprofilename)){@{ Description = "DNS Profile Name"; Value = $CS.dnsprofilename }}
				If (![string]::IsNullOrWhiteSpace($CS.quicprofilename)){@{ Description = "QUIC Profile Name"; Value = $CS.quicprofilename }}

				# Push
				If ($CS.push -eq "ENABLED"){
					@{ Description = "Push"; Value = $CS.push}
					If (![string]::IsNullOrWhiteSpace($CS.pushmulticlients)){@{ Description = "Push Multiple Clients"; Value = $CS.pushmulticlients}}
					If (![string]::IsNullOrWhiteSpace($CS.pushvserver)){@{ Description = "Push Virtual Server"; Value = $CS.pushvserver}}
					If (![string]::IsNullOrWhiteSpace($CS.pushlabel)){@{ Description = "Push Label Expression"; Value = $CS.pushlabel}}
				}
			)
			$Params = $null
			$Params = @{
				Hashtable	= $CSVSERVERSH
				Columns		= "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			#region SSL Settings
			#Don't process SSL unless we are using an SSL based Service Type
			If ($CS.ServiceType -match "SSL" ) {
				New-BindingTable -Name $CS.Name -BindingType "sslvserver_sslcertkey_binding" -BindingTypeName "Certificates" -Properties "certkeyname,ca,crlcheck,snicert,ocspcheck,cleartextport" -Headers "Certificate Name,CA Certificate,CRL Checks Enabled,SNI Enabled,OCSP Enabled,Clear Text Port" -Style 4
				New-SSLSettings $CS.Name sslvserver 4
			}
			#endregion SSL Settings

			#region Content Switching Policies
			New-BindingTable -Name $CS.name -BindingType "csvserver_cspolicy_binding" -BindingTypeName "Content Switching Policies" -Properties "priority,policyname,rule,gotopriorityexpression,targetlbvserver" -Headers "Priority,Policy Name,Rule,Goto Expression,Target LB vServer" -Style 4
			New-BindingTable -Name $CS.name -BindingType "csvserver_responderpolicy_binding" -BindingTypeName "Responder Policies" -Properties "priority,policyname,gotopriorityexpression" -Headers "Priority,Policy Name,Go To Expression" -Style 4
			New-BindingTable -Name $CS.name -BindingType "csvserver_rewritepolicy_binding" -BindingTypeName "Rewrite Policies" -Properties "priority,policyname,bindpoint,gotopriorityexpression" -Headers "Priority,Policy Name,Bindpoint,Go To Expression" -Style 4
			New-BindingTable -Name $CS.name -BindingType "csvserver_tmtrafficpolicy_binding" -BindingTypeName "Traffic Policies" -Properties "priority,policyname" -Headers "Priority,Policy Name" -Style 4
			New-BindingTable -Name $CS.name -BindingType "csvserver_analyticsprofile_binding" -BindingTypeName "Analytics Policies" -Properties "analyticsprofile" -Headers "Analytics Profile" -Style 4
			New-BindingTable -Name $CS.name -BindingType "csvserver_cachepolicy_binding" -BindingTypeName "Cache Policies" -Properties "priority,policyname,bindpoint,gotopriorityexpression" -Headers "Priority,Policy Name,BindPoint,Go To Expression" -Style 4
			#endregion Content Switching policies
		} # end if
		#endregion Content Switching vServers
	}
}
#endregion Content Switches

#region Cache Redirection
If ($FEATCR -eq "Enabled"){
	If ((Get-vNetScalerObjectCount -Type crvserver).__count -ge 1) {
		$selection.InsertNewPage()
		WriteWordLine 2 0 "Cache Redirection"
		$crservers = Get-vNetScalerObject -Type crvserver
		foreach ($crserver in $crservers) {
			$crname = $crserver.name
			WriteWordLine 2 0 "Cache Redirection Server $crname"
			$Params = $null
			$Params = @{
				Hashtable = @{
					PROT       = $crserver.servicetype
					IP         = $crserver.ip
					Port       = $crserver.port
					CACHETYPE  = $crserver.cachetype
					REDIRECT   = $crserver.redirect
					CLTTIEMOUT = $crserver.clttimeout
				}
				Columns   = "PROT", "IP", "PORT", "CACHETYPE", "REDIRECT", "CLTTIEMOUT"
				Headers   = "Protocol", "IP", "Port", "Cache Type", "Redirect", "Client Time-out"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
}
#endregion Cache Redirection

#region DNS Configuration
$selection.InsertNewPage()
WriteWordLine 2 0 "DNS Configuration"

#region DNS name servers
If ((Get-vNetScalerObjectCount -Type dnsnameserver).__count -ge 1) {
	WriteWordLine 3 0 "DNS Name Servers"
	$dnsnameservers = Get-vNetScalerObject -Type dnsnameserver
    [System.Collections.Hashtable[]] $DNSNAMESERVERH = @()
    foreach ($DNSNAMESERVER in $DNSNAMESERVERS) {
        $DNSNAMESERVERH += @{
            DNSServer = $dnsnameserver.ip
            State     = $dnsnameserver.state
            Prot      = $dnsnameserver.type
        }
    }
    If ($DNSNAMESERVERH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $DNSNAMESERVERH
            Columns   = "DNSServer", "State", "Prot"
            Headers   = "DNS Name Server", "State", "Protocol" 
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion DNS name servers

#region DNS Name Suffix
If ((Get-vNetScalerObjectCount -Type dnssuffix).__count -ge 1) {
	WriteWordLine 3 0 "DNS Name Suffixes"
	$dnssuffixes = Get-vNetScalerObject -Type dnssuffix
    [System.Collections.Hashtable[]] $DNSSUFFIXCONFIGH = @()
    foreach ($dnssuffix in $dnssuffixes) {
        $DNSSUFFIXCONFIGH += @{
            DNSSUFFIX = $dnssuffix.dnssuffix
        }
    }
    If ($DNSSUFFIXCONFIGH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $DNSSUFFIXCONFIGH
            Columns   = "DNSSUFFIX"
            Headers   = "DNS Suffix"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion DNS Address Records

#region DNS Address Records
If ((Get-vNetScalerObjectCount -Type dnsaddrec).__count -ge 1) {
	WriteWordLine 3 0 "DNS Address Records"
	$dnsaddrecs = Get-vNetScalerObject -Type dnsaddrec
    [System.Collections.Hashtable[]] $DNSRECORDCONFIGH = @()
    foreach ($dnsaddrec in $dnsaddrecs) {
        $DNSRECORDCONFIGH += @{
            DNSRecord = $dnsaddrec.hostname
            IPAddress = $dnsaddrec.ipaddress
            TTL       = $dnsaddrec.ttl
            AUTHTYPE  = $dnsaddrec.authtype
        }
    }
    If ($DNSRECORDCONFIGH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $DNSRECORDCONFIGH
            Columns   = "DNSRecord", "IPAddress", "TTL", "AUTHTYPE"
            Headers   = "DNS Record", "IP Address", "TTL", "Authentication Type"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion DNS Address Records

#region DNS AAA Records
If ((Get-vNetScalerObjectCount -Type dnsaaaarec).__count -ge 1) {
	WriteWordLine 3 0 "DNS AAA Records"
	$dnsaaaarecs = Get-vNetScalerObject -Type dnsaaaarec
    [System.Collections.Hashtable[]] $DNSRECORDCONFIGH = @()
    foreach ($dnsaaarec in $dnsaaaarecs) {
        $DNSRECORDCONFIGH += @{
            DNSRecord = $dnsaaarec.hostname
            IPAddress = $dnsaaarec.ipv6address
            TTL       = $dnsaaarec.ttl
            AUTHTYPE  = $dnsaaarec.authtype
        }
    }
    If ($DNSRECORDCONFIGH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $DNSRECORDCONFIGH
            Columns   = "DNSRecord", "IPAddress", "TTL", "AUTHTYPE"
            Headers   = "DNS Record", "IP Address", "TTL", "Authentication Type"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion DNS AAA Records

#region DNS CNAME Records
If ((Get-vNetScalerObjectCount -Type dnscnamerec).__count -ge 1) {
	WriteWordLine 3 0 "DNS CNAME Records"
	$dnscnamerecs = Get-vNetScalerObject -Type dnscnamerec    
    [System.Collections.Hashtable[]] $DNSRECORDCONFIGH = @()
    foreach ($dnscnamerec in $dnscnamerecs) {
        $DNSRECORDCONFIGH += @{
            DNSRecord = $dnscnamerec.aliasname
            IPAddress = $dnscnamerec.canonicalname
            TTL       = $dnscnamerec.ttl
            AUTHTYPE  = $dnscnamerec.authtype
        }
    }
    If ($DNSRECORDCONFIGH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $DNSRECORDCONFIGH
            Columns   = "DNSRecord", "IPAddress", "TTL", "AUTHTYPE"
            Headers   = "Alias Name", "Canonical Name", "TTL", "Authentication Type"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion DNS CNAME Records

#region DNS MX Records
If ((Get-vNetScalerObjectCount -Type dnsmxrec).__count -ge 1) {
	WriteWordLine 3 0 "DNS MX Records"
	$dnsmxrecs = Get-vNetScalerObject -Type dnsmxrec
    [System.Collections.Hashtable[]] $DNSMXRECORDCONFIGH = @()
    foreach ($dnsmxrec in $dnsmxrecs) {
        $DNSMXRECORDCONFIGH += @{
            DOMAIN   = $dnsmxrec.domain
            MX       = $dnsmxrec.mx
            TTL      = $dnsmxrec.ttl
            AUTHTYPE = $dnsmxrec.authtype
        }
    }
    If ($DNSMXRECORDCONFIGH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $DNSMXRECORDCONFIGH
            Columns   = "DOMAIN", "MX", "TTL", "AUTHTYPE"
            Headers   = "Domain", "MX", "TTL", "Authentication Type"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion DNS MX Records

#region DNS NS Records
If ((Get-vNetScalerObjectCount -Type dnsnsrec).__count -ge 1) {
	WriteWordLine 3 0 "DNS NS Records"
	$dnsnsrecs = Get-vNetScalerObject -Type dnsnsrec
    [System.Collections.Hashtable[]] $DNSNSRECORDCONFIGH = @()
    foreach ($dnsnsrec in $dnsnsrecs) {
        $DNSNSRECORDCONFIGH += @{
            DOMAIN   = $dnsnsrec.domain
            NS       = $dnsnsrec.nameserver
            TTL      = $dnsnsrec.ttl
            AUTHTYPE = $dnsnsrec.authtype
        }
    }
    If ($DNSNSRECORDCONFIGH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $DNSNSRECORDCONFIGH
            Columns   = "DOMAIN", "NS", "TTL", "AUTHTYPE"
            Headers   = "Domain", "NameServer", "TTL", "Authentication Type"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion DNS NS Records

#region DNS SOA Records
If ((Get-vNetScalerObjectCount -Type dnssoarec).__count -ge 1) {
	WriteWordLine 3 0 "DNS SOA Records"
	$dnssoarecs = Get-vNetScalerObject -Type dnssoarec
    [System.Collections.Hashtable[]] $DNSSOARECORDCONFIGH = @()
    foreach ($dnssoarec in $dnssoarecs) {
        $DNSSOARECORDCONFIGH += @{
            DOMAIN   = $dnssoarec.domain
            ORIGIN   = $dnssoarec.originserver
            CONTACT  = $dnssoarec.contact
            SERIAL   = $dnssoarec.serial
            TTL      = $dnssoarec.ttl
            AUTHTYPE = $dnssoarec.authtype
        }
    }
    If ($DNSSOARECORDCONFIGH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $DNSSOARECORDCONFIGH
            Columns   = "DOMAIN", "ORIGIN", "CONTACT", "SERIAL", "TTL", "AUTHTYPE"
            Headers   = "Domain", "Origin Server", "Admin Contact", "Serial Number", "TTL", "Authentication Type"
        }
        $Table = AddWordTable @Params
        FindWordDocumentEnd
    }
}
#endregion DNS SOA Records
#endregion DNS Configuration

#region Global Server Load Balancing
If ($FEATGSLB -eq "Enabled"){
	$selection.InsertNewPage()
	WriteWordLine 2 0 "Global Server Load Balancing"

	#region GSLB Parameters
	WriteWordLine 3 0 "GSLB Parameters"
	$gslbparameters = Get-vNetScalerObject -Type gslbparameter
	[System.Collections.Hashtable[]] $GSLBParameterDetails = @(
		@{ Description = "Description"; Value = "Value"}
		@{ Description = "LDNS Entry Timeout"; Value = $gslbparameters.ldnsentrytimeout}
		@{ Description = "RTT Tolerance"; Value = $gslbparameters.rtttolerance}
		@{ Description = "IPv4 LDNS Mask"; Value = $gslbparameters.ldnsmask}
		@{ Description = "IPv6 LDNS Mask"; Value = $gslbparameters.v6ldnsmasklen}
		@{ Description = "LDNS Probe Order"; Value = $gslbparameters.ldnsprobeorder -Join ", "}
		@{ Description = "Drop LDNS Requests"; Value = $gslbparameters.dropldnsreq}
		@{ Description = "GSLB Service State Delay Time (secs)"; Value = $gslbparameters.gslbsvcstatedelaytime}
		@{ Description = "Automatic Config Sync"; Value = $gslbparameters.automaticconfigsync}
	)
	$Params = $null
	$Params = @{
		Hashtable = $GSLBParameterDetails
		Columns   = "Description", "Value"
	}
	$Table = AddWordTable @Params -List
	FindWordDocumentEnd
	#endregion GSLB Parameters

	#region GSLB vServers
	$gslbvservercounter = Get-vNetScalerObjectCount -Type gslbvserver 
	If ($gslbvservercount -ge 1) {
		WriteWordLine 3 0 "GSLB Virtual Servers"
		$gslbvservers = Get-vNetScalerObject -Type gslbvserver
		foreach ($gslbvserver in $gslbvservers) {
			WriteWordLine 4 0 "$($gslbvserver.name)"
			[System.Collections.Hashtable[]] $GSLBvServerDetails = @(
				@{ Description = "Description"; Value = "Value"}
				@{ Description = "Service Type"; Value = $gslbvserver.servicetype}
				@{ Description = "State"; Value = $gslbvserver.state}
				@{ Description = "Status"; Value = $gslbvserver.status}
				@{ Description = "IP Type"; Value = $gslbvserver.iptype}
				@{ Description = "DNS Record Type"; Value = $gslbvserver.dnsrecordtype}
				@{ Description = "Persistence Type"; Value = $gslbvserver.persistencetype}
				@{ Description = "Persistence ID"; Value = $gslbvserver.persistenceid}
				@{ Description = "Load Balancing Method"; Value = $gslbvserver.lbmethod}
				@{ Description = "Backup Load Balancing Method"; Value = $gslbvserver.backuplbmethod}
				@{ Description = "Tolerance"; Value = $gslbvserver.tolerance}
				@{ Description = "Timeout"; Value = $gslbvserver.timeout}
				@{ Description = "Netmask"; Value = $gslbvserver.netmask}
				@{ Description = "IPv6 Netmask"; Value = $gslbvserver.v6netmasklen}
				@{ Description = "Persistence mask"; Value = $gslbvserver.persistmask}
				@{ Description = "IPv6 Persistence mask"; Value = $gslbvserver.v6persistmasklen}
				@{ Description = "Bound Services"; Value = $gslbvserver.servicename}
				@{ Description = "Weight"; Value = $gslbvserver.weight}
				@{ Description = "Domain Name"; Value = $gslbvserver.domainname}
				@{ Description = "TTL"; Value = $gslbvserver.ttl}
				@{ Description = "Backup IP Address"; Value = $gslbvserver.backupip}
				@{ Description = "Cookie Domain"; Value = $gslbvserver.cookiedomain}
				@{ Description = "Cookie Timeout"; Value = $gslbvserver.cookietimeout}
				@{ Description = "Domain TTL"; Value = $gslbvserver.sitedomainttl}
				@{ Description = "Backup vServer"; Value = $gslbvserver.backupvserver}
				@{ Description = "Disable Primary when down"; Value = $gslbvserver.disableprimaryondown}
				@{ Description = "Dynamic Weight"; Value = $gslbvserver.dynamicweight}
				@{ Description = "ISC Weight"; Value = $gslbvserver.iscweight}
				@{ Description = "Site Persistence"; Value = $gslbvserver.sitepersistence}
				@{ Description = "Comment"; Value = $gslbvserver.comment}
				@{ Description = "vServer Bind Service IP"; Value = $gslbvserver.vsvrbindsvcip}
				@{ Description = "vServer Bind Service Port"; Value = $gslbvserver.vsvrbindsvcport}
				@{ Description = "EDNS Client Subnet"; Value = $gslbvserver.ecs}
				@{ Description = "Validate ECS Addresses"; Value = $gslbvserver.ecsaddrvalidation}
				#TODO: Spillover Policies
			)
			$Params = $null
			$Params = @{
				Hashtable = $GSLBvServerDetails
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List 
			FindWordDocumentEnd
			
			#region GSLB vServer Service Bindings
			If ((Get-vNetScalerObjectCount -Type gslbvserver_gslbservice_binding -Name $gslbvservername).__count -ge 1) {
				WriteWordLine 5 0 "Services"
				$GSLBServiceBinds = Get-vNetScalerObject -Type gslbvserver_gslbservice_binding -Name $gslbvservername
				[System.Collections.Hashtable[]] $GSLBServices = @()
				foreach ($GSLBServiceBind in $GSLBServiceBinds) {
					$GSLBServices += @{ ServiceName = $GSLBServiceBind.servicename; Weight = $GSLBServiceBind.weight}
				}
				$Params = $null
				$Params = @{
					Hashtable = $GSLBServices 
					Columns   = "ServiceName", "Weight"
					Headers   = "Service Name", "Service Weight"
				}
				FindWordDocumentEnd
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
			#endregion GSLB vServer Service Bindings
			
			#region GSLB Domain Bindings
			If ((Get-vNetScalerObjectCount -Type gslbvserver_domain_binding -Name $gslbvservername).__count -ge 1) {
				WriteWordLine 5 0 "Domain Bindings"
				$GSLBDomainBinds = Get-vNetScalerObject -Type gslbvserver_domain_binding -Name $gslbvservername
				[System.Collections.Hashtable[]] $GSLBDomains = @()
				foreach ($GSLBDomainBind in $GSLBDomainBinds) {
					$GSLBDomains += @{ DomainName = $GSLBDomainBind.domainname; TTL = $GSLBDomainBind.ttl; CookieDomain = $GSLBDomainBind.cookie_domain; CookieTimeout = $GSLBDomainBind.cookietimeout}
				}
				$Params = $null
				$Params = @{
					Hashtable = $GSLBDomains 
					Columns   = "DomainName", "TTL", "CookieDomain", "CookieTimeout"
					Headers   = "Domain Name", "TTL", "Cookie Domain", "Cookie Timeout"
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
			#endregion GSLB Domain Bindings
		}
	}
	#endregion GSLB vServers

	#region GSLB Services
	If ((Get-vNetScalerObjectCount -Type gslbservice).__count -ge 1) {
		WriteWordLine 3 0 "GSLB Services"
		$gslbservicesall = Get-vNetScalerObject -Type gslbservice
		foreach ($gslbservice in $gslbservicesall) {
			WriteWordLine 4 0 "$($gslbservice.servicename)"
			[System.Collections.Hashtable[]] $GSLBServiceDetails = @(
				@{ Description = "Description"; Value = "Value"}
				@{ Description = "Service Location"; Value = $gslbservice.gslb}
				@{ Description = "GSLB Site"; Value = $gslbservice.sitename}
				@{ Description = "IP Address"; Value = $gslbservice.ipaddress}
				@{ Description = "IP"; Value = $gslbservice.ip}
				@{ Description = "Server Name"; Value = $gslbservice.servername}
				@{ Description = "Port"; Value = $gslbservice.port}
				@{ Description = "Public IP"; Value = $gslbservice.publicip}
				@{ Description = "Public Port"; Value = $gslbservice.publicport}
				@{ Description = "Max Clients"; Value = $gslbservice.maxclient}
				@{ Description = "Max AAA Users"; Value = $gslbservice.maxaaausers}
				@{ Description = "Monitor Threshold"; Value = $gslbservice.monthreshold}
				@{ Description = "State"; Value = $gslbservice.state}
				@{ Description = "Insert Client IP"; Value = $gslbservice.cip}
				@{ Description = "Client IP Header"; Value = $gslbservice.cipheader}
				@{ Description = "Site Persistence"; Value = $gslbservice.sitepersistence}
				@{ Description = "Site Prefix"; Value = $gslbservice.siteprefix}
				@{ Description = "Client Timeout"; Value = $gslbservice.clttimeout}
				@{ Description = "Server Timeout"; Value = $gslbservice.svrtimeout}
				@{ Description = "Preferred Location"; Value = $gslbservice.preferredlocation}
				@{ Description = "Maximum bandwidth, in Kbps"; Value = $gslbservice.maxbandwidth}
				@{ Description = "Flush active transactions for DOWN service"; Value = $gslbservice.downstateflush}
				@{ Description = "CNAME Entry"; Value = $gslbservice.cnameentry}
				@{ Description = "Comment"; Value = $gslbservice.comment}
			)
			$Params = $null
			$Params = @{
				Hashtable = $GSLBServiceDetails
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List 
			FindWordDocumentEnd
			
			#region GSLB Service Monitors
			If ((Get-vNetScalerObjectCount -Type gslbservice_lbmonitor_binding -Name $gslbservicename).__count -ge 1) {
				WriteWordLine 5 0 "Monitors"
				$GSLBMonitorBinds = Get-vNetScalerObject -Type gslbservice_lbmonitor_binding -Name $gslbservicename
				[System.Collections.Hashtable[]] $GSLBMonitors = @()
				foreach ($GSLBMonitorBind in $GSLBMonitorBinds) {
					$GSLBServices += @{ MonitorName = $GSLBMonitorBind.monitor_name; Weight = $GSLBMonitorBind.weight}
				}
				$Params = $null
				$Params = @{
					Hashtable = $GSLBMonitors 
					Columns   = "MonitorName", "Weight"
					Headers   = "Monitor Name", "Weight"
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
			#endregion GSLB Service Monitors

			#region GSLB Service DNS View
			If ((Get-vNetScalerObjectCount -Type gslbservice_dnsview_binding -Name $gslbservicename).__count -ge 1) {
				WriteWordLine 5 0 "DNS Views"
				$GSLBDNSViewBinds = Get-vNetScalerObject -Type gslbservice_dnsview_binding -Name $gslbservicename
				[System.Collections.Hashtable[]] $GSLBDNSViews = @()
				foreach ($GSLBDNSViewBind in $GSLBDNSViewBinds) {
					$GSLBDNSViews += @{ ViewName = $GSLBDNSViewBind.viewname; ViewIP = $GSLBDNSViewBind.viewip}
				}
				$Params = $null
				$Params = @{
					Hashtable = $GSLBMonitors 
					Columns   = "ViewName", "ViewIP"
					Headers   = "View Name", "View IP"
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd            
			}
			#endregion GSLB Service DNS View
		}
	}
	#endregion GSLB Services

	#region GSLB Sites
	If ((Get-vNetScalerObjectCount -Type gslbsite).__count -ge 1) {
		WriteWordLine 3 0 "GSLB Sites"
		$gslbsitesall = Get-vNetScalerObject -Type gslbsite
		foreach ($gslbsite in $gslbsitesall) {
			WriteWordLine 4 0 "$($gslbsite.sitename)"
			[System.Collections.Hashtable[]] $GSLBSiteDetails = @(
				@{ Description = "Description"; Value = "Value"}
				@{ Description = "Site Type"; Value = $gslbsite.sitetype}
				@{ Description = "Site IP Address"; Value = $gslbsite.siteipaddress}
				@{ Description = "Site Public IP"; Value = $gslbsite.publicip}
				@{ Description = "Metric Exchange"; Value = $gslbsite.metricexchange}
				@{ Description = "Persistence Exchange"; Value = $gslbsite.persistencemepstatus}
				@{ Description = "Network Metric Exchange"; Value = $gslbsite.nwmetricexchange}
				@{ Description = "Session Exchange"; Value = $gslbsite.sessionexchange}
				@{ Description = "Parent Site"; Value = $gslbsite.parentsite}
				@{ Description = "Cluster IP"; Value = $gslbsite.clip}
				@{ Description = "Public Cluster IP"; Value = $gslbsite.publicclip}
				@{ Description = "Backup Parent Sites"; Value = $gslbsite.backupparentlist -Join ", "}
			)
			$Params = $null
			$Params = @{
				Hashtable = $GSLBSiteDetails
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
		}
	}
	#endregion GSLB Sites

	$selection.InsertNewPage()
}

#endregion Global Server Load Balancing

#region SSL
$selection.InsertNewPage()
WriteWordLine 2 0 "SSL"

#region SSL Certificates
WriteWordLine 3 0 "SSL Certificates"
$sslcerts = Get-vNetScalerObject -Type sslcertkey
foreach ($sslcert in $sslcerts) {
    $sslcert1 = Get-vNetScalerObject -Type sslcertkey -Name $sslcert.certkey
    $subject = $sslcert1.subject
    $subject1 = $subject.Split(',')[-1]
    $sslfqdn = ($subject1 -replace 'CN=', '')
    $sslcertname = $sslcert.certkey
    WriteWordLine 4 0 "SSL Certificate: $sslcertname"
    [System.Collections.Hashtable[]] $SSLCertDetails = @(
        @{ Description = "Description"; Value = "Value"}
        @{ Description = "Name"; Value = $sslcert.certkey}
        @{ Description = "FQDN"; Value = $sslfqdn}
        @{ Description = "Issuer"; Value = $sslcert.issuer}
        @{ Description = "Certificate File"; Value = $sslcert.cert}
        If (![string]::IsNullOrWhiteSpace($sslcert.key)){@{ Description = "Key File"; Value = $sslcert.key}}
        @{ Description = "Key Size"; Value = $sslcert.publickeysize}
        @{ Description = "Valid From"; Value = $sslcert.clientcertnotbefore}
        @{ Description = "Valid Until"; Value = $sslcert.clientcertnotafter}
        @{ Description = "Days to Expiry"; Value = $sslcert.daystoexpiration}
        @{ Description = "Certificate Type"; Value = $sslcert.certificatetype -join ", "}
        If (![string]::IsNullOrWhiteSpace($sslcert.linkcertkeyname)){@{ Description = "Linked Certificate"; Value = $sslcert.linkcertkeyname}}
    )
    $Params = $null
    $Params = @{
        Hashtable = $SSLCertDetails
        Columns   = "Description", "Value"
    }
    $Table = AddWordTable @Params -List 
    FindWordDocumentEnd
}
#endregion SSL Certificates

#region SSL Cipher Groups
WriteWordLine 3 0 "SSL Cipher Groups"
$SSLCiphers = Get-vNetScalerObject -Type sslcipher
foreach ($SSLCipher in $SSLCiphers) {
	If ($SSLCipher.description -eq "User Defined Cipher Group"){
		WriteWordLine 4 0 "$($SSLCipher.ciphergroupname)"
		$SSLCipherGroups = Get-vNetScalerObject -Type sslcipher_sslciphersuite_binding -Name $SSLCipher.ciphergroupname
		[System.Collections.Hashtable[]] $SSLCIPHERGROUPH = @()
		foreach ($SSLCipherGroup in $SSLCipherGroups){
			$SSLCIPHERGROUPH += @{ PRIORITY = $SSLCipherGroup.cipherpriority; CIPHER = $SSLCipherGroup.ciphername}
		}
		If ($SSLCIPHERGROUPH.Length -gt 0) {
			$Params = $null
			$Params = @{
				Hashtable = $SSLCIPHERGROUPH 
				Columns   = "PRIORITY", "CIPHER"
				Headers	  = "Priority", "Cipher"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
}
#endregion SSL Cipher Groups
#endregion SSL
#endregion traffic management

#region Security
$aaavserverscount = (Get-vNetScalerObjectCount -Type authenticationvserver).__count
If ($FEATAAA -eq "Enabled" -and $aaavserverscount -ge 1 -or $FEATAppFw -eq "Enabled"){$selection.InsertNewPage();WriteWordLine 1 0 "Security"}

#region AAA
If ($FEATAAA -eq "Enabled"){
	#region AAA vServers
	If ($aaavserverscount -ge 1) {
		$selection.InsertNewPage()
		WriteWordLine 2 0 "AAA - Virtual Servers"
		$aaavservers = Get-vNetScalerObject -Type authenticationvserver
		foreach ($aaavserver in $aaavservers) {
			WriteWordLine 3 0 "$($aaavserver.name)"
			
			#region AAA vServer Basic Config
			[System.Collections.Hashtable[]] $AAAVSH = @(
				@{ Description = "IP Address"; Value = $aaavserver.ip}
				If (![string]::IsNullOrWhiteSpace($aaavserver.value)){@{ Description = "Value"; Value = $aaavserver.value}}
				@{ Description = "Port"; Value = $aaavserver.port}
				@{ Description = "Service Type"; Value = $aaavserver.servicetype}
				@{ Description = "Type"; Value = $aaavserver.type}
				@{ Description = "State"; Value = $aaavserver.curstate}
				@{ Description = "Status"; Value = $aaavserver.status}
				@{ Description = "Cache Type"; Value = $aaavserver.cachetype}
				If (![string]::IsNullOrWhiteSpace($aaavserver.redirect)){@{ Description = "Redirect"; Value = $aaavserver.redirect}}
				@{ Description = "Precedence"; Value = $aaavserver.precedence}
				If (![string]::IsNullOrWhiteSpace($aaavserver.redirecturl)){@{ Description = "Redirect URL"; Value = $aaavserver.redirecturl}}
				@{ Description = "Authentication"; Value = $aaavserver.authentication}
				If (![string]::IsNullOrWhiteSpace($aaavserver.authenticationdomain)){@{ Description = "Authentication Domain"; Value = $aaavserver.authenticationdomain}}
				If (![string]::IsNullOrWhiteSpace($aaavserver.rule)){@{ Description = "Rule"; Value = $aaavserver.rule}}
				If (![string]::IsNullOrWhiteSpace($aaavserver.policyname)){@{ Description = "Policy Name"; Value = $aaavserver.policyname}}
				If (![string]::IsNullOrWhiteSpace($aaavserver.policy)){@{ Description = "Policy"; Value = $aaavserver.policy}}
				If (![string]::IsNullOrWhiteSpace($aaavserver.servicename)){@{ Description = "Service Name"; Value = $aaavserver.servicename}}
				@{ Description = "Weight"; Value = $aaavserver.weight}
				If (![string]::IsNullOrWhiteSpace($aaavserver.cachevserver)){@{ Description = "Caching vServer"; Value = $aaavserver.cachevserver}}
				If (![string]::IsNullOrWhiteSpace($aaavserver.backupvserver)){@{ Description = "Backup vServer"; Value = $aaavserver.backupvserver}}
				@{ Description = "Client Timeout"; Value = $aaavserver.clttimeout}
				@{ Description = "Spillover Method"; Value = $aaavserver.somethod}
				@{ Description = "Spillover Threshold"; Value = $aaavserver.sothreshold}
				@{ Description = "Spillover Persistence"; Value = $aaavserver.sopersistence}
				@{ Description = "Spillover Persistence Timeout"; Value = $aaavserver.sopersistencetimeout}
				If (![string]::IsNullOrWhiteSpace($aaavserver.priority)){@{ Description = "Priority"; Value = $aaavserver.priority}}
				@{ Description = "Downstate Flush"; Value = $aaavserver.downstateflush}
				@{ Description = "Disable Primary When Down"; Value = $aaavserver.disableprimaryondown}
				@{ Description = "Listen Policy"; Value = $aaavserver.listenpolicy}
				If (![string]::IsNullOrWhiteSpace($aaavserver.listenpriority)){@{ Description = "Listen Priority"; Value = $aaavserver.listenpriority}}
				If (![string]::IsNullOrWhiteSpace($aaavserver.tcpprofilename)){@{ Description = "TCP Profile Name"; Value = $aaavserver.tcpprofilename}}
				If (![string]::IsNullOrWhiteSpace($aaavserver.httpprofilename)){@{ Description = "HTTP Profile Name"; Value = $aaavserver.httpprofilename}}
				If (![string]::IsNullOrWhiteSpace($aaavserver.comment)){@{ Description = "Comment"; Value = $aaavserver.comment}}
				@{ Description = "Enable AppFlow"; Value = $aaavserver.appflowlog}
				@{ Description = "Virtual Server Type"; Value = $aaavserver.vstype}
				If (![string]::IsNullOrWhiteSpace($aaavserver.ngname)){@{ Description = "NetScaler Gateway Name"; Value = $aaavserver.ngname}}
				If (![string]::IsNullOrWhiteSpace($aaavserver.maxloginattempts)){@{ Description = "Max Login Attempts"; Value = $aaavserver.maxloginattempts}}
				If (![string]::IsNullOrWhiteSpace($aaavserver.failedlogintimeout)){@{ Description = "Failed Login Timeout"; Value = $aaavserver.failedlogintimeout}}
				@{ Description = "Secondary"; Value = $aaavserver.secondary}
				@{ Description = "Group Extraction Enabled"; Value = $aaavserver.groupextraction}
			)
			$Params = $null
			$Params = @{
				Hashtable = $AAAVSH
				Columns   = "Description", "Value"
				headers	  = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			#endregion AAA vServer Basic Config

			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationcertpolicy_binding" -BindingTypeName "Certificate Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationldappolicy_binding" -BindingTypeName "LDAP Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationlocalpolicy_binding" -BindingTypeName "Local Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationloginschemapolicy_binding" -BindingTypeName "Login Schema Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationnegotiatepolicy_binding" -BindingTypeName "Negotiate Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationoauthidppolicy_binding" -BindingTypeName "OAuth IDP Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationradiuspolicy_binding" -BindingTypeName "Radius Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationsamlidppolicy_binding" -BindingTypeName "SAML IDP Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationsamlpolicy_binding" -BindingTypeName "SAML Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationtacacspolicy_binding" -BindingTypeName "TACACS Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "authenticationvserver_authenticationwebauthpolicy_binding" -BindingTypeName "WebAuth Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $aaavserver.name -BindingType "sslvserver_sslcertkey_binding" -BindingTypeName "Certificates" -Properties "certkeyname,ca,crlcheck,snicert,ocspcheck,cleartextport" -Headers "Certificate Name,CA Certificate,CRL Checks Enabled,SNI Enabled,OCSP Enabled,Clear Text Port" -Style 4
			New-SSLSettings $aaavserver.name sslvserver 4
			$selection.InsertNewPage()
		}
	}
	#endregion AAA vServers
	
	#region AAA Authentication Profiles
	If ((Get-vNetScalerObjectCount -Type authenticationauthnprofile).__count -ge 1) {
		WriteWordLine 3 0 "Authentication Profiles"
		$Authprfs = Get-vNetScalerObject -Type authenticationauthnprofile
		[System.Collections.Hashtable[]] $AUTHPRFH = @()
		foreach ($Authprf in $Authprfs) {
			WriteWordLine 3 0 "$($Authprf.name)"
			$AUTHPRFH += @{
				authnvsname			= $Authprf.authnvsname
				authenticationhost	= $Authprf.authenticationhost
				authenticationdomain= $Authprf.authenticationdomain
				authenticationlevel	= $Authprf.authenticationlevel
			}
			$Params = $null
			$Params = @{
				Hashtable = $AUTHPRFH
				Columns   = "authnvsname", "authenticationhost", "authenticationdomain", "authenticationlevel"
				Headers   = "Auth vServername", "Authentication Host", "Authentication Domain", "Authentication Level"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion AAA Authentication Profiles
}
#endregion AAA

#region AppFW
If ($FEATAppFw -eq "Enabled"){
	$selection.InsertNewPage()
	WriteWordLine 2 0 "Application Firewall"
		
	#region AppFW Profiles
	If ((Get-vNetScalerObjectCount -Type appfwprofile).__count -ge 1) {
		WriteWordLine 3 0 "Application Firewall Profiles"
		$fwprofiles = Get-vNetScalerObject -Type appfwprofile
		[System.Collections.Hashtable[]] $AFWPROFH = @() 
		foreach ($fwprofile in $fwprofiles) {
			WriteWordLine 4 0 "$($fwprofile.name)"
			[System.Collections.Hashtable[]] $AFWPROFH = @(
				@{ Description = "Description"; Value = "Value"}
				@{ Description = "Profile Type"; Value = $fwprofile.type}
				@{ Description = "StartURL Action"; Value = $fwprofile.starturlaction -join ", "}
				@{ Description = "Content Type Action"; Value = $fwprofile.contenttypeaction -join ", "}
				@{ Description = "Inspect Content Types"; Value = $fwprofile.inspectcontenttypes -join ", "}
				@{ Description = "Start URL Closure"; Value = $fwprofile.starturlclosure}
				@{ Description = "Deny URL Action"; Value = $fwprofile.denyurlaction -join ","}
				@{ Description = "Referer Header Check"; Value = $fwprofile.refererheadercheck}
				@{ Description = "Cookie Consistency Action"; Value = $fwprofile.cookieconsistencyaction -join ", "}
				@{ Description = "Cookie Transformation"; Value = $fwprofile.cookietransforms}
				@{ Description = "Cookie Encryption"; Value = $fwprofile.cookieencryption}
				@{ Description = "Proxy Cookies"; Value = $fwprofile.cookieproxying}
				@{ Description = "Add Cookie Flags"; Value = $fwprofile.addcookieflags}
				@{ Description = "Field Consistency Check"; Value = $fwprofile.fieldconsistencyaction}
				@{ Description = "Cross Site Request Forgery Tag Check"; Value = $fwprofile.csrftagaction -join ", "}
				@{ Description = "XSS (Cross-Site Scripting) Check"; Value = $fwprofile.crosssitescriptingaction}
				@{ Description = "Transform Cross-Site Scripts"; Value = $fwprofile.crosssitescriptingtransformunsafehtml}
				@{ Description = "XSS - Check complete URLs"; Value = $fwprofile.crosssitescriptingcheckcompleteurls}
				@{ Description = "SQL Injection Action"; Value = $fwprofile.sqlinjectionaction -join ", "}
				@{ Description = "SQL Injection - Transform Special Characters"; Value = $fwprofile.sqlinjectiontransformspecialchars}
				@{ Description = "SQL Injection - Only check fields with SQL Characters"; Value = $fwprofile.sqlinjectiononlycheckfieldswithsqlchars}
				@{ Description = "SQL Injection Type"; Value = $fwprofile.sqlinjectiontype}
				@{ Description = "SQL Injection - Check SQL wild characters"; Value = $fwprofile.sqlinjectionchecksqlwildchars}
				@{ Description = "Field Format Actions"; Value = $fwprofile.fieldformataction}
				@{ Description = "Default Field Format Type"; Value = $fwprofile.defaultfieldformattype}
				@{ Description = "Default Field Format minimum length"; Value = $fwprofile.defaultfieldformatminlength}
				@{ Description = "Default Field Format maximum length"; Value = $fwprofile.defaultfieldformatmaxlength}
				@{ Description = "Buffer Overflow Actions"; Value = $fwprofile.bufferoverflowaction}
				@{ Description = "Buffer Overflow - Maximum URL Length"; Value = $fwprofile.bufferoverflowmaxurllength}
				@{ Description = "Buffer Overflow - Maximum Header Length"; Value = $fwprofile.bufferoverflowmaxheaderlength}
				@{ Description = "Buffer Overflow - Maximum Cookie Length"; Value = $fwprofile.bufferoverflowmaxcookielength}
				@{ Description = "Credit Card Action"; Value = $fwprofile.creditcardaction -join ", "}
				@{ Description = "Credit Card Types to protect"; Value = $fwprofile.creditcard -join ", "}
				@{ Description = "Maximum number of Credit Cards per page"; Value = $fwprofile.creditcardmaxallowed}
				@{ Description = "X-Out Credit Card Numbers"; Value = $fwprofile.creditcardxout}
				@{ Description = "Log Credit Card Numbers when matched"; Value = $fwprofile.dosecurecreditcardlogging}
				@{ Description = "Request Streaming"; Value = $fwprofile.streaming}
				@{ Description = "Trace Status"; Value = $fwprofile.trace}
				@{ Description = "Request Content Type"; Value = $fwprofile.requestcontenttype}
				@{ Description = "Response Content Type"; Value = $fwprofile.responsecontenttype}
				@{ Description = "XML Denial of Service Action"; Value = $fwprofile.xmldosaction -join ", "}
				@{ Description = "XML Format Action"; Value = $fwprofile.xmlformataction -join ", " }
				@{ Description = "XML SQL Injection Action"; Value = $fwprofile.xmlsqlinjectionaction -join ", "}
				@{ Description = "XML SQL Injection - Only check fields with SQL characters"; Value = $fwprofile.xmlsqlinjectiononlycheckfieldswithsqlchars}
				@{ Description = "XML SQL Injection - Type"; Value = $fwprofile.xmlsqlinjectiontype}
				@{ Description = "XML SQL Injection - Check fields with SQL Wild characters"; Value = $fwprofile.xmlsqlinjectionchecksqlwildchars}
				@{ Description = "XML SQL Injection - Parse Comments"; Value = $fwprofile.xmlsqlinjectionparsecomments}
				@{ Description = "XML XSS (Cross-Site Scripting) Action"; Value = $fwprofile.xmlxssaction -join ", "}
				@{ Description = "XML WSI (Web Services Interoperability) Action"; Value = $fwprofile.xmlwsiaction -join ", "}
				@{ Description = "XML Attachments Action"; Value = $fwprofile.xmlattachmentaction -join ", "}
				@{ Description = "XML validation Action"; Value = $fwprofile.xmlvalidationaction -join ", "}
				@{ Description = "XML Error Object Name"; Value = $fwprofile.xmlerrorobject}
				@{ Description = "Custom Settings"; Value = $fwprofile.customsettings}
				@{ Description = "Signatures"; Value = $fwprofile.signatures}
				@{ Description = "XML SOAP Fault Action"; Value = $fwprofile.xmlsoapfaultaction -join ", "}
				@{ Description = "Use HTML Error Object"; Value = $fwprofile.usehtmlerrorobject}
				@{ Description = "Error URL"; Value = $fwprofile.errorurl}
				@{ Description = "HTML Error Object Name"; Value = $fwprofile.htmlerrorobject}
				@{ Description = "Log Every Policy Hit"; Value = $fwprofile.logeverypolicyhit}
				@{ Description = "Strip Comments"; Value = $fwprofile.stripcomments}
				@{ Description = "Strip HTML Comments"; Value = $fwprofile.striphtmlcomments}
				@{ Description = "Strip XML Comments"; Value = $fwprofile.sttripxmlcomments}
				@{ Description = "Exempt URLS passing the Start URL Closure check from Security Checks"; Value = $fwprofile.exemptclosureurlsfromsecuritychecks}
				@{ Description = "Default Character Set"; Value = $fwprofile.defaultcharset}
				@{ Description = "Maximum Post Body Size (bytes)"; Value = $fwprofile.postbodylimit}
				@{ Description = "Maximum number of file uploads per form submission"; Value = $fwprofile.fileuploadmaxnum}
				@{ Description = "Perform Entity encoding for special response characters"; Value = $fwprofile.canonicalizehtmlresponse}
				@{ Description = "Enable Form Tagging"; Value = $fwprofile.enableformtagging}
				@{ Description = "Perform Sessionless Field Consistency Checks"; Value = $fwprofile.sessionlessfieldconsistency}
				@{ Description = "Enable Sessionless URL Closure Checks"; Value = $fwprofile.sessionlessurlclosure}
				@{ Description = "Allow Semi-Colon field separator in URL"; Value = $fwprofile.semicolonfieldseparator}
				@{ Description = "Exclude Uploaded Files from Checks"; Value = $fwprofile.excludefileuploadfromchecks}
				@{ Description = "HTML SQL Injection - Parse Comments"; Value = $fwprofile.sqlinjectionparsecomments}
				@{ Description = "Method for handling Percent encoded names"; Value = $fwprofile.invalidpercenthandling}
				@{ Description = "Check Request Headers for SQL Injection and XSS"; Value = $fwprofile.checkrequestheaders}
				@{ Description = "Optimize Partial Requests"; Value = $fwprofile.optimizepartialreqs}
				@{ Description = "URL decode Request Cookies"; Value = $fwprofile.urldecoderequestcookies}
				@{ Description = "Comment"; Value = $fwprofile.comment}
				@{ Description = "Archive Name"; Value = $fwprofile.archivename}
				@{ Description = "State"; Value = $fwprofile.state}
			)
			$Params = $null
			$Params = @{
				Hashtable = $AFWPROFH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
		}
		$selection.InsertNewPage()
	}
	#endregion AppFw Profiles

	#region AppFW Policies
	If ((Get-vNetScalerObjectCount -Type appfwpolicy).__count -ge 1) {
		WriteWordLine 3 0 "Application Firewall Policies"
		$fwpolicies = Get-vNetScalerObject -Type appfwpolicy
		[System.Collections.Hashtable[]] $AFWPOLH = @() 
		foreach ($fwpolicy in $fwpolicies) {
			WriteWordLine 4 0 "$($fwpolicy.name)"     
			[System.Collections.Hashtable[]] $AFWPOLH = @(
				@{ Description = "Description"; Value = "Value"}
				@{ Description = "Rule"; Value = $fwpolicy.rule}
				@{ Description = "Profile Name"; Value = $fwpolicy.profilename}
				@{ Description = "Comment"; Value = $fwpolicy.comment}
				@{ Description = "Log Action"; Value = $fwpolicy.logaction}
			)
			$Params = $null
			$Params = @{
				Hashtable = $AFWPOLH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
		}
		$selection.InsertNewPage()
	}
	#endregion AppFw Policies
}
#endregion AppFW
#endregion Security

#region NetScaler Gateway
If ($FEATSSLVPN -eq "Enabled"){
	$selection.InsertNewPage()
	WriteWordLine 1 0 "NetScaler Gateway"

	#region NetScaler Gateway Global
	WriteWordLine 2 0 "Global Settings"
	$cagglobalclient = Get-vNetScalerObject -Type vpnparameter

	#region Global Network
	WriteWordLine 3 0 "Network"
	[System.Collections.Hashtable[]] $NsGlobalNetwork = @(
		@{ Description = "Description"; Value = "Value"}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.dnsvservername)){@{ Description = "DNS vServer Name"; Value = $cagglobalclient.dnsvservername}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.winsip)){@{ Description = "WINS IP"; Value = $cagglobalclient.winsip}}
		@{ Description = "Kill Connections"; Value = $cagglobalclient.killconnections}
		@{ Description = "ICA Session Timeout"; Value = $cagglobalclient.icasessiontimeout}
		@{ Description = "Use Mapped IP"; Value = $cagglobalclient.usemip}
		@{ Description = "Use Intranet IP"; Value = $cagglobalclient.useiip}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.iipdnssuffix)){@{ Description = "Intranet IP DNS Suffix"; Value = $cagglobalclient.iipdnssuffix}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.httpport)){@{ Description = "HTTP Ports"; Value = $cagglobalclient.httpport}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.forcedtimeout)){@{ Description = "Forced Timeout"; Value = $cagglobalclient.forcedtimeout}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.forcedtimeoutwarning)){@{ Description = "Forced Time-out Warning (mins)"; Value = $cagglobalclient.forcedtimeoutwarning}}
		@{ Description = "Backend Server SNI"; Value = $cagglobalclient.backendserversni}
		@{ Description = "Backend Server Certificate Validation"; Value = $cagglobalclient.backendcertvalidation}
	)
	$Params = $null
	$Params = @{
		Hashtable = $NsGlobalNetwork
		Columns   = "Description", "Value"
	}
	$Table = AddWordTable @Params -List
	FindWordDocumentEnd
	#endregion Global Network

	#region Global Client Experience
	WriteWordLine 3 0 "Client Experience"
	[System.Collections.Hashtable[]] $NsGlobalClientExperience = @(
		@{ Description = "Description"; Value = "Value"}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.homepage)){@{ Description = "Homepage"; Value = $cagglobalclient.homepage}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.emailhome)){@{ Description = "URL for Web Based Email"; Value = $cagglobalclient.emailhome}}
		@{ Description = "Split Tunnel"; Value = $cagglobalclient.splittunnel}
		@{ Description = "Session Time-Out"; Value = $cagglobalclient.sesstimeout}
		@{ Description = "Client-Idle Time-Out"; Value = $cagglobalclient.clientidletimeoutwarning}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.windowsclienttype)){@{ Description = "Plug-in Type"; Value = $cagglobalclient.windowsclienttype}}
		@{ Description = "Windows Plugin Upgrade"; Value = $cagglobalclient.windowspluginupgrade}
		@{ Description = "Linux Plugin Upgrade"; Value = $cagglobalclient.linuxpluginupgrade}
		@{ Description = "MAC Plugin Upgrade"; Value = $cagglobalclient.macpluginupgrade}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.alwaysonprofilename)){@{ Description = "AlwaysON Profile Name"; Value = $cagglobalclient.alwaysonprofilename}}
		@{ Description = "Clientless Access"; Value = $cagglobalclient.clientlessvpnmode}
		@{ Description = "Clientless URL Encoding"; Value = $cagglobalclient.clientlessmodeurlencoding}
		@{ Description = "Clientless Persistent Cookie"; Value = $cagglobalclient.clientlesspersistentcookie}
		@{ Description = "Advanced Clientless VPN Mode"; Value = $cagglobalclient.advancedclientlessvpnmode}
		@{ Description = "Single Sign-On to Web Applications"; Value = $cagglobalclient.sso}
		@{ Description = "Credential Index"; Value = $cagglobalclient.ssocredential}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.kcdaccount)){@{ Description = "KCD Account"; Value = $cagglobalclient.kcdaccount}}
		@{ Description = "Single Sign-On with Windows"; Value = $cagglobalclient.windowsautologon}
		@{ Description = "Client Cleanup Prompt"; Value = $cagglobalclient.clientcleanupprompt}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.uitheme)){@{ Description = "UI Theme"; Value = $cagglobalclient.uitheme}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.loginscript)){@{ Description = "Login Script"; Value = $cagglobalclient.loginscript}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.logoutscript)){@{ Description = "Logout Script"; Value = $cagglobalclient.logoutscript}}
		@{ Description = "Split DNS"; Value = $cagglobalclient.splitdns}
		@{ Description = "Application Token Timeout"; Value = $cagglobalclient.apptokentimeout}
		@{ Description = "MDX Token Timeout"; Value = $cagglobalclient.mdxtokentimeout}
		@{ Description = "Allow Users to Change Log Levels"; Value = $cagglobalclient.clientconfiguration}
		@{ Description = "Local LAN Access"; Value = $cagglobalclient.locallanaccess}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.windowsclienttype)){@{ Description = "Allow access to private network IP addresses only"; Value = $cagglobalclient.windowsclienttype}}
		@{ Description = "Client Choices"; Value = $cagglobalclient.clientchoices}
		@{ Description = "Show VPN Plugin icon"; Value = $cagglobalclient.iconwithreceiver}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.fqdnspoofedip)){@{ Description = "Spoofed IP Address"; Value = $cagglobalclient.fqdnspoofedip}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.netmask)){@{ Description = "Spoofed IP Netmask"; Value = $cagglobalclient.netmask}}
		@{ Description = "Client Force Cleanup"; Value = $cagglobalclient.forcecleanup}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.proxy)){@{ Description = "Proxy"; Value = $cagglobalclient.proxy}}
		If ($cagglobalclient.proxy -eq "BROWSER" -or $cagglobalclient.proxy -eq "NS"){
			If ([string]::IsNullOrWhiteSpace($cagglobalclient.autoproxyurl)){
				If ([string]::IsNullOrWhiteSpace($cagglobalclient.allprotocolproxy)){
					@{ Description = "HTTP"; Value = $cagglobalclient.httpproxy}
					@{ Description = "HTTPS"; Value = $cagglobalclient.sslproxy}
					@{ Description = "FTP"; Value = $cagglobalclient.ftpproxy}
					@{ Description = "Socks"; Value = $cagglobalclient.socksproxy}
					@{ Description = "Gopher"; Value = $cagglobalclient.gopherproxy}
					@{ Description = "Use the same proxy server for all protocols"; Value = "DISABLED"}
				} Else {
					@{ Description = "HTTP"; Value = $cagglobalclient.allprotocolproxy}
					@{ Description = "Use the same proxy server for all protocols"; Value = "ENABLED"}
				}
				@{ Description = "Proxy Exception"; Value = $cagglobalclient.proxyexception}
				@{ Description = "Bypass proxy server for local addresses"; Value = $cagglobalclient.proxylocalbypass}
			} Else {
				@{ Description = "Use Automatic Configuration"; Value = "ENABLED"}
				@{ Description = "URL To Auto Proxy Config File"; Value = $cagglobalclient.autoproxyurl}
			}
		}
	)
	$Params = $null
	$Params = @{
		Hashtable = $NsGlobalClientExperience
		Columns   = "Description", "Value"
	}
	$Table = AddWordTable @Params -List
	FindWordDocumentEnd
	#endregion Global Client Experience

	#region Global Security
	WriteWordLine 3 0 "Security"
	[System.Collections.Hashtable[]] $NsGlobalSecurity = @(
		@{ Description = "Description"; Value = "Value"}
		@{ Description = "Default Authorization Action"; Value = $cagglobalclient.defaultauthorizationaction}
		@{ Description = "Client Security Encryption"; Value = $cagglobalclient.encryptcsecexp}
		@{ Description = "Secure Browse"; Value = $cagglobalclient.securebrowse}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.smartgroup)){@{ Description = "Smartgroup"; Value = $cagglobalclient.smartgroup}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.clientsecurity)){@{ Description = "Client Security"; Value = $cagglobalclient.clientsecurity}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.clientsecuritymessage)){@{ Description = "Error Message"; Value = $cagglobalclient.clientsecuritymessage}}
		@{ Description = "Enable Client Security Logging"; Value = $cagglobalclient.clientsecuritylog}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.authorizationgroup)){@{ Description = "Authorization Group"; Value = $cagglobalclient.authorizationgroup}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.allowedlogingroups)){@{ Description = "Groups Allowed To Login"; Value = $cagglobalclient.allowedlogingroups}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.samesite)){@{ Description = "SameSite"; Value = $cagglobalclient.samesite}}
	)
	$Params = $null
	$Params = @{
		Hashtable = $NsGlobalSecurity
		Columns   = "Description", "Value"
	}
	$Table = AddWordTable @Params -List
	FindWordDocumentEnd
	#endregion Global Security

	#region Global Published Apps
	WriteWordLine 3 0 "Global Settings Published Applications"
	[System.Collections.Hashtable[]] $NsGlobalPublishedApps = @(
		@{ Description = "Description"; Value = "Value"}
		@{ Description = "ICA Proxy"; Value = $cagglobalclient.icaproxy}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.icauseraccounting)){@{ Description = "ICA RADIUS User Accounting"; Value = $cagglobalclient.icauseraccounting}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.wihome)){@{ Description = "Web Interface Address"; Value = $cagglobalclient.wihome}}
		@{ Description = "Web Interface Address Type"; Value = $cagglobalclient.wihomeaddresstype}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.wiportalmode)){@{ Description = "Web Interface Portal Mode"; Value = $cagglobalclient.wiportalmode}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.ntdomain)){@{ Description = "Single Sign-on Domain"; Value = $cagglobalclient.ntdomain}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.citrixreceiverhome)){@{ Description = "Citrix Receiver Home Page"; Value = $cagglobalclient.citrixreceiverhome}}
		If (![string]::IsNullOrWhiteSpace($cagglobalclient.storefronturl)){@{ Description = "Account Services Address"; Value = $cagglobalclient.storefronturl}}
	)
	$Params = $null
	$Params = @{
		Hashtable = $NsGlobalPublishedApps
		Columns   = "Description", "Value"
	}
	$Table = AddWordTable @Params -List
	FindWordDocumentEnd
	#endregion Global Published Apps

	#region Global STA
	If ((Get-vNetScalerObjectCount -Type vpnglobal_staserver_binding).__count -ge 1) {
		WriteWordLine 3 0 "Secure Ticket Authority Configuration"
		$vpnglobalstas = Get-vNetScalerObject -Type vpnglobal_staserver_binding
		[System.Collections.Hashtable[]] $STASH = @()
		foreach ($vpnglobalsta in $vpnglobalstas) {
			$STASH += @{ 
				STA            = $vpnglobalsta.staserver 
				STAADDRESSTYPE = $vpnglobalsta.staaddresstype
				STAAUTHID      = $vpnglobalsta.STAAUTHID
			}
		} 
		$Params = $null
		$Params = @{
			Hashtable = $STASH
			Columns   = "STA", "STAADDRESSTYPE", "STAAUTHID"
			Headers   = "Secure Ticket Authority", "Address Type", "Authentication ID"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
	#endregion Global STA

	#region Global AppController
	If ((Get-vNetScalerObjectCount -Type vpnglobal_appcontroller_binding).__count -ge 1) {
		WriteWordLine 3 0 "App Controller Configuration"    
		$vpnglobalappcs = Get-vNetScalerObject -Type vpnglobal_appcontroller_binding
		[System.Collections.Hashtable[]] $APPCH = @()
		foreach ($vpnglobalappc in $vpnglobalappcs) {
			$APPCH += @{
				APPController = $vpnglobalappc.appController
			} 
		}
		$Params = $null
		$Params = @{
			Hashtable = $APPCH
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
	#endregion Global AppController

	#region Global AAAParams
	WriteWordLine 3 0 "AAA Parameters"
	$cagaaa = Get-vNetScalerObject -Type aaaparameter
	[System.Collections.Hashtable[]] $NsGlobalAAAH = @(
		@{ Description = "Description"; Value = "Value"}
		@{ Description = "Maximum number of Users"; Value = $cagaaa.maxaaausers}
		If (![string]::IsNullOrWhiteSpace($cagaaa.maxloginattempts)){@{ Description = "MaxLogin Attempts"; Value = $cagaaa.maxloginattempts}}
		@{ Description = "NAT IP Address"; Value = $cagaaa.aaadnatip}
		If (![string]::IsNullOrWhiteSpace($cagaaa.failedlogintimeout)){@{ Description = "Failed login timeout"; Value = $cagaaa.failedlogintimeout}}
		@{ Description = "Default Authentication Type"; Value = $cagaaa.defaultauthtype}
		@{ Description = "AAA Session Log Levels"; Value = $cagaaa.aaasessionloglevel}
		@{ Description = "Enable Static Page Caching"; Value = $cagaaa.enablestaticpagecaching}
		@{ Description = "Enable Enhanced Authentication Feedback"; Value = $cagaaa.enableenhancedauthfeedback}
		@{ Description = "Enable Session Stickiness"; Value = $cagaaa.enablesessionstickiness}   
	)
	$Params = $null
	$Params = @{
		Hashtable = $NsGlobalAAAH
		Columns   = "Description", "Value"
	}
	$Table = AddWordTable @Params -List
	FindWordDocumentEnd
	#endregion Global AAAParams
	#endregion NetScaler Gateway Global

	#region NetScaler vServers
	If ((Get-vNetScalerObjectCount -Type vpnvserver).__count -gt 0) {
		$selection.InsertNewPage()
		WriteWordLine 2 0 "Virtual Servers"
		$vpnvservers = Get-vNetScalerObject -Type vpnvserver
		foreach ($vpn in $vpnvservers) {
			WriteWordLine 3 0 "$($vpn.name)"

			#region NetScaler vServer basic configuration
			[System.Collections.Hashtable[]] $VPNVSERVERH = @(
				@{ Description = "Description"; Value = "Configuration" }
				
				# Basic Settings
				@{ Description = "Protocol"; Value = $vpn.servicetype }
				If ($vpn.ipv46 -eq 0.0.0.0){
					@{ Description = "IP Address"; Value = $vpn.ipv46 }
					@{ Description = "Port"; Value = $vpn.port }
				}
				If ($vpn.servicetype -eq "SSL"){
					If (![string]::IsNullOrWhiteSpace($vpn.rdpserverprofilename)){@{ Description = "RDP Server Profile"; Value = $vpn.rdpserverprofilename }}
					If (![string]::IsNullOrWhiteSpace($vpn.pcoipvserverprofilename)){@{ Description = "PCOIP Server Profile"; Value = $vpn.pcoipvserverprofilename }}
					If (![string]::IsNullOrWhiteSpace($vpn.maxaaausers)){@{ Description = "Maximum Users"; Value = $vpn.maxaaausers }}
					If (![string]::IsNullOrWhiteSpace($vpn.maxloginattempts)){@{ Description = "Max Login Attempts"; Value = $vpn.maxloginattempts }}
					If (![string]::IsNullOrWhiteSpace($vpn.failedlogintimeout)){@{ Description = "Failed Login Timeout"; Value = $vpn.failedlogintimeout }}
					If ($vpn.icaonly -eq "ON"){@{ Description = "ICA Only [off]"; Value = $vpn.icaonly }}
					If ($vpn.authentication -eq "OFF"){@{ Description = "Enable Authentication [on]"; Value = $vpn.authentication }}
				}
				If ($vpn.doublehop -eq "ENABLED"){@{ Description = "Double Hop [disabled]"; Value = $vpn.doublehop }}
				If ($vpn.downstateflush -eq "DISABLED"){@{ Description = "Down State Flush [enabled]"; Value = $vpn.downstateflush }}
				If (![string]::IsNullOrWhiteSpace($vpn.range) -and $vpn.range -ne "1"){@{ Description = "IP Range [1]"; Value = $vpn.range }}
				If (![string]::IsNullOrWhiteSpace($vpn.ipset)){@{ Description = "IP Set"; Value = $vpn.ipset }}
				If ($vpn.logoutonsmartcardremoval -eq "ENABLED"){@{ Description = "Logout On Smart Card Removal [disabled]"; Value = $vpn.logoutonsmartcardremoval }}
				If ($vpn.loginonce -eq "ENABLED"){@{ Description = "Login Once [disabled]"; Value = $vpn.loginonce }}
				If ($vpn.servicetype -eq "SSL"){
					If (![string]::IsNullOrWhiteSpace($vpn.windowsepapluginupgrade)){@{ Description = "Windows EPA Plugin Upgrade"; Value = $vpn.windowsepapluginupgrade }}
					If (![string]::IsNullOrWhiteSpace($vpn.linuxepapluginupgrade)){@{ Description = "Linux EPA Plugin Upgrade"; Value = $vpn.linuxepapluginupgrade }}
					If (![string]::IsNullOrWhiteSpace($vpn.macepapluginupgrade)){@{ Description = "Mac EPA Plugin Upgrade"; Value = $vpn.macepapluginupgrade }}
					If ($vpn.dtls -eq "DISABLED"){@{ Description = "DTLS [enabled]"; Value = $vpn.dtls }}
				}
				If ($vpn.appflowlog -eq "DISABLED"){@{ Description = "AppFlow Logging [enabled]"; Value = $vpn.appflowlog }}
				If ($vpn.icaproxysessionmigration -eq "ON"){@{ Description = "ICA Proxy Session Migration [off]"; Value = $vpn.icaproxysessionmigration }}
				If ($vpn.state -eq "DISABLED"){@{ Description = "State [enabled]"; Value = $vpn.state }}
				If (![string]::IsNullOrWhiteSpace($vpn.samesite)){@{ Description = "SameSite"; Value = $vpn.samesite }}
				If ($vpn.servicetype -eq "SSL"){
					If ($vpn.devicecert -eq "ENABLED"){@{ Description = "Enable Device Certificates [disabled]"; Value = $vpn.devicecert }}
					If (![string]::IsNullOrWhiteSpace($vpn.certkeynames)){@{ Description = "CA for Device Certificate"; Value = $vpn.certkeynames -Join ", " }}
				}
				If (![string]::IsNullOrWhiteSpace($vpn.comment)){@{ Description = "Comments"; Value = $vpn.comment }}
				
				# Profiles
				If (![string]::IsNullOrWhiteSpace($vpn.authnprofile)){@{ Description = "Authentication profile"; Value = $vpn.authnprofile }}
				If (![string]::IsNullOrWhiteSpace($vpn.netprofile)){@{ Description = "Net profile"; Value = $vpn.netprofile }}
				If (![string]::IsNullOrWhiteSpace($vpn.tcoprofilename)){@{ Description = "TCP profile"; Value = $vpn.tcoprofilename }}
				If (![string]::IsNullOrWhiteSpace($vpn.httpprofilename)){@{ Description = "HTTP profile"; Value = $vpn.httpprofilename }}
				
				# Other Settings
				If ($vpn.icmpvsrresponse -eq "ACTIVE"){@{ Description = "ICMP Virtual Server Response"; Value = $vpn.icmpvsrresponse }}
				If ($vpn.rhistate -eq "ACTIVE"){@{ Description = "RHI State [passive]"; Value = $vpn.rhistate }}
				If ($vpn.cginfrahomepageredirect -eq "DISABLED"){@{ Description = "Redirect to Home page [enabled]"; Value = $vpn.cginfrahomepageredirect }}
				If (![string]::IsNullOrWhiteSpace($vpn.listenpriority)){@{ Description = "Listen Priority"; Value = $vpn.listenpriority }}
				If ($vpn.listenpolicy -ne "NONE"){@{ Description = "Listen Policy Expression"; Value = $vpn.listenpolicy }}
				If ($vpn.l2conn -eq "ON"){@{ Description = "L2 Connection [off]"; Value = $vpn.l2conn }}
			)
			$Params = $null
			$Params = @{
				Hashtable = $VPNVSERVERH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			#endregion NetScaler vServer basic configuration

			#region NetScaler Policies
			New-BindingTable -Name $vpn.Name -BindingType "sslvserver_sslcertkey_binding" -BindingTypeName "Certificates" -Properties "certkeyname,ca,crlcheck,snicert,ocspcheck,cleartextport" -Headers "Certificate Name,CA Certificate,CRL Checks Enabled,SNI Enabled,OCSP Enabled,Clear Text Port" -Style 4
			New-SSLSettings $vpn.Name sslvserver 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_cachepolicy_binding" -BindingTypeName "Cache Policies" -Properties "priority,policy,bindpoint,gotopriorityexpression" -Headers "Priority,Policy Name,BindPoint,Go To Expression" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_cspolicy_binding" -BindingTypeName "Content Switching Policies" -Properties "priority,policyname,rule,gotopriorityexpression,targetlbvserver" -Headers "Priority,Policy Name,Rule,Goto Expression,Target LB vServer" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationcertpolicy_binding" -BindingTypeName "Authentication Certificate Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationdfapolicy_binding" -BindingTypeName "Authentication DFA Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationldappolicy_binding" -BindingTypeName "Authentication LDAP Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationlocalpolicy_binding" -BindingTypeName "Authentication LOCAL Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationloginschemapolicy_binding" -BindingTypeName "Authentication Loginschema Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationnegotiatepolicy_binding" -BindingTypeName "Authentication Negotiate Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationoauthidppolicy_binding" -BindingTypeName "Authentication NOAUTH Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationpolicy_binding" -BindingTypeName "Authentication Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationradiuspolicy_binding" -BindingTypeName "Authentication Radius Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationsamlidppolicy_binding" -BindingTypeName "Authentication SAML IDP Policies" -Properties "priority,policy" -Headers "Priority,Name" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationsamlpolicy_binding" -BindingTypeName "Authentication SAML Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationtacacspolicy_binding" -BindingTypeName "Authentication Tacacs Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_authenticationwebauthpolicy_binding" -BindingTypeName "Authentication Web Auth Policies" -Properties "priority,policy,secondary" -Headers "Priority,Name,Secondary" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_vpnportaltheme_binding" -BindingTypeName "Portal Theme" -Properties "portaltheme" -Headers "Portal Theme" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_vpnsessionpolicy_binding" -BindingTypeName "Session Policies" -Properties "priority,policy" -Headers "Priority,Policy Name" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_vpntrafficpolicy_binding" -BindingTypeName "Traffic Policies" -Properties "priority,policy" -Headers "Priority,Policy Name" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_staserver_binding" -BindingTypeName "Secure Ticket Authority" -Properties "staserver,staaddresstype,staauthid" -Headers "Secure Ticket Authority,Address Type,Authentication ID" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_responderpolicy_binding" -BindingTypeName "Responder Policies" -Properties "priority,policyname,gotopriorityexpression" -Headers "Priority,Policy Name,Go To Expression" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_rewritepolicy_binding" -BindingTypeName "Rewrite Policies" -Properties "priority,policyname,bindpoint,gotopriorityexpression" -Headers "Priority,Policy Name,Bindpoint,Go To Expression" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_sharefileserver_binding" -BindingTypeName "ShareFile" -Properties "sharefile" -Headers "ShareFile" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_vpnintranetapplication_binding" -BindingTypeName "Intranet Applications" -Properties "intranetapplication" -Headers "Name" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_vpnintranetip_binding" -BindingTypeName "Intranet IP's" -Properties "intranetip,netmask" -Headers "Intranet IP,NetMask" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_vpnintranetip6_binding" -BindingTypeName "Intranet v6 IP's" -Properties "intranetip6,numaddr" -Headers "Intranet IPv6,Number of IPv6 addresses" -Style 4
			New-BindingTable -Name $vpn.Name -BindingType "vpnvserver_vpnurl_binding" -BindingTypeName "Bookmarks" -Properties "urlname" -Headers "Name" -Style 4
			#endregion NetScaler Policies
		}
	}
	#endregion NetScaler vServers

	#region NetScaler Gateway Portal Themes
	WriteWordLine 2 0 "Portal Themes"
	$vpnportalthemes = Get-vNetScalerObject -Type vpnportaltheme
	[System.Collections.Hashtable[]] $THEMESH = @()
	foreach ($vpnportaltheme in $vpnportalthemes) {
		If ($vpnportaltheme.basetheme) {
			$THEMESH += @{ 
				NAME      = $vpnportaltheme.name 
				BASETHEME = $vpnportaltheme.basetheme
			}
		}
		Else {
			$THEMESH += @{ 
				NAME      = $vpnportaltheme.name 
				BASETHEME = "(BUILTIN)"
			}
		}
	}
	$Params = $null
	$Params = @{
		Hashtable = $THEMESH
		Columns   = "NAME", "BASETHEME"
		Headers   = "Theme Name", "Base Theme"
	}
	$Table = AddWordTable @Params
	FindWordDocumentEnd

	#Iterate through each portal theme to extract details
	foreach ($vpnportaltheme in $vpnportalthemes) {
		If ($vpnportaltheme.basetheme) {
			$portalthemename = $vpnportaltheme.name
			WriteWordLine 3 0 "$portalthemename"

			$customcssfile = Get-vNetScalerFile -FileName "custom.css" -FileLocation "/var/netscaler/logon/themes/$portalthemename/css"
			$themecssfile = Get-vNetScalerFile -FileName "theme.css" -FileLocation "/var/netscaler/logon/themes/$portalthemename/css"
			$enxmlfile = Get-vNetScalerFile -FileName "en.xml" -FileLocation "/var/netscaler/logon/themes/$portalthemename/resources"
			$frxmlfile = Get-vNetScalerFile -FileName "fr.xml" -FileLocation "/var/netscaler/logon/themes/$portalthemename/resources"
			$jaxmlfile = Get-vNetScalerFile -FileName "ja.xml" -FileLocation "/var/netscaler/logon/themes/$portalthemename/resources"
			$dexmlfile = Get-vNetScalerFile -FileName "de.xml" -FileLocation "/var/netscaler/logon/themes/$portalthemename/resources"
			$esxmlfile = Get-vNetScalerFile -FileName "es.xml" -FileLocation "/var/netscaler/logon/themes/$portalthemename/resources"

			$customcsscontents = [System.Text.Encoding]::ASCII.Getstring([System.convert]::FromBase64String($customcssfile.filecontent))
			$themecsscontents = [System.Text.Encoding]::ASCII.Getstring([System.convert]::FromBase64String($themecssfile.filecontent))
			[xml]$enxmlcontents = [System.Text.Encoding]::UTF8.Getstring([System.convert]::FromBase64String($enxmlfile.filecontent))
			[xml]$dexmlcontents = [System.Text.Encoding]::UTF8.Getstring([System.convert]::FromBase64String($dexmlfile.filecontent))
			[xml]$esxmlcontents = [System.Text.Encoding]::UTF8.Getstring([System.convert]::FromBase64String($esxmlfile.filecontent))
			[xml]$jaxmlcontents = [System.Text.Encoding]::UTF8.Getstring([System.convert]::FromBase64String($jaxmlfile.filecontent))
			[xml]$frxmlcontents = [System.Text.Encoding]::UTF8.Getstring([System.convert]::FromBase64String($frxmlfile.filecontent))

			#Look and Feel - Home Page
			#Body Background Colour
			$PTBackgroundColour = Get-AttributeFromCSS -SearchPattern "body {" -Attribute "background-color" -Lines 1 -Inputstring $customcsscontents
			#Navigation Pane Background Colour
			$PTNavPaneBackgroundColour = Get-AttributeFromCSS -SearchPattern ".website_section#homepage b:after {" -Attribute "background" -Lines 1 -Inputstring $customcsscontents
			#Navigation Pane Font Colour
			$PTNavPaneFontColour = Get-AttributeFromCSS -SearchPattern ".nav {" -Attribute "color" -Lines 1 -Inputstring $customcsscontents
			#Navigation Selected Tab Background Color	
			$PTNavSelectedTabBackgroundColour = Get-AttributeFromCSS -SearchPattern ".nav .primary li.selected {" -Attribute "background-color :" -Lines 1 -Inputstring $customcsscontents
			#Navigation Selected Tab Font Color
			$PTNavSelectedTabFontColour = Get-AttributeFromCSS -SearchPattern ".nav .primary li.selected {" -Attribute " color :" -Lines 1 -Inputstring $customcsscontents
			#Content Pane Background Color
			$PTContentPaneBackgroundColour = Get-AttributeFromCSS -SearchPattern "#commonBox {" -Attribute "background" -Lines 1 -Inputstring $customcsscontents
			#Button Background Color
			$PTButtonBackgroundColour = Get-AttributeFromCSS -SearchPattern "input.Apply_Cancel_OK {" -Attribute " background:" -Lines 1 -Inputstring $customcsscontents
			#Content Pane Font Color
			$PTContentPaneFontColour = Get-AttributeFromCSS -SearchPattern ".website_section .NUI_Icon table td.cell3 a.bookmark_icon_anchor {" -Attribute "color" -Lines 1 -Inputstring $customcsscontents
			#Content Pane Title Font Color
			$PTContentPaneTitleFontColour = Get-AttributeFromCSS -SearchPattern "#homepage b {" -Attribute "color" -Lines 1 -Inputstring $customcsscontents
			#Bookmarks Description Font Color
			$PTBookmarksDescriptionFontColour = Get-AttributeFromCSS -SearchPattern ".personal_fileshare_section .NUI_Icon table td span.descr {" -Attribute "color" -Lines 1 -Inputstring $customcsscontents
			#Show Enterprise Websites Section
			$PTShowEnterpriseWebsites = Get-AttributeFromCSS -SearchPattern ".enterprise_website_section {" -Attribute "display" -Lines 1 -Inputstring $customcsscontents
			If ($PTShowEnterpriseWebsites -eq "none") { $PTShowEnterpriseWebsites = "Disabled" } Else { $PTShowEnterpriseWebsites = "Enabled" }
			#Show Personal Websites Section
			$PTShowPersonalWebsites = Get-AttributeFromCSS -SearchPattern ".personal_websites_section {" -Attribute "display" -Lines 1 -Inputstring $customcsscontents
			If ($PTShowPersonalWebsites -eq "none" ) { $PTShowPersonalWebsites = "Disabled" } Else { $PTShowPersonalWebsites = "Enabled" }
			#Show File Transfer Tab
			$PTShowFileTransferTab = Get-AttributeFromCSS -SearchPattern ".files-icon {" -Attribute "display" -Lines 1 -Inputstring $customcsscontents
			If ($PTShowFileTransferTab -eq "none") { $PTShowFileTransferTab = "Disabled" } Else { $PTShowFileTransferTab = "Enabled" }
			#Show Enterprise File Shares Section
			$PTShowEnterpriseFileShares = Get-AttributeFromCSS -SearchPattern ".enterprise_fileshare_section {" -Attribute "display" -Lines 1 -Inputstring $customcsscontents
			If (($PTShowEnterpriseFileShares -eq "none") -or ($PTShowFileTransferTab -eq "Disabled")) { $PTShowEnterpriseFileShares = "Disabled" } Else { $PTShowEnterpriseFileShares = "Enabled" }
			#Show Personal File Shares Section
			$PTShowPersonalFileShares = Get-AttributeFromCSS -SearchPattern ".personal_fileshare_section {" -Attribute "display" -Lines 1 -Inputstring $customcsscontents
			If (($PTShowPersonalFileShares -eq "none") -or ($PTShowFileTransferTab -eq "Disabled")) { $PTShowPersonalFileShares = "Disabled" } Else { $PTShowPersonalFileShares = "Enabled" }

			#Look and Feel - Common
			#Background Image
			$PTBackgroundImage = Get-AttributeFromCSS -SearchPattern "body {" -Attribute "background-image" -Lines 1 -Inputstring $customcsscontents
			#Header Background Colour
			$PTHeaderBackgroundColour = Get-AttributeFromCSS -SearchPattern ".header {" -Attribute " background-color :" -Lines 1 -Inputstring $customcsscontents
			#Header Font Colour
			$PTHeaderFontColour = Get-AttributeFromCSS -SearchPattern ".header {" -Attribute " color :" -Lines 1 -Inputstring $customcsscontents
			#Header Border-Bottom Colour
			$PTHeaderBorderBottomColour = Get-AttributeFromCSS -SearchPattern ".header {" -Attribute "border-bottom" -Lines 1 -Inputstring $customcsscontents
			#Header Logo
			$PTHeaderLogo = Get-AttributeFromCSS -SearchPattern ".custom_logo{" -Attribute "background:" -Lines 1 -Inputstring $customcsscontents
			#Center Logo
			$PTCenterLogo = Get-AttributeFromCSS -SearchPattern "#logonbox-logoimage {" -Attribute "background-image" -Lines 1 -Inputstring $customcsscontents
			#Watermark Image
			$PTWatermarkImage = Get-AttributeFromCSS -SearchPattern "#logonbelt-bottomshadow {" -Attribute "background-image" -Lines 1 -Inputstring $customcsscontents
			#Form Font Size
			$PTFormFontSize = Get-AttributeFromCSS -SearchPattern ".form_text {" -Attribute " font-size :" -Lines 1 -Inputstring $customcsscontents
			#Form Font Colour
			$PTFormFontColour = Get-AttributeFromCSS -SearchPattern ".form_text  {" -Attribute " color:" -Lines 1 -Inputstring $customcsscontents
			#Button Image
			$PTButtonImage = Get-AttributeFromCSS -SearchPattern ".custombutton {" -Attribute "background-image" -Lines 1 -Inputstring $customcsscontents
			#Button Hover Image
			$PTButtonHoverImage = Get-AttributeFromCSS -SearchPattern ".custombutton:hover {" -Attribute "background-image" -Lines 1 -Inputstring $customcsscontents
			#Form Title Font Size
			$PTFormTitleFontSize = Get-AttributeFromCSS -SearchPattern ".CTX_ContentTitleHeader {" -Attribute "font-size" -Lines 1 -Inputstring $customcsscontents
			#Form Title Font Colour
			$PTFormTitleFontColour = Get-AttributeFromCSS -SearchPattern ".CTX_ContentTitleHeader { " -Attribute "color" -Lines 1 -Inputstring $customcsscontents
			#Form Background Colour
			$PTFormBackgroundColour = Get-AttributeFromCSS -SearchPattern "#logonbox-innerbox {" -Attribute "background" -Lines 1 -Inputstring $customcsscontents
			#EULA Title Font Size
			$PTEULATitleFontSize = Get-AttributeFromCSS -SearchPattern ".eula_title {" -Attribute "font-size" -Lines 1 -Inputstring $customcsscontents
			#Language

			#region English Strings
			#Login Page
			$enLogonPage = $enxmlcontents.Resources.Partition | ? { $_.id -eq "logon" }
			$ENPageTitle = $enLogonPage.Title
			$ENFormTitle = $enLogonPage.String | ? { $_.id -eq "ctl08_loginAgentCdaHeaderText2" } 
			$ENUserName = $enLogonPage.String | ? { $_.id -eq "User_name" } 
			$ENPassword = $enLogonPage.String | ? { $_.id -eq "Password" } 
			$ENPassword2 = $enLogonPage.String | ? { $_.id -eq "Password2" } 

			# Home Page
			$enPortalPage = $enxmlcontents.Resources.Partition | ? { $_.id -eq "portal" }
			$enFTPage = $enxmlcontents.Resources.Partition | ? { $_.id -eq "ftlist" }
			$enBookmark = $enxmlcontents.Resources.Partition | ? { $_.id -eq "bookmark" }

			$ENWebApps = $enPortalPage.String | ? { $_.id -eq "ctl00_webSites_label" } 
			$ENEntWebApps = $enBookmark.String | ? { $_.id -eq "id_EnterpriseWebSites" }
			$ENPersWebApps = $enBookmark.String | ? { $_.id -eq "id_PersonalWebSites" }
			$ENApps = $enPortalPage.String | ? { $_.id -eq "ctl00_applications_label" }
			$ENFileTrans = $enPortalPage.String | ? { $_.id -eq "id_FileTransfer" }
			$ENEntFile = $enFTPage.String | ? { $_.id -eq "id_EnterpriseFileShares" }
			$ENPersFile = $enFTPage.String | ? { $_.id -eq "id_PersonalFileShares" }
			$ENEmail = $enPortalPage.String | ? { $_.id -eq "id_Email" }

			#VPN Connection
			$enVPNpage = $enxmlcontents.Resources.Partition | ? { $_.id -eq "f_services" }
			$enVPNpagemac = $enxmlcontents.Resources.Partition | ? { $_.id -eq "m_services" }
			$enVPNWaitmsg = $enVPNPage.String | ? { $_.id -eq "waitingmsg" }
			$enVPNproxy = $enVPNPage.String | ? { $_.id -eq "If a proxy server is configured" }
			$enVPNnoplugin = $enVPNPage.String | ? { $_.id -eq "If the Access Gateway Plug-in is not installed" }
			$enVPNnopluginmac = $enVPNPagemac.String | ? { $_.id -eq "If the Access Gateway Plug-in is not installed" }
			$enVPNnopluginlinux = $enVPNPagemac.String | ? { $_.id -eq "If the Access Gateway Linux-Plug-in is not installed" }

			#EPA Page
			$enEPApage = $enxmlcontents.Resources.Partition | ? { $_.id -eq "epa" }
			$enEPATitle = $enEPApage.String | ? { $_.id -eq "loginAgentCdaHeaderText" }
			$enEPAIntro = $enEPApage.String | ? { $_.id -eq "The Access Gateway must confirm that you have the minimum requirements on your device before you can log on." }
			$enEPAPlugin = $enEPApage.String | ? { $_.id -eq "AppINFO" }
			$enEPADownload = $enEPApage.String | ? { $_.id -eq "You do not have the latest version of Endpoint Analysis plug-in installed. Please download the updated plug-in from the link provided" }
			$enEPAPluginError = $enEPApage.String | ? { $_.id -eq "Endpoint Analysis plug-in is either not launched/installed. Please launch or click on the download link provided." }
			$enEPASoftDownload = $enEPApage.String | ? { $_.id -eq "Your device is checked automatically if the Citrix Endpoint Analysis Plug-in software is already installed." }
		  
			#EPA Error Page
			$enEPAErrorpage = $enxmlcontents.Resources.Partition | ? { $_.id -eq "epaerrorpage" }
			$enEPAErrorTitle = $enEPAErrorpage.String | ? { $_.id -eq "Access Denied" }
			$enEPADeviceReqs = $enEPAErrorpage.String | ? { $_.id -eq "Your device does not meet the requirements for logging on." }
			$enEPAMacError = $enEPApage.String | ? { $_.id -eq "End point analysis failed" }
			$enEPAErrorMessage = $enEPAErrorpage.String | ? { $_.id -eq "For more information, contact your help desk and provide the following information:" }
			$enEPAErrorCert = $enEPApage.String | ? { $_.id -eq "Device certificate check failed" }

			#Post EPA Page
			$enEPAPostpage = $enxmlcontents.Resources.Partition | ? { $_.id -eq "postepa" }
			$enEPAPostTitle = $enEPAPostpage.String | ? { $_.id -eq "Checking Your Device" }
			$enEPAPostFail = $enEPAPostpage.String | ? { $_.id -eq "The Endpoint Analysis Plug-in failed to start. " }
			$enEPAPostSkipped = $enEPAPostpage.String | ? { $_.id -eq "The user skipped the scan" }
			#endregion English Strings

			#region French Strings
			#Login Page
			$frLogonPage = $frxmlcontents.Resources.Partition | ? { $_.id -eq "logon" }
			$FRPageTitle = $frLogonPage.Title
			$FRFormTitle = $frLogonPage.String | ? { $_.id -eq "ctl08_loginAgentCdaHeaderText2" } 
			$frUserName = $frLogonPage.String | ? { $_.id -eq "User_name" } 
			$frPassword = $frLogonPage.String | ? { $_.id -eq "Password" } 
			$frPassword2 = $frLogonPage.String | ? { $_.id -eq "Password2" } 

			# Home Page
			$frPortalPage = $frxmlcontents.Resources.Partition | ? { $_.id -eq "portal" }
			$frFTPage = $frxmlcontents.Resources.Partition | ? { $_.id -eq "ftlist" }
			$frBookmark = $frxmlcontents.Resources.Partition | ? { $_.id -eq "bookmark" }

			$frWebApps = $frPortalPage.String | ? { $_.id -eq "ctl00_webSites_label" } 
			$frEntWebApps = $frBookmark.String | ? { $_.id -eq "id_EnterpriseWebSites" }
			$frPersWebApps = $frBookmark.String | ? { $_.id -eq "id_PersonalWebSites" }
			$frApps = $frPortalPage.String | ? { $_.id -eq "ctl00_applications_label" }
			$frFileTrans = $frPortalPage.String | ? { $_.id -eq "id_FileTransfer" }
			$frEntFile = $frFTPage.String | ? { $_.id -eq "id_EnterpriseFileShares" }
			$frPersFile = $frFTPage.String | ? { $_.id -eq "id_PersonalFileShares" }
			$frEmail = $frPortalPage.String | ? { $_.id -eq "id_Email" }

			#VPN Connection
			$frVPNpage = $frxmlcontents.Resources.Partition | ? { $_.id -eq "f_services" }
			$frVPNpagemac = $frxmlcontents.Resources.Partition | ? { $_.id -eq "m_services" }
			$frVPNWaitmsg = $frVPNPage.String | ? { $_.id -eq "waitingmsg" }
			$frVPNproxy = $frVPNPage.String | ? { $_.id -eq "If a proxy server is configured" }
			$frVPNnoplugin = $frVPNPage.String | ? { $_.id -eq "If the Access Gateway Plug-in is not installed" }
			$frVPNnopluginmac = $frVPNPagemac.String | ? { $_.id -eq "If the Access Gateway Plug-in is not installed" }
			$frVPNnopluginlinux = $frVPNPagemac.String | ? { $_.id -eq "If the Access Gateway Linux-Plug-in is not installed" }

			#EPA Page
			$frEPApage = $frxmlcontents.Resources.Partition | ? { $_.id -eq "epa" }
			$frEPATitle = $frEPApage.String | ? { $_.id -eq "loginAgentCdaHeaderText" }
			$frEPAIntro = $frEPApage.String | ? { $_.id -eq "The Access Gateway must confirm that you have the minimum requirements on your device before you can log on." }
			$frEPAPlugin = $frEPApage.String | ? { $_.id -eq "AppINFO" }
			$frEPADownload = $frEPApage.String | ? { $_.id -eq "You do not have the latest version of Endpoint Analysis plug-in installed. Please download the updated plug-in from the link provided" }
			$frEPAPluginError = $frEPApage.String | ? { $_.id -eq "Endpoint Analysis plug-in is either not launched/installed. Please launch or click on the download link provided." }
			$frEPASoftDownload = $frEPApage.String | ? { $_.id -eq "Your device is checked automatically if the Citrix Endpoint Analysis Plug-in software is already installed." }
		  
			#EPA Error Page
			$frEPAErrorpage = $frxmlcontents.Resources.Partition | ? { $_.id -eq "epaerrorpage" }
			$frEPAErrorTitle = $frEPAErrorpage.String | ? { $_.id -eq "Access Denied" }
			$frEPADeviceReqs = $frEPAErrorpage.String | ? { $_.id -eq "Your device does not meet the requirements for logging on." }
			$frEPAMacError = $frEPApage.String | ? { $_.id -eq "End point analysis failed" }
			$frEPAErrorMessage = $frEPAErrorpage.String | ? { $_.id -eq "For more information, contact your help desk and provide the following information:" }
			$frEPAErrorCert = $frEPApage.String | ? { $_.id -eq "Device certificate check failed" }

			#Post EPA Page
			$frEPAPostpage = $frxmlcontents.Resources.Partition | ? { $_.id -eq "postepa" }
			$frEPAPostTitle = $frEPAPostpage.String | ? { $_.id -eq "Checking Your Device" }
			$frEPAPostFail = $frEPAPostpage.String | ? { $_.id -eq "The Endpoint Analysis Plug-in failed to start. " }
			$frEPAPostSkipped = $frEPAPostpage.String | ? { $_.id -eq "The user skipped the scan" }
			#endregion French Strings

			#region German Strings
			#Login Page
			$deLogonPage = $dexmlcontents.Resources.Partition | ? { $_.id -eq "logon" }
			$dePageTitle = $deLogonPage.Title
			$deFormTitle = $deLogonPage.String | ? { $_.id -eq "ctl08_loginAgentCdaHeaderText2" } 
			$deUserName = $deLogonPage.String | ? { $_.id -eq "User_name" } 
			$dePassword = $deLogonPage.String | ? { $_.id -eq "Password" } 
			$dePassword2 = $deLogonPage.String | ? { $_.id -eq "Password2" } 

			# Home Page
			$dePortalPage = $dexmlcontents.Resources.Partition | ? { $_.id -eq "portal" }
			$deFTPage = $dexmlcontents.Resources.Partition | ? { $_.id -eq "ftlist" }
			$deBookmark = $dexmlcontents.Resources.Partition | ? { $_.id -eq "bookmark" }

			$deWebApps = $dePortalPage.String | ? { $_.id -eq "ctl00_webSites_label" } 
			$deEntWebApps = $deBookmark.String | ? { $_.id -eq "id_EnterpriseWebSites" }
			$dePersWebApps = $deBookmark.String | ? { $_.id -eq "id_PersonalWebSites" }
			$deApps = $dePortalPage.String | ? { $_.id -eq "ctl00_applications_label" }
			$deFileTrans = $dePortalPage.String | ? { $_.id -eq "id_FileTransfer" }
			$deEntFile = $deFTPage.String | ? { $_.id -eq "id_EnterpriseFileShares" }
			$dePersFile = $deFTPage.String | ? { $_.id -eq "id_PersonalFileShares" }
			$deEmail = $dePortalPage.String | ? { $_.id -eq "id_Email" }

			#VPN Connection
			$deVPNpage = $dexmlcontents.Resources.Partition | ? { $_.id -eq "f_services" }
			$deVPNpagemac = $dexmlcontents.Resources.Partition | ? { $_.id -eq "m_services" }
			$deVPNWaitmsg = $deVPNPage.String | ? { $_.id -eq "waitingmsg" }
			$deVPNproxy = $deVPNPage.String | ? { $_.id -eq "If a proxy server is configured" }
			$deVPNnoplugin = $deVPNPage.String | ? { $_.id -eq "If the Access Gateway Plug-in is not installed" }
			$deVPNnopluginmac = $deVPNPagemac.String | ? { $_.id -eq "If the Access Gateway Plug-in is not installed" }
			$deVPNnopluginlinux = $deVPNPagemac.String | ? { $_.id -eq "If the Access Gateway Linux-Plug-in is not installed" }

			#EPA Page
			$deEPApage = $dexmlcontents.Resources.Partition | ? { $_.id -eq "epa" }
			$deEPATitle = $deEPApage.String | ? { $_.id -eq "loginAgentCdaHeaderText" }
			$deEPAIntro = $deEPApage.String | ? { $_.id -eq "The Access Gateway must confirm that you have the minimum requirements on your device before you can log on." }
			$deEPAPlugin = $deEPApage.String | ? { $_.id -eq "AppINFO" }
			$deEPADownload = $deEPApage.String | ? { $_.id -eq "You do not have the latest version of Endpoint Analysis plug-in installed. Please download the updated plug-in from the link provided" }
			$deEPAPluginError = $deEPApage.String | ? { $_.id -eq "Endpoint Analysis plug-in is either not launched/installed. Please launch or click on the download link provided." }
			$deEPASoftDownload = $deEPApage.String | ? { $_.id -eq "Your device is checked automatically if the Citrix Endpoint Analysis Plug-in software is already installed." }
		  
			#EPA Error Page
			$deEPAErrorpage = $dexmlcontents.Resources.Partition | ? { $_.id -eq "epaerrorpage" }
			$deEPAErrorTitle = $deEPAErrorpage.String | ? { $_.id -eq "Access Denied" }
			$deEPADeviceReqs = $deEPAErrorpage.String | ? { $_.id -eq "Your device does not meet the requirements for logging on." }
			$deEPAMacError = $deEPApage.String | ? { $_.id -eq "End point analysis failed" }
			$deEPAErrorMessage = $deEPAErrorpage.String | ? { $_.id -eq "For more information, contact your help desk and provide the following information:" }
			$deEPAErrorCert = $deEPApage.String | ? { $_.id -eq "Device certificate check failed" }

			#Post EPA Page
			$deEPAPostpage = $dexmlcontents.Resources.Partition | ? { $_.id -eq "postepa" }
			$deEPAPostTitle = $deEPAPostpage.String | ? { $_.id -eq "Checking Your Device" }
			$deEPAPostFail = $deEPAPostpage.String | ? { $_.id -eq "The Endpoint Analysis Plug-in failed to start. " }
			$deEPAPostSkipped = $deEPAPostpage.String | ? { $_.id -eq "The user skipped the scan" }
			#endregion German Strings

			#region Spanish Strings
			#Login Page
			$esLogonPage = $esxmlcontents.Resources.Partition | ? { $_.id -eq "logon" }
			$esPageTitle = $esLogonPage.Title
			$esFormTitle = $esLogonPage.String | ? { $_.id -eq "ctl08_loginAgentCdaHeaderText2" } 
			$esUserName = $esLogonPage.String | ? { $_.id -eq "User_name" } 
			$esPassword = $esLogonPage.String | ? { $_.id -eq "Password" } 
			$esPassword2 = $esLogonPage.String | ? { $_.id -eq "Password2" } 

			# Home Page
			$esPortalPage = $esxmlcontents.Resources.Partition | ? { $_.id -eq "portal" }
			$esFTPage = $esxmlcontents.Resources.Partition | ? { $_.id -eq "ftlist" }
			$esBookmark = $esxmlcontents.Resources.Partition | ? { $_.id -eq "bookmark" }

			$esWebApps = $esPortalPage.String | ? { $_.id -eq "ctl00_webSites_label" } 
			$esEntWebApps = $esBookmark.String | ? { $_.id -eq "id_EnterpriseWebSites" }
			$esPersWebApps = $esBookmark.String | ? { $_.id -eq "id_PersonalWebSites" }
			$esApps = $esPortalPage.String | ? { $_.id -eq "ctl00_applications_label" }
			$esFileTrans = $esPortalPage.String | ? { $_.id -eq "id_FileTransfer" }
			$esEntFile = $esFTPage.String | ? { $_.id -eq "id_EnterpriseFileShares" }
			$esPersFile = $esFTPage.String | ? { $_.id -eq "id_PersonalFileShares" }
			$esEmail = $esPortalPage.String | ? { $_.id -eq "id_Email" }

			#VPN Connection
			$esVPNpage = $esxmlcontents.Resources.Partition | ? { $_.id -eq "f_services" }
			$esVPNpagemac = $esxmlcontents.Resources.Partition | ? { $_.id -eq "m_services" }
			$esVPNWaitmsg = $esVPNPage.String | ? { $_.id -eq "waitingmsg" }
			$esVPNproxy = $esVPNPage.String | ? { $_.id -eq "If a proxy server is configured" }
			$esVPNnoplugin = $esVPNPage.String | ? { $_.id -eq "If the Access Gateway Plug-in is not installed" }
			$esVPNnopluginmac = $esVPNPagemac.String | ? { $_.id -eq "If the Access Gateway Plug-in is not installed" }
			$esVPNnopluginlinux = $esVPNPagemac.String | ? { $_.id -eq "If the Access Gateway Linux-Plug-in is not installed" }

			#EPA Page
			$esEPApage = $esxmlcontents.Resources.Partition | ? { $_.id -eq "epa" }
			$esEPATitle = $esEPApage.String | ? { $_.id -eq "loginAgentCdaHeaderText" }
			$esEPAIntro = $esEPApage.String | ? { $_.id -eq "The Access Gateway must confirm that you have the minimum requirements on your device before you can log on." }
			$esEPAPlugin = $esEPApage.String | ? { $_.id -eq "AppINFO" }
			$esEPADownload = $esEPApage.String | ? { $_.id -eq "You do not have the latest version of Endpoint Analysis plug-in installed. Please download the updated plug-in from the link provided" }
			$esEPAPluginError = $esEPApage.String | ? { $_.id -eq "Endpoint Analysis plug-in is either not launched/installed. Please launch or click on the download link provided." }
			$esEPASoftDownload = $esEPApage.String | ? { $_.id -eq "Your device is checked automatically if the Citrix Endpoint Analysis Plug-in software is already installed." }
		  
			#EPA Error Page
			$esEPAErrorpage = $esxmlcontents.Resources.Partition | ? { $_.id -eq "epaerrorpage" }
			$esEPAErrorTitle = $esEPAErrorpage.String | ? { $_.id -eq "Access Denied" }
			$esEPADeviceReqs = $esEPAErrorpage.String | ? { $_.id -eq "Your device does not meet the requirements for logging on." }
			$esEPAMacError = $esEPApage.String | ? { $_.id -eq "End point analysis failed" }
			$esEPAErrorMessage = $esEPAErrorpage.String | ? { $_.id -eq "For more information, contact your help desk and provide the following information:" }
			$esEPAErrorCert = $esEPApage.String | ? { $_.id -eq "Device certificate check failed" }

			#Post EPA Page
			$esEPAPostpage = $esxmlcontents.Resources.Partition | ? { $_.id -eq "postepa" }
			$esEPAPostTitle = $esEPAPostpage.String | ? { $_.id -eq "Checking Your Device" }
			$esEPAPostFail = $esEPAPostpage.String | ? { $_.id -eq "The Endpoint Analysis Plug-in failed to start. " }
			$esEPAPostSkipped = $esEPAPostpage.String | ? { $_.id -eq "The user skipped the scan" }
			#endregion Spanish Strings

			#region Japanese Strings
			#Login Page
			$jaLogonPage = $jaxmlcontents.Resources.Partition | ? { $_.id -eq "logon" }
			$jaPageTitle = $jaLogonPage.Title
			$jaFormTitle = $jaLogonPage.String | ? { $_.id -eq "ctl08_loginAgentCdaHeaderText2" } 
			$jaUserName = $jaLogonPage.String | ? { $_.id -eq "User_name" } 
			$jaPassword = $jaLogonPage.String | ? { $_.id -eq "Password" } 
			$jaPassword2 = $jaLogonPage.String | ? { $_.id -eq "Password2" } 

			# Home Page
			$jaPortalPage = $jaxmlcontents.Resources.Partition | ? { $_.id -eq "portal" }
			$jaFTPage = $jaxmlcontents.Resources.Partition | ? { $_.id -eq "ftlist" }
			$jaBookmark = $jaxmlcontents.Resources.Partition | ? { $_.id -eq "bookmark" }

			$jaWebApps = $jaPortalPage.String | ? { $_.id -eq "ctl00_webSites_label" } 
			$jaEntWebApps = $jaBookmark.String | ? { $_.id -eq "id_EnterpriseWebSites" }
			$jaPersWebApps = $jaBookmark.String | ? { $_.id -eq "id_PersonalWebSites" }
			$jaApps = $jaPortalPage.String | ? { $_.id -eq "ctl00_applications_label" }
			$jaFileTrans = $jaPortalPage.String | ? { $_.id -eq "id_FileTransfer" }
			$jaEntFile = $jaFTPage.String | ? { $_.id -eq "id_EnterpriseFileShares" }
			$jaPersFile = $jaFTPage.String | ? { $_.id -eq "id_PersonalFileShares" }
			$jaEmail = $jaPortalPage.String | ? { $_.id -eq "id_Email" }

			#VPN Connection
			$jaVPNpage = $jaxmlcontents.Resources.Partition | ? { $_.id -eq "f_services" }
			$jaVPNpagemac = $jaxmlcontents.Resources.Partition | ? { $_.id -eq "m_services" }
			$jaVPNWaitmsg = $jaVPNPage.String | ? { $_.id -eq "waitingmsg" }
			$jaVPNproxy = $jaVPNPage.String | ? { $_.id -eq "If a proxy server is configured" }
			$jaVPNnoplugin = $jaVPNPage.String | ? { $_.id -eq "If the Access Gateway Plug-in is not installed" }
			$jaVPNnopluginmac = $jaVPNPagemac.String | ? { $_.id -eq "If the Access Gateway Plug-in is not installed" }
			$jaVPNnopluginlinux = $jaVPNPagemac.String | ? { $_.id -eq "If the Access Gateway Linux-Plug-in is not installed" }

			#EPA Page
			$jaEPApage = $jaxmlcontents.Resources.Partition | ? { $_.id -eq "epa" }
			$jaEPATitle = $jaEPApage.String | ? { $_.id -eq "loginAgentCdaHeaderText" }
			$jaEPAIntro = $jaEPApage.String | ? { $_.id -eq "The Access Gateway must confirm that you have the minimum requirements on your device before you can log on." }
			$jaEPAPlugin = $jaEPApage.String | ? { $_.id -eq "AppINFO" }
			$jaEPADownload = $jaEPApage.String | ? { $_.id -eq "You do not have the latest version of Endpoint Analysis plug-in installed. Please download the updated plug-in from the link provided" }
			$jaEPAPluginError = $jaEPApage.String | ? { $_.id -eq "Endpoint Analysis plug-in is either not launched/installed. Please launch or click on the download link provided." }
			$jaEPASoftDownload = $jaEPApage.String | ? { $_.id -eq "Your device is checked automatically if the Citrix Endpoint Analysis Plug-in software is already installed." }
		  
			#EPA Error Page
			$jaEPAErrorpage = $jaxmlcontents.Resources.Partition | ? { $_.id -eq "epaerrorpage" }
			$jaEPAErrorTitle = $jaEPAErrorpage.String | ? { $_.id -eq "Access Denied" }
			$jaEPADeviceReqs = $jaEPAErrorpage.String | ? { $_.id -eq "Your device does not meet the requirements for logging on." }
			$jaEPAMacError = $jaEPApage.String | ? { $_.id -eq "End point analysis failed" }
			$jaEPAErrorMessage = $jaEPAErrorpage.String | ? { $_.id -eq "For more information, contact your help desk and provide the following information:" }
			$jaEPAErrorCert = $jaEPApage.String | ? { $_.id -eq "Device certificate check failed" }

			#Post EPA Page
			$jaEPAPostpage = $jaxmlcontents.Resources.Partition | ? { $_.id -eq "postepa" }
			$jaEPAPostTitle = $jaEPAPostpage.String | ? { $_.id -eq "Checking Your Device" }
			$jaEPAPostFail = $jaEPAPostpage.String | ? { $_.id -eq "The Endpoint Analysis Plug-in failed to start. " }
			$jaEPAPostSkipped = $jaEPAPostpage.String | ? { $_.id -eq "The user skipped the scan" }
			#endregion Japanese Strings

			#Look And Feel Table
			WriteWordLine 4 0 "Look and Feel - Home Page"
			
			$PTLANDFH = $null
			[System.Collections.Hashtable[]] $PTLANDFH = @(
				@{ Description = "Attribute"; Value = "Setting"}
				@{ Description = "Body Backgound Colour"; Value = $PTBackgroundColour}
				@{ Description = "Navigation Pane Background Colour"; Value = $PTNavPaneBackgroundColour}
				@{ Description = "Navigation Pane Font Colour"; Value = $PTNavPaneFontColour}
				@{ Description = "Navigation Selected Tab Background Color"; Value = $PTNavSelectedTabBackgroundColour}
				@{ Description = "Navigation Selected Tab Font Color"; Value = $PTNavSelectedTabFontColour}
				@{ Description = "Content Pane Background Color"; Value = $PTContentPaneBackgroundColour}
				@{ Description = "Button Background Color"; Value = $PTButtonBackgroundColour}
				@{ Description = "Content Pane Font Color"; Value = $PTContentPaneFontColour}
				@{ Description = "Content Pane Title Font Color"; Value = $PTContentPaneTitleFontColour}
				@{ Description = "Bookmarks Description Font Color"; Value = $PTBookmarksDescriptionFontColour}
				@{ Description = "Show Enterprise Websites Section"; Value = $PTShowEnterpriseWebsites}
				@{ Description = "Show Personal Websites Section"; Value = $PTShowPersonalWebsites}
				@{ Description = "Show File Transfer Tab"; Value = $PTShowFileTransferTab}
				@{ Description = "Show Enterprise File Shares Section"; Value = $PTShowEnterpriseFileShares}
				@{ Description = "Show Personal File Shares Section"; Value = $PTShowPersonalFileShares}    
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTLANDFH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			WriteWordLine 4 0 "Look and Feel - Common"
			$PTCOMMONH = $null
			[System.Collections.Hashtable[]] $PTCOMMONH = @(
				@{ Description = "Attribute"; Value = "Setting"}
				@{ Description = "Background Image"; Value = $PTBackgroundImage}
				@{ Description = "Header Background Colour"; Value = $PTHeaderBackgroundColour}
				@{ Description = "Header Font Colour"; Value = $PTHeaderFontColour}
				@{ Description = "Header Border-Bottom Colour"; Value = $PTHeaderBorderBottomColour}
				@{ Description = "Header Logo"; Value = $PTHeaderLogo}
				@{ Description = "Center Logo"; Value = $PTCenterLogo}
				@{ Description = "Watermark Image"; Value = $PTWatermarkImage}
				@{ Description = "Form Font Size"; Value = $PTFormFontSize}
				@{ Description = "Form Font Colour"; Value = $PTFormFontColour}
				@{ Description = "Button Image"; Value = $PTButtonImage}
				@{ Description = "Button Hover Image"; Value = $PTButtonHoverImage}
				@{ Description = "Form Title Font Size"; Value = $PTFormTitleFontSize}
				@{ Description = "Form Title Font Colour"; Value = $PTFormTitleFontColour}
				@{ Description = "Form Background Colour"; Value = $PTFormBackgroundColour}
				@{ Description = "EULA Title Font Size"; Value = $PTEULATitleFontSize}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTCOMMONH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			#region English Language
			WriteWordLine 4 0 "English Language"
			$PTENLOGINH = $null
			[System.Collections.Hashtable[]] $PTENLOGINH = @(
				@{ Description = "Login Page"; Value = ""}
				@{ Description = "Page Title"; Value = $ENPageTitle.InnerText}
				@{ Description = "Form Title"; Value = $ENFormTitle.InnerText}
				@{ Description = "User Name Field Title"; Value = $ENUserName.InnerText}
				@{ Description = "Password Field Title"; Value = $ENPassword.InnerText}
				@{ Description = "Password Field2 Title"; Value = $ENPassword2.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTENLOGINH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTENHOMEH = $null
			[System.Collections.Hashtable[]] $PTENHOMEH = @(
				@{ Description = "Home Page"; Value = ""}
				@{ Description = "Web Apps Tab Label"; Value = $ENWebApps.InnerText}
				@{ Description = "Enterprise Web Sites Label"; Value = $ENEntWebapps.InnerText}
				@{ Description = "Personal Web Sites Label"; Value = $ENPersWebapps.InnerText}
				@{ Description = "Applications Tab Label"; Value = $ENApps.InnerText}
				@{ Description = "File Transfer Tab Label"; Value = $ENFileTrans.InnerText}
				@{ Description = "Enterprise File Shares Label"; Value = $ENEntFile.InnerText}
				@{ Description = "Personal File Shares Label"; Value = $ENPersFile.InnerText}
				@{ Description = "Email Tab Label"; Value = $ENEmail.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTENHOMEH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTENVPNH = $null
			[System.Collections.Hashtable[]] $PTENVPNH = @(
				@{ Description = "VPN Connection"; Value = ""}
				@{ Description = "Waiting Message"; Value = $ENVPNWaitmsg.InnerText}
				@{ Description = "Proxy Configured message"; Value = $ENVPNproxy.InnerText}
				@{ Description = "Windows Plug-in Not Installed Message"; Value = $ENVPNnoplugin.InnerText}
				@{ Description = "MAC Plug-in Not Installed Message"; Value = $ENVPNnopluginmac.InnerText}
				@{ Description = "Linux Plug-in Not Installed Message"; Value = $ENVPNnopluginlinux.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTENVPNH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTENEPAH = $null
			[System.Collections.Hashtable[]] $PTENEPAH = @(
				@{ Description = "EPA Page"; Value = ""}
				@{ Description = "Title"; Value = $ENEPATitle.InnerText}
				@{ Description = "Introductory Message"; Value = $ENEPAIntro.InnerText}
				@{ Description = "Plug-in Check Message"; Value = $ENEPAPlugin.InnerText}
				@{ Description = "Download Plug-In Message"; Value = $ENEPADownload.InnerText}
				@{ Description = "Plug-in Launch Error Message"; Value = $ENEPAPluginError.InnerText}
				@{ Description = "Download Software Message"; Value = $ENEPASoftDownload.InnerText}    
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTENEPAH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTENEPAERRH = $null
			[System.Collections.Hashtable[]] $PTENEPAERRH = @(
				@{ Description = "EPA Error Page"; Value = ""}
				@{ Description = "Error Title"; Value = $ENEPAErrorTitle.InnerText}
				@{ Description = "Device Requirements Not Matching Message"; Value = $ENEPADeviceReqs.InnerText}
				@{ Description = "Mac Failure Message"; Value = $ENEPAMacError.InnerText}
				@{ Description = "Error More Info Message"; Value = $ENEPAErrorMessage.InnerText}
				@{ Description = "Device Certificate Check Failure Message"; Value = $ENEPAErrorCert.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTENEPAERRH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTENPOSTEPAH = $null
			[System.Collections.Hashtable[]] $PTENPOSTEPAH = @(
				@{ Description = "Post EPA Page"; Value = ""}
				@{ Description = "Title"; Value = $ENEPAPostTitle.InnerText}
				@{ Description = "Failure To Start Message"; Value = $ENEPAPostFail.InnerText}
				@{ Description = "User Skipped Scan Message"; Value = $ENEPAPostSkipped.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTENPOSTEPAH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			#endregion English Language

			#region French Language
			WriteWordLine 4 0 "French Language"
			$PTFRLOGINH = $null
			[System.Collections.Hashtable[]] $PTFRLOGINH = @(
				@{ Description = "Login Page"; Value = ""}
				@{ Description = "Page Title"; Value = $FRPageTitle.InnerText}
				@{ Description = "Form Title"; Value = $FRFormTitle.InnerText}
				@{ Description = "User Name Field Title"; Value = $frUserName.InnerText}
				@{ Description = "Password Field Title"; Value = $frPassword.InnerText}
				@{ Description = "Password Field2 Title"; Value = $frPassword2.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTFRLOGINH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTFRHOMEH = $null
			[System.Collections.Hashtable[]] $PTFRHOMEH = @(
				@{ Description = "Home Page"; Value = ""}
				@{ Description = "Web Apps Tab Label"; Value = $FRWebApps.InnerText}
				@{ Description = "Enterprise Web Sites Label"; Value = $FREntWebapps.InnerText}
				@{ Description = "Personal Web Sites Label"; Value = $FRPersWebapps.InnerText}
				@{ Description = "Applications Tab Label"; Value = $FRApps.InnerText}
				@{ Description = "File Transfer Tab Label"; Value = $FRFileTrans.InnerText}
				@{ Description = "Enterprise File Shares Label"; Value = $FREntFile.InnerText}
				@{ Description = "Personal File Shares Label"; Value = $FRPersFile.InnerText}
				@{ Description = "Email Tab Label"; Value = $FREmail.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTFRHOMEH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTFRVPNH = $null
			[System.Collections.Hashtable[]] $PTFRVPNH = @(
				@{ Description = "VPN Connection"; Value = ""}
				@{ Description = "Waiting Message"; Value = $FRVPNWaitmsg.InnerText}
				@{ Description = "Proxy Configured message"; Value = $FRVPNproxy.InnerText}
				@{ Description = "Windows Plug-in Not Installed Message"; Value = $FRVPNnoplugin.InnerText}
				@{ Description = "MAC Plug-in Not Installed Message"; Value = $FRVPNnopluginmac.InnerText}
				@{ Description = "Linux Plug-in Not Installed Message"; Value = $FRVPNnopluginlinux.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTFRVPNH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTFREPAH = $null
			[System.Collections.Hashtable[]] $PTFREPAH = @(
				@{ Description = "EPA Page"; Value = ""}
				@{ Description = "Title"; Value = $FREPATitle.InnerText}
				@{ Description = "Introductory Message"; Value = $FREPAIntro.InnerText}
				@{ Description = "Plug-in Check Message"; Value = $FREPAPlugin.InnerText}
				@{ Description = "Download Plug-In Message"; Value = $FREPADownload.InnerText}
				@{ Description = "Plug-in Launch Error Message"; Value = $FREPAPluginError.InnerText}
				@{ Description = "Download Software Message"; Value = $FREPASoftDownload.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTFREPAH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTFREPAERRH = $null
			[System.Collections.Hashtable[]] $PTFREPAERRH = @(
				@{ Description = "EPA Error Page"; Value = ""}
				@{ Description = "Error Title"; Value = $FREPAErrorTitle.InnerText}
				@{ Description = "Device Requirements Not Matching Message"; Value = $FREPADeviceReqs.InnerText}
				@{ Description = "Mac Failure Message"; Value = $FREPAMacError.InnerText}
				@{ Description = "Error More Info Message"; Value = $FREPAErrorMessage.InnerText}
				@{ Description = "Device Certificate Check Failure Message"; Value = $FREPAErrorCert.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTFREPAERRH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTFRPOSTEPAH = $null
			[System.Collections.Hashtable[]] $PTFRPOSTEPAH = @(
				@{ Description = "Post EPA Page"; Value = ""}
				@{ Description = "Title"; Value = $FREPAPostTitle.InnerText}
				@{ Description = "Failure To Start Message"; Value = $FREPAPostFail.InnerText}
				@{ Description = "User Skipped Scan Message"; Value = $FREPAPostSkipped.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTFRPOSTEPAH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			#endregion French Language

			#region German Language
			WriteWordLine 4 0 "German Language"
			$PTDELOGINH = $null
			[System.Collections.Hashtable[]] $PTDELOGINH = @(
				@{ Description = "Login Page"; Value = ""}
				@{ Description = "Page Title"; Value = $DEPageTitle.InnerText}
				@{ Description = "Form Title"; Value = $DEFormTitle.InnerText}
				@{ Description = "User Name Field Title"; Value = $DEUserName.InnerText}
				@{ Description = "Password Field Title"; Value = $DEPassword.InnerText}
				@{ Description = "Password Field2 Title"; Value = $DEPassword2.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTDELOGINH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTDEHOMEH = $null
			[System.Collections.Hashtable[]] $PTDEHOMEH = @(
				@{ Description = "Home Page"; Value = ""}
				@{ Description = "Web Apps Tab Label"; Value = $DEWebApps.InnerText}
				@{ Description = "Enterprise Web Sites Label"; Value = $DEEntWebapps.InnerText}
				@{ Description = "Personal Web Sites Label"; Value = $DEPersWebapps.InnerText}
				@{ Description = "Applications Tab Label"; Value = $DEApps.InnerText}
				@{ Description = "File Transfer Tab Label"; Value = $DEFileTrans.InnerText}
				@{ Description = "Enterprise File Shares Label"; Value = $DEEntFile.InnerText}
				@{ Description = "Personal File Shares Label"; Value = $DEPersFile.InnerText}
				@{ Description = "Email Tab Label"; Value = $DEEmail.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTDEHOMEH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTDEVPNH = $null
			[System.Collections.Hashtable[]] $PTDEVPNH = @(
				@{ Description = "VPN Connection"; Value = ""}
				@{ Description = "Waiting Message"; Value = $DEVPNWaitmsg.InnerText}
				@{ Description = "Proxy Configured message"; Value = $DEVPNproxy.InnerText}
				@{ Description = "Windows Plug-in Not Installed Message"; Value = $DEVPNnoplugin.InnerText}
				@{ Description = "MAC Plug-in Not Installed Message"; Value = $DEVPNnopluginmac.InnerText}
				@{ Description = "Linux Plug-in Not Installed Message"; Value = $DEVPNnopluginlinux.InnerText}    
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTDEVPNH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTDEEPAH = $null
			[System.Collections.Hashtable[]] $PTDEEPAH = @(
				@{ Description = "EPA Page"; Value = ""}
				@{ Description = "Title"; Value = $DEEPATitle.InnerText}
				@{ Description = "Introductory Message"; Value = $DEEPAIntro.InnerText}
				@{ Description = "Plug-in Check Message"; Value = $DEEPAPlugin.InnerText}
				@{ Description = "Download Plug-In Message"; Value = $DEEPADownload.InnerText}
				@{ Description = "Plug-in Launch Error Message"; Value = $DEEPAPluginError.InnerText}
				@{ Description = "Download Software Message"; Value = $DEEPASoftDownload.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTDEEPAH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTDEEPAERRH = $null
			[System.Collections.Hashtable[]] $PTDEEPAERRH = @(
				@{ Description = "EPA Error Page"; Value = ""}
				@{ Description = "Error Title"; Value = $DEEPAErrorTitle.InnerText}
				@{ Description = "Device Requirements Not Matching Message"; Value = $DEEPADeviceReqs.InnerText}
				@{ Description = "Mac Failure Message"; Value = $DEEPAMacError.InnerText}
				@{ Description = "Error More Info Message"; Value = $DEEPAErrorMessage.InnerText}
				@{ Description = "Device Certificate Check Failure Message"; Value = $DEEPAErrorCert.InnerText}    
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTDEEPAERRH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTDEPOSTEPAH = $null
			[System.Collections.Hashtable[]] $PTDEPOSTEPAH = @(
				@{ Description = "Post EPA Page"; Value = ""}
				@{ Description = "Title"; Value = $DEEPAPostTitle.InnerText}
				@{ Description = "Failure To Start Message"; Value = $DEEPAPostFail.InnerText}
				@{ Description = "User Skipped Scan Message"; Value = $DEEPAPostSkipped.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTDEPOSTEPAH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			#endregion German Language

			#region Spanish Language
			WriteWordLine 4 0 "Spanish Language"
			$PTESLOGINH = $null
			[System.Collections.Hashtable[]] $PTESLOGINH = @(
				@{ Description = "Login Page"; Value = ""}
				@{ Description = "Page Title"; Value = $ESPageTitle.InnerText}
				@{ Description = "Form Title"; Value = $ESFormTitle.InnerText}
				@{ Description = "User Name Field Title"; Value = $ESUserName.InnerText}
				@{ Description = "Password Field Title"; Value = $ESPassword.InnerText}
				@{ Description = "Password Field2 Title"; Value = $ESPassword2.InnerText}    
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTESLOGINH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTESHOMEH = $null
			[System.Collections.Hashtable[]] $PTESHOMEH = @(
				@{ Description = "Home Page"; Value = ""}
				@{ Description = "Web Apps Tab Label"; Value = $ESWebApps.InnerText}
				@{ Description = "Enterprise Web Sites Label"; Value = $ESEntWebapps.InnerText}
				@{ Description = "Personal Web Sites Label"; Value = $ESPersWebapps.InnerText}
				@{ Description = "Applications Tab Label"; Value = $ESApps.InnerText}
				@{ Description = "File Transfer Tab Label"; Value = $ESFileTrans.InnerText}
				@{ Description = "Enterprise File Shares Label"; Value = $ESEntFile.InnerText}
				@{ Description = "Personal File Shares Label"; Value = $ESPersFile.InnerText}
				@{ Description = "Email Tab Label"; Value = $ESEmail.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTESHOMEH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTESVPNH = $null
			[System.Collections.Hashtable[]] $PTESVPNH = @(
				@{ Description = "VPN Connection"; Value = ""}
				@{ Description = "Waiting Message"; Value = $ESVPNWaitmsg.InnerText}
				@{ Description = "Proxy Configured message"; Value = $ESVPNproxy.InnerText}
				@{ Description = "Windows Plug-in Not Installed Message"; Value = $ESVPNnoplugin.InnerText}
				@{ Description = "MAC Plug-in Not Installed Message"; Value = $ESVPNnopluginmac.InnerText}
				@{ Description = "Linux Plug-in Not Installed Message"; Value = $ESVPNnopluginlinux.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTESVPNH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTESEPAH = $null
			[System.Collections.Hashtable[]] $PTESEPAH = @(
				@{ Description = "EPA Page"; Value = ""}
				@{ Description = "Title"; Value = $ESEPATitle.InnerText}
				@{ Description = "Introductory Message"; Value = $ESEPAIntro.InnerText}
				@{ Description = "Plug-in Check Message"; Value = $ESEPAPlugin.InnerText}
				@{ Description = "Download Plug-In Message"; Value = $ESEPADownload.InnerText}
				@{ Description = "Plug-in Launch Error Message"; Value = $ESEPAPluginError.InnerText}
				@{ Description = "Download Software Message"; Value = $ESEPASoftDownload.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTESEPAH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTESEPAERRH = $null
			[System.Collections.Hashtable[]] $PTESEPAERRH = @(
				@{ Description = "EPA Error Page"; Value = ""}
				@{ Description = "Error Title"; Value = $ESEPAErrorTitle.InnerText}
				@{ Description = "Device Requirements Not Matching Message"; Value = $ESEPADeviceReqs.InnerText}
				@{ Description = "Mac Failure Message"; Value = $ESEPAMacError.InnerText}
				@{ Description = "Error More Info Message"; Value = $ESEPAErrorMessage.InnerText}
				@{ Description = "Device Certificate Check Failure Message"; Value = $ESEPAErrorCert.InnerText}    
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTESEPAERRH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTESPOSTEPAH = $null
			[System.Collections.Hashtable[]] $PTESPOSTEPAH = @(
				@{ Description = "Post EPA Page"; Value = ""}
				@{ Description = "Title"; Value = $ESEPAPostTitle.InnerText}
				@{ Description = "Failure To Start Message"; Value = $ESEPAPostFail.InnerText}
				@{ Description = "User Skipped Scan Message"; Value = $ESEPAPostSkipped.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTESPOSTEPAH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			#endregion Spanish Language

			#region Japanese
			WriteWordLine 4 0 "Japanese Language"
			$PTJALOGINH = $null
			[System.Collections.Hashtable[]] $PTJALOGINH = @(
				@{ Description = "Login Page"; Value = ""}
				@{ Description = "Page Title"; Value = $JAPageTitle.InnerText}
				@{ Description = "Form Title"; Value = $JAFormTitle.InnerText}
				@{ Description = "User Name Field Title"; Value = $JAUserName.InnerText}
				@{ Description = "Password Field Title"; Value = $JAPassword.InnerText}
				@{ Description = "Password Field2 Title"; Value = $JAPassword2.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTJALOGINH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTJAHOMEH = $null
			[System.Collections.Hashtable[]] $PTJAHOMEH = @(
				@{ Description = "Home Page"; Value = ""}
				@{ Description = "Web Apps Tab Label"; Value = $JAWebApps.InnerText}
				@{ Description = "Enterprise Web Sites Label"; Value = $JAEntWebapps.InnerText}
				@{ Description = "Personal Web Sites Label"; Value = $JAPersWebapps.InnerText}
				@{ Description = "Applications Tab Label"; Value = $JAApps.InnerText}
				@{ Description = "File Transfer Tab Label"; Value = $JAFileTrans.InnerText}
				@{ Description = "Enterprise File Shares Label"; Value = $JAEntFile.InnerText}
				@{ Description = "Personal File Shares Label"; Value = $JAPersFile.InnerText}
				@{ Description = "Email Tab Label"; Value = $JAEmail.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTJAHOMEH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTJAVPNH = $null
			[System.Collections.Hashtable[]] $PTJAVPNH = @(
				@{ Description = "VPN Connection"; Value = ""}
				@{ Description = "Waiting Message"; Value = $JAVPNWaitmsg.InnerText}
				@{ Description = "Proxy Configured message"; Value = $JAVPNproxy.InnerText}
				@{ Description = "Windows Plug-in Not Installed Message"; Value = $JAVPNnoplugin.InnerText}
				@{ Description = "MAC Plug-in Not Installed Message"; Value = $JAVPNnopluginmac.InnerText}
				@{ Description = "Linux Plug-in Not Installed Message"; Value = $JAVPNnopluginlinux.InnerText}    
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTJAVPNH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTJAEPAH = $null
			[System.Collections.Hashtable[]] $PTJAEPAH = @(
				@{ Description = "EPA Page"; Value = ""}
				@{ Description = "Title"; Value = $JAEPATitle.InnerText}
				@{ Description = "Introductory Message"; Value = $JAEPAIntro.InnerText}
				@{ Description = "Plug-in Check Message"; Value = $JAEPAPlugin.InnerText}
				@{ Description = "Download Plug-In Message"; Value = $JAEPADownload.InnerText}
				@{ Description = "Plug-in Launch Error Message"; Value = $JAEPAPluginError.InnerText}
				@{ Description = "Download Software Message"; Value = $JAEPASoftDownload.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTJAEPAH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTJAEPAERRH = $null
			[System.Collections.Hashtable[]] $PTJAEPAERRH = @(
				@{ Description = "EPA Error Page"; Value = ""}
				@{ Description = "Error Title"; Value = $JAEPAErrorTitle.InnerText}
				@{ Description = "Device Requirements Not Matching Message"; Value = $JAEPADeviceReqs.InnerText}
				@{ Description = "Mac Failure Message"; Value = $JAEPAMacError.InnerText}
				@{ Description = "Error More Info Message"; Value = $JAEPAErrorMessage.InnerText}
				@{ Description = "Device Certificate Check Failure Message"; Value = $JAEPAErrorCert.InnerText}    
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTJAEPAERRH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd

			$PTJAPOSTEPAH = $null
			[System.Collections.Hashtable[]] $PTJAPOSTEPAH = @(
				@{ Description = "Post EPA Page"; Value = ""}
				@{ Description = "Title"; Value = $JAEPAPostTitle.InnerText}
				@{ Description = "Failure To Start Message"; Value = $JAEPAPostFail.InnerText}
				@{ Description = "User Skipped Scan Message"; Value = $JAEPAPostSkipped.InnerText}
			)
			$Params = $null
			$Params = @{
				Hashtable = $PTJAPOSTEPAH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			#endregion Japanese
		}
	} #End Foreach portal theme

	#endregion NetScaler Gateway Portal Themes

	#region User Administration
	#region AAA Groups
	$AAAGroupscount = (Get-vNetScalerObjectCount -Type aaagroup).__count
	$AAAUserscount = (Get-vNetScalerObjectCount -Type aaauser).__count
	If ($AAAGroupscount -gt 0 -or $AAAUserscount -gt 0){WriteWordLine 2 0 "User Administration"}
	If ($AAAGroupscount -gt 0) {
		WriteWordLine 3 0 "AAA Groups"
		$AAAGroups = Get-vNetScalerObject -Type aaagroup
		foreach ($AAAGroup in $AAAGroups) {
			WriteWordLine 4 0 "$($AAAGroup.groupname) ($($AAAGroup.weight))"			
			New-BindingTable -Name $AAAGroup.groupname -BindingType "aaagroup_aaauser_binding" -BindingTypeName "Members" -Properties "username" -Headers "Username" -Style 5
			New-BindingTable -Name $AAAGroup.groupname -BindingType "aaagroup_authorizationpolicy_binding" -BindingTypeName "Authorization Policies" -Properties "priority,policy" -Headers "Priority,Policy Name" -Style 5
			New-BindingTable -Name $AAAGroup.groupname -BindingType "aaagroup_vpnintranetapplication_binding" -BindingTypeName "Intranet Applications" -Properties "intranetapplication" -Headers "Name" -Style 5
			New-BindingTable -Name $AAAGroup.groupname -BindingType "aaagroup_vpnsessionpolicy_binding" -BindingTypeName "Session Policies" -Properties "priority,policy" -Headers "Priority,Policy Name" -Style 5
			New-BindingTable -Name $AAAGroup.groupname -BindingType "aaagroup_vpntrafficpolicy_binding" -BindingTypeName "Traffic Policies" -Properties "priority,policy" -Headers "Priority,Policy Name" -Style 5
			New-BindingTable -Name $AAAGroup.groupname -BindingType "aaagroup_intranetip_binding" -BindingTypeName "Intranet IP Addresses" -Properties "intranetip,netmask" -Headers "Intranet IP, Netmask" -Style 5
			New-BindingTable -Name $AAAGroup.groupname -BindingType "aaagroup_intranetip6_binding" -BindingTypeName "Intranet IP v6 Addresses" -Properties "intranetip,netmask" -Headers "Intranet IP, Netmask" -Style 5
			New-BindingTable -Name $AAAGroup.groupname -BindingType "aaagroup_vpnurl_binding" -BindingTypeName "Bookmarks" -Properties "urlname" -Headers "Name" -Style 5
		}
	}
	#endregion AAA Groups

	#region AAA Users
	If ($AAAUserscount -gt 0) {
		WriteWordLine 3 0 "AAA Users"
		$AAAUsers = Get-vNetScalerObject -Type aaauser
		foreach ($AAAUser in $AAAUsers) {
			WriteWordLine 4 0 "$($AAAUser.username)"
			New-BindingTable -Name $AAAUser.username -BindingType "aaagroup_aaauser_binding" -BindingTypeName "Member of AAA groups" -Properties "groupname" -Headers "Groups" -Style 5
			New-BindingTable -Name $AAAUser.username -BindingType "aaauser_authorizationpolicy_binding" -BindingTypeName "Authorization Policies" -Properties "priority,policy" -Headers "Priority,Policy Name" -Style 5
			New-BindingTable -Name $AAAUser.username -BindingType "aaauser_vpnintranetapplication_binding" -BindingTypeName "Intranet Applications" -Properties "intranetapplication" -Headers "Name" -Style 5
			New-BindingTable -Name $AAAUser.username -BindingType "aaauser_vpnsessionpolicy_binding" -BindingTypeName "Session Policies" -Properties "priority,policy" -Headers "Priority,Policy Name" -Style 5
			New-BindingTable -Name $AAAUser.username -BindingType "aaauser_vpntrafficpolicy_binding" -BindingTypeName "Traffic Policies" -Properties "priority,policy" -Headers "Priority,Policy Name" -Style 5
			New-BindingTable -Name $AAAUser.username -BindingType "aaauser_intranetip_binding" -BindingTypeName "Intranet IP Addresses" -Properties "intranetip,netmask" -Headers "Intranet IP, Netmask" -Style 5
			New-BindingTable -Name $AAAUser.username -BindingType "aaauser_vpnurl_binding" -BindingTypeName "Bookmarks" -Properties "urlname" -Headers "Name" -Style 5
			New-BindingTable -Name $AAAUser.username -BindingType "aaauser_intranetip6_binding" -BindingTypeName "Intranet IP v6 Addresses" -Properties "intranetip,netmask" -Headers "Intranet IP, Netmask" -Style 5
		}
	}
	#endregion AAA Users
	#endregion User Administration

	#region NetScaler Gateway KCD Accounts
	If ((Get-vNetScalerObjectCount -Type aaakcdaccount).__count -ge 1) {
		WriteWordLine 3 0 "KCD Accounts"
		$kcdaccounts = Get-vNetScalerObject -Type aaakcdaccount
		[System.Collections.Hashtable[]] $KCDH = @() 
		foreach ($kcdaccount in $kcdaccounts) {           
			$kcdname = $kcdaccount.kcdaccount
			WriteWordLine 3 0 "KCD Account: $kcdname"     
			[System.Collections.Hashtable[]] $KCDH = @(
				@{ Description = "KeyTab File"; Value = $kcdaccount.keytab}
				@{ Description = "Principle"; Value = $kcdaccount.principle}
				@{ Description = "SPN"; Value = $kcdaccount.kcdspn}
				@{ Description = "Realm"; Value = $kcdaccount.realmstr}
				@{ Description = "User Realm"; Value = $kcdaccount.userrealm}
				@{ Description = "Enterprise Realm"; Value = $kcdaccount.enterpriserealm}
				@{ Description = "Delegated User"; Value = $kcdaccount.delegateduser}
				@{ Description = "KCD Password"; Value = $kcdaccount.kcdpassword}
				@{ Description = "User Certificate"; Value = $kcdaccount.usercert}
				@{ Description = "CA Certificate"; Value = $kcdaccount.cacert}
				@{ Description = "Service SPN"; Value = $kcdaccount.servicespn}
			)
			$Params = $null
			$Params = @{
				Hashtable	= $KCDH
				Columns		= "Description", "Value"
				Headers		= "Description", "Configuration"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
		}
	}
	#endregion NetScaler Gateway KCD Accounts

	#region NetScaler Gateway Policies
	$selection.InsertNewPage()
	WriteWordLine 2 0 "Policies"
	#region NetScaler Gateway Session Policies
	If ((Get-vNetScalerObjectCount -Type vpnsessionpolicy).__count -ge 1) {
		WriteWordLine 3 0 "Session Policies"
		$vpnsessionpolicies = Get-vNetScalerObject -Type vpnsessionpolicy
		[System.Collections.Hashtable[]] $VPNSESPOLH = @()
		foreach ($vpnsessionpolicy in $vpnsessionpolicies) {
`			$VPNSESPOLH += @{
				NAME	= $vpnsessionpolicy.name
				RULE	= $vpnsessionpolicy.rule
				ACTION	= $vpnsessionpolicy.action
				ACTIVE	= $vpnsessionpolicy.activepolicy
			}
		}
		If ($VPNSESPOLH.Length -gt 0) {
			$Params = $null
			$Params = @{
				Hashtable = $VPNSESPOLH
				Columns   = "NAME", "RULE", "ACTION", "ACTIVE"
				Headers   = "Policy", "Rule", "Action", "Active"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion NetScaler Gateway Session Policies
	
	#region NetScaler Gateway Session Profiles
	If ((Get-vNetScalerObjectCount -Type vpnsessionaction).__count -ge 1) {
		WriteWordLine 3 0 "Session Profiles"
		$vpnsessionactions = Get-vNetScalerObject -Type vpnsessionaction
		foreach ($vpnsessionaction in $vpnsessionactions) {
			WriteWordLine 4 0 "$($vpnsessionaction.name)"
			
			#region ClientExperience
			WriteWordLine 5 0 "Client Experience"
			[System.Collections.Hashtable[]] $VPNACTCEXH = @(
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.homepage)){@{ Description = "Homepage"; Value = $vpnsessionaction.homepage}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.emailhome)){@{ Description = "URL for Web Based Email"; Value = $vpnsessionaction.emailhome}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.sesstimeout)){@{ Description = "Session Time-Out"; Value = $vpnsessionaction.sesstimeout}}
				If ($vpnsessionaction.clientidletimeoutwarning -ne "0"){@{ Description = "Client-Idle Time-Out [0]"; Value = $vpnsessionaction.clientidletimeoutwarning}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.sso)){@{ Description = "Single Sign-On to Web Applications"; Value = $vpnsessionaction.sso}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.windowsautologon)){@{ Description = "Single Sign-On with Windows"; Value = $vpnsessionaction.windowsautologon}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.splittunnel)){@{ Description = "Split Tunnel"; Value = $vpnsessionaction.splittunnel}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.locallanaccess)){@{ Description = "Local LAN Access"; Value = $vpnsessionaction.locallanaccess}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.windowsclienttype)){@{ Description = "Plug-in Type"; Value = $vpnsessionaction.windowsclienttype}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.windowspluginupgrade)){@{ Description = "Windows Plugin Upgrade"; Value = $vpnsessionaction.windowspluginupgrade}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.macpluginupgrade)){@{ Description = "MAC Plugin Upgrade"; Value = $vpnsessionaction.macpluginupgrade}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.linuxpluginupgrade)){@{ Description = "Linux Plugin Upgrade"; Value = $vpnsessionaction.linuxpluginupgrade}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.alwaysonprofilename)){@{ Description = "AlwaysON Profile Name"; Value = $vpnsessionaction.alwaysonprofilename}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.clientlessvpnmode)){@{ Description = "Clientless Access"; Value = $vpnsessionaction.clientlessvpnmode}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.clientlessmodeurlencoding)){@{ Description = "Clientless URL Encoding"; Value = $vpnsessionaction.clientlessmodeurlencoding}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.clientlesspersistentcookie)){@{ Description = "Clientless Persistent Cookie"; Value = $vpnsessionaction.clientlesspersistentcookie}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.ssocredential)){@{ Description = "Credential Index"; Value = $vpnsessionaction.ssocredential}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.kcdaccount)){@{ Description = "KCD Account"; Value = $vpnsessionaction.kcdaccount}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.clientcleanupprompt)){@{ Description = "Client Cleanup Prompt"; Value = $vpnsessionaction.clientcleanupprompt}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.uitheme)){@{ Description = "UI Theme"; Value = $vpnsessionaction.uitheme}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.loginscript)){@{ Description = "Login Script"; Value = $vpnsessionaction.loginscript}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.logoutscript)){@{ Description = "Logout Script"; Value = $vpnsessionaction.logoutscript}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.apptokentimeout)){@{ Description = "Application Token Timeout"; Value = $vpnsessionaction.apptokentimeout}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.mdxtokentimeout)){@{ Description = "MDX Token Timeout"; Value = $vpnsessionaction.mdxtokentimeout}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.clientconfiguration)){@{ Description = "Allow Users to Change Log Levels"; Value = $vpnsessionaction.clientconfiguration}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.windowsclienttype)){@{ Description = "Allow access to private network IP addresses only"; Value = $vpnsessionaction.windowsclienttype}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.clientchoices)){@{ Description = "Client Choices"; Value = $vpnsessionaction.clientchoices}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.iconwithreceiver)){@{ Description = "Show VPN Plugin icon"; Value = $vpnsessionaction.iconwithreceiver}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.pcoipprofilename)){@{ Description = "PCOIP Profile Name"; Value = $vpnsessionaction.pcoipprofilename}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.autoproxyurl)){@{ Description = "AutoProxy URL"; Value = $vpnsessionaction.autoproxyurl}}
			)
			If ($VPNACTCEXH.Length -gt 0) {
				$Params = $null
				$Params = @{
					Hashtable	= $VPNACTCEXH
					Columns		= "Description", "Value"
					Headers		= "Description", "Configuration"
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
			#endregion ClientExperience

			#region Security
			WriteWordLine 5 0 "Security"
			[System.Collections.Hashtable[]] $VPNACTSECH = @(
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.defaultauthorizationaction)){@{ Description = "Default Authorization Action"; Value = $vpnsessionaction.defaultauthorizationaction}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.securebrowse)){@{ Description = "Secure Browse"; Value = $vpnsessionaction.securebrowse}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.clientsecurity)){@{ Description = "Client Security Check String"; Value = $vpnsessionaction.clientsecurity}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.clientsecuritygroup)){@{ Description = "Quarantine Group"; Value = $vpnsessionaction.clientsecuritygroup}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.clientsecuritymessage)){@{ Description = "Error Message"; Value = $vpnsessionaction.clientsecuritymessage}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.clientsecuritylog)){@{ Description = "Enable Client Security Logging"; Value = $vpnsessionaction.clientsecuritylog}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.authorizationgroup)){@{ Description = "Authorization Groups"; Value = $vpnsessionaction.authorizationgroup}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.allowedlogingroups)){@{ Description = "Groups allowed to login"; Value = $vpnsessionaction.allowedlogingroups}}
			)
			If ($VPNACTSECH.Length -gt 0) {
				$Params = $null
				$Params = @{
					Hashtable	= $VPNACTSECH
					Columns		= "Description", "Value"
					Headers		= "Description", "Configuration"
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
			#endregion Security

			#region Published Applications  
			WriteWordLine 5 0 "Published Applications"
			[System.Collections.Hashtable[]] $VPNACTPAH = @(
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.icaproxy)){@{ Description = "ICA Proxy"; Value = $vpnsessionaction.icaproxy}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.wihome)){
					@{ Description = "Web Interface Address"; Value = $vpnsessionaction.wihome}
					@{ Description = "Web Interface Address Type"; Value = $vpnsessionaction.wihomeaddresstype}
				}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.ntdomain)){@{ Description = "Single Sign-on Domain"; Value = $vpnsessionaction.ntdomain}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.citrixreceiverhome)){@{ Description = "Citrix Receiver Home Page"; Value = $vpnsessionaction.citrixreceiverhome}}
				If (![string]::IsNullOrWhiteSpace($vpnsessionaction.storefronturl)){@{ Description = "Account Services Address"; Value = $vpnsessionaction.storefronturl}}
			)
			If ($VPNACTPAH.Length -gt 0) {
				$Params = $null
				$Params = @{
					Hashtable	= $VPNACTPAH
					Columns		= "Description", "Value"
					Headers		= "Description", "Configuration"
				}
				$Table = AddWordTable @Params
				FindWordDocumentEnd
			}
			#endregion Published Applications
		}
	}
	#endregion NetScaler Gateway Session Profiles

	#region NetScaler Gateway Traffic
	#region NetScaler Gateway Traffic Policies
	If ((Get-vNetScalerObjectCount -Type vpntrafficpolicy).__count -ge 1) {
		WriteWordLine 3 0 "Traffic Policies"
		$vpntrafficpolicys = Get-vNetScalerObject -Type vpntrafficpolicy
		[System.Collections.Hashtable[]] $vpntrafficpolicyH = @()
		foreach ($vpntrafficpolicy in $vpntrafficpolicys) {
			$vpntrafficpolicyH += @{ 
				NAME    = $vpntrafficpolicy.name
				ACTION = $vpntrafficpolicy.action
				RULE  = $vpntrafficpolicy.Rule
			}
		} 
		$Params = $null
		$Params = @{
			Hashtable = $vpntrafficpolicyH
			Columns   = "NAME", "ACTION", "RULE"
			Headers   = "Name", "Request Profile", "Expression"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
	#endregion NetScaler Gateway Traffic Policies

	#region NetScaler Gateway Traffic Profiles
	If ((Get-vNetScalerObjectCount -Type vpntrafficaction).__count -ge 1) {
		WriteWordLine 3 0 "Traffic Profiles"
		$vpntrafficactions = Get-vNetScalerObject -Type vpntrafficaction
		foreach ($vpntrafficaction in $vpntrafficactions) {
			WriteWordLine 4 0 "$($vpntrafficaction.name)"
			[System.Collections.Hashtable[]] $vpntrafficactionH = @(
				@{ Description = "Description"; Value = "Configuration" }
				@{ Description = "Protocol"; Value = $vpntrafficaction.qual }
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.apptimeout)){@{ Description = "AppTimeout (minutes)"; Value = $vpntrafficaction.apptimeout }}
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.sso)){@{ Description = "Single Sign-on"; Value = $vpntrafficaction.sso }}
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.formssoaction)){@{ Description = "Form SSO Profile"; Value = $vpntrafficaction.formssoaction }}
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.samlssoprofile)){@{ Description = "SAML SSO Action"; Value = $vpntrafficaction.samlssoprofile }}
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.fta)){@{ Description = "File Type Association"; Value = $vpntrafficaction.fta }}
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.hdx)){@{ Description = "HDX Proxy"; Value = $vpntrafficaction.hdx }}
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.proxy)){@{ Description = "Proxy"; Value = $vpntrafficaction.proxy }}
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.wanscaler)){@{ Description = "CloudBridge"; Value = $vpntrafficaction.wanscaler }}
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.kcdaccount)){@{ Description = "KCD Account"; Value = $vpntrafficaction.kcdaccount }}
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.userexpression)){@{ Description = "SSO User Expression"; Value = $vpntrafficaction.userexpression }}
				If (![string]::IsNullOrWhiteSpace($vpntrafficaction.passwdexpression)){@{ Description = "SSO Password Expression"; Value = $vpntrafficaction.passwdexpression }}
			)
		} 
		$Params = $null
		$Params = @{
			Hashtable = $vpntrafficactionH
			Columns   = "Description", "Value"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
	#endregion NetScaler Gateway Traffic Profiles
	#endregion NetScaler Gateway Traffic

	#region NetScaler Gateway RDP
	#region NetScaler Gateway RDP Server Profiles
	If ((Get-vNetScalerObjectCount -Type rdpserverprofile).__count -ge 1) {
		WriteWordLine 3 0 "RDP Server Profiles"
		$rdpsrvprofiles = Get-vNetScalerObject -Type rdpserverprofile
		[System.Collections.Hashtable[]] $RDPSRVPROFH = @()
		foreach ($rdpsrvprofile in $rdpsrvprofiles) {
			$PSK = Get-NonEmptyString $rdpsrvprofile.psk
			$RDPSRVPROFH += @{ 
				NAME  = $rdpsrvprofile.name 
				IP    = $rdpsrvprofile.rdpip
				PORT  = $rdpsrvprofile.rdpport
				REDIR = $rdpsrvprofile.rdpredirection
			}
		} 
		$Params = $null
		$Params = @{
			Hashtable = $RDPSRVPROFH
			Columns   = "NAME", "IP", "PORT", "REDIR"
			Headers   = "Name", "RDP Listener IP", "RDP Port", "RDP Redirection Support (Broker)"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
	#endregion NetScaler Gateway RDP Server Profiles

	#region NetScaler Gateway RDP Client Profiles
	If ((Get-vNetScalerObjectCount -Type rdpclientprofile).__count -ge 1) {
		WriteWordLine 3 0 "RDP Client Profiles"
		$rdpcltprofiles = Get-vNetScalerObject -Type rdpclientprofile
		foreach ($rdpcltprofile in $rdpcltprofiles) {
			WriteWordLine 4 0 "$($rdpcltprofile.name)"
			[System.Collections.Hashtable[]] $RDPCLTH = @(
				@{ Description = "Override RDP URL"; Value = $rdpcltprofile.rdpurloverride}
				@{ Description = "Redirect Clipboard"; Value = $rdpcltprofile.redirectclipboard}
				@{ Description = "Redirect Drives"; Value = $rdpcltprofile.redirectdrives}
				@{ Description = "Redirect Printers"; Value = $rdpcltprofile.redirectprinters}
				@{ Description = "Redirect COM Ports"; Value = $rdpcltprofile.redirectcomports}
				@{ Description = "Redirect PnP Devices"; Value = $rdpcltprofile.redirectpnpdevices}
				@{ Description = "Keyboard Hook"; Value = $rdpcltprofile.keyboardhook}
				@{ Description = "Audio Capture Mode"; Value = $rdpcltprofile.audiocapturemode}
				@{ Description = "Video Playback Mode"; Value = $rdpcltprofile.videoplaybackmode}
				If ($rdpcltprofile.rdpcookievalidity -ne "60"){@{ Description = "RDP Cookie Validity"; Value = $rdpcltprofile.rdpcookievalidity}}
				@{ Description = "Include Username in RDP File"; Value = $rdpcltprofile.addusernameinrdpfile}
				If (![string]::IsNullOrWhiteSpace($rdpcltprofile.rdpfilename)){@{ Description = "RDP File Name"; Value = $rdpcltprofile.rdpfilename}}
				If (![string]::IsNullOrWhiteSpace($rdpcltprofile.rdphost)){@{ Description = "RDP Host"; Value = $rdpcltprofile.rdphost}}
				If (![string]::IsNullOrWhiteSpace($rdpcltprofile.rdplistener)){@{ Description = "RDP Listener"; Value = $rdpcltprofile.rdplistener}}
				@{ Description = "Multiple Monitor Support"; Value = $rdpcltprofile.multimonitorsupport}
				If (![string]::IsNullOrWhiteSpace($rdpcltprofile.rdpcustomparams)){@{ Description = "RDP Custom Parameters"; Value = $rdpcltprofile.rdpcustomparams}}
				If (![string]::IsNullOrWhiteSpace($rdpcltprofile.psk)){@{ Description = "Pre-Shared Key"; Value = $rdpcltprofile.psk}}
				@{ Description = "Randomize RDP File Name"; Value = $rdpcltprofile.randomizerdpfilename}
				If (![string]::IsNullOrWhiteSpace($rdpcltprofile.rdplinkattribute)){@{ Description = "RDP Link Attribute (fetch from AD)"; Value = $rdpcltprofile.rdplinkattribute}}
			)
			$Params = $null
			$Params = @{
				Hashtable	= $RDPCLTH
				Columns		= "Description", "Value"
				Headers		= "Description", "Configuration"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion NetScaler Gateway RDP Client Profiles
	#endregion NetScaler Gateway RDP

	#region NetScaler Gateway PCOIP
	#region NetScaler Gateway PCoIP vServer Profiles
	If ((Get-vNetScalerObjectCount -Type vpnpcoipvserverprofile).__count -ge 1) {
		WriteWordLine 3 0 "PCoIP vServer Profiles"
		$pcoipvprofiles = Get-vNetScalerObject -Type vpnpcoipvserverprofile
		[System.Collections.Hashtable[]] $PCoIPVSRVH = @()
		foreach ($pcoipvprofile in $pcoipvprofiles) {
			$PCoIPVSRVH += @{ 
				NAME    = $pcoipvprofile.name 
				DOMAIN  = $pcoipvprofile.logindomain
				UDPPORT = $pcoipvprofile.udpport
			}
		} 
		$Params = $null
		$Params = @{
			Hashtable = $PCoIPVSRVH
			Columns   = "NAME", "DOMAIN", "UDPPORT"
			Headers   = "Name", "Logon Domain", "UDP Port"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
	#endregion NetScaler Gateway PCOIP vServer Profiles

	#region NetScaler Gateway PCOIP Profiles
	If ((Get-vNetScalerObjectCount -Type vpnpcoipprofile).__count -ge 1) {
		WriteWordLine 3 0 "PCoIP Profiles"
		$pcoipprofiles = Get-vNetScalerObject -Type vpnpcoipprofile
		[System.Collections.Hashtable[]] $PCoIPPROFH = @()
		foreach ($pcoipprofile in $pcoipprofiles) {
			$PCoIPPROFH += @{ 
				NAME      = $pcoipprofile.name 
				CONSERVER = $pcoipprofile.conserverurl
				ICV       = $pcoipprofile.icvverification
				IDLE      = $pcoipprofile.sessionidletimeout
			}
		} 
		$Params = $null
		$Params = @{
			Hashtable = $PCoIPPROFH
			Columns   = "NAME", "CONSERVER", "ICV", "IDLE"
			Headers   = "Name", "Connection Server URL", "ICV Verification", "Session Idle Timeout"
		}
		$Table = AddWordTable @Params
		FindWordDocumentEnd
	}
	#endregion NetScaler Gateway PCOIP Profiles
	#endregion NetScaler Gateway PCOIP

	#region NetScaler Gateway AlwaysOn policies
	If ((Get-vNetScalerObjectCount -Type vpnalwaysonprofile).__count -ge 1) {
		WriteWordLine 3 0 "AlwaysON Policies"
		$vpnalwaysonpolicies = Get-vNetScalerObject -Type vpnalwaysonprofile
		foreach ($vpnalwaysonpolicy in $vpnalwaysonpolicies) {
			WriteWordLine 4 0 "$($vpnalwaysonpolicy.name)"
			[System.Collections.Hashtable[]] $AOPOLCONFH = @(
				@{ Description = "Location Based VPN"; Value = $vpnalwaysonpolicy.locationbasedvpn}
				@{ Description = "Client Control"; Value = $vpnalwaysonpolicy.clientcontrol}
				@{ Description = "Network Access On VPN Failure"; Value = $vpnalwaysonpolicy.networkaccessonvpnfailure}
			)
			$Params = $null
			$Params = @{
				Hashtable	= $AOPOLCONFH
				Columns		= "Description", "Value"
				Headers		= "Description", "Configuration"
			}
			$Table = AddWordTable @Params
			FindWordDocumentEnd
		}
	}
	#endregion NetScaler Gateway AlwaysOn policies
	#endregion NetScaler Gateway Policies
	
	#region NetScaler Gateway Resources
	$selection.InsertNewPage()
	WriteWordLine 2 0 "Resources"
	#region NetScaler Gateway Intranet Applications
	If ((Get-vNetScalerObjectCount -Type vpnintranetapplication).__count -ge 1) {
		WriteWordLine 3 0 "Intranet Applications"
		$vpnintapps = Get-vNetScalerObject -Type vpnintranetapplication
		foreach ($vpnintapp in $vpnintapps) {
			WriteWordLine 4 0 "$($vpnintapp.intranetapplication)"
			[System.Collections.Hashtable[]] $VPNINTAPPH = @(
				@{ Description = "Protocol"; Value = $vpnintapp.protocol}
				If (![string]::IsNullOrWhiteSpace($vpnintapp.destip)){@{ Description = "Destination IP Address"; Value = $vpnintapp.destip}}
				If (![string]::IsNullOrWhiteSpace($vpnintapp.netmask)){@{ Description = "Netmask"; Value = $vpnintapp.netmask}}
				If (![string]::IsNullOrWhiteSpace($vpnintapp.iprange)){@{ Description = "IP Range"; Value = $vpnintapp.iprange}}
				If (![string]::IsNullOrWhiteSpace($vpnintapp.hostname)){@{ Description = "Hostname"; Value = $vpnintapp.hostname}}
				If (![string]::IsNullOrWhiteSpace($vpnintapp.clientapplication)){@{ Description = "Client Application"; Value = $vpnintapp.clientapplication}}
				If (![string]::IsNullOrWhiteSpace($vpnintapp.spoofiip)){@{ Description = "Spoof IP"; Value = $vpnintapp.spoofiip}}
				If (![string]::IsNullOrWhiteSpace($vpnintapp.destport)){@{ Description = "Destination Port"; Value = $vpnintapp.destport}}
				If (![string]::IsNullOrWhiteSpace($vpnintapp.interception)){@{ Description = "Interception Mode"; Value = $vpnintapp.interception}}
				If (![string]::IsNullOrWhiteSpace($vpnintapp.srcip)){@{ Description = "Source IP"; Value = $vpnintapp.srcip}}
				If (![string]::IsNullOrWhiteSpace($vpnintapp.srcprt)){@{ Description = "Source Port"; Value = $vpnintapp.srcprt}}
			)
			$Params = $null
			$Params = @{
				Hashtable	= $VPNINTAPPH
				Columns		= "Description", "Value"
				Headers		= "Description", "Configuration"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
		}
	}
	#endregion NetScaler Gateway Intranet Applications

	#region NetScaler Gateway Bookmarks
	If ((Get-vNetScalerObjectCount -Type vpnurl).__count -ge 1) {
		WriteWordLine 3 0 "Bookmarks"
		$vpnurls = Get-vNetScalerObject -Type vpnurl
		foreach ($vpnurl in $vpnurls) {
			WriteWordLine 4 0 "$($vpnurl.urlname)"
			[System.Collections.Hashtable[]] $VPNURLH = @(
				@{ Description = "Description"; Value = $vpnurl.linkname}
				If (![string]::IsNullOrWhiteSpace($vpnurl.actualurl)){@{ Description = "URL"; Value = $vpnurl.actualurl}}
				If (![string]::IsNullOrWhiteSpace($vpnurl.vservername)){@{ Description = "Virtual Server Name"; Value = $vpnurl.vservername}}
				If (![string]::IsNullOrWhiteSpace($vpnurl.clientlessaccess)){@{ Description = "Clientless Access"; Value = $vpnurl.clientlessaccess}}
				If (![string]::IsNullOrWhiteSpace($vpnurl.comment)){@{ Description = "Comment"; Value = $vpnurl.comment}}
				If (![string]::IsNullOrWhiteSpace($vpnurl.iconurl)){@{ Description = "Icon URL"; Value = $vpnurl.iconurl}}
				If (![string]::IsNullOrWhiteSpace($vpnurl.ssotype)){@{ Description = "SSO Type"; Value = $vpnurl.ssotype}}
				If (![string]::IsNullOrWhiteSpace($vpnurl.applicationtype)){@{ Description = "Application Type"; Value = $vpnurl.applicationtype}}
				If (![string]::IsNullOrWhiteSpace($vpnurl.samlssoprofile)){@{ Description = "SAML SSO Profile"; Value = $vpnurl.samlssoprofile}}
			)
			$Params = $null
			$Params = @{
				Hashtable = $VPNURLH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
		}
	}
	#endregion NetScaler Gateway Bookmarks
	#endregion NetScaler Gateway Resources
	
	#region Unified Gateway SaaS Templates
	WriteWordLine 2 0 "NetScaler Unified Gateway SaaS Templates"
	WriteWordLine 3 0 "System Templates"
	$SystemContent = Get-vNetScalerFile -FileName system.json -FileLocation "/var/app_catalog" | Select -ExpandProperty filecontent
	$SystemTemplates = $null
	$SystemTemplates = Get-StringFromBase64 -Object $SystemContent -Encoding UTF8 | ConvertFrom-Json
	If (!($systemTemplates)) { writewordline 0 0 "No System SaaS Templates were found on the appliance" } Else {
		foreach ($app in $SystemTemplates.apps) {
			[System.Collections.Hashtable[]] $SYSAPPH = @(
				@{ Description = "Setting"; Value = "Value"}
				@{ Description = "Display Name"; Value = $app.displayname}
				@{ Description = "Description"; Value = $app.description}
				@{ Description = "URL"; Value = $app.url}
				@{ Description = "Related URL"; Value = $app.relatedURLs}    
				@{ Description = "SAML Type"; Value = $app.SAMLType.value}
				@{ Description = "Assertion Consumer Service (ACS) URL"; Value = $app.sso.saml.assertionConsumerServiceURL.value}
				@{ Description = "Name ID Format"; Value = $app.sso.saml.nameID.format}
				@{ Description = "Name ID Value"; Value = $app.sso.saml.nameID.value}
				@{ Description = "Signature Algorithm"; Value = $app.sso.saml.signatureAlg.value}
				@{ Description = "Digest Method"; Value = $app.sso.saml.digestMethod.value }
				@{ Description = "Sign Assertion"; Value = $app.sso.saml.signAssertion.value}
				@{ Description = "Reject Unsigned Requests"; Value = $app.sso.saml.rejectUnsignedRequests.value}
				@{ Description = "SAML SP Certificate Name"; Value = $app.sso.saml.samlSpCertName.value}
				$samlattrcount = 0
				foreach ($attribute in $app.sso.saml.attributes) {
					$samlattrcount++
					@{ Description = "SAML Attribute $samlattrcount"; Value = " "}
					@{ Description = "$($attribute.name)"; Value = "$($attribute.value)"}
				}    
			)
			$Params = $null
			$Params = @{
				Hashtable = $SYSAPPH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			WriteWordLine 0 0 ""
		}
	}

	WriteWordLine 3 0 "User Templates"
	$UserContent = Get-vNetScalerFile -FileName user.json -FileLocation "/var/app_catalog" | Select -ExpandProperty filecontent
	$UserTemplates = $null
	$UserTemplates = Get-StringFromBase64 -Object $UserContent -Encoding UTF8 | ConvertFrom-Json
	If (!($UserTemplates)) { writewordline 0 0 "No User SaaS Templates were found on the appliance" } Else {
		foreach ($app in $UserTemplates.apps) {
			[System.Collections.Hashtable[]] $USERAPPH = @(
				@{ Description = "Setting"; Value = "Value"}
				@{ Description = "Display Name"; Value = $app.displayname}
				@{ Description = "Description"; Value = $app.description}
				@{ Description = "URL"; Value = $app.url}
				@{ Description = "Related URL"; Value = $app.relatedURLs}    
				@{ Description = "SAML Type"; Value = $app.SAMLType.value}
				@{ Description = "Assertion Consumer Service (ACS) URL"; Value = $app.sso.saml.assertionConsumerServiceURL.value}
				@{ Description = "Name ID Format"; Value = $app.sso.saml.nameID.format}
				@{ Description = "Name ID Value"; Value = $app.sso.saml.nameID.value}
				@{ Description = "Signature Algorithm"; Value = $app.sso.saml.signatureAlg.value}
				@{ Description = "Digest Method"; Value = $app.sso.saml.digestMethod.value }
				@{ Description = "Sign Assertion"; Value = $app.sso.saml.signAssertion.value}
				@{ Description = "Reject Unsigned Requests"; Value = $app.sso.saml.rejectUnsignedRequests.value}
				@{ Description = "SAML SP Certificate Name"; Value = $app.sso.saml.samlSpCertName.value}
				$samlattrcount = 0
				foreach ($attribute in $app.sso.saml.attributes) {
					$samlattrcount++
					@{ Description = "SAML Attribute $samlattrcount"; Value = " "}
					@{ Description = "$($attribute.name)"; Value = "$($attribute.value)"}
				}
			)
			$Params = $null
			$Params = @{
				Hashtable = $USERAPPH
				Columns   = "Description", "Value"
			}
			$Table = AddWordTable @Params -List
			FindWordDocumentEnd
			WriteWordLine 0 0 ""
		}
	}
	#endregion Unified Gateway SaaS Templatess
}
#endregion NetScaler Gateway

#region Logout
Set-Progress "Logging out of NetScaler"
Logout-vNetScalerSession
#endregion Logout

#endregion Documentation Script Complete

#region restore SSL validation to normal behavior
# Many thanks go out to Esther Barthel for fixing this!
If ($USENSSSL) {
    Write-Verbose "Rollback of required change for SSL certificate trust of NetScaler System Certificate"
    # source: blogs.technet.microsoft.com/bshukla/2010/0… 
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $false }
}
#endregion restore SSL validation to normal behavior

#region script template 2
Write-Verbose "$(Get-Date): Finishing up document"
Write-Log "Finishing up document"
#end of document processing

###Change the two lines below for your script
$AbstractTitle = "NetScaler Documentation Report"
$SubjectTitle = "NetScaler Documentation Report"

If (!$Offline) {
    Set-Progress "Finalising Document"
    Write-Log "Finalising Document"
    UpdateDocumentProperties $AbstractTitle $SubjectTitle
    Write-Log "Processing Document Output"
    ProcessDocumentOutput
}

ProcessScriptEnd
Write-Log "Script Completed"
Set-Progress "Script Completed"
#recommended by webster
#$error
#endregion script template 2