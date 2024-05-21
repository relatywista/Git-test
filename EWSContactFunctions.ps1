function Connect-Exchange
{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
    )  
 	Begin
		 {
		## Load Managed API dll  
		###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
		$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
		if (Test-Path $EWSDLL)
		    {
		    Import-Module $EWSDLL
		    }
		else
		    {
		    "$(get-date -format yyyyMMddHHmmss):"
		    "This script requires the EWS Managed API 1.2 or later."
		    "Please download and install the current version of the EWS Managed API from"
		    "http://go.microsoft.com/fwlink/?LinkId=255472"
		    ""
		    "Exiting Script."
		    $exception = New-Object System.Exception ("Managed Api missing")
			throw $exception
		    } 
  
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  
		  
		## Create Exchange Service Object  
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		  
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		  
		#Credentials Option 1 using UPN for the windows Account  
		#$psCred = Get-Credential  
		$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
		$service.Credentials = $creds      
		#Credentials Option 2  
		#service.UseDefaultCredentials = $true  
		 #$service.TraceEnabled = $true
		## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		  
		## Code From http://poshcode.org/624
		## Create a compilation environment
		$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
		$Compiler=$Provider.CreateCompiler()
		$Params=New-Object System.CodeDom.Compiler.CompilerParameters
		$Params.GenerateExecutable=$False
		$Params.GenerateInMemory=$True
		$Params.IncludeDebugInformation=$False
		$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
		$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
		$TAAssembly=$TAResults.CompiledAssembly

		## We now create an instance of the TrustAll and attach it to the ServicePointManager
		$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
		[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

		## end code from http://poshcode.org/624
		  
		## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
		  
		#CAS URL Option 1 Autodiscover  
		$service.AutodiscoverUrl($MailboxName,{$true})  
		Write-host ("Using CAS Server : " + $Service.url)   
		   
		#CAS URL Option 2 Hardcoded  
		  
		#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
		#$service.Url = $uri    
		  
		## Optional section for Exchange Impersonation  
		  
		#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		if(!$service.URL){
			throw "Error connecting to EWS"
		}
		else
		{		
			return $service
		}
	}
}
####################### 
<# 
.SYNOPSIS 
 Creates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Creates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To create a contact in the default contacts folder 
	Create-Contact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe"
	
	Example 2 To create a contact and add a contact picture
	Create-Contact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe" -photo 'c:\photo\Jdoe.jpg'

	Example 3 To create a contact in a user created subfolder 
	Create-Contact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe" -Folder "\MyCustomContacts"
    
	This cmdlet uses the EmailAddress as unique key so it wont let you create a contact with that email address if one already exists.
#> 
########################
function Create-Contact 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
 		[Parameter(Position=1, Mandatory=$true)] [string]$DisplayName,
		[Parameter(Position=2, Mandatory=$true)] [string]$FirstName,
		[Parameter(Position=3, Mandatory=$true)] [string]$LastName,
		[Parameter(Position=4, Mandatory=$true)] [string]$EmailAddress,
		[Parameter(Position=5, Mandatory=$false)] [string]$CompanyName,
		[Parameter(Position=6, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=7, Mandatory=$false)] [string]$Department,
		[Parameter(Position=8, Mandatory=$false)] [string]$Office,
		[Parameter(Position=9, Mandatory=$false)] [string]$BusinssPhone,
		[Parameter(Position=10, Mandatory=$false)] [string]$MobilePhone,
		[Parameter(Position=11, Mandatory=$false)] [string]$HomePhone,
		[Parameter(Position=12, Mandatory=$false)] [string]$IMAddress,
		[Parameter(Position=13, Mandatory=$false)] [string]$Street,
		[Parameter(Position=14, Mandatory=$false)] [string]$City,
		[Parameter(Position=15, Mandatory=$false)] [string]$State,
		[Parameter(Position=16, Mandatory=$false)] [string]$PostalCode,
		[Parameter(Position=17, Mandatory=$false)] [string]$Country,
		[Parameter(Position=18, Mandatory=$false)] [string]$JobTitle,
		[Parameter(Position=19, Mandatory=$false)] [string]$Notes,
		[Parameter(Position=20, Mandatory=$false)] [string]$Photo,
		[Parameter(Position=21, Mandatory=$false)] [string]$FileAs,
		[Parameter(Position=22, Mandatory=$false)] [string]$WebSite,
		[Parameter(Position=23, Mandatory=$false)] [string]$Title,
		[Parameter(Position=24, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=25, Mandatory=$false)] [string]$EmailAddressDisplayAs,
		[Parameter(Position=26, Mandatory=$false)] [switch]$useImpersonation

		
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		if($service.URL){
			$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
			$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
			$ParentFolderIds = [Activator]::CreateInstance($type)
			$ParentFolderIds.Add($Contacts.Id)
			$Error.Clear();
			$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
			$ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts,$true,$cnpsPropset);
			$createContactOkay = $false
			if($Error.Count -eq 0){
				if ($ncCol.Count -eq 0) {
					$createContactOkay = $true;	
				}
				else{
					foreach($Result in $ncCol){
						if($Result.Contact -eq $null){
							Write-host "Contact already exists " + $Result.Mailbox.Name
							throw ("Contact already exists")
						}
						else{
							if((Validate-EmailAddres -EmailAddress $EmailAddress)){
								if($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox){
									$UserDn = Get-UserDN -Credentials $Credentials -EmailAddress $Result.Mailbox.Address
									$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
									$ncCola = $service.ResolveName($UserDn,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly,$true,$cnpsPropset);
									if ($ncCola.Count -eq 0) {  
										$createContactOkay = $true;		
									}
									else
									{
										Write-Host -ForegroundColor  Red ("Number of existing Contacts Found " + $ncCola.Count)
										foreach($Result in $ncCola){
											Write-Host -ForegroundColor  Red ($ncCola.Mailbox.Name)
										}
										throw ("Contact already exists")
									}
								}
							}
							else{
								Write-Host -ForegroundColor Yellow ("Email Address is not valid for GAL match")
							}
						}
					}
				}
				if($createContactOkay){
					$Contact = New-Object Microsoft.Exchange.WebServices.Data.Contact -ArgumentList $service 
					#Set the GivenName
					$Contact.GivenName = $FirstName
					#Set the LastName
					$Contact.Surname = $LastName
					#Set Subject  
					$Contact.Subject = $DisplayName
					$Contact.FileAs = $DisplayName
					if($Title -ne ""){
						$PR_DISPLAY_NAME_PREFIX_W = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3A45,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
						$Contact.SetExtendedProperty($PR_DISPLAY_NAME_PREFIX_W,$Title)						
					}
					$Contact.CompanyName = $CompanyName
					$Contact.DisplayName = $DisplayName
					$Contact.Department = $Department
					$Contact.OfficeLocation = $Office
					$Contact.CompanyName = $CompanyName
					$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $BusinssPhone
					$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $MobilePhone
					$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] = $HomePhone
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] = New-Object  Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street = $Street
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State = $State
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City = $City
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion = $Country
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode = $PostalCode
					$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $EmailAddress
					if([string]::IsNullOrEmpty($EmailAddressDisplayAs)-eq $false){
						$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name = $EmailAddressDisplayAs
					} 
					$Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = $IMAddress 
					$Contact.FileAs = $FileAs
					$Contact.BusinessHomePage = $WebSite
					#Set any Notes  
					$Contact.Body = $Notes
					$Contact.JobTitle = $JobTitle
					if($Photo){
						$fileAttach = $Contact.Attachments.AddFileAttachment($Photo)
						$fileAttach.IsContactPhoto = $true
					}
			   		$Contact.Save($Contacts.Id)				
					Write-Host ("Contact Created")
				}
			}
		}
	}
}
####################### 
<# 
.SYNOPSIS 
 Gets a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Gets a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To get a Contact from a Mailbox's default contacts folder
	Get-Contact -MailboxName mailbox@domain.com -EmailAddress contact@email.com
	
	Example 2  The Partial Switch can be used to do partial match searches. Eg to return all the contacts that contain a particular word (note this could be across all the properties that are searched) you can use
	Get-Contact -MailboxName mailbox@domain.com -EmailAddress glen -Partial

	Example 3 By default only the Primary Email of a contact is checked when you using ResolveName if you want it to search the multivalued Proxyaddressses property you need to use something like the following
	Get-Contact -MailboxName  mailbox@domain.com -EmailAddress smtp:info@domain.com -Partial

    Example 4 Or to search via the SIP address you can use
	Get-Contact -MailboxName  mailbox@domain.com -EmailAddress sip:info@domain.com -Partial

#> 
########################
function Get-Contact 
{
   [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [string]$EmailAddress,
		[Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=3, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=4, Mandatory=$false)] [switch]$SearchGal,
		[Parameter(Position=5, Mandatory=$false)] [switch]$Partial,
		[Parameter(Position=6, Mandatory=$false)] [switch]$useImpersonation
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($SearchGal)
		{
			$Error.Clear();
			$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
			$ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly,$true,$cnpsPropset);
			if($Error.Count -eq 0){
				foreach($Result in $ncCol){	
					if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
						Write-Output $ncCol.Contact
					}
					else{
						Write-host -ForegroundColor Yellow ("Partial Match found but not returned because Primary Email Address doesn't match consider using -Partial " + $ncCol.Contact.DisplayName + " : Subject-" + $ncCol.Contact.Subject + " : Email-" + $Result.Mailbox.Address)
					}
				}
			}
		}
		else
		{
			if($Folder){
				$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
			}
			else{
				$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
			}
			if($service.URL){
				$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
				$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
				$ParentFolderIds = [Activator]::CreateInstance($type)
				$ParentFolderIds.Add($Contacts.Id)
				$Error.Clear();
				$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
				$ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts,$true,$cnpsPropset);
				if($Error.Count -eq 0){
					if ($ncCol.Count -eq 0) {
						Write-Host -ForegroundColor Yellow ("No Contact Found")		
					}
					else{
						$ResultWritten = $false
						foreach($Result in $ncCol){
							if($Result.Contact -eq $null){
								if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
									$Contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$Result.Mailbox.Id)
									Write-Output $Contact  
									$ResultWritten = $true
								}
							}
							else{
							
								if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
									if($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox){
										$ResultWritten = $true
										$UserDn = Get-UserDN -EmailAddress $Result.Mailbox.Address -Credentials $Credentials 
										$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
										$ncCola = $service.ResolveName($UserDn,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly,$true,$cnpsPropset);
										if ($ncCola.Count -eq 0) {  
											#Write-Host -ForegroundColor Yellow ("No Contact Found")			
										}
										else
										{
											$ResultWritten = $true
											Write-Host ("Number of matching Contacts Found " + $ncCola.Count)
											foreach($aResult in $ncCola){
												$Contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$aResult.Mailbox.Id)
												Write-Output $Contact
											}
											
										}
									}
								}
							}
							
						}
						if(!$ResultWritten){
							Write-Host -ForegroundColor Yellow ("No Contract Found")
						}
					}
				}

				
			}
		}
	}
}
####################### 
<# 
.SYNOPSIS 
 Gets a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Gets a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To get a Contact from a Mailbox's default contacts folder
	Get-Contacts -MailboxName mailbox@domain.com 
	
	Example 2 To get all the Contacts from subfolder of the Mailbox's default contacts folder
	Get-Contacts -MailboxName mailbox@domain.com -Folder \Contact\test
	
#> 
########################
function Get-Contacts 
{
   [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=3, Mandatory=$false)] [switch]$useImpersonation
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		if($service.URL){
			$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,"IPM.Contact") 
			#Define ItemView to retrive just 1000 Items    
			$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
			$fiItems = $null    
			do{    
			    $fiItems = $service.FindItems($Contacts.Id,$SfSearchFilter,$ivItemView)    
			    if($fiItems.Items.Count -gt 0){
					$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
					[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
				    foreach($Item in $fiItems.Items){      
						Write-Output $Item    
				    }
				}
			    $ivItemView.Offset += $fiItems.Items.Count    
			}while($fiItems.MoreAvailable -eq $true) 

		}
	}
}


####################### 
<# 
.SYNOPSIS 
 Updates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Updates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example1 Update the phone number of an existing contact
	Update-Contact  -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -MobilePhone 023213421 

 	Example 2 Update the phone number of a contact in a users subfolder
	Update-Contact  -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -MobilePhone 023213421 -Folder "\MyCustomContacts"
#> 
########################
function Update-Contact
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
 		[Parameter(Position=1, Mandatory=$false)] [string]$DisplayName,
		[Parameter(Position=2, Mandatory=$false)] [string]$FirstName,
		[Parameter(Position=3, Mandatory=$false)] [string]$LastName,
		[Parameter(Position=4, Mandatory=$true)] [string]$EmailAddress,
		[Parameter(Position=5, Mandatory=$false)] [string]$CompanyName,
		[Parameter(Position=6, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=7, Mandatory=$false)] [string]$Department,
		[Parameter(Position=8, Mandatory=$false)] [string]$Office,
		[Parameter(Position=9, Mandatory=$false)] [string]$BusinssPhone,
		[Parameter(Position=10, Mandatory=$false)] [string]$MobilePhone,
		[Parameter(Position=11, Mandatory=$false)] [string]$HomePhone,
		[Parameter(Position=12, Mandatory=$false)] [string]$IMAddress,
		[Parameter(Position=13, Mandatory=$false)] [string]$Street,
		[Parameter(Position=14, Mandatory=$false)] [string]$City,
		[Parameter(Position=15, Mandatory=$false)] [string]$State,
		[Parameter(Position=16, Mandatory=$false)] [string]$PostalCode,
		[Parameter(Position=17, Mandatory=$false)] [string]$Country,
		[Parameter(Position=18, Mandatory=$false)] [string]$JobTitle,
		[Parameter(Position=19, Mandatory=$false)] [string]$Notes,
		[Parameter(Position=20, Mandatory=$false)] [string]$Photo,
		[Parameter(Position=21, Mandatory=$false)] [string]$FileAs,
		[Parameter(Position=22, Mandatory=$false)] [string]$WebSite,
		[Parameter(Position=23, Mandatory=$false)] [string]$Title,
		[Parameter(Position=24, Mandatory=$false)] [string]$Folder,
		[Parameter(Mandatory=$false)] [switch]$Partial,
		[Parameter(Mandatory=$false)] [switch]$force,
		[Parameter(Position=25, Mandatory=$false)] [string]$EmailAddressDisplayAs,
		[Parameter(Position=26, Mandatory=$false)] [switch]$useImpersonation
    )  
 	Begin
	{
		if($Partial.IsPresent){$force = $false}
		if($Folder){
			if($Partial.IsPresent){
				$Contacts = Get-Contact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Folder $Folder -Partial
			}
			else{
				$Contacts = $Contacts = Get-Contact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Folder $Folder
			}
		}
		else{
			if($Partial.IsPresent){
				$Contacts = Get-Contact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials  -Partial
			}
			else{
				$Contacts = $Contacts = Get-Contact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials 
			}
		}	

		$Contacts | ForEach-Object{
			$Contact = $_
			if(($Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
				$updateOkay = $false
				if($force){
					$updateOkay = $true
				}
				else
				{
					$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""  
		            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","" 
		            $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)  
		            $message = "Do you want to update contact with DisplayName " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address 
		            $result = $Host.UI.PromptForChoice($caption,$message,$choices,1)  
		            if($result -eq 0) {                       
						$updateOkay = $true
		            } 
					else{
						Write-Host ("No Action Taken")
					}				
				}
				if($updateOkay){
					if($FirstName -ne ""){
						$Contact.GivenName = $FirstName
					}
					if($LastName -ne ""){
						$Contact.Surname = $LastName
					}
					if($DisplayName -ne ""){
						$Contact.Subject = $DisplayName
					}
					if($Title -ne ""){
						$PR_DISPLAY_NAME_PREFIX_W = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3A45,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
						$Contact.SetExtendedProperty($PR_DISPLAY_NAME_PREFIX_W,$Title)						
					}
					if($CompanyName -ne ""){
						$Contact.CompanyName = $CompanyName
					}
					if($DisplayName -ne ""){
						$Contact.DisplayName = $DisplayName
					}
					if($Department -ne ""){
						$Contact.Department = $Department
					}
					if($Office -ne ""){
						$Contact.OfficeLocation = $Office
					}
					if($CompanyName -ne ""){
						$Contact.CompanyName = $CompanyName
					}
					if($BusinssPhone -ne ""){
						$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $BusinssPhone
					}
					if($MobilePhone -ne ""){
						$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $MobilePhone
					}
					if($HomePhone -ne ""){
						$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] = $HomePhone
					}
					if($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business]  -eq $null){
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] = New-Object  Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
					}
					if($Street -ne ""){
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street = $Street
					}
					if($State -ne ""){
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State = $State
					}
					if($City -ne ""){
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City = $City
					}
					if($Country -ne ""){
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion = $Country
					}
					if($PostalCode -ne ""){
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode = $PostalCode
					}
					if($EmailAddress -ne ""){
						$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $EmailAddress
					}
					if([string]::IsNullOrEmpty($EmailAddressDisplayAs)-eq $false){
						$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name = $EmailAddressDisplayAs
					} 
					if($IMAddress -ne ""){
						$Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = $IMAddress 
					}
					if($FileAs -ne ""){
						$Contact.FileAs = $FileAs
					}
					if($WebSite -ne ""){
						$Contact.BusinessHomePage = $WebSite
					}
					if($Notes -ne ""){  
						$Contact.Body = $Notes
					}
					if($JobTitle -ne ""){
						$Contact.JobTitle = $JobTitle
					}
					if($Photo){
						$fileAttach = $Contact.Attachments.AddFileAttachment($Photo)
						$fileAttach.IsContactPhoto = $true
					}
					$Contact.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
					"Contact updated " + $Contact.Subject
				
				}
			}
		}
	}
}

####################### 
<# 
.SYNOPSIS 
 Deletes a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Deletes a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE 
	Example 1 To delete a contact from the default contacts folder
	Delete-Contact -MailboxName mailbox@domain.com -EmailAddress email@domain.com 

	Example2 To delete a contact from a non user subfolder
	Delete-Contact -MailboxName mailbox@domain.com -EmailAddress email@domain.com -Folder \Contacts\Subfolder
#> 
########################
function Delete-Contact 
{

   [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [string]$EmailAddress,
		[Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=3, Mandatory=$false)] [switch]$force,
		[Parameter(Position=4, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=5, Mandatory=$false)] [switch]$Partial
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		if($service.URL){
			$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
			$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
			$ParentFolderIds = [Activator]::CreateInstance($type)
			$ParentFolderIds.Add($Contacts.Id)
			$Error.Clear();
			$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
			$ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts,$true,$cnpsPropset);
			if($Error.Count -eq 0){
				if ($ncCol.Count -eq 0) {
					Write-Host -ForegroundColor Yellow ("No Contact Found")		
				}
				else{
					foreach($Result in $ncCol){
						if($Result.Contact -eq $null){
							$contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$Result.Mailbox.Id) 
							if($force){
								if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower())){
									$contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)  
									Write-Host ("Contact Deleted " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address)
								}
								else
								{
									Write-Host ("This script won't allow you to force the delete of partial matches")
								}
							}
							else{
								if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
								    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""  
		                            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","" 
		                           
		                            $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)  
		                            $message = "Do you want to Delete contact with DisplayName " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address
		                            $result = $Host.UI.PromptForChoice($caption,$message,$choices,1)  
		                            if($result -eq 0) {                       
		                                $contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete) 
										Write-Host ("Contact Deleted")
		                            } 
									else{
										Write-Host ("No Action Taken")
									}
								}
								
							}
						}
						else{
							if((Validate-EmailAddres -EmailAddress $Result.Mailbox.Address)){
							    if($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox){
									$UserDn = Get-UserDN -Credentials $Credentials -EmailAddress $Result.Mailbox.Address
									$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
									$ncCola = $service.ResolveName($UserDn,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly,$true,$cnpsPropset);
									if ($ncCola.Count -eq 0) {  
										Write-Host -ForegroundColor Yellow ("No Contact Found")			
									}
									else
									{
										Write-Host ("Number of matching Contacts Found " + $ncCola.Count)
										$rtCol = @()
										foreach($aResult in $ncCola){
											if(($aResult.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
												$contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$aResult.Mailbox.Id) 
												if($force){
													if($aResult.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()){
														$contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)  
														Write-Host ("Contact Deleted " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address)
													}
													else
													{
														Write-Host ("This script won't allow you to force the delete of partial matches")
													}
												}
												else{
												    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""  
						                            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","" 
						                            $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)  
						                            $message = "Do you want to Delete contact with DisplayName " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address 
						                            $result = $Host.UI.PromptForChoice($caption,$message,$choices,1)  
						                            if($result -eq 0) {                       
						                                $contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete) 
														Write-Host ("Contact Deleted ")
						                            } 
													else{
														Write-Host ("No Action Taken")
													}
													
												}
											}
											else{
												Write-Host ("Skipping Matching because email address doesn't match address on match " + $aResult.Mailbox.Address.ToLower())
											}
										}								
									}
								}
							}
							else
							{
								Write-Host -ForegroundColor Yellow ("Email Address is not valid for GAL match")
							}
						}
					}
				}
			}	
			
		}
	}
}
function Make-UniqueFileName{
    param(
		[Parameter(Position=0, Mandatory=$true)] [string]$FileName
	)
	Begin
	{
	
	$directoryName = [System.IO.Path]::GetDirectoryName($FileName)
    $FileDisplayName = [System.IO.Path]::GetFileNameWithoutExtension($FileName);
    $FileExtension = [System.IO.Path]::GetExtension($FileName);
    for ($i = 1; ; $i++){
            
            if (![System.IO.File]::Exists($FileName)){
				return($FileName)
			}
			else{
					$FileName = [System.IO.Path]::Combine($directoryName, $FileDisplayName + "(" + $i + ")" + $FileExtension);
			}                
            
			if($i -eq 10000){throw "Out of Range"}
        }
	}
}

function Get-ContactFolder{
	param (
	        [Parameter(Position=0, Mandatory=$true)] [string]$FolderPath,
			[Parameter(Position=1, Mandatory=$true)] [string]$SmptAddress,
			[Parameter(Position=2, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
		  )
	process{
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$SmptAddress)   
		$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
		#Split the Search path into an array  
		$fldArray = $FolderPath.Split("\") 
		 #Loop through the Split Array and do a Search for each level of folder 
		for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
	        #Perform search based on the displayname of each folder level 
	        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
	        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$fldArray[$lint]) 
	        $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView) 
	        if ($findFolderResults.TotalCount -gt 0){ 
	            foreach($folder in $findFolderResults.Folders){ 
	                $tfTargetFolder = $folder                
	            } 
	        } 
	        else{ 
	            Write-host ("Error Folder Not Found check path and try again")  
	            $tfTargetFolder = $null  
	            break  
	        }     
	    }  
		if($tfTargetFolder -ne $null){
			return [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$tfTargetFolder.Id)
		}
		else{
			throw ("Folder Not found")
		}
	}
}
####################### 
<# 
.SYNOPSIS 
 Exports a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API to a VCF File 
 
.DESCRIPTION 
  Exports a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE 

	Example 1 To Export a contact to local file
	Export-Contact -MailboxName mailbox@domain.com -EmailAddress address@domain.com -FileName c:\export\filename.vcf
	If the file already exists it will handle creating a unique filename

	Example 2 To export from a contacts subfolder use
	Export-Contact -MailboxName mailbox@domain.com -EmailAddress address@domain.com -FileName c:\export\filename.vcf -folder \contacts\subfolder

#> 
########################
function Export-Contact 
{
   [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [string]$EmailAddress,
		[Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=3, Mandatory=$true)] [string]$FileName,
		[Parameter(Position=4, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=5, Mandatory=$false)] [switch]$Partial
		
    )  
 	Begin
	{
		if($Folder){
			if($Partial.IsPresent){
				$Contacts = Get-Contact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Folder $Folder -Partial
			}
			else{
				$Contacts = $Contacts = Get-Contact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Folder $Folder
			}
		}
		else{
			if($Partial.IsPresent){
				$Contacts = Get-Contact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials  -Partial
			}
			else{
				$Contacts = $Contacts = Get-Contact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials 
			}
		}	

		$Contacts | ForEach-Object{
			$Contact = $_
			$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)    
		  	$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent); 
			$Contact.load($psPropset)
			$FileName = Make-UniqueFileName -FileName $FileName
			[System.IO.File]::WriteAllBytes($FileName,$Contact.MimeContent.Content) 
		    write-host ("Exported " + $FileName)  
		
		}
		

	}
}
####################### 
<# 
.SYNOPSIS 
 Exports a Contact from the Global Address List on an Exchange Server using the  Exchange Web Services API to a VCF File 
 
.DESCRIPTION 
  Exports a Contact from the Global Address List on an Exchange Server using the  Exchange Web Services API to a VCF File 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE 

	Example 1 To export a GAL Entry to a vcf file 
	Export-GalContact -MailboxName user@domain.com -EmailAddress email@domain.com -FileName c:\export\export.vcf

	Example 2 To export a GAL Entry to vcf including the users photo
	Export-GalContact -MailboxName user@domain.com -EmailAddress email@domain.com -FileName c:\export\export.vcf -IncludePhoto

#> 
########################
function Export-GALContact 
{
   [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [string]$EmailAddress,
		[Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=3, Mandatory=$false)] [switch]$IncludePhoto,
		[Parameter(Position=4, Mandatory=$true)] [string]$FileName,
		[Parameter(Position=5, Mandatory=$false)] [switch]$Partial
    )  
 	Begin
	{
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		$Error.Clear();
		$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
		$ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly,$true,$cnpsPropset);
		if($Error.Count -eq 0){
			foreach($Result in $ncCol){				
				if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
					$ufilename = Make-UniqueFileName -FileName $FileName
					Set-content -path $ufilename "BEGIN:VCARD" 
					add-content -path $ufilename "VERSION:2.1"
					$givenName = ""
					if($ncCol.Contact.GivenName -ne $null){
						$givenName = $ncCol.Contact.GivenName
					}
					$surname = ""
					if($ncCol.Contact.Surname -ne $null){
						$surname = $ncCol.Contact.Surname
					}
					add-content -path $ufilename ("N:" + $surname + ";" + $givenName)
					add-content -path $ufilename ("FN:" + $ncCol.Contact.DisplayName)
					$Department = "";
					if($ncCol.Contact.Department -ne $null){
						$Department = $ncCol.Contact.Department
					}
				
					$CompanyName = "";
					if($ncCol.Contact.CompanyName -ne $null){
						$CompanyName = $ncCol.Contact.CompanyName
					}
					add-content -path $ufilename ("ORG:" + $CompanyName + ";" + $Department)	
					if($ncCol.Contact.JobTitle -ne $null){
						add-content -path $ufilename ("TITLE:" + $ncCol.Contact.JobTitle)
					}
					if($ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] -ne $null){
						add-content -path $ufilename ("TEL;CELL;VOICE:" + $ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone])		
					}
					if($ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] -ne $null){
						add-content -path $ufilename ("TEL;HOME;VOICE:" + $ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone])		
					}
					if($ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] -ne $null){
						add-content -path $ufilename ("TEL;WORK;VOICE:" + $ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone])		
					}
					if($ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessFax] -ne $null){
						add-content -path $ufilename ("TEL;WORK;FAX:" + $ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessFax])
					}
					if($ncCol.Contact.BusinessHomePage -ne $null){
						add-content -path $ufilename ("URL;WORK:" + $ncCol.Contact.BusinessHomePage)
					}
					if($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] -ne $null){
						if($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion -ne $null){
							$Country = $ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion.Replace("`n","")
						}
						if($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City -ne $null){
							$City = $ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City.Replace("`n","")
						}
						if($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street -ne $null){
							$Street = $ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street.Replace("`n","")
						}
						if($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State -ne $null){
							$State = $ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State.Replace("`n","")
						}
						if($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode -ne $null){
							$PCode = $ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode.Replace("`n","")
						}
						$addr = "ADR;WORK;PREF:;" + $Country + ";" + $Street + ";" + $City + ";" + $State + ";" + $PCode + ";" + $Country
						add-content -path $ufilename $addr
					}
					if($ncCol.Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] -ne $null){
						add-content -path $ufilename ("X-MS-IMADDRESS:" + $ncCol.Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1])
					}
					add-content -path $ufilename ("EMAIL;PREF;INTERNET:" + $ncCol.Mailbox.Address)
					
					
					if($IncludePhoto){
						$PhotoURL = AutoDiscoverPhotoURL -EmailAddress $MailboxName  -Credentials $Credentials
						$PhotoSize = "HR120x120" 
						$PhotoURL= $PhotoURL + "/GetUserPhoto?email="  + $ncCol.Mailbox.Address + "&size=" + $PhotoSize;
						$wbClient = new-object System.Net.WebClient
						$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
						$wbClient.Credentials = $creds
						$photoBytes = $wbClient.DownloadData($PhotoURL);
						add-content -path $ufilename "PHOTO;ENCODING=BASE64;TYPE=JPEG:"
						$ImageString = [System.Convert]::ToBase64String($photoBytes,[System.Base64FormattingOptions]::InsertLineBreaks)
						add-content -path $ufilename $ImageString
						add-content -path $ufilename "`r`n"	
					}
					add-content -path $ufilename "END:VCARD"	
					Write-Host ("Contact exported to " + $ufilename)			
				}						
			}
		}
	}
}

function Export-ContactsFolderToCSV{
	    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=3, Mandatory=$false)] [switch]$force,
		[Parameter(Position=4, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=5, Mandatory=$false)] [switch]$Partial,
		[Parameter(Position=6, Mandatory=$true)] [string]$FileName
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		Invoke-ExportContactFolderToCSV -Contacts $Contacts -FileName $FileName
		
	}
}

function Export-PublicFolderContactsFolderToCSV{
	    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=3, Mandatory=$false)] [switch]$force,
		[Parameter(Position=4, Mandatory=$false)] [string]$PublicFolderPath,
		[Parameter(Position=5, Mandatory=$false)] [switch]$Partial,
		[Parameter(Position=6, Mandatory=$true)] [string]$FileName
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName)
		$fldId = PublicFolderIdFromPath -FolderPath $PublicFolderPath  -SmtpAddress $MailboxName -service $service		
		$ContactsId =  new-object Microsoft.Exchange.WebServices.Data.FolderId($fldId)
		$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$ContactsId)
		Get-PublicFolderRoutingHeader -service $service -Credentials $Credentials -MailboxName $MailboxName -Header "X-AnchorMailbox"
		Invoke-ExportContactFolderToCSV -Contacts $Contacts -FileName $FileName
		
	}
}
function Get-PublicFolderRoutingHeader
{
    param (
	        [Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
            [Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		    [Parameter(Position=2, Mandatory=$true)] [string]$MailboxName,
            [Parameter(Position=3, Mandatory=$true)] [string]$Header
          )
	process
    {
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
        $AutoDiscoverService =  New-Object  Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($ExchangeVersion);
        $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
        $AutoDiscoverService.Credentials = $creds
        $AutoDiscoverService.EnableScpLookup = $false;
        $AutoDiscoverService.RedirectionUrlValidationCallback = {$true};
        $AutoDiscoverService.PreAuthenticate = $true;
        $AutoDiscoverService.KeepAlive = $false;      
        if($Header -eq "X-AnchorMailbox")
        {
            $gsp = $AutoDiscoverService.GetUserSettings($MailboxName,[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation);
            $PublicFolderInformation = $null
            if ($gsp.Settings.TryGetValue([Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation, [ref] $PublicFolderInformation))
            {
				write-host ("Public Folder Routing Information Header : " + $PublicFolderInformation) 
				if(!$service.HttpHeaders.ContainsKey($Header)){
					$service.HttpHeaders.Add($Header, $PublicFolderInformation) 
				} 
				else{
					$service.HttpHeaders[$Header] = $PublicFolderInformation
				}                         
                        
            } 
            
        }

       
    }
    
}

function Get-PublicFolderContentRoutingHeader
{
    param (
	        [Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
            [Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		    [Parameter(Position=2, Mandatory=$true)] [string]$MailboxName,
            [Parameter(Position=3, Mandatory=$true)] [string]$pfAddress
     )
	process
    {
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
        $AutoDiscoverService =  New-Object  Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($ExchangeVersion);
        $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
        $AutoDiscoverService.Credentials = $creds
        $AutoDiscoverService.EnableScpLookup = $false;
        $AutoDiscoverService.RedirectionUrlValidationCallback = {$true};
        $AutoDiscoverService.PreAuthenticate = $true;
        $AutoDiscoverService.KeepAlive = $false;      
        $gsp = $AutoDiscoverService.GetUserSettings($MailboxName,[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AutoDiscoverSMTPAddress);
        #Write-Host $AutoDiscoverService.url
        $auDisXML = "<Autodiscover xmlns=`"http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006`"><Request>`r`n" +
        "<EMailAddress>" + $pfAddress + "</EMailAddress>`r`n" +
        "<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>`r`n" +
        "</Request>`r`n" +
        "</Autodiscover>`r`n";
        $AutoDiscoverRequest = [System.Net.HttpWebRequest]::Create($AutoDiscoverService.url.ToString().replace(".svc",".xml"));
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($auDisXML);
        $AutoDiscoverRequest.ContentLength = $bytes.Length;
        $AutoDiscoverRequest.ContentType = "text/xml";
        $AutoDiscoverRequest.UserAgent = "Microsoft Office/16.0 (Windows NT 6.3; Microsoft Outlook 16.0.6001; Pro)";            
        $AutoDiscoverRequest.Headers.Add("Translate", "F");
        $AutoDiscoverRequest.Method = "POST";
        $AutoDiscoverRequest.Credentials = $creds;
        $RequestStream = $AutoDiscoverRequest.GetRequestStream();
        $RequestStream.Write($bytes, 0, $bytes.Length);
        $RequestStream.Close();
        $AutoDiscoverRequest.AllowAutoRedirect = $truee;
        $Response = $AutoDiscoverRequest.GetResponse().GetResponseStream()
        $sr = New-Object System.IO.StreamReader($Response)
        [XML]$xmlReposne = $sr.ReadToEnd()
        if($xmlReposne.Autodiscover.Response.User.AutoDiscoverSMTPAddress -ne $null)
        {
            write-host ("Public Folder Content Routing Information Header : " + $xmlReposne.Autodiscover.Response.User.AutoDiscoverSMTPAddress)  
            $service.HttpHeaders["X-AnchorMailbox"] = $xmlReposne.Autodiscover.Response.User.AutoDiscoverSMTPAddress    
            $service.HttpHeaders["X-PublicFolderMailbox"] = $xmlReposne.Autodiscover.Response.User.AutoDiscoverSMTPAddress              
        }

    }
    
}
function PublicFolderIdFromPath{
	param (
            [Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
            [Parameter(Position=1, Mandatory=$true)] [String]$FolderPath,
            [Parameter(Position=2, Mandatory=$true)] [String]$SmtpAddress
		  )
	process{
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)   
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $PR_REPLICA_LIST = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6698,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary); 
        $psPropset.Add($PR_REPLICA_LIST)
		$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)  
        $PR_REPLICA_LIST_Value = $null
        if($tfTargetFolder.TryGetProperty($PR_REPLICA_LIST,[ref]$PR_REPLICA_LIST_Value)){
                 $GuidAsString = [System.Text.Encoding]::ASCII.GetString($PR_REPLICA_LIST_Value, 0, 36);
                 $HeaderAddress = new-object System.Net.Mail.MailAddress($service.HttpHeaders["X-AnchorMailbox"])
                 $pfHeader = $GuidAsString + "@" + $HeaderAddress.Host
                 write-host ("Root Public Folder Routing Information Header : " + $pfHeader )  
                 $service.HttpHeaders.Add("X-PublicFolderMailbox", $pfHeader)    
        }
		#Split the Search path into an array  
		$fldArray = $FolderPath.Split("\") 
		 #Loop through the Split Array and do a Search for each level of folder 
		for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
	        #Perform search based on the displayname of each folder level 
	        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
            $fvFolderView.PropertySet = $psPropset
	        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$fldArray[$lint]) 
	        $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView) 
	        if ($findFolderResults.TotalCount -gt 0){ 
	            foreach($folder in $findFolderResults.Folders){ 
	                $tfTargetFolder = $folder                
	            } 
	        } 
	        else{ 
	            "Error Folder Not Found"  
	            $tfTargetFolder = $null  
	            break  
	        }     
	    }  
		if($tfTargetFolder -ne $null){
            $PR_REPLICA_LIST_Value = $null
            if($tfTargetFolder.TryGetProperty($PR_REPLICA_LIST,[ref]$PR_REPLICA_LIST_Value)){
                    $GuidAsString = [System.Text.Encoding]::ASCII.GetString($PR_REPLICA_LIST_Value, 0, 36);
                    $HeaderAddress = new-object System.Net.Mail.MailAddress($service.HttpHeaders["X-AnchorMailbox"])
                    $pfHeader = $GuidAsString + "@" + $HeaderAddress.Host
                    write-host ("Target Public Folder Routing Information Header : " + $pfHeader )  
                    Get-PublicFolderContentRoutingHeader -service $service -Credentials $Credentials -MailboxName $SmtpAddress -pfAddress $pfHeader
            }            
			return $tfTargetFolder.Id.UniqueId.ToString()
		}
		else{
			throw "Folder not found"
		}
	}
}
function Invoke-ExportContactFolderToCSV{

	   [CmdletBinding()] 
    param( 
		[Parameter(Position=1, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.Folder]$Contacts,		
		[Parameter(Position=2, Mandatory=$true)] [string]$FileName
    )  
 	Begin
	{
		    $ExportCollection = @()
			$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)   
			$PR_Gender = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(14925,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Short)  
			$psPropset.Add($PR_Gender)  			
			#Define ItemView to retrive just 1000 Items      
			$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)      
			$fiItems = $null      
			do{      
				$fiItems = $service.FindItems($Contacts.Id,$ivItemView)   
				[Void]$Contacts.Service.LoadPropertiesForItems($fiItems,$psPropset)    
				foreach($Item in $fiItems.Items){       
					if($Item -is [Microsoft.Exchange.WebServices.Data.Contact]){  
						$expObj = "" | select DisplayName,GivenName,Surname,Gender,Email1DisplayName,Email1Type,Email1EmailAddress,BusinessPhone,MobilePhone,HomePhone,BusinessStreet,BusinessCity,BusinessState,HomeStreet,HomeCity,HomeState  
						$expObj.DisplayName = $Item.DisplayName  
						$expObj.GivenName = $Item.GivenName  
						$expObj.Surname = $Item.Surname  
						$expObj.Gender = ""  
						$Gender = $null  
						if($item.TryGetProperty($PR_Gender,[ref]$Gender)){  
							if($Gender -eq 2){  
								$expObj.Gender = "Male"   
							}  
							if($Gender -eq 1){  
								$expObj.Gender = "Female"   
							}  
						}  
						$BusinessPhone = $null  
						$MobilePhone = $null  
						$HomePhone = $null  
						if($Item.PhoneNumbers -ne $null){  
							if($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone,[ref]$BusinessPhone)){  
								$expObj.BusinessPhone = $BusinessPhone  
							}  
							if($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone,[ref]$MobilePhone)){  
								$expObj.MobilePhone = $MobilePhone  
							}     
							if($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone,[ref]$HomePhone)){  
								$expObj.HomePhone = $HomePhone  
							}     
						}             
						if($Item.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1)){                  
							$expObj.Email1DisplayName = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name  
							$expObj.Email1Type = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].RoutingType  
							$expObj.Email1EmailAddress = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address  
						}  
						$HomeAddress = $null  
						$BusinessAddress = $null  
						if($item.PhysicalAddresses -ne $null){  
							if($item.PhysicalAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home,[ref]$HomeAddress)){  
								$expObj.HomeStreet = $HomeAddress.Street  
								$expObj.HomeCity = $HomeAddress.City  
								$expObj.HomeState = $HomeAddress.State  
							}  
							if($item.PhysicalAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business,[ref]$BusinessAddress)){  
								$expObj.BusinessStreet = $BusinessAddress.Street  
								$expObj.BusinessCity = $BusinessAddress.City  
								$expObj.BusinessState = $BusinessAddress.State  
							}  
						}  
						
						$ExportCollection += $expObj  
					}  
				}      
				$ivItemView.Offset += $fiItems.Items.Count      
			}while($fiItems.MoreAvailable -eq $true)   		

			$ExportCollection | Export-Csv -NoTypeInformation -Path $FileName
			"Exported to " + $FileName 
	} 
}


function AutoDiscoverPhotoURL{
       param (
              $EmailAddress="$( throw 'Email is a mandatory Parameter' )",
              $Credentials="$( throw 'Credentials is a mandatory Parameter' )"
              )
       process{
              $version= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
              $adService= New-Object Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($version);
			  $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
              $adService.Credentials = $creds
              $adService.EnableScpLookup=$false;
              $adService.RedirectionUrlValidationCallback= {$true}
              $adService.PreAuthenticate=$true;
              $UserSettings= new-object Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[] 1
              $UserSettings[0] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::ExternalPhotosUrl
              $adResponse=$adService.GetUserSettings($EmailAddress, $UserSettings)
              $PhotoURI= $adResponse.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::ExternalPhotosUrl]
              return $PhotoURI.ToString()
       }
}
Function Validate-EmailAddres
{
	 param( 
	 	[Parameter(Position=0, Mandatory=$true)] [string]$EmailAddress
	 )
	 Begin
	{
 		try
		{
  			$check = New-Object System.Net.Mail.MailAddress($EmailAddress)
 			 return $true
 		}
 		catch
 		{
  			return $false
 		}
   }
}
####################### 
<# 
.SYNOPSIS 
 Copies a Contact from the Global Address List to a Local Mailbox Contacts folder using the  Exchange Web Services API  
 
.DESCRIPTION 
  Copies a Contact from the Global Address List to a Local Mailbox Contacts folder using the  Exchange Web Services API
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE 

	Example 1 To Copy a Gal contacts to local Contacts folder
	Copy-Contacts.GalToMailbox -MailboxName mailbox@domain.com -EmailAddress email@domain.com  

 	Example 2 Copy a GAL contact to a Contacts subfolder
	Copy-Contacts.GalToMailbox -MailboxName mailbox@domain.com -EmailAddress email@domain.com  -Folder \Contacts\UnderContacts

#> 
########################
function Copy-Contacts.GalToMailbox
{
   [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [string]$EmailAddress,
		[Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=3, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=4, Mandatory=$false)] [switch]$IncludePhoto,
		[Parameter(Position=5, Mandatory=$false)] [switch]$useImpersonation
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		$Error.Clear();
		$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
		$ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly,$true,$cnpsPropset);
		if($Error.Count -eq 0){
			foreach($Result in $ncCol){				
				if($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()){					
					$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
					$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
					$ParentFolderIds = [Activator]::CreateInstance($type)
					$ParentFolderIds.Add($Contacts.Id)
					$Error.Clear();
					$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
					$ncCola = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts,$true,$cnpsPropset);
					$createContactOkay = $false
					if($Error.Count -eq 0){
						if ($ncCola.Count -eq 0) {							
						    $createContactOkay = $true;	
						}
						else{
							foreach($aResult in $ncCola){
								if($aResult.Contact -eq $null){
									Write-host "Contact already exists " + $aResult.Contact.DisplayName
									throw ("Contact already exists")
								}
								else{
									if((Validate-EmailAddres -EmailAddress $Result.Mailbox.Address)){
									    if($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox){
											$UserDn = Get-UserDN -Credentials $Credentials -EmailAddress $Result.Mailbox.Address
											$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
											$ncColb = $service.ResolveName($UserDn,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly,$true,$cnpsPropset);
											if ($ncColb.Count -eq 0) {  
												$createContactOkay = $true;		
											}
											else
											{
												Write-Host -ForegroundColor  Red ("Number of existing Contacts Found " + $ncColb.Count)
												foreach($Result in $ncColb){
													Write-Host -ForegroundColor  Red ($ncColb.Mailbox.Name)
												}
												throw ("Contact already exists")
											}
										}
									}
									else{
										Write-Host -ForegroundColor Yellow ("Email Address is not valid for GAL match")
									}
								}
							}
						}
						if($createContactOkay){
							#check for SipAddress
							$IMAddress = ""
							$emailVal = $null;
							if($ncCol.Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1,[ref]$emailVal)){
								$email1 = $emailVal.Address
								if($email1.tolower().contains("sip:")){
									$IMAddress = $email1
								}
							}
							$emailVal = $null;
							if($ncCol.Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2,[ref]$emailVal)){
								$email2 = $emailVal.Address
								if($email2.tolower().contains("sip:")){
									$IMAddress = $email2
									$ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2] = $null
								}
							}
							$emailVal = $null;
							if($ncCol.Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3,[ref]$emailVal)){
								$email3 =  $emailVal.Address
								if($email3.tolower().contains("sip:")){
									$IMAddress = $email3
									$ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3] = $null
								}
							}
							if($IMAddress -ne ""){
								$ncCol.Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = $IMAddress
							}	
    						$ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2] = $null
							$ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3] = $null
							$ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address = $ncCol.Mailbox.Address.ToLower()
							$ncCol.Contact.FileAs = $ncCol.Contact.DisplayName
							if($IncludePhoto){					
								$PhotoURL = AutoDiscoverPhotoURL -EmailAddress $MailboxName  -Credentials $Credentials
								$PhotoSize = "HR120x120" 
								$PhotoURL= $PhotoURL + "/GetUserPhoto?email="  + $ncCol.Mailbox.Address + "&size=" + $PhotoSize;
								$wbClient = new-object System.Net.WebClient
								$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
								$wbClient.Credentials = $creds
								$photoBytes = $wbClient.DownloadData($PhotoURL);
								$fileAttach = $ncCol.Contact.Attachments.AddFileAttachment("contactphoto.jpg",$photoBytes)
								$fileAttach.IsContactPhoto = $true
							}
							$ncCol.Contact.Save($Contacts.Id);
							Write-Host ("Contact copied")
						}
					}
				}
			}
		}
	}
}

function Get-UserDN{
	param (
			[Parameter(Position=0, Mandatory=$true)] [string]$EmailAddress,
			[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
		  )
	process{
		$ExchangeVersion= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
		$adService = New-Object Microsoft.Exchange.WebServices.AutoDiscover.AutodiscoverService($ExchangeVersion);
		$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
		$adService.Credentials = $creds
		$adService.EnableScpLookup = $false;
		$adService.RedirectionUrlValidationCallback = {$true}
		$UserSettings = new-object Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[] 1
		$UserSettings[0] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDN
		$adResponse = $adService.GetUserSettings($EmailAddress , $UserSettings);
		return $adResponse.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDN]
	}
}

####################### 
<# 
.SYNOPSIS 
 Creates a Contact Group in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Creates a Contact Group in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To create a Contact Group in the default contacts folder 
	Create-ContactGroup  -Mailboxname mailbox@domain.com -GroupName GroupName -Members ("member1@domain.com","member2@domain.com")
    Example 2 To create a Contact Group in a subfolder of default contacts folder 
	Create-ContactGroup  -Mailboxname mailbox@domain.com -GroupName GroupName -Folder \Contacts\Folder1 -Members ("member1@domain.com","member2@domain.com")

#> 
########################
function Create-ContactGroup 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=3, Mandatory=$true)] [string]$GroupName,
		[Parameter(Position=4, Mandatory=$true)] [PsObject]$Members,
		[Parameter(Position=5, Mandatory=$false)] [switch]$useImpersonation
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		if($service.URL){
			$ContactGroup = New-Object Microsoft.Exchange.WebServices.Data.ContactGroup -ArgumentList $service
			$ContactGroup.DisplayName = $GroupName
			foreach($Member in $Members){
				$ContactGroup.Members.Add($Member)
			}
			$ContactGroup.Save($Contacts.Id)
			Write-Host ("Contact Group created " + $GroupName)
		}
	}
}
####################### 
<# 
.SYNOPSIS 
 Gets a Contact Group in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Gets a Contact Group in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To Get a Contact Group in the default contacts folder 
	Get-ContactGroup  -Mailboxname mailbox@domain.com -GroupName GroupName 
    Example 2 To Get a Contact Group in a subfolder of default contacts folder 
	Get-ContactGroup  -Mailboxname mailbox@domain.com -GroupName GroupName -Folder \Contacts\Folder1 

#> 
########################
function Get-ContactGroup 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=3, Mandatory=$true)] [string]$GroupName,
		[Parameter(Position=6, Mandatory=$false)] [switch]$useImpersonation
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		if($service.URL){
			$SfSearchFilter1 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ContactGroupSchema]::DisplayName,$GroupName) 
			$SfSearchFilter2 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,"IPM.DistList") 
			$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);  
			$sfCollection.add($SfSearchFilter1)  
			$sfCollection.add($SfSearchFilter2)  
			#Define ItemView to retrive just 1000 Items    
			$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
			$fiItems = $null    
			do{    
			    $fiItems = $service.FindItems($Contacts.Id,$sfCollection,$ivItemView)    
				if($fiItems.Item.Count -eq 0){
					Write-Host ("No Groups Found with that Name")
				}
			    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
			    foreach($Item in $fiItems.Items){      
					Write-Output $Item
			    }    
			    $ivItemView.Offset += $fiItems.Items.Count    
			}while($fiItems.MoreAvailable -eq $true) 
		}
	}
}
####################### 
<# 
.SYNOPSIS 
 Search Contacts in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Searches Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951


#> 
########################
function Search-ContactsForCCNumbers 
{
   [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=3, Mandatory=$false)] [switch]$useImpersonation
    )  
 	Begin
	{
		$Script:rptCollection = @()
		Import-Module .\CreditCardValidator.dll -Force
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		if($service.URL){
			$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,"IPM.Contact") 
			#Define ItemView to retrive just 1000 Items    
			$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
			$fiItems = $null    
			do{    
			    $fiItems = $service.FindItems($Contacts.Id,$SfSearchFilter,$ivItemView)    
			    if($fiItems.Items.Count -gt 0){
					$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
					[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
				    foreach($Contact in $fiItems.Items){      
						if($Contact -is [Microsoft.Exchange.WebServices.Data.Contact]){
							$DnName = $Contact.DisplayName
							write-host ("Processing " + $DnName)
							$BusinssPhone = $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] 
							$MobilePhone = $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] 
					    	$HomePhone = $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone]
							if($BusinssPhone -ne $null){
								Write-host ('Check BusinessPhone')
								$CheckObj = DetectCCNumber -Number $BusinssPhone -Property "BusinessPhone" -MailboxName $MailboxName -DisplayName $DnName 								
							}
							if($MobilePhone -ne $null){
								Write-host ('Check MobilePhone')
								$CheckObj = DetectCCNumber -Number $MobilePhone -Property "MobilePhone" -MailboxName $MailboxName -DisplayName $DnName 
								
							}
							if($HomePhone -ne $null){
								Write-host ('Check HomePhone')
								$CheckObj = DetectCCNumber -Number $HomePhone -Property "HomePhone" -MailboxName $MailboxName -DisplayName $DnName 
								
							}
    						$Email1 = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1]
							$Email2 = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2]
							$Email3 = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3]
							if($Email1 -ne $null){
								Write-host ('Check Email1')
								if(![string]::IsNullOrEmpty($Email1.Address)){
									if([String]::IsNullOrEmpty($DnName)){
										$DnName =$Email1.Address						
									}
									$CheckObj = DetectCCNumber -Number $Email1.Address -Property "Email" -MailboxName $MailboxName -DisplayName $DnName 
								}
								
							}
						    if($Email2 -ne $null){
								Write-host ('Check Email2')
								if(![string]::IsNullOrEmpty($Email2.Address)){
									$CheckObj = DetectCCNumber -Number $Email2.Address -Property "Email2" -MailboxName $MailboxName -DisplayName $DnName 
								}	
								
							}
						    if($Email3 -ne $null){
								Write-host ('Check Email3')
								if(![string]::IsNullOrEmpty($Email3.Address)){
									$CheckObj = DetectCCNumber -Number $Email3.Address -Property "Email3" -MailboxName $MailboxName -DisplayName $DnName 
								}								
							}

						}
				    }
				}
			    $ivItemView.Offset += $fiItems.Items.Count    
			}while($fiItems.MoreAvailable -eq $true) 

		}
		write-output $Script:rptCollection
	}
}

function DetectCCNumber
{
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$Number,
		[Parameter(Position=1, Mandatory=$true)] [string]$Property,
	    [Parameter(Position=2, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=3, Mandatory=$true)] [string]$DisplayName
    )  
	Begin
	{
		$Number = ExtractNumbers($Number)
		$Number = $Number.Replace("-","").Trim()
		if($Number -ne 0){
			write-host $Number
			if($Number.Length -gt 4){
				$detector = new-object CreditCardValidator.CreditCardDetector($Number)
				if($detector.IsValid()){
					$rptObj = "" | Select Mailbox,Contact,Property,Number,Brand,BrandName,IssuerCategory
					$rptObj.Mailbox = $MailboxName
					$rptObj.Contact = $DisplayName
					$rptObj.Property = $Property
					$rptObj.Number = $Number
					$rptObj.Brand = $detector.Brand
					$rptObj.BrandName = $detector.BrandName
					$rptObj.IssuerCategory = $detector.IssuerCategory
					$Script:rptCollection += $rptObj
				}
				else{
					$SSN_Regex = "^(?!000)([0-6]\d{2}|7([0-6]\d|7[012]))([ -]?)(?!00)\d\d\3(?!0000)\d{4}$"
					$Matches = $Number | Select-String -Pattern $SSN_Regex
					if($Matches.Matches.Count -gt 0){
						$rptObj = "" | Select Mailbox,Contact,Property,Number,Brand,BrandName,IssuerCategory
						$rptObj.Mailbox = $MailboxName
						$rptObj.Contact = $DisplayName
						$rptObj.Property = $Property
						$rptObj.Number = $Number
						$rptObj.Brand = "Social Security Number"
						$Script:rptCollection += $rptObj
					}
				}				
			}
		}
		return $detector
	}
}

Function ExtractNumbers ([string]$InStr){
   $Out = $InStr -replace("[^\d]")
   try{return [int]$Out}
       catch{}
   try{return [uint64]$Out}
       catch{return 0}}