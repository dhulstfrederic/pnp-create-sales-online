##################################################################
# 08.01.20 - P6805
# Create sites from templates. Call function "SetConfig" to start
# Last update : 08.01.2020
# Creation command:  .\CreateSalesTemplate.ps1 -Title salestest2 -Url salestest2 -EmailRequester P06805 -MainGroupToCreate salestest2 
#Login = "shptdevsyncsap@thenrbgroup.onmicrosoft.com";
#Password = "dphljpmjbjbgngxy";
##################################################################

#================================================================== Get Params ================================================================
param (

    [string]$Title = $(throw "-Title is required."),
    [string]$Url = $(throw "-Url is required."),
    [string]$Vertical = "",
    [string]$Horizontal = "",
    [string]$EmailRequester = $(throw "-EmailRequester is required."),
    [string]$MainGroupToCreate = $(throw "-GroupName is required.")
)

# #======================== ============= =================== ======  . Global Variables  .  =========== ================= ====================
#region Global Variables
$global:SPUrl = "Demo";
Write-Host -ForegroundColor Yellow "Declaring global variables"
$global:SPUrl_UAT_PROD = "https://thenrbgroup.sharepoint.com/sites/Sales";

$global:LogFile = "C:\TEMP\template_online_sales.log";
$global:scriptPath = "D:\Scripts\PortalScripts\Template-Online\Templates Creation\";
$global:Template = "Default";

$global:SPUrl = "empty";
$global:SiteURL =   "empty"                                      
$global:SiteTitle = "empty"                                    
#endregion


Write-Host -ForegroundColor Yellow "Global variables have been declared."

function init([string]$_siteUrl, [string]$_siteTitle) {

    #replace unsual chars 
   $pattern = '[^a-zA-Z0-9]'
   $_siteUrl = $_siteUrl -replace $pattern, '' 

    Set-Variable -Name SiteURL    -Value $_siteUrl      -Scope Global   #  Site url 
    Set-Variable -Name SiteTitle -Value $_siteTitle    -Scope Global   #  Site Title
    Set-Variable -Name SPUrl -Value $SPUrl_UAT_PROD -Scope Global;          
                
    main;

} 

function main() {

    #####################[DON'T TOUCH]######################
    Start-Transcript -path $LogFile -append
    Write-Host 

    ########################################################
    Write-Host -ForegroundColor Yellow "Config done."
    Write-Host -ForegroundColor Yellow "Starting main function"
    ### Login
    Write-Host "Connecting to $SPUrl"; 
    
    #Set Credentials
    try{
        $Credentials  = Get-Credential
        Connect-PnPOnline -Url $SPUrl -Credentials $Credentials
    }
    catch{
        Write-Host "something went wrong during connexion";
        break;
    }
    
  
    # Create Site  site
    Write-Host "Getting site..."
    $SiteError = $null
    $Site = $null
    $Site = Get-PnPWeb -Identity $SPUrl/$SiteURL -ErrorAction SilentlyContinue -ErrorVariable SiteError
    if ($SiteError -ne $null) {
        Write-Host "  Site doesn't exists" -ForegroundColor Yellow

        # Creating Site 
        Write-Host "Creating site..."
        $Site = New-PnPWeb -Title $SiteTitle -Url $SiteURL -Locale 1033 -Template "STS#3" -BreakInheritance  -InheritNavigation -Description "Sales"

    }
    else {
        Write-Host "  Site already exists" -ForegroundColor Red
            Exit;
    }

    ## connect to the site 
    Disconnect-PnPOnline
    Connect-PnPOnline -Url $SPUrl/$SiteURL -Credentials $Credentials  
    # enable publishing feature
    Enable-PnPFeature -Identity 94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb
    

    # activate custom master page
    #Set-PnPMasterPage -MasterPageServerRelativeUrl /projects/portfolio/_catalogs/masterpage/portal-afelio.master 

    configSite 

}

function configSite() {

    createLists
    configGroups ;    
    
    # remove webparts from homepage
    $subSiteRelativePath = "/sites/Sales/" + $SiteURL;
    $homePage = $subSiteRelativePath + "/sitepages/home.aspx";
    Remove-PnPWebPart -ServerRelativePageUrl $homePage -Title "Documents"
    Remove-PnPWebPart -ServerRelativePageUrl $homePage -Title "Get started with your site"
    Remove-PnPWebPart -ServerRelativePageUrl $homePage -Title "Site Feed"
    Remove-PnPNavigationNode -Title "Site Contents" -Location QuickLaunch -force
    Remove-PnPNavigationNode -Title "Notebook" -Location QuickLaunch -force
    Remove-PnPNavigationNode -Title "Pages" -Location QuickLaunch -force
    Remove-PnPNavigationNode -Title "Documents" -Location QuickLaunch -force

    #Create Folders in Libs
    $lib = Get-PnPList -Identity BIDMANAGEMENT
    $Thissite = Get-PnPSite
    $fullUrl = ($Thissite.Url + '/' + $SiteURL)
 
    if ($lib.Id) 
    {
        $path = ($fullUrl + "/BIDMANAGEMENT/1. Tender Plan");
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "1. Tender Plan" -Folder "BIDMANAGEMENT" -ErrorAction Continue
        }
        $path = ($fullUrl + "/BIDMANAGEMENT/2. Templates")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "2. Templates" -Folder "BIDMANAGEMENT"-ErrorAction Continue
        }
        $path = ($fullUrl + "/BIDMANAGEMENT/3. Inspiration Documents")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "3. Inspiration Documents" -Folder "BIDMANAGEMENT"-ErrorAction Continue
        }

        $path = ($fullUrl + "/RFPQA/1. RFP");
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "1. RFP" -Folder "RFPQA" -ErrorAction Continue
        }
        $path = ($fullUrl + "/RFPQA/2. Q&A")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "2. Q&A" -Folder "RFPQA"-ErrorAction Continue
        }


        $path = ($fullUrl + "/ADMINISTRATIVESELECTION/1. Identification of the Tenderer");
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "1. Identification of the Tenderer" -Folder "ADMINISTRATIVESELECTION" -ErrorAction Continue
        }
        $path = ($fullUrl + "/ADMINISTRATIVESELECTION/2. Declaration of Honour")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "2. Declaration of Honour" -Folder "ADMINISTRATIVESELECTION"-ErrorAction Continue
        }
        $path = ($fullUrl + "/ADMINISTRATIVESELECTION/3. Legal and Regulatory")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "3. Legal and Regulatory" -Folder "ADMINISTRATIVESELECTION"-ErrorAction Continue
        }
        $path = ($fullUrl + "/ADMINISTRATIVESELECTION/4. Economic and Financial Capacity")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "4. Economic and Financial Capacity" -Folder "ADMINISTRATIVESELECTION"-ErrorAction Continue
        }
        $path = ($fullUrl + "/ADMINISTRATIVESELECTION/5. Other documents")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "5. Other documents" -Folder "ADMINISTRATIVESELECTION"-ErrorAction Continue
        }

        $path = ($fullUrl + "/TECHNICALANDPROFESSIONALCAPACITY/1. ManPower");
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "1. ManPower" -Folder "TECHNICALANDPROFESSIONALCAPACITY" -ErrorAction Continue
        }
        $path = ($fullUrl + "/TECHNICALANDPROFESSIONALCAPACITY/2. References")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "2. References" -Folder "TECHNICALANDPROFESSIONALCAPACITY"-ErrorAction Continue
        }
        $path = ($fullUrl + "/TECHNICALANDPROFESSIONALCAPACITY/3. Certificates")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "3. Certificates" -Folder "TECHNICALANDPROFESSIONALCAPACITY"-ErrorAction Continue
        }
        $path = ($fullUrl + "/TECHNICALANDPROFESSIONALCAPACITY/4. Oracle Partnership")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "4. Oracle Partnership" -Folder "TECHNICALANDPROFESSIONALCAPACITY"-ErrorAction Continue
        }


        $path = ($fullUrl + "/CVS/01. Project Manager");
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "01. Project Manager" -Folder "CVS" -ErrorAction Continue
        }
        $path = ($fullUrl + "/CVS/02. Scrum Master")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "02. Scrum Master" -Folder "CVS"-ErrorAction Continue
        }
        $path = ($fullUrl + "/CVS/11. System Integrations Architect")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "11. System Integrations Architect" -Folder "CVS"-ErrorAction Continue
        }
        $path = ($fullUrl + "/CVS/15. Change Manager")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "15. Change Manager" -Folder "CVS"-ErrorAction Continue
        }
        $path = ($fullUrl + "/CVS/17. UX Designer")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "17. UX Designer" -Folder "CVS"-ErrorAction Continue
        }
        $path = ($fullUrl + "/CVS/18. UI Designer")
        $folder = Get-PnPFolder -Url $path -ErrorAction SilentlyContinue
        if (!$folder) {
            Add-PnPFolder -Name "18. UX Designer" -Folder "CVS"-ErrorAction Continue
        }

    }


    # set "All documents" views
    Set-PnPView -List "BIDMANAGEMENT" -Identity  "All Documents" -Fields @("Type", "File Size", "Name", "Version", "Remarks", "Modified", "Modified By", "_CheckinComment", "Checked Out To");     
    Set-PnPView -List "RFPQA" -Identity  "All Documents" -Fields @("Type", "File Size", "Name", "Version", "Modified", "Modified By", "Checked Out To");    
    Set-PnPView -List "ADMINISTRATIVESELECTION" -Identity  "All Documents" -Fields @("Type", "File Size", "Question", "Name", "Workflow", "Version", "Doc Owner", "Doc Reviewer", "Remarks", "Modified", "Modified By", "Checked Out To");      
    Set-PnPView -List "TECHNICALANDPROFESSIONALCAPACITY" -Identity  "All Documents" -Fields @("Type", "File Size", "Question", "Name", "Workflow", "Version", "Doc Owner", "Doc Reviewer", "Remarks", "Modified", "Modified By", "_CheckinComment", "Checked Out To");      
    Set-PnPView -List "CVS" -Identity  "All Documents" -Fields @("Type", "File Size", "Name", "Doc Status", "Profile", "Level", "Workflow", "Version", "Doc Owner", "Doc Reviewer", "Remarks", "Modified", "Modified By", "_CheckinComment", "Checked Out To");      
    Set-PnPView -List "TECHNICALTENDER" -Identity  "All Documents" -Fields @("Type", "File Size", "Question", "Name", "Workflow", "Version", "Doc Owner", "Doc Reviewer", "Remarks", "Modified", "Modified By", "_CheckinComment", "Checked Out To");      
    Set-PnPView -List "CONTACTS" -Identity  "All contacts" -Fields @("Last Name", "First Name", "Company", "Business Phone", "Mobile Number", "Role", "Email Address", "Address", "City", "ZIP/Postal Code", "Job Title");    

 


    $web = Get-PnPWeb;
    $fullurl = $web.Url;  


        
    #Add menus  nav       
    $parent = Add-PnPNavigationNode -Title "BID MANAGEMENT" -Url $fullurl"/BIDMANAGEMENT" -Location "QuickLaunch" -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "1. Tender Plan" -Url $fullurl"/BIDMANAGEMENT/1. Tender Plan" -Location "QuickLaunch" -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "2. Templates" -Url $fullurl"/BIDMANAGEMENT/2. Templates" -Location "QuickLaunch" -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "3. Inspiration Documents" -Url $fullurl"/BIDMANAGEMENT/3. Inspiration Documents" -Location "QuickLaunch" -ErrorAction SilentlyContinue
    Add-PnPNavigationNode -Title "Contacts" -Url $fullurl"/Lists/Contacts" -Location "QuickLaunch"  -ErrorAction SilentlyContinue
    $parent = Add-PnPNavigationNode -Title "RFP and Q&A" -Url $fullurl"/RFPQA" -Location "QuickLaunch" -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "1. RFP" -Url $fullurl"/RFPQA/1. RFP" -Location "QuickLaunch" -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "2. Q&A" -Url $fullurl"/RFPQA/2. Q&A" -Location "QuickLaunch" -ErrorAction SilentlyContinue
    $parent = Add-PnPNavigationNode -Title "ADMINISTRATIVE SELECTION" -Url $fullurl"/ADMINISTRATIVESELECTION" -Location "QuickLaunch"  -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "1. Identification of the Tenderer" -Url $fullurl"/ADMINISTRATIVESELECTION/1. Identification of the Tenderer" -Location "QuickLaunch"  -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "2. Declaration of Honour" -Url $fullurl"/ADMINISTRATIVESELECTION/2. Declaration of Honour" -Location "QuickLaunch"  -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "3. Legal and Regulatory" -Url $fullurl"/ADMINISTRATIVESELECTION/3. Legal and Regulatory" -Location "QuickLaunch"  -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "4. Economic and Financial Capacity" -Url $fullurl"/ADMINISTRATIVESELECTION/4. Economic and Financial Capacity" -Location "QuickLaunch"  -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "5. Other documents" -Url $fullurl"/ADMINISTRATIVESELECTION/5. Other documents" -Location "QuickLaunch"  -ErrorAction SilentlyContinue
    $parent = Add-PnPNavigationNode -Title "TECHNICAL SELECTION" -Url $fullurl -Location "QuickLaunch"  -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "Technical And Professional Capacity" -Url $fullurl"/TECHNICALANDPROFESSIONALCAPACITY" -Location "QuickLaunch"  -ErrorAction SilentlyContinue
        Add-PnPNavigationNode -Parent $parent.Id -Title "CVs" -Url $fullurl"/CVS" -Location "QuickLaunch"  -ErrorAction SilentlyContinue
    Add-PnPNavigationNode -Title "TECHNICAL TENDER" -Url $fullurl"/TECHNICALTENDER" -Location "QuickLaunch"  -ErrorAction SilentlyContinue

    #Add-PnPNavigationNode -Title "Consortium Discussion Forum" -Url $fullurl"/Lists/CONSORTIUMDISCUSSIONFORUM" -Location "QuickLaunch"  -ErrorAction SilentlyContinue
    #Add-PnPNavigationNode -Title "Bid Calendar" -Url $fullurl"/Lists/BIDCALENDAR" -Location "QuickLaunch"  -ErrorAction SilentlyContinue

    # remove menus
    Remove-PnPNavigationNode -Title "Notebook" -Location QuickLaunch -force
    Remove-PnPNavigationNode -Title "Recent" -Location QuickLaunch -force

    setHomePage;


     #$xmlLinksWebpart = $scriptPath + "\Sales_Webparts\Sales.webpart";
     #$content = Get-Content $xmlLinksWebpart -Raw
     #Add-PnPWebPartToWikiPage -ServerRelativePageUrl $homePage -xml $content -Row 1 -Column 1

     #$xmlLinksWebpart = $scriptPath + "\Sales_Webparts\Consortium1.webpart";
     #$content = Get-Content $xmlLinksWebpart -Raw
     #Add-PnPWebPartToWikiPage -ServerRelativePageUrl $homePage -xml $content -Row 1 -Column 2

    # $xmlLinksWebpart = $scriptPath + "\Sales_Webparts\Consortium2.webpart";
    # $content = Get-Content $xmlLinksWebpart -Raw
     #$currentlist = Get-PnPList -Identity "CONSORTIUMDISCUSSIONFORUM"
     #$view = Get-PnPView -List $currentlist -Identity "Recent" -ErrorAction SilentlyContinue;
     #$content = $content.replace("My-ListId", $currentlist.Id);
     #$content = $content.replace("My-ListName", $currentlist.DefaultViewUrl);
     #$content = $content.replace("My-ViewId", $view.Id);
     #$content = $content.replace("My-Site", $SiteTitle);
     #Add-PnPWebPartToWikiPage -ServerRelativePageUrl $homePage -xml $content -Row 1 -Column 2
  
    
}

function createLists() {

    # BID MANAGEMENT
    $lib1 = @{ };
    $lib1.ListName = "BIDMANAGEMENT";
    $lib1.ListTitle = "BID MANAGEMENT";
    $lib1.Template = "DocumentLibrary"
    $lib1.CT = "BIDMANAGEMENT_CT";

    # RFP and QA
    $lib2 = @{ };
    $lib2.ListName = "RFPQA";
    $lib2.ListTitle = "RFPQA";
    $lib2.Template = "DocumentLibrary"    
    $lib2.CT = "RFPQA_CT";


    # ADMINISTRATIVE SELECTION
    $lib3 = @{ };
    $lib3.ListName = "ADMINISTRATIVESELECTION";
    $lib3.ListTitle = "ADMINISTRATIVE SELECTION";
    $lib3.Template = "DocumentLibrary"
    $lib3.CT = "ADMINISTRATIVESELECTION_CT";

    # TECHNICALANDPROFESSIONALCAPACITY
    $lib4 = @{ };
    $lib4.ListName = "TECHNICALANDPROFESSIONALCAPACITY";
    $lib4.ListTitle = "TECHNICAL AND PROFESSIONAL CAPACITY";
    $lib4.Template = "DocumentLibrary"
    $lib4.CT = "TECHNICALANDPROFESSIONALCAPACITY_CT";

    #CVS
    $lib5 = @{ };
    $lib5.ListName = "CVS";
    $lib5.ListTitle = "CVS";
    $lib5.Template = "DocumentLibrary"
    $lib5.CT = "CVS_CT";
  

    #TECHNICALTENDER
    $lib6 = @{ };
    $lib6.ListName = "TECHNICALTENDER";
    $lib6.ListTitle = "TECHNICAL TENDER";
    $lib6.Template = "DocumentLibrary"
    $lib6.CT = "TECHNICALTENDER_CT";


    #CONTACTS
    $list1 = @{ };
    $list1.ListName = "CONTACTS";
    $list1.ListTitle = "CONTACTS";
    $list1.Template = "Contacts";
    $list1.CT = "CONTACTS_CT";
  
    #BIDCALENDAR
    #$list2 = @{ };
    #$list2.ListName = "BIDCALENDAR";
    #$list2.ListTitle = "BID CALENDAR";
    #$list2.Template = "Events";
    #$list2.CT = "BIDCALENDAR_CT";

    #CONSORTIUMDISCUSSIONFORUM
    #$list3 = @{ };
    #$list3.ListName = "CONSORTIUMDISCUSSIONFORUM";
    #$list3.ListTitle = "CONSORTIUM DISCUSSION FORUM";
    #$list3.Template = "108";
    #$list3.CT = "CONSORTIUMDISCUSSIONFORUM_CT";

    
    $lists = @();
    $lists = ($lib1, $lib2, $lib3, $lib4, $lib5, $lib6, $list1);

    #Create Lists and fields
    associateContentTypes($lists);
    setPermissions($lists) 
}

function configGroups() {

    #Create New Groups NRB

    #Create MainGroup
    $GroupError = $null;
    $newGroupName = ("NRB_" +$MainGroupToCreate+"_OWNER")  
    New-PnPGroup -Title $newGroupName  -Owner $EmailRequester -ErrorVariable GroupError -ErrorAction SilentlyContinue
     $MainGroup = Get-PnPGroup -Identity $newGroupName
    if($MainGroup){
        #Add requester to group
        Add-PnPUserToGroup -LoginName $EmailRequester -Identity $MainGroup
        #Set permissions to group
        Set-PnPGroupPermissions -Web $Site.Id -Identity $MainGroup   -AddRole "Contribute";
    }

    # Create Members group
    $GroupError = $null;
    $newGroupName = ("NRB_" +$MainGroupToCreate+"_MEMBERS")  
    New-PnPGroup -Title $newGroupName  -Owner $EmailRequester -ErrorVariable GroupError -ErrorAction SilentlyContinue
     $MainGroup = Get-PnPGroup -Identity $newGroupName
    if($MainGroup){
        #Add requester to group
        Add-PnPUserToGroup -LoginName $EmailRequester -Identity $MainGroup
        #Set permissions to group
        Set-PnPGroupPermissions -Web $Site.Id -Identity $MainGroup   -AddRole "Contribute";
    }
    # Create award group
    $GroupError = $null;
    $newGroupName = ("NRB_" +$MainGroupToCreate+"_AWARD")  
    New-PnPGroup -Title $newGroupName  -Owner $EmailRequester -ErrorVariable GroupError -ErrorAction SilentlyContinue
     $MainGroup = Get-PnPGroup -Identity $newGroupName
    if($MainGroup){
        #Add requester to group
        Add-PnPUserToGroup -LoginName $EmailRequester -Identity $MainGroup
        #Set permissions to group
        Set-PnPGroupPermissions -Web $Site.Id -Identity $MainGroup   -AddRole "Contribute";
    }

    # Create Visitors group
    $GroupError = $null;
    $newGroupName = ("NRB_" +$MainGroupToCreate+"_VISITORS")  
    New-PnPGroup -Title $newGroupName  -Owner $EmailRequester -ErrorVariable GroupError -ErrorAction SilentlyContinue
     $MainGroup = Get-PnPGroup -Identity $newGroupName
    if($MainGroup){
        Set-PnPGroupPermissions -Web $Site.Id -Identity $MainGroup   -AddRole "Read";
    }

    #Create New Groups EXT

    #Create MainGroup
    $GroupError = $null;
    $newGroupName = ("EXT_" +$MainGroupToCreate+"_OWNER")  
    New-PnPGroup -Title $newGroupName  -Owner $EmailRequester -ErrorVariable GroupError -ErrorAction SilentlyContinue
     $MainGroup = Get-PnPGroup -Identity $newGroupName
    if($MainGroup){
        #Add requester to group
        Add-PnPUserToGroup -LoginName $EmailRequester -Identity $MainGroup
        #Set permissions to group
        Set-PnPGroupPermissions -Web $Site.Id -Identity $MainGroup   -AddRole "Contribute";
    }

    # Create Members group
    $GroupError = $null;
    $newGroupName = ("EXT_" +$MainGroupToCreate+"_MEMBERS")  
    New-PnPGroup -Title $newGroupName  -Owner $EmailRequester -ErrorVariable GroupError -ErrorAction SilentlyContinue
     $MainGroup = Get-PnPGroup -Identity $newGroupName
    if($MainGroup){
        #Add requester to group
        Add-PnPUserToGroup -LoginName $EmailRequester -Identity $MainGroup
        #Set permissions to group
        Set-PnPGroupPermissions -Web $Site.Id -Identity $MainGroup   -AddRole "Contribute";
    }

    # Create Visitors group
    $GroupError = $null;
    $newGroupName = ("EXT_" +$MainGroupToCreate+"_VISITORS")  
    New-PnPGroup -Title $newGroupName  -Owner $EmailRequester -ErrorVariable GroupError -ErrorAction SilentlyContinue
     $MainGroup = Get-PnPGroup -Identity $newGroupName
    if($MainGroup){
        Set-PnPGroupPermissions -Web $Site.Id -Identity $MainGroup   -AddRole "Read";
    }
          
}

function associateContentTypes($listsColl) {

    foreach ($list in $listsColl) {
        Write-Host -ForegroundColor Cyan "List " $list.ListName;       
        #Get content type from parent
        $ct = Get-PnPContentType -Identity $list.CT -InSiteHierarchy  -ErrorAction SilentlyContinue;
        #Create list
        New-PnPList -Title $list.ListName -Template $list.Template ;  
        $currentlist = get-PnPList -Identity $list.ListName
        
        # rename du title de la lib
        Set-PnPList -Identity $currentlist -Title $list.ListTitle;
        
        #Add ct to list
        Add-PnPContentTypeToList -List $currentlist -ContentType $ct
        Remove-PnPContentTypeFromList -List $currentlist -ContentType "Item" 
        Remove-PnPContentTypeFromList -List $currentlist -ContentType "Document" 

        #Activate Versioning
        Set-PnPList -Identity $currentlist -EnableContentTypes $true -EnableVersioning $true -EnableMinorVersions $true -ErrorAction SilentlyContinue;

        if ($list.ListName -eq "Deliverables") {
            $ct = Get-PnPContentType -Identity Deliverables_CT -InSiteHierarchy  -ErrorAction SilentlyContinue;
            Add-PnPContentTypeToList -List $currentlist -ContentType $ct
        }

    }
}


function setPermissions($alllibs) {
                 
    foreach ($list in $alllibs) {

        if ($list.Permissions) {

            foreach ($permElem in $list.Permissions) {
                $group = Get-PnPGroup -Identity ($permElem.Name) -ErrorAction SilentlyContinue
                if ($group) {
                       
                    Write-Host -ForegroundColor Green "found group " $group.Title 
                    Write-Host -ForegroundColor Yellow "Breaking inheridance for list "$list.ListName
                    Set-PnPList -BreakRoleInheritance -Identity $list.ListName
                    Set-PnPListPermission -Identity $list.ListName -Group $group -AddRole $permElem.Access
                }
                else {
                    Write-Host -ForegroundColor Yellow "did not find group " ($permElem.Name)
                }
            }
        }
    }
}


function setHomePage() {
    $htmlToInject = '
    <style>
    .classH2 {
        color: #0072C6!important;
        line-height:1!important;
    }
    </style
<div id="container">
    <div style="float:left">
        <h2 class="classH2">Consortium leader :</h2>
        <h2 class="classH2"> Partners : </h2>
        <h2 class="classH2"> Subcontractors :</h2>
        <h2 class="classH2"> Bid Manager&#160;:</h2>
           <strong>Identity card</strong> </p>
        <div> 
           <table width="527" class=" " cellspacing="0" style="width: 527px; height: 189px;"> 
              <tbody> 
                 <tr class=""> 
                    <td class="" style="width: 30%;">​Customer<br/></td> 
                    <td class="" style="width: 50%;"><br/></td> 
                 </tr> 
                 <tr style="background-color: #C0E4FF;"> 
                    <td class="" style="width: 30%;">​Customer profile<br/></td> 
                    <td class="">​<br/></td> 
                 </tr> 
                 <tr class=""> 
                    <td class="" style="width: 30%;">​Title of RFP<br/></td> 
                    <td class=""><br/></td> 
                 </tr> 
                 <tr style="background-color: #C0E4FF;"> 
                    <td class="" style="width: 30%;">Ref customer<br/></td> 
                    <td class=""><br/></td> 
                 </tr> 
                 <tr class=""> 
                    <td class="" style="width: 30%;">​CRM number<br/></td> 
                    <td class="">​<br/></td> 
                 </tr> 
                 <tr style="background-color: #C0E4FF;"> 
                    <td class="" style="width: 30%;">​URL TED<br/></td> 
                    <td class=""><br/></td> 
                 </tr> 
              </tbody> 
           </table> 
           <br/> 
           <div> 
              <strong>Calendar</strong> 
              <div> 
                 <table width="528" class="" cellspacing="0" style="width: 528px; height: 157px;"> 
                    <tbody> 
                       <tr class=""> 
                          <td class="" style="width: 158px;">Information session​<br/></td> 
                          <td class="" style="width: 272px;">​<br/></td> 
                       </tr> 
                       <tr style="background-color: #C0E4FF;"> 
                          <td class="" style="width: 158px;">​Deadline questions<br/></td> 
                          <td class="">​</td> 
                       </tr> 
                       <tr class=""> 
                          <td class="" rowspan="1" style="width: 158px;">​Deadline offer<br/></td> 
                          <td class="" rowspan="1"><br/></td> 
                       </tr> 
                       <tr style="background-color: #C0E4FF;"> 
                          <td class="" style="width: 158px;">Internal deadline offer​<br/></td> 
                          <td class=""><br/></td> 
                       </tr> 
                       <tr class=""> 
                          <td class="" style="width: 158px;">Opening session<br/></td> 
                          <td class="">​<br/></td> 
                       </tr> 
                    </tbody> 
                 </table> 
              </div> 
              <strong>
                 <br/></strong></div> 
           <div> 
              <strong>Main facts about the RFP</strong> 
              <div> 
                 <table width="527" class="" cellspacing="0" style="width: 527px; height: 408px;"> 
                    <tbody> 
                       <tr class=""> 
                          <td class="" style="width: 158px;">Type of procedure<br/></td> 
                          <td class="" style="width: 272px;">​<br/></td> 
                       </tr> 
                       <tr style="background-color: #C0E4FF;"> 
                          <td class="" style="width: 158px;">Duration<br/></td> 
                          <td class="">​</td> 
                       </tr> 
                       <tr class=""> 
                          <td class="" style="width: 158px;">Estimated budget<br/></td> 
                          <td class="">€​</td> 
                       </tr> 
                       <tr style="background-color: #C0E4FF;"> 
                          <td class="" style="width: 158px;">Lots + short descr<br/></td> 
                          <td class=""><br/></td> 
                       </tr> 
                       <tr class=""> 
                          <td class="" rowspan="1" style="width: 158px;">​Intra/extra-muros<br/></td> 
                          <td class="" rowspan="1">​<br/></td> 
                       </tr> 
                       <tr class=""> 
                          <td class="" rowspan="1" style="width: 158px;">​Contract type<br/></td> 
                          <td class="" rowspan="1">​<br/></td> 
                       </tr> 
                       <tr style="background-color: #C0E4FF;"> 
                          <td class="" rowspan="1" style="width: 158px;">​Number of contractors<br/></td> 
                          <td class="" rowspan="1">​<br/></td> 
                       </tr> 
                       <tr class=""> 
                          <td class="" rowspan="1" style="width: 158px;">​Cascade or reopening<br/></td> 
                          <td class="" rowspan="1">​</td> 
                       </tr> 
                       <tr style="background-color: #C0E4FF;"> 
                          <td class="" rowspan="1" style="width: 158px;">Technical/price ratio<br/></td> 
                          <td class="" rowspan="1">​</td> 
                       </tr> 
                       <tr class=""> 
                          <td class="" rowspan="1" style="width: 158px;">Thresholds for awards<br/></td> 
                          <td class="" rowspan="1"><br/></td> 
                       </tr> 
                       <tr style="background-color: #C0E4FF;"> 
                          <td class="" rowspan="1" style="width: 158px;"> 
                             <span class="">Partners admitted?<br/></span></td> 
                          <td class="" rowspan="1">​<br/></td> 
                       </tr> 
                       <tr class=""> 
                          <td class="" rowspan="1" style="width: 158px;"> 
                             <span class="">​Subco s admitted?</span><br/></td> 
                          <td class="" rowspan="1">​</td> 
                       </tr> 
                    </tbody> 
                 </table> 
              </div> 
              <br/> 
           </div> 
        </div>
    </div>
    <div style="float:right">
        <h2 class="class="classH2"">Consortium Name :</h2>
        <h2 class="classH2">Consortium Logo : </h2>
    </div>
</div>
    ';
    #Add-PnPClientSideText -Page "Home" -Text $htmlToInject
    
    $page = Get-PnPClientSidePage -Identity "home.aspx"
    # ajouter une section avec 2 colonnes 2/3 et 1/3 voir https://docs.microsoft.com/en-us/dotnet/api/officedevpnp.core.pages.canvassectiontemplate?view=sharepointpnpcoreonline-2.18.1709.0
    $section = $page.AddSection(4,5); 
    $page.Save()

    Add-PnPClientSideWebPart -Page home.aspx -Component “NRB - Sales HomePage - Identity Card” -Section 1 -Column 1
    Add-PnPClientSideWebPart -Page home.aspx -Component “NRB - Sales HomePage - Logo” -Section 1 -Column 2

    #Set-PnPClientSidePage -Identity "Home" -Publish

    Set-PnPClientSidePage -Identity "Home” -CommentsEnabled
}





init -_siteUrl $Url -_siteTitle $Title;



$Connexion = Get-PnPConnection -ErrorAction SilentlyContinue
$linkNewSite = $Connexion.Url
if($linkNewSite){
Write-Host -ForegroundColor Yellow "New site created at "$linkNewSite
Write-Host -ForegroundColor Yellow "Disconnecting"
Write-Host -ForegroundColor red "Please check web permissions"
Write-Host -ForegroundColor red "Don't forget to add recent documents and select language=english"

#####set an environment variable to have it in script
[environment]::SetEnvironmentVariable('SiteURL',$linkNewSite,'MACHINE')
}
### Disconnect form site
$temp = Disconnect-PnPOnline
Stop-Transcript 


