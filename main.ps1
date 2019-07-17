##############################################
#
#   ������ �������� ������ ��������������� ��������
#
#   ����: Main.ps1 
#   �������� : ������� ����������� ������. ������� ����� 
#              ��������������� ������ � ��������������� ��������.
#
#   �����: ������� ��������� ���������, ����� ��������������, ���: 223-1338
#          29AstahovAB@cbr.ru
#
#   ��������� ������, �� ����� ������ �� ���
#   ������� ����������: ������ ��������� ������� ��������� �� ������� Sharepoint 
#                       �� ��������� ������� Sharepoint 2010
#   ����: ������ 2017.
#   ������: 17.01.11
#
##############################################
function Add-newInforModuleWebSite ($webObject)
{

      $webSite        = $webObject.Site 
      $webHierarchy   = $webObject.Hierarchy
      $newWeb         = $webObject.InfoModule
      $webTitle       = "�������������� ������"
      $WebDescription = "�������������� ������"

      $siteTemplate = "STS#1" # blank Site
      
      $webUrl=$webSite+$webHierarchy+$newWeb
      
      
      
      $webUrl=$webSite+$webHierarchy+$newWeb


      # $webUrl
      Remove-SPWeb $webUrl -Confirm:$false -ErrorAction SilentlyContinue 


      write-host $("��� ���� : " + $webUrl)  -ForegroundColor Green
      write-host "�������� ����� �������������� ������..." -nonewline
      # �������� ���-�����
      $nooutput = new-spweb $webUrl -Template $siteTemplate  -Name $webTitle  -UseParentTopNav -Description $WebDescription # -Description $WebDescription

      $web = get-spweb -Identity $webUrl
      Change-Theme $webSite $web
          

      $web.Update()
      $web.Dispose()
          
      write-host "[OK]" -ForegroundColor Yellow
      $web = get-spweb -Identity $webUrl
      write-host "�������� ������� � ���������..." -nonewline
      #�������� ������� � ��������� ����������
      $splistAdminsGuid = New-InfoModuleDocLib $web

      write-host "[OK]" -ForegroundColor Yellow

      write-host "�������� ������ ���������..." -nonewline
      $web.QuickLaunchEnabled=$true

      Costumize-InfoQuickLaunchMenu $web $webObject
      write-host "[OK]" -ForegroundColor Yellow

      write-host "���������� ������� ��������..." -nonewline
      Update-InfoWebParts $web $oWeb
      write-host "[OK]" -ForegroundColor Yellow
      $web.Update()
      write-host "������ ���� �������������� ������...  "  -ForegroundColor Green
      write-host $webUrl
      write-host
      write-host ===================================
      write-host      
}

function Update-InfoWebParts ($spweb,$webObject)
{
     $listToAdd = "����������� ����"
     Add-XsltListViewWebPart $spweb  $listToAdd 
	 
}
function New-InfoModuleDocLib( $spweb)
{


    $NameOfListEng     = "Ndocs"
    $NameOfListRus     = "����������� ����"
    $ListDescriptionRus= "����������� ���������"


    $spSite = Get-SPSite $spWeb.Site.Url    

    # ������� ���������� ����������
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "DocumentLibrary")
    
    $SpList = $spWeb.Lists[$NameOfListEng];
          
    
    $Listguid   = $SpList.Id

    $Listguid   = $SpList.GUID
    foreach($culture in $spweb.SupportedUiCultures)
    {


      [System.Threading.Thread]::CurrentThread.CurrentUICulture=$culture
      $SpList = $Web.Lists[$NameOfListEng]
      $SpList.Title = $NameOfListRus
      $SpList.Update()

    }
    $ListFields = @()
    
    $ListFields += '<Field Type="Number" DisplayName="� �/�" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" StaticName="fNPP" Name="fNPP" ColName="fNPP"  />'
    $ListFields += '<Field   Name="FirstName" StaticName="FirstName"  DisplayName="��������" Type="Text" ColName="FirstName"  EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" />'

    # ��������� ����
    foreach ($newColumn in $ListFields)
    {
          $disableoutput = $SpList.Fields.AddFieldAsXml($newColumn,$true,[Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView)
    }
    
     
    $spweb.Dispose()
    
    #$cType = $SpList.ContentsTypes["����� ����������"]
    
    return $Listguid
        
       
}

function Costumize-InfoQuickLaunchMenu ($spWeb, $oWeb){
        $url = $spWeb.URL

	$qlmenu = @()
	$qlmenu += 'parent,�����������'
	$qlmenu +=       'child, �����������,����������� ����,' + $url +'/Ndocs/'
	$qlmenu += 'parent,��������������� �������'
	for ($kk=1; $kk -le $oWeb.SiteCount; $kk++)
	{
	      $qlmenu += 'child,��������������� �������, ��������������� ������� ' + $kk.ToString().Trim() + ',' + $oweb.Site + $oweb.Hierarchy + $oweb.NewWebName + $kk.ToString().Trim()
	}
	
	

        foreach($el in $qlmenu)
        {
             $menulevel = $el.split(",")[0].trim()
             
             if ($menulevel -eq "parent")
             {
                  $header = $el.split(",")[1].trim()
                  Add-QLHeading $spweb $header $url
             }
        }

        foreach($el in $qlmenu)
        {
             $menulevel = $el.split(",")[0].trim()
             
             if ($menulevel -eq "child")
             {
                  $header = $el.split(",")[1].trim()
                  $itemName = $el.split(",")[2].trim()
                  $linkUrl  = $($el.split(",")[3].trim())
                  # write $header, $itemName, $linkUrl
                  
		              Add-QLLink $spweb $header $itemName  $linkUrl $true
             }
        }

        # Delete-QuickLaunchMenu $spweb $qlmenu

}

function Clear-TempFolder
{
   $TempFolder = $env:Temp
   $allfile = $TempFolder+"\*"

   write-host $("������� ��������� ������� :"+$allfile)
   
   remove-Item $allfile -recurse -ErrorAction SilentlyContinue 


}
function New-SummaryLinks( $webObject)
{
            $slinks = @()

            $groupLinkA = New-Object Microsoft.SharePoint.Publishing.SummaryLink "�������������� ������"
            $groupLinkA.IsGroupHeader = $true;
            $slinks  += $groupLinkA
            
            $sumLink = New-Object Microsoft.SharePoint.Publishing.SummaryLink "���� ��������������� ������"
           
            $sumLink.LinkUrl = $webObject.Site + $webObject.Hierarchy + $webObject.infoModule
            $sumLink.Style = "TitleWithBackground";
            $slinks  += $sumLink
            
            
            $groupLinkB = New-Object Microsoft.SharePoint.Publishing.SummaryLink "����� ��������������� ��������"
            $groupLinkB.IsGroupHeader = $true;
            $slinks  += $groupLinkB



            for ($i=1; $i -le $webObject.SiteCount ; $i++)
            {
                 $url =  $webObject.Site + $webObject.Hierarchy + $webObject.NewWebName + $i.ToString().Trim()
                 $webTitle   = $webObject.Title+" � "+ $($i.ToString().PadLeft(2,"0"))
                 $WebDescription = $webObject.Description + " � " + $($i.ToString().PadLeft(2,"0"))

                 $alink = New-Object Microsoft.SharePoint.Publishing.SummaryLink $webTitle
                  
                 $alink.Style = "TitleWithBackground";
                 $alink.Description =  $WebDescription
                 $alink.LinkUrl = $url
                 $slinks  += $alink
           }

           return $slinks


}
function Add-SummaryLinkWebPart($spWeb , $wpName ,$webObject, [string]$PageName="default.aspx")
{
      $page = $spweb.GetFile($PageName)
      $page.CheckOut()
      $pubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($spweb) 

      $webpart = New-Object "Microsoft.SharePoint.Publishing.WebControls.SummaryLinkWebPart"
      $webpart.Title   = $wpName
      
      $webpart.ChromeType = "BorderOnly"

      $slink  = New-SummaryLinks $webObject
      foreach($el in $slink)
      {
             $webpart.SummaryLinkValue.SummaryLinks.Add($el)
      }

      $webpart.SummaryLinkValue = $webpart.SummaryLinkValue
       
      $oWebPartManager = $page.Web.GetLimitedWebPartManager($PageName,[System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
      $oWebPartManager.AddWebPart($webpart,"Rigth",0) # RightZone LeftZone
      
      $page.CheckIn("�������� ��������� �������������.")      

}


function Add-XsltListViewWebPart($spWeb,$spListName,[string]$PageName="default.aspx")
{
     $list = $spWeb.Lists[$spListName]
     $page = $spweb.GetFile($PageName)
     $page.CheckOut()
  
     
     $oWebPartManager = $page.Web.GetLimitedWebPartManager($page.Url,[System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
     
     $oListViewWebPart = New-Object "Microsoft.Sharepoint.webPartPages.XsltListViewWebPart"
     $oListViewWebPart.Title = $spListName
     $oListViewWebPart.ListName = $list.ID.ToString("B").ToUpper()
    
     $oListViewWebPart.ViewGuid =  ($list.DefaultView.ID).ToString("B").ToUpper() 
     $oListViewWebPart.ZoneID = "Left"  
     $oListViewWebPart.ChromeType = "BorderOnly"
     $oWebPartManager.AddWebPart($oListViewWebPart,"Left",0) # RightZone LeftZone
 
     $page.CheckIn("�������� ��������� �������������.")

     
}
function Add-SimpleFormWebPart ($spWeb , $wpName, $HTMLCONTENT,[string]$PageName="default.aspx")
{
      $page = $spweb.GetFile($PageName)
      $page.CheckOut()
      $pubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($spweb) 

      $webpart = New-Object "Microsoft.SharePoint.WebPartPages.SimpleFormWebPart"
      $webpart.Title   = $wpName
      $webpart.Content = $HTMLCONTENT
      $webpart.ChromeType = "BorderOnly"
 
       
      $oWebPartManager = $page.Web.GetLimitedWebPartManager($PageName,[System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
      $oWebPartManager.AddWebPart($webpart,"Left",0) # RightZone LeftZone
      
      $page.CheckIn("�������� ��������� �������������.")      
}
function Add-ContentEditorWebPart($spWeb,$cewpName,$HTMLCONTENT,[string]$PageName="default.aspx")
{
     
       $lvwpGUID = [System.Guid]::NewGuid().ToString()
       $lvwpKey  = "g_"+ $lvwpGUID.Replace("-","_")

       $contentEditor=New-Object "Microsoft.Sharepoint.webPartPages.ContentEditorWebPart"

       $xmlDoc = New-Object System.Xml.XmlDocument

       $xmlElement = $xmlDoc.CreateElement("HtmlContent");
       $xmlElement.InnerText = $HTMLCONTENT;
	 
       $contentEditor.Title = $cewpName
       $contentEditor.ID = $lvwpKey
       $contentEditor.Visible = $true
       $contentEditor.ChromeType = "None"   # Default TitleAndBorder None TitleOnly BorderOnly
       $contentEditor.ZoneID = "Left" 
       $contentEditor.HorizontalAlign = "Center"
       $contentEditor.Content = $xmlElement;

       $pubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($spweb) 

       $oWebPartManager = $pubWeb.Web.GetLimitedWebPartManager($PageName,[System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)   
       $oWebPartManager.AddWebPart($contentEditor,"Left",0)
     
}


function Update-WebParts ($spweb,$webObject)
{
     $listToAdd = "������ � ��� ��"
     Add-XsltListViewWebPart $spweb  $listToAdd 
     
     
                  
      $urllist = $spweb.Url.ToString()
      
     
	
      $content ='<div class="ms-rteThemeBackColor-6-3"><hr class="ms-rteElement-Hr"/></div>'+
                '<div class="ms-rteThemeBackColor-6-3"><span class="ms-rteFontSize-3">�������� ����������, ���������� �&#160;����������� ���������&#160;���������������� ������� ���</span> </div>'+
                '<div class="ms-rteThemeBackColor-6-3">&#160;</div>'+
                '<div class="ms-rteThemeBackColor-6-3"><span>������&#160;����������� ��� �� ��� ��� ����� ���������� ������� ������������� ������� ���������������� �������</span></div>'+
                '<div class="ms-rteThemeBackColor-6-3">&#160;</div>'+
                '<div class="ms-rteThemeBackColor-6-3"><em><font color="#cc3300">(�����&#160;�����������&#160;������ ���������� ���������</font><span style="color: #cc3300"> </span></em>'+
                '<a class="ms-rteFontSize-2" href="' + $urllist + '/Lists/TUUsr/" target="_blank" style="text-decoration: underline">'+
                '<em>���������� �������������</em></a><em><span class="ms-rteFontSize-2"> </span><span class="ms-rteFontSize-2" style="color: #cc3300"><span style="color: #cc3300">�</span></span>'+
                '<span class="ms-rteForeColor-4 ms-rteFontSize-2"> </span></em>'+
                '<a class="ms-rteFontSize-2" href="' + $urllist+ '/Lists/TM/" target="_blank" style="text-decoration: underline"><em>���������� ��-���������������</em></a>)</div>'+
                '<div class="ms-rteThemeBackColor-6-3">&#160;</div>'



     Add-ContentEditorWebPart $spWeb "�������� ��� ��" $content



     $listToAdd = "�������� ���"
     Add-XsltListViewWebPart $spweb  $listToAdd
 
     $content ='<div class="ms-rteThemeBackColor-6-3"><hr class="ms-rteElement-Hr"/></div>'+
                '<div class="ms-rteThemeBackColor-6-3"><span class="ms-rteFontSize-3">��������&#160;����������� �������</span><span class="ms-rteFontSize-3">&#160;���������������� ������� ���</span></div>'+
                '<div class="ms-rteThemeBackColor-6-3"><span class="ms-rteFontSize-1"></span><em></em>&#160;</div>'+
                '<div class="ms-rteThemeBackColor-6-3"><em><span>������&#160;����������� ��� </span>�� ��� ��� �����&#160;����������&#160;������� ��� ������� ���������������� �������</em>&#160;</div>'+
                '<div class="ms-rteThemeBackColor-6-3">&#160;</div>'
     
     Add-ContentEditorWebPart $spWeb "�������� ���" $content
     
     $listToAdd = "�������� ��������� ���������������� �������"
     Add-XsltListViewWebPart $spweb  $listToAdd     



      $content ='<div class="ms-rteThemeBackColor-6-3"><hr class="ms-rteElement-Hr"/></div>'+
                '<div class="ms-rteFontSize-3 ms-rteThemeBackColor-6-3">�������� ��������� ���������������� ������� ���&#160;</div>'+
                '<div class="ms-rteFontSize-1 ms-rteThemeBackColor-6-3"><em></em>&#160;</div>'+
                '<div class="ms-rteFontSize-1 ms-rteThemeBackColor-6-3"><em>������&#160;����������� ��� �� ��� ��� �����&#160;���������� ������� ��������� ������� ���������������� �������</em></div>'+
                '<div class="ms-rteFontSize-1 ms-rteThemeBackColor-6-3"><em></em>&#160;</div>'



     Add-ContentEditorWebPart $spWeb "�������� ���������" $content 
     
     
     $listToAdd = "������ ��������������� ���������������� �������"
     Add-XsltListViewWebPart $spweb  $listToAdd     
     
    $content ='<div class="ms-rteThemeBackColor-6-3"><hr class="ms-rteElement-Hr"/></div>'+
              '<div class="ms-rteFontSize-3 ms-rteThemeBackColor-6-3"><span>��������&#160;��������������� �������������� ������������&#160; </span>���������������� ������� ���&#160;</div>'+
              '<div class="ms-rteThemeBackColor-6-3" style="color: #080808"><span class="ms-rteStyle-Normal"><span><span><em></em></span></span></span>&#160;</div>'+
              '<div class="ms-rteThemeBackColor-6-3" style="color: #080808"><span class="ms-rteStyle-Normal ms-rteThemeBackColor-6-3"><span><span>'+
              '<em>������&#160;����������� ��� �� ��� ��� �����&#160;���������� �������&#160;��������������� <span>������� ���������������� </span>�������</em></span></span></span></div>'+
              '<div class="ms-rteThemeBackColor-6-3" style="color: #080808"><span class="ms-rteStyle-Normal ms-rteThemeBackColor-6-3"><span><span><em></em></span></span></span>&#160;</div>'

     
     Add-ContentEditorWebPart $spWeb "�������� ���������������" $content 
     
     
 
    $content ='<script id="RightWPzone" type="text/javascript">'+
                  'function GetParentTable(elem)'+
                  '{'+
                       'while(elem !=null && elem.tagName != "TABLE") {elem=elem.parentNode;}'+
                       'return elem;'+
                  '}'+
                  'var thisWPzone = GetParentTable(GetParentTable(document.getElementById("RightWPzone")).parentNode);'+
                  'if (thisWPzone.id != "MSOZone")'+
                  '{'+
                      'var thisWPzone=thisWPzone.parentNode;'+
                      'thisWPzone.style.width="100%"'+
                  '}</script>'

     Add-ContentEditorWebPart $spWeb "�������� ������ ����" $content 
		 
}
     

function Costumize-QuickLaunchMenu ($oweb, $CurrentSiteIndex, $spWeb){
        $url = $spWeb.URL
                   
	$qlmenu = @()
	$qlmenu += 'parent,�����'
	$qlmenu +=       'child,�����,���� ��������������� ������,'+$($oWeb.Site + $oWeb.Hierarchy + $oWeb.infoModule)
    	
            for ($i=1; $i -le $oweb.SiteCount ; $i++)
            {
                 if ($i -eq $CurrentSiteIndex)
                 {
                 
                       $url1 =  $webObject.Site + $webObject.Hierarchy + $webObject.NewWebName + $i.ToString().Trim()
                       $webTitle   = $webObject.Title+" � "+ $($i.ToString().PadLeft(2,"0"))

                       $qlmenu +=       'child,�����,' + $webTitle +','+$url1
                 }      
           }

	
	
	$qlmenu += 'parent,�����������'
	$qlmenu +=       'child, �����������,������������ ���������������� �������,' + $url + '/Lists/TUUsr/'
	$qlmenu +=       'child, �����������,�� ��������������,' + $url + '/Lists/TM/'
	$qlmenu += 'parent,���. �������'
	$qlmenu +=     'child,���. �������,�������������� ���������������� �������,' + $url + '/Lists/Admins/'
	$qlmenu +=     'child,���. �������, ������ � ��� ��,' + $url + '/Lists/SVTRestrict/'
	$qlmenu +=     'child,���. �������, �������� ���,' + $url + '/SVTs/'
	$qlmenu +=     'child,���. �������, ��������� ���������������� �������,' + $url + '/Lists/Rooms/'
	$qlmenu +=     'child,���. �������, ���������� �����,' + $url + '/Lists/DailyReport/'
	
	$qlmenu += 'parent,������'
	$qlmenu +=     'child,������,�������� �����,' + $url + '/Reports'
	$qlmenu +=     'child,������,���������� ������ ��� � ������� MSWORD,' + $url + '/DReports'


        foreach($el in $qlmenu)
        {
             $menulevel = $el.split(",")[0].trim()
             
             if ($menulevel -eq "parent")
             {
                  $header = $el.split(",")[1].trim()
                  Add-QLHeading $spweb $header $url
             }
        }

        foreach($el in $qlmenu)
        {
             $menulevel = $el.split(",")[0].trim()
             
             if ($menulevel -eq "child")
             {
                  $header = $el.split(",")[1].trim()
                  $itemName = $el.split(",")[2].trim()
                  $linkUrl  =  $($el.split(",")[3].trim())
                  
		              Add-QLLink $web $header $itemName  $linkUrl $true
             }
        }

        Delete-QuickLaunchMenu $spweb $qlmenu

}

function Add-QLHeading ($spweb, $HeadingName,$Link)
{
    $pubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($spweb)
    $CreateSPNavigationNode = [Microsoft.sharepoint.Publishing.Navigation.SPNavigationSiteMapNode]::CreateSPNavigationNode
    $qlNav = $pubweb.Navigation.CurrentNavigationNodes

    $headingNode = $CreateSPNavigationNode.Invoke($HeadingName ,"",[Microsoft.sharepoint.Publishing.NodeTypes]::None,$qlNav)


    
}

function Add-QLLink ($spweb, $HeadingName, $DisplayName, $URL,$external)
{

   Start-Sleep -Seconds 5
   $qlNav = $spweb.Navigation.QuickLaunch
   
   $qLink = $qlNav | Where {$_.Title.ToUpper().Trim() -eq $HeadingName.ToUpper().Trim()}
   
   $linkNode = new-Object Microsoft.Sharepoint.Navigation.SPNavigationNode($DisplayName, $URL,$external)
   $noresult=$qLink.Children.AddAsLast($linkNode)


}
function Delete-QuickLaunchMenu ($spweb, $qlmenu)
{

  
    for ($i=$web.Navigation.QuickLaunch.count-1;$i -ge 1;$i--)
    {     
          $IsMenuForDeleting = $true
          foreach ($el in $qlmenu)
          {
              $header = $el.split(",")[1].trim()
              
              if ( $header.ToUpper() -eq  $spweb.Navigation.QuickLaunch[$i].Title.Trim().ToUpper() )
              {
                   $IsMenuForDeleting = $false
		               break
              }
         
          
          }
          if ($IsMenuForDeleting){
              $spweb.Navigation.QuickLaunch[$i].Delete()
              }
    }

}

function Delete-QuickLaunchMenuExclude ($spweb, $MenuTitle)
{

  
    for ($i=$web.Navigation.QuickLaunch.count-1;$i -ge 0;$i--)
    {  
          if ($spweb.Navigation.QuickLaunch[$i].Title -ne $MenuTitle) {  
                #write-host $($($spweb.Navigation.QuickLaunch[$i].Title) + " Deleting...")
                $spweb.Navigation.QuickLaunch[$i].Delete()
          }

    }

}
function Convert-ToInternalName ($fieldname)
{

 $arr=@()
 $b=$fieldname.ToCharArray()
 
 $c=""
 Foreach ($element in $b) {$arr +=  [System.String]::Format("{0:X}", [System.Convert]::ToUInt32($element))}
 foreach ($e in $arr) {$c += "_x" + $e.PadLeft( 4, "0") +"_"}
 return $c.toLower()

}

function Define-LookupField ($spListParent, $spParentLookUpFieldNameRus, $spLookUpFieldNameRus, [string]$FieldRequired="TRUE") {
    
    $fild = '<Field Type="Lookup" DisplayName="' + $spLookUpFieldNameRus + '" RelationshipDeleteBehavior="None" Required="' + $FieldRequired +'" List="{' +$spListParent.ID.ToString()+'}" '

    $spFieldInternalName =  Convert-ToInternalName $spParentLookUpFieldNameRus
    $fild += ' ShowField="' + $spFieldInternalName +'" />' 
    
    return $fild
}
function Define-LookupFieldMult ($spListParent, $spParentLookUpFieldNameRus, $spLookUpFieldNameRus, [string]$FieldRequired="TRUE") {
    
    $fild = '<Field Type="Lookup" DisplayName="' + $spLookUpFieldNameRus + '" RelationshipDeleteBehavior="None" Required="' + $FieldRequired +'" List="{' +$spListParent.ID.ToString()+'}" Mult="TRUE" Sortable="TRUE"'

    $spFieldInternalName =  Convert-ToInternalName $spParentLookUpFieldNameRus
    
    $fild += ' ShowField="' + $spFieldInternalName +'" />' 
    
    return $fild
}

function Define-LookupFieldTitle ($spListParent, $spParentLookUpFieldNameRus, $spLookUpFieldNameRus, [string]$FieldRequired="TRUE",[string]$LookupMulti="FALSE") {
    
    $isRequired = [boolean]::Parse($FieldRequired)
    $isMultu    = [boolean]::Parse($LookupMulti)
    
    $fild = '<Field  '
    if ($isMultu)
    {
          $fild += ' Type="LookupMulti" Mult="TRUE" Sortable="TRUE" '
    }
    else
    {
          $fild += ' Type="Lookup" '
    }
    $fild += ' DisplayName="' + $spLookUpFieldNameRus + '" RelationshipDeleteBehavior="None" Required="' + $FieldRequired +'" List="{' +$spListParent.ID.ToString()+'}" '
    $spFieldInternalName =   $spParentLookUpFieldNameRus
    $fild += ' ShowField="' + $spFieldInternalName +'" />' 
    
    return $fild
}
function Define-DependedLookupField ($spListParent, $spParentMainLookUpFieldNameRus,$spParentDependedLookUpFieldNameRus, $spLookUpFieldNameRus ) {
    
    
    $fieldid = $spListParent.Fields[$spParentMainLookUpFieldNameRus].ID.ToString()
    $spFieldInternalName =  Convert-ToInternalName $spParentDependedLookUpFieldNameRus
    $fild = '<Field Type="Lookup" DisplayName="' + $spLookUpFieldNameRus + '" RelationshipDeleteBehavior="None"  ReadOnly="TRUE" Required="FALSE" List="{' +$spListParent.ID.ToString()+'}" '
    $fild += 'FieldRef="'+$fieldid+'" '  # ������� ���� Lookup
    $fild += ' ShowField="' + $spFieldInternalName +'" />'  
    # Write-Host $fild
    return $fild
}

function Rename-ListFieldTitle(
    $spListObject,
    [string]$NameOfField = "������������",
    [string]$Default = ""
)
{
      $titlenew= $NameOfField
      $title = $spListObject.Fields.GetFieldByInternalName("Title")
      $title.DefaultValue = $Default

      $tr = $title.TitleResource
      $tr.SetValueForUiCulture($web.UICulture,$titlenew)
      $tr.Update()

      $title.Title = $titlenew
      $title.PushChangesToLists = $true;
      $title.Update($true)
      $spListObject.Update();
}
function Rename-ListField(
    $spListObject,
    [string]$OldFieldName,
    [string]$NameOfField = "������������",
    [string]$Default = ""
)
{
      $titlenew= $NameOfField
      $title = $spListObject.Fields.GetFieldByInternalName($OldFieldName)
      $title.DefaultValue = $Default

      $tr = $title.TitleResource
      $tr.SetValueForUiCulture($web.UICulture,$titlenew)
      $tr.Update()

      $title.Title = $titlenew
      $title.PushChangesToLists = $true;
      $title.Update($true)
      $spListObject.Update();
}
function Change-DoclibTemplate( $splist,$templateName)
{
$templateFile = Get-ChildItem $templateName
$PrntWeb = $splist.ParentWeb
$templateFolderPath=$PrntWeb.Url.ToString()+"/"+$splist.RootFolder.URL.ToString()+"/Forms"
$templateFolder = $PrntWeb.GetFolder($templateFolderPath)
$uploadPath = $templateFolder.Url + "/"+ $templateFile.Name
$spNewTemplateFile = $templateFolder.Files.Add($uploadPath ,$templateFile.OpenRead(), $true)

$splist.DocumentTemplateUrl = $uploadPath
$splist.Update()
}
function Rename-ListNameRus($spWeb,$NameOfListEng,$NameOfListRus)
{
   
    foreach($culture in $spweb.SupportedUiCultures)
    {


      [System.Threading.Thread]::CurrentThread.CurrentUICulture=$culture
      $SpList = $Web.Lists[$NameOfListEng]
      $SpList.Title = $NameOfListRus
      $SpList.Update()

    }
    
    $spweb.Dispose()


}


function New-UsersList( $spweb)
{
    # ��� ������
    $NameOfListEng     = "TUUsr"
    $NameOfListRus     = "������ ������������� ���������������� �������"
    $ListDescriptionRus= "������ ������������� ���������������� �������. �������� ���������� � ������������� ���.��������"
    
    # ����� ������������ ���� Title
    $TitleNewName      = "��������� �.�.�."
    
 
    # ������� ������
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # ������ ������������ ������
    $SpList = $spWeb.Lists[$NameOfListEng];
    $Listguid   = $SpList.Id
    # ��������������� ������    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus
    
    
    # ������ ������������ ���� Title
    Rename-ListFieldTitle $splist $TitleNewName

   
    
    $SpList.Update();
    $spweb.Update();
    $spweb.Dispose();
    
    
    Add-TuUsersSampleData($SpList)
    
    return $Listguid
    
}

function New-TMList( $spweb, $oWeb)
{
    # ��� ������
    $NameOfListEng     = "TM"
    $NameOfListRus     = "������ �� ���������������"
    $ListDescriptionRus= "������ ������ �� ��������������� ���������������� �������. �������� ���������� � ������ �� ��������������� ���.�������."
    
    # ����� ������������ ���� Title
    $TitleNewName      = "� ��-��������������"
    
 
    # ������� ������
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # ������ ������������ ������
    $SpList = $spWeb.Lists[$NameOfListEng];
    $Listguid   = $SpList.Id
    # ��������������� ������    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus
    
    
    # ������ ������������ ���� Title
    Rename-ListFieldTitle $splist $TitleNewName

   
    
    $SpList.Update();
    $spweb.Update();
    $spweb.Dispose();
    
    Add-TMSampleData($SpList)
    
    return $Listguid
    
}
function New-AdminsList( $spweb)
{
    # ��� ������
    $NameOfListEng     = "Admins"
    $NameOfListRus     = "������ ��������������� ���������������� �������"
    $ListDescriptionRus= "������ ��������������� ���������������� �������. �������� ���������� �� ��������������� ���.��������"
    
    # ����� ������������ ���� Title
    $TitleNewName      = "��������� �.�.�."
    
 
    # ������� ������
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # ������ ������������ ������
    $SpList = $spWeb.Lists[$NameOfListEng];
    $Listguid   = $SpList.Id
    # ��������������� ������    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus
    
    
    # ������ ������������ ���� Title
    Rename-ListFieldTitle $splist $TitleNewName

    # ���������� ����
    $ListFields = @()
    
    $ListFields += '<Field Type="DateTime" DisplayName="���� ������������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly" Name="DateOfSigned" Description="���� ������������ ���������� � ��������"/>'
    $ListFields += '<Field Type="Choice"   DisplayName="�����������" Required="FALSE" EnforceUniqueValues="FALSE" Format="RadioButtons" FillInChoice="FALSE"  StaticName="IsSigned" Name="IsSigned" ><Default>���</Default><CHOICES><CHOICE>��</CHOICE><CHOICE>���</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Choice"   DisplayName="����" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" StaticName="AdminRole" Name="AdminRole" ><Default>���</Default><CHOICES><CHOICE>���</CHOICE><CHOICE>��������� ���</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Text"     DisplayName="���������" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="PersonellPosition" Name="PersonellPosition" />'
    $ListFields += '<Field Type="DateTime" DisplayName="���� ������� � ����������" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly" StaticName="dpONazn" Name="dpONazn" />'
    $ListFields += '<Field Type="Text"     DisplayName="����� ������� � ����������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="npONazn" Name="npONazn" />'
    
       
    # ��������� ����
    foreach ($newColumn in $ListFields)
    {
          $disableoutput = $SpList.Fields.AddFieldAsXml($newColumn,$true,[Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView)
    }
    
    $SpList.Update();
    $spweb.Update();
    $spweb.Dispose();
    
    Add-AdminsSampleData($SpList)
    Create-IQY $spList $oWeb $NameOfListEng
    Change-AdminsDefaultView $spweb $spList
    
    return $Listguid
    
}

function Change-AdminsDefaultView( $spweb, $spList)
{

      
      $viewTitle  = $splist.Title

      $viewFields = New-Object System.Collections.Specialized.StringCollection
      

      
      
      $viewFields.Add("����")
      $viewFields.Add("LinkTitle")
      $viewFields.Add("����� ������� � ����������")
      $viewFields.Add("���� ������� � ����������")
      
      $viewRowLimit    = 300
      $viewDefaultView = $true
      $viewPaged       = $true
      $groupBy         = ''
      
      $viewQuery       = '<OrderBy><FieldRef Name="Title" /></OrderBy>'
    
      $newView = $splist.Views.Add($viewTitle, $viewFields, $($groupBy + $viewQuery), $viewRowLimit, $viewPaged, $viewDefaultView)
      $newView.ApplyStyle($spWeb.ViewStyles.StyleById(17)) # � ����������
      $newView.Update()
    

}

function New-RoomList( $spweb, $oWeb)
{
    # ��� ������
    $NameOfListEng     = "Rooms"
    $NameOfListRus     = "�������� ��������� ���������������� �������"
    $ListDescriptionRus= "�������� ��������� ���������������� �������. �������� ���������� �� ��������� ���.��������"
    
    # ����� ������������ ���� Title
    $TitleNewName      = "����� ���������"
    

    
    # ������� ������
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # ������ ������������ ������
    $SpList = $spWeb.Lists[$NameOfListEng];
    
    $Listguid   = $SpList.Id
    
    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus

    
    # ������ ������������ ���� Title
    Rename-ListFieldTitle $splist $TitleNewName

    # ���������� ����
    $ListFields = @()
    
    
    $ListFields += '<Field Type="DateTime" DisplayName="���� ����������" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly"  StaticName="dattDate" Name="dattDate" ColName="dattDate" RowOrdinal="0" Version="1"/>'
    $ListFields += '<Field Type="Text"     DisplayName="����� ���� ����������" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="tNumAkt" Name="tNumAkt" ColName="tNumAkt" />'
    $ListFields += '<Field Type="DateTime" DisplayName="������� ������ �������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly"  StaticName="dSpisDost" Name="dSpisDost" ColName="dSpisDost" RowOrdinal="0" Version="1" />'
    $ListFields += '<Field Type="Choice"   DisplayName="������� �������� � ���������� ��������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrSkud" Name="bhrSkud" ColName="bhrSkud" RowOrdinal="0" Version="1"><Default>�� �����������</Default><CHOICES><CHOICE>�� �����������</CHOICE><CHOICE>�����������</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Choice"   DisplayName="������� �����" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrLocker" Name="bhrLocker" ColName="bhrLocker" RowOrdinal="0" Version="1"><Default>�� ����������</Default><CHOICES><CHOICE>�� ����������</CHOICE><CHOICE>����������</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Choice"   DisplayName="������������" Description="������������ ���������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrIsLocked" Name="bhrIsLocked" ColName="bhrIsLocked" RowOrdinal="0" Version="1"><Default>�� �����������</Default><CHOICES><CHOICE>�� �����������</CHOICE><CHOICE>�����������</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Choice"   DisplayName="����� ��������� ��� ������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrTypeLocked" Name="bhrTypeLocked" ColName="bhrTypeLocked" RowOrdinal="0" Version="1"><Default>����������� ��������</Default><CHOICES><CHOICE>���</CHOICE><CHOICE>����������� ��������</CHOICE><CHOICE>�� ���� ������</CHOICE></CHOICES></Field>'

    
    # ��������� ����
    foreach ($newColumn in $ListFields)
    {
          $disableoutput = $SpList.Fields.AddFieldAsXml($newColumn,$true,[Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView)
    }
    
    $SpList.Update();
    $spweb.Update();
    $spweb.Dispose();
    
    Create-IQY $spList $oWeb $NameOfListEng
    Add-UsersSampleData $SpList
    Change-RoomsDefaultView  $spweb $spList
    
    return $Listguid
    
}

function Change-RoomsDefaultView( $spweb, $spList)
{

      
      $viewTitle  = $splist.Title

      $viewFields = New-Object System.Collections.Specialized.StringCollection
   
      $viewFields.Add("LinkTitle")
      $viewFields.Add("���� ����������")
      $viewFields.Add("����� ���� ����������")
      $viewFields.Add("������� ������ �������")
      $viewFields.Add("������� �������� � ���������� ��������")
      $viewFields.Add("������� �����")
      
      
      $viewRowLimit    = 300
      $viewDefaultView = $true
      $viewPaged       = $true
      $groupBy         = ''
      
      $viewQuery       = '<OrderBy><FieldRef Name="Title" /></OrderBy>'
    
      $newView = $splist.Views.Add($viewTitle, $viewFields, $($groupBy + $viewQuery), $viewRowLimit, $viewPaged, $viewDefaultView)
      $newView.ApplyStyle($spWeb.ViewStyles.StyleById(17)) # � ����������
      $newView.Update()
    
      
      
    

}
function New-SVTRestrictionList( $spweb, $splist_SVTID, $oWeb)
{
    # ��� ������
    $NameOfListEng     = "SVTRestrict"
    $NameOfListRus     = "������ � ��� ��"
    $ListDescriptionRus= "�������� ����������� ���������� � ���. �������� ���������� � �����������"
    
    # ����� ������������ ���� Title

    $TitleNewName      = "��� ��. ������ � ��� �� ���"
    

    
    # ������� ������
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # ������ ������������ ������
    $SpList = $spWeb.Lists[$NameOfListEng];
    
    $Listguid   = $SpList.GUID
    
    $Fieldtitle  = $title = $spList.Fields.GetFieldByInternalName("Title")
    $Fieldtitle.Required = $false
    $Fieldtitle.Update()
    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus

    
    # ������ ������������ ���� Title
    Rename-ListFieldTitle $splist $TitleNewName

    # ���������� ����
    $ListFields = @()
    
    
    
    
    $spFieldLookup = "����������� ����� ���" 
    $spListParent  = $spweb.Lists["�������� ���"]
    $spFieldParent = "����������� �����"
    # $spFieldOriginal = "tINVN"
    

    $ListFields += Define-LookupField $spListParent $spFieldParent $spFieldLookup 
    # ($spListParent, $spParentLookUpFieldNameRus, $spLookUpFieldNameRus)
    
    $spFieldLookup = "�.�.�. ����������� � ���"
    $spListParent  = $spweb.Lists["������ ������������� ���������������� �������"]
    $spFieldParent = "��������� �.�.�."
    $spFieldParent = "Title"
    
    $ListFields += Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup 
    
    #$ListFields += '<Field Type="Text" DisplayName="��� ��. ������ � ��� �� ���" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="tUserNameSZINSD" Name="tUserNameSZINSD" ColName="tUserNameSZINSD" ><Default/></Field>'

    
    $spFieldLookup = "� ��-��������������"
    $spListParent  = $spweb.Lists["������ �� ���������������"]
    $spFieldParent = "� ��-��������������"
    $spFieldParent = "��������"
    $spFieldParent = "Title"
    #Define-LookupFieldTitle

    
    $ListFields += Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup 
    
    
    $ListFields += '<Field Type="Choice" DisplayName="����� �������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrAccessRigth" Name="bhrAccessRigth" ColName="bhrAccessRigth" ><Default>������������</Default><CHOICES><CHOICE>�������������</CHOICE><CHOICE>���. �������������</CHOICE><CHOICE>������������</CHOICE></CHOICES></Field>'

    $spFieldLookup = "���������� ���" 
    $spListParent  = $spweb.Lists["�������� ���"]
    $spFieldParent = "����������� �����"
    $spFieldDepended = "��"
    
        
    $ListFields += Define-DependedLookupField   $spListParent $spFieldParent $spFieldDepended $spFieldLookup  
    # ��������� ����
    foreach ($newColumn in $ListFields)
    {
          $disableoutput = $SpList.Fields.AddFieldAsXml($newColumn,$true,[Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView)
    }
    
    $SpList.Update();
    $spweb.Update();
    $spweb.Dispose();
    
    Create-IQY $spList  $oWeb $NameOfListEng
    Add-SVTRestrictSampleData  $SpList
    
    Change-SVTRestrictDefaultView $spweb $spList
    
    return $Listguid
    
}
function Change-SVTRestrictDefaultView($spweb, $splist)
{

      
      $viewTitle  = $splist.Title

      $viewFields = New-Object System.Collections.Specialized.StringCollection
      
      $viewFields.Add("�.�.�. ����������� � ���")
      $viewFields.Add("LinkTitle")
      $viewFields.Add("� ��-��������������")
      $viewFields.Add("����� �������")
      $viewFields.Add("���������� ���")
      
      $viewRowLimit    = 300
      $viewDefaultView = $true
      $viewPaged       = $true
      $groupBy         = '<GroupBy Collapse="FALSE" GroupLimit="100"><FieldRef Name="'
      $groupBy        += $(Convert-ToInternalName "����������� ����� ���").Substring(0,32)   # ������-�� ���
      $groupBy        += '" /></GroupBy>'
      $viewQuery       = '<OrderBy><FieldRef Name="ID" /></OrderBy>'
      $newView = $splist.Views.Add($viewTitle, $viewFields, $($groupBy + $viewQuery), $viewRowLimit, $viewPaged, $viewDefaultView)
      
      $newView.ApplyStyle($spWeb.ViewStyles.StyleById(17)) # � ����������
      $newView.Update()

}

function Add-DocumentSetContentType($spSite,$splist)
{
    $spList.ContentTypesEnabled = $true
    $ctTpAdd = $spSite.RootWeb.ContentTypes["����� ����������"]
    $ctTpRemove=$spSite.RootWeb.ContentTypes["��������"]
    $spList.ContentTypes.Add($ctTpAdd)
    $SpList.Update()  

}

function Change-ListOrder ($list)
{
	    $CTOrder = "����� ����������","��������"
            $currentListOrder = $list.RootFolder.ContentTypeOrder
            
            for($i = $currentListOrder.Count; $i -gt 0; $i--)
            {
                $lct = $currentListOrder[$i]
                $disableoutput = $currentListOrder.Remove($lct)
            }

            # add the first contenttype to the order so we can remove the old left over content type
            $currentListOrder.Add($list.ContentTypes[$CTOrder[0]])
            # delete the left over content type
            $disableoutput = $currentListOrder.Remove($currentListOrder[0])
            $list.RootFolder.UniqueContentTypeOrder = $currentListOrder
            $list.RootFolder.Update()
          
}


function Remove-DocLibContentType($docLibrary,$contentType)
{
    $ctToRemove = $docLibrary.ContentTypes[$contentType]
    $docLibrary.ContentTypes.Delete($ctToRemove.Id)
    $docLibrary.Update()
}

function New-DocumentSet( $spweb, $oWeb)
{


    $NameOfListEng     = "SVTs"
    $NameOfListRus     = "�������� ���"
    $ListDescriptionRus= "�������� ���. ����� ���������� ��� �������� ������ �� ���"


    $spSite = Get-SPSite $spWeb.Site.Url    

    # ������� ���������� ����������
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "DocumentLibrary")
    
    $SpList = $spWeb.Lists[$NameOfListEng];
          
    Add-DocumentSetContentType $spSite $SpList

    
    $Listguid   = $SpList.Id

    # �������� ������� ���� �������� content types
    Change-ListOrder $spList

    Remove-DocLibContentType  $spList "��������"  

    $Listguid   = $SpList.GUID
    foreach($culture in $spweb.SupportedUiCultures)
    {


      [System.Threading.Thread]::CurrentThread.CurrentUICulture=$culture
      $SpList = $Web.Lists[$NameOfListEng]
      $SpList.Title = $NameOfListRus
      $SpList.Update()

    }
    $ListFields = @()
    


    $ListFields += '<Field Type="Text"   DisplayName="����������� �����" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tINVN" Name="tINVN" ColName="tINVN" />'
    $ListFields += '<Field Type="Text"   DisplayName="��������� �����" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tSN" Name="tSN" ColName="tSN" />'
    $ListFields += '<Field Type="Text"   DisplayName="��� ���" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tCompType" Name="tCompType" ColName="tCompType" />'    
    $ListFields += '<Field Type="Text"   DisplayName="��" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tAS" Name="tAS" ColName="tAS" />'
    $ListFields += '<Field Type="Text"   DisplayName="�.�.�. ��������������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tFIO" Name="tFIO" ColName="tFIO" />'    
    $ListFields += '<Field Type="Choice" DisplayName="BBK" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" ColName="tBBK" ><Default>��������� ����������� 6.0</Default><CHOICES><CHOICE>��������� ����������� 6.0</CHOICE><CHOICE>��������� ����������� 8.0</CHOICE><CHOICE>��������� ����������� 10</CHOICE><CHOICE>Dr.Web</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Choice" DisplayName="������������ �������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" StaticName="bhrOS" Name="bhrOS" ColName="bhrOS" ><Default>MS DOS</Default><CHOICES><CHOICE>MS DOS</CHOICE><CHOICE>Windows XP</CHOICE><CHOICE>Windows 7</CHOICE><CHOICE>Windows 8</CHOICE><CHOICE>Windows 2003</CHOICE><CHOICE>Windows 2008</CHOICE><CHOICE>Windows 2012</CHOICE></CHOICES></Field>'    
    $ListFields += '<Field Type="Choice" DisplayName="��� �� ���" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrSZINSD" Name="bhrSZINSD" ColName="bhrSZINSD" ><Default>���</Default><CHOICES><CHOICE>���</CHOICE><CHOICE>������-Win32</CHOICE><CHOICE>������-Win64</CHOICE><CHOICE>������-����</CHOICE><CHOICE>SecretNet</CHOICE><CHOICE>������</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Text"   DisplayName="���. ����� ���" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="tSZIInvN" Name="tSZIInvN" ColName="tSZIInvN" />'
    $ListFields += '<Field Type="Text"   DisplayName="��������� ����� ���" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="tSZISN" Name="tSZISN" ColName="tSZISN" />'
    $ListFields += '<Field Type="Text"   DisplayName="����� ���������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tRoomNumber" Name="tRoomNumber" ColName="tRoomNumber" />' 
    $ListFields += '<Field Type="Calculated" DisplayName="�� ����"  ResultType="Text" ><Formula>=1</Formula></Field>'




    
    # ��������� ����
    foreach ($newColumn in $ListFields)
    {
          $disableoutput = $SpList.Fields.AddFieldAsXml($newColumn,$true,[Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView)
    }
    
    $fieldc = $splist.Fields["�� ����"]
    $fieldc.Formula = '=�������(" [";[���. ����� ���];" ";[��];" ")'
    $fieldc.Update()
         
    $spweb.Dispose()
    
    
    
    Create-IQY $spList $oWeb $NameOfListEng
    Add-SVTsampleData $SpList
    Change-DocumentsetDefaultView $spweb  $splist
    return $Listguid
        
       
}
function Change-DocumentsetDefaultView($spweb, $splist)
{
      #$defaultView = $spList.DefaultView
      
      $viewTitle  = $splist.Title
      #write-host $ViewTitle
      $viewFields = New-Object System.Collections.Specialized.StringCollection
      

      $viewFields.Add("DocIcon")   
      $viewFields.Add("LinkFilename")
      $viewFields.Add("����������� �����")
      $viewFields.Add("����� ���������")    
      $viewFields.Add("��������� �����")    
      $viewFields.Add("��� ���")   
      $viewFields.Add("�.�.�. ��������������")   
      $viewFields.Add("��")    
      $viewFields.Add("BBK")    
      
      $viewRowLimit    = 300
      $viewDefaultView = $true
      $viewPaged       = $true
      $viewQuery       = '<OrderBy><FieldRef Name="FileLeafRef" /></OrderBy>'
      $newView = $splist.Views.Add($viewTitle, $viewFields,  $viewQuery, $viewRowLimit, $viewPaged, $viewDefaultView)
      
      $newView.ApplyStyle($spWeb.ViewStyles.StyleById(17)) # � ����������
      $newView.Update()
  
}
function Change-Theme ($webSite, $spweb)
{
    $tcount = [Microsoft.Sharepoint.Utilities.ThmxTheme]::GetManagedThemes($webSite).Count
    If ($tcount -gt 1) # 2010
    {

          $theme = [Microsoft.Sharepoint.Utilities.ThmxTheme]::GetManagedThemes($webSite)[6]

          $theme.DarkColor1.DefaultColor="#000000"
          $theme.DarkColor2.DefaultColor="#3E5354"


          $theme.LightColor1.DefaultColor = "#CCECFF"
          $theme.LightColor2.DefaultColor = "#DDDDDD"


          $theme.AccentColor1.DefaultColor = "#859E9D"
          $theme.AccentColor2.DefaultColor = "#9EB5B4"
          $theme.AccentColor3.DefaultColor = "#5DD1C6"
          $theme.AccentColor4.DefaultColor = "#7DC3D9"
          $theme.AccentColor5.DefaultColor = "#8FDBD4"
          $theme.AccentColor6.DefaultColor = "#A1BAC7"
          $theme.HyperlinkColor.DefaultColor = "#0000FF"
          $theme.FollowedHyperlinkColor.DefaultColor = "#660033"
          $theme.MajorFont.LatinFont="Agency FB"
          $theme.MinorFont.LatinFont="Agency FB"


          $theme.ApplyTo($spweb,$false)
    }
}

function add-newwebTU($oweb, $CurrentSiteIndex)
{
      
      $webSite        = $oWeb.Site 
      $webHierarchy   = $oWeb.Hierarchy
      $newWeb         = $oWeb.NewWebName  + $CurrentSiteIndex.ToString().Trim()
      $webTitle       = $oWeb.Title  + " � "+$($CurrentSiteIndex.ToString().PadLeft(2,"0"))
      $WebDescription = $oWeb.Description  + " � "+$($CurrentSiteIndex.ToString().PadLeft(2,"0"))
      
 
      $siteTemplate = "STS#1" # blank Site
      $webUrl=$webSite+$webHierarchy+$newWeb


      # $webUrl
      Remove-SPWeb $webUrl -Confirm:$false -ErrorAction SilentlyContinue 

      write-host $("��� ���� : " + $webUrl)  -ForegroundColor Green
      write-host "�������� ����� ���������������� �������..." -nonewline
      # �������� ���-�����
      $nooutput = new-spweb $webUrl -Template $siteTemplate  -Name $webTitle  -UseParentTopNav -Description $WebDescription # -Description $WebDescription

      $web = get-spweb -Identity $webUrl

      Change-Theme $webSite $web

      $web.Update()
      $web.Dispose()

      write-host "[OK]" -ForegroundColor Yellow
      $web = get-spweb -Identity $webUrl

      write-host "�������� ������� � ���������..." -nonewline
      #�������� ������� � ��������� ����������
      $splistAdminsGuid = New-AdminsList $web $oWeb
     

      $splistUsersGuid = New-UsersList $web

      $splistRoomsGuid = New-RoomList  $web $oWeb


      $splistTMGuid = New-TMList $web 
      $splistSVTsGuid = New-DocumentSet $web $oWeb

      $splistSVTRestrictions = New-SVTRestrictionList $web $splistSVTsGuid $oWeb
      # "�������� ����������� ���������� � ���. �������� ���������� � �����������"
      $splistDailyReportGuid = New-DailyReport $web $oWeb
      
      $splistDReportLibGuid = New-DReportLib  $web $oWeb 
      
      $splistReports = New-ReportDocLib $web
      
      
      write-host "[OK]" -ForegroundColor Yellow

      write-host "�������� ������ ���������..." -nonewline
      $web.QuickLaunchEnabled=$true

      Costumize-QuickLaunchMenu $oweb $CurrentSiteIndex $web
      write-host "[OK]" -ForegroundColor Yellow

      write-host "���������� ������� ��������..." -nonewline
      Update-WebParts $web $oWeb
      write-host "[OK]" -ForegroundColor Yellow
      $web.Update()
      write-host "������ ���� ���������������� ������� ...  "  -ForegroundColor Green
      write-host $($webSite + $webHierarchy + $newWeb)
      write-host
      write-host ===================================
      write-host �������� Excel... -nonewline
      $ExcelExists = Test-Excel
      if ($ExcelExists)
      {
            write-host "[OK]" -ForegroundColor Yellow
            write-host "�������� ������ ���������������� ������� ...  "  -ForegroundColor Yellow
            $reportName = "Report" + $CurrentSiteIndex.ToString() + ".xlsx"
            $ReportFileName = Create-Excel $reportName $oweb
            
            
            $spReportLib= $web.GetFolder("Reports")
            Write-host $("�������� ���� "+$ReportFileName+ " � ���������� �������.") 
            Add-DocumentToDocLib $spReportLib $ReportFileName "������� ��.xlsx"
      }
      else
      {
            write-host "[ERROR]" -ForegroundColor Red
            
            write-host "�� ������ ���������� �� ���������� Excel ������ 2010 ��� ����."  -ForegroundColor Yellow
            write-host $("����� ��������� ������ ������� ��������� ������� ") -ForegroundColor Yellow 
            
            write-host $oweb.iqyFolder -ForegroundColor Cyan 
           
            write-host $("�� ����� ������������� �� ���������� �� ������� ���������� Excel") -ForegroundColor Yellow
            write-host $("� ��������� �� ����� �������� ��������� ���������� ������� �������")  -ForegroundColor Yellow
            write-host $("PvtExcel.ps1")  -ForegroundColor Cyan
      }
      $web.Update()
      write-host ===================================
      write-host
        
      
}

function Add-DocumentToDocLib($spDocLib,$FileToStore,$nameToStore)
{
     
     $File = Get-ChildItem $FileToStore
     $spDocLib.Files.Add($nameToStore,$File.OpenRead(),$false) | Out-Null
}
function New-ReportDocLib( $spweb)
{


    $NameOfListEng     = "Reports"
    $NameOfListRus     = "������"
    $ListDescriptionRus= "������ �� ���������������� �������"


    $spSite = Get-SPSite $spWeb.Site.Url    

    # ������� ���������� ����������
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "DocumentLibrary")
    
    $SpList = $spWeb.Lists[$NameOfListEng];
          
    #Add-DocumentSetContentType $spSite $SpList

 
    
    
    $Listguid   = $SpList.Id

   

    $Listguid   = $SpList.GUID
    foreach($culture in $spweb.SupportedUiCultures)
    {


      [System.Threading.Thread]::CurrentThread.CurrentUICulture=$culture
      $SpList = $Web.Lists[$NameOfListEng]
      $SpList.Title = $NameOfListRus
      $SpList.Update()

    }
     
    $spweb.Dispose()
    

    
    return $Listguid
        
       
}

function Check-SPWebSite($webSite,$webHierarchy)
{
$s = $webSite+$webHierarchy

$webt = Get-SPWeb -identity $s -ErrorAction SilentlyContinue

$SiteExists = !([string]::isnullorempty($webt.URL))

return $SiteExists

}



start-transcript "TU.log"




[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Publishing")
[System.Reflection.Assembly]::LoadWithPartialName("System")
[System.Reflection.Assembly]::LoadWithPartialName("System.IO")
Add-PsSnapin Microsoft.SharePoint.PowerShell
cls
$0 = $myInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$IniFile = "$dp0\TU.INI"   
. "$dp0\Utils-Excel.ps1"
. "$dp0\Utils-Ini.ps1"
. "$dp0\Utils-SampleData.ps1"
. "$dp0\List-DailyReport"

# �������������
$isIniExists = Test-Ini $IniFile $dp0

if ($isIniExists)
{

        $webOBool = Test-WebIniObject  $IniFile
        $IsParamFileGood = Show-BadIniParams $webOBool $IniFile


        if ($IsParamFileGood)
        {
              $webObject = Get-WebIniObject $IniFile 
              $webObject.WrkDir = $dp0
              #$webObject | fl
              
              $webSite = $webObject.Site
              $webHierarchy = $webObject.Hierarchy  
              $newWeb = $webObject.NewWebName  
              $infoModule = $webObject.InfoModule  
              $webTitle = $webObject.Title       
              $WebDescription = $webObject.Description 
              $websiteCount = $webObject.SiteCount   
              
              $SiteExists = Check-SPWebSite  $webSite $webHierarchy
              if ($SiteExists)
              {
 
                    write-host "������ �� �������� Web ������ ��������������� ��������."
                    write-host "������ ������ ������� ��� ����� � ����� ��������������."
                    Write-host "���������� ������� ������� ��������� � ������������ � ������� �� �� ��� "
                    Write-host "� 19-2-18/178560 �� 16.12.2016"
                    Write-host 

                    Write-host "-------------------------------------------------------------"
                    Write-host "�� ���� �������� ��������� � ������� ������� �������, ����������� ��� ������ "
                    write-host "���������� ���������� � ��������� ������ �� �� ���."
                    write-host "��������� �����:����� ������ �������,�.8-223-1261 - ������������ �����������;"
                    write-host "��������� �� :������� ��������� ���������,�. 8-223-1338 - ����������� ���������."
                    write-host "-------------------------------------------------------------"
                    write-host "��� ����������� ������� <ENTER>" -nonewline
                    read-host
                    write-host 
                    write-host "��������!!!! ����� ������� ��������� ��� �����:" -foreground green
                    write-host "��� - ���� ��������������� ������:" -foreground green
                    write-host  $($webSite + $webHierarchy + $infoModule) -foreground yellow
                    write-host 
                    write-host "��� - ����� ��������������� ��������:" -foreground green
                    
                    
                    for ($jj = 1; $jj -le $websiteCount; $jj++)
                    {

                        write-host  $($webSite + $webHierarchy + $newWeb + $jj.Tostring().Trim()) -foreground yellow
                    }

                    write-host "���� ����� ����� ��� ����������, ��� ����� �������, ��� ������ �� ��� ����� ����������!" -foreground green
                    write-host "-------------------------------------------------------------"
                    write-host 
                    write-host "������ ����� ����������� " -foreground yellow -nonewline
                    write-host "���-���" -foreground cyan -nonewline
                    write-host ", ������� �� ������ ������ ������," -foreground yellow 
                    write-host "���� �� �� �������, ��� �� ���������� ���������, ������ ������� Enter," 
                    write-host "������ �������� ������, ������ �� ����� ������� � ����� �� ����� �������."
                    write-host 
                    write-host "���� ������� ���-���� ����� �������� ���� �������� �� �������� ������."
                    write-host "-------------------------------------------------------------"




                    $randomToken = Get-Random -Maximum 9999 -Minimum 1

                    $Pin = $randomToken.ToString().PadLeft(4, "0")
                    Write-Host $("���-��� :" + $Pin) -foreground cyan
                    Write-Host "������� ���-���:" -nonewline
                    $userPin = read-host  


                    If ($userPin -eq $Pin){

                            Clear-TempFolder

                            Enable-SPFeature -identity "DocumentSet" -Url $webSite -ErrorAction SilentlyContinue 
                            Add-newInforModuleWebSite $webObject

                            for ($kk=1;$kk -le $webObject.SiteCount; $kk++)
                            {
                                
                                  add-newwebTU $webObject $kk
                            }

                    }
                    else
                    {
                        Write-Host  "������ �� ���� �������. �������� ���!"

                    }
        }
        else
        {
                    write-host $("� ����������������� ����� " +$($IniFile )+ " ������� ���������") -foreground cyan 
                    write-host $("Site="+$webSite) -foreground Yellow
                    write-host $("Hierarchy="+$webHierarchy) -foreground Yellow
                    write-host $("��� ������� ��������� � ����� "+$webSite+$webHierarchy+" ��������� ������.") -foreground Yellow
                    write-host $("������ ���� �� ���������� ��� ��� ���� ������� � ����� �����. ") -foreground cyan 
        }     
        }
        
        else
        {
             write-host $("���������� �������� ���������...")
        }
}
else
{
        Write-host
        Write-Host $("��������!  �� ��� ������ ����������������� ����.") -foreground cyan
        Write-host $("��������� ������������� ������� ����������������� ���� � ������ : ") -foreground cyan 
        Write-host $($IniFile ) -foreground Yellow -nonewline
        Write-host $(" � ����������� �� ���������.") -foreground cyan 
        Write-host
        Write-host
        Write-Host $("��������� ��������� ������� ����� � ��������� ��������� ��� ���.") -foreground cyan
        
}

Stop-Transcript

