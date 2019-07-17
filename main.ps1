##############################################
#
#   Модуль создания сайтов технологических участков
#
#   Файл: Main.ps1 
#   Описание : Главный исполняемый модуль. Создает сайты 
#              информационного модуля и технологических участков.
#
#   Автор: Астахов Александр Борисович, отдел информатизации, тел: 223-1338
#          29AstahovAB@cbr.ru
#
#   Отделение Калуга, ГУ Банка России по ЦФО
#   Условия выполнения: Данную процедуру следует запускать на сервере Sharepoint 
#                       из командной консоли Sharepoint 2010
#   Дата: Январь 2017.
#   Версия: 17.01.11
#
##############################################
function Add-newInforModuleWebSite ($webObject)
{

      $webSite        = $webObject.Site 
      $webHierarchy   = $webObject.Hierarchy
      $newWeb         = $webObject.InfoModule
      $webTitle       = "Информационный модуль"
      $WebDescription = "Информационный модуль"

      $siteTemplate = "STS#1" # blank Site
      
      $webUrl=$webSite+$webHierarchy+$newWeb
      
      
      
      $webUrl=$webSite+$webHierarchy+$newWeb


      # $webUrl
      Remove-SPWeb $webUrl -Confirm:$false -ErrorAction SilentlyContinue 


      write-host $("Веб сайт : " + $webUrl)  -ForegroundColor Green
      write-host "Создание сайта Информационный модуль..." -nonewline
      # Создание веб-сайта
      $nooutput = new-spweb $webUrl -Template $siteTemplate  -Name $webTitle  -UseParentTopNav -Description $WebDescription # -Description $WebDescription

      $web = get-spweb -Identity $webUrl
      Change-Theme $webSite $web
          

      $web.Update()
      $web.Dispose()
          
      write-host "[OK]" -ForegroundColor Yellow
      $web = get-spweb -Identity $webUrl
      write-host "Создание списков и библиотек..." -nonewline
      #создание списков и библиотек документов
      $splistAdminsGuid = New-InfoModuleDocLib $web

      write-host "[OK]" -ForegroundColor Yellow

      write-host "Создание панели навигации..." -nonewline
      $web.QuickLaunchEnabled=$true

      Costumize-InfoQuickLaunchMenu $web $webObject
      write-host "[OK]" -ForegroundColor Yellow

      write-host "Обновление главной страницы..." -nonewline
      Update-InfoWebParts $web $oWeb
      write-host "[OK]" -ForegroundColor Yellow
      $web.Update()
      write-host "Создан сайт Информационный модуль...  "  -ForegroundColor Green
      write-host $webUrl
      write-host
      write-host ===================================
      write-host      
}

function Update-InfoWebParts ($spweb,$webObject)
{
     $listToAdd = "Нормативная база"
     Add-XsltListViewWebPart $spweb  $listToAdd 
	 
}
function New-InfoModuleDocLib( $spweb)
{


    $NameOfListEng     = "Ndocs"
    $NameOfListRus     = "Нормативная база"
    $ListDescriptionRus= "Нормативные документы"


    $spSite = Get-SPSite $spWeb.Site.Url    

    # создаем Библиотеку документов
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
    
    $ListFields += '<Field Type="Number" DisplayName="№ п/п" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" StaticName="fNPP" Name="fNPP" ColName="fNPP"  />'
    $ListFields += '<Field   Name="FirstName" StaticName="FirstName"  DisplayName="Документ" Type="Text" ColName="FirstName"  EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" />'

    # добавляем поля
    foreach ($newColumn in $ListFields)
    {
          $disableoutput = $SpList.Fields.AddFieldAsXml($newColumn,$true,[Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView)
    }
    
     
    $spweb.Dispose()
    
    #$cType = $SpList.ContentsTypes["Набор документов"]
    
    return $Listguid
        
       
}

function Costumize-InfoQuickLaunchMenu ($spWeb, $oWeb){
        $url = $spWeb.URL

	$qlmenu = @()
	$qlmenu += 'parent,СПРАВОЧНИКИ'
	$qlmenu +=       'child, СПРАВОЧНИКИ,Нормативная база,' + $url +'/Ndocs/'
	$qlmenu += 'parent,Технологические участки'
	for ($kk=1; $kk -le $oWeb.SiteCount; $kk++)
	{
	      $qlmenu += 'child,Технологические участки, Технологический участок ' + $kk.ToString().Trim() + ',' + $oweb.Site + $oweb.Hierarchy + $oweb.NewWebName + $kk.ToString().Trim()
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

   write-host $("Очищаем временный каталог :"+$allfile)
   
   remove-Item $allfile -recurse -ErrorAction SilentlyContinue 


}
function New-SummaryLinks( $webObject)
{
            $slinks = @()

            $groupLinkA = New-Object Microsoft.SharePoint.Publishing.SummaryLink "Информационный модуль"
            $groupLinkA.IsGroupHeader = $true;
            $slinks  += $groupLinkA
            
            $sumLink = New-Object Microsoft.SharePoint.Publishing.SummaryLink "Сайт Информационного модуля"
           
            $sumLink.LinkUrl = $webObject.Site + $webObject.Hierarchy + $webObject.infoModule
            $sumLink.Style = "TitleWithBackground";
            $slinks  += $sumLink
            
            
            $groupLinkB = New-Object Microsoft.SharePoint.Publishing.SummaryLink "Сайты технологических участков"
            $groupLinkB.IsGroupHeader = $true;
            $slinks  += $groupLinkB



            for ($i=1; $i -le $webObject.SiteCount ; $i++)
            {
                 $url =  $webObject.Site + $webObject.Hierarchy + $webObject.NewWebName + $i.ToString().Trim()
                 $webTitle   = $webObject.Title+" № "+ $($i.ToString().PadLeft(2,"0"))
                 $WebDescription = $webObject.Description + " № " + $($i.ToString().PadLeft(2,"0"))

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
      
      $page.CheckIn("Страница обновлена автоматически.")      

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
 
     $page.CheckIn("Страница обновлена автоматически.")

     
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
      
      $page.CheckIn("Страница обновлена автоматически.")      
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
     $listToAdd = "Допуск к СВТ ТУ"
     Add-XsltListViewWebPart $spweb  $listToAdd 
     
     
                  
      $urllist = $spweb.Url.ToString()
      
     
	
      $content ='<div class="ms-rteThemeBackColor-6-3"><hr class="ms-rteElement-Hr"/></div>'+
                '<div class="ms-rteThemeBackColor-6-3"><span class="ms-rteFontSize-3">Перечень работников, допущенных к&#160;техническим средствам&#160;технологического участка УОС</span> </div>'+
                '<div class="ms-rteThemeBackColor-6-3">&#160;</div>'+
                '<div class="ms-rteThemeBackColor-6-3"><span>Список&#160;редактирует АИБ ТУ УОС при любых изменениях состава пользователей данного технологического участка</span></div>'+
                '<div class="ms-rteThemeBackColor-6-3">&#160;</div>'+
                '<div class="ms-rteThemeBackColor-6-3"><em><font color="#cc3300">(Перед&#160;наполнением&#160;списка необходимо заполнить</font><span style="color: #cc3300"> </span></em>'+
                '<a class="ms-rteFontSize-2" href="' + $urllist + '/Lists/TUUsr/" target="_blank" style="text-decoration: underline">'+
                '<em>справочник пользователей</em></a><em><span class="ms-rteFontSize-2"> </span><span class="ms-rteFontSize-2" style="color: #cc3300"><span style="color: #cc3300">и</span></span>'+
                '<span class="ms-rteForeColor-4 ms-rteFontSize-2"> </span></em>'+
                '<a class="ms-rteFontSize-2" href="' + $urllist+ '/Lists/TM/" target="_blank" style="text-decoration: underline"><em>справочник ТМ-идентификаторов</em></a>)</div>'+
                '<div class="ms-rteThemeBackColor-6-3">&#160;</div>'



     Add-ContentEditorWebPart $spWeb "Сведения СВТ ТУ" $content



     $listToAdd = "Перечень СВТ"
     Add-XsltListViewWebPart $spweb  $listToAdd
 
     $content ='<div class="ms-rteThemeBackColor-6-3"><hr class="ms-rteElement-Hr"/></div>'+
                '<div class="ms-rteThemeBackColor-6-3"><span class="ms-rteFontSize-3">Перечень&#160;технических средств</span><span class="ms-rteFontSize-3">&#160;технологического участка УОС</span></div>'+
                '<div class="ms-rteThemeBackColor-6-3"><span class="ms-rteFontSize-1"></span><em></em>&#160;</div>'+
                '<div class="ms-rteThemeBackColor-6-3"><em><span>Список&#160;редактирует АИБ </span>ТУ УОС при любых&#160;изменениях&#160;перечня СВТ данного технологического участка</em>&#160;</div>'+
                '<div class="ms-rteThemeBackColor-6-3">&#160;</div>'
     
     Add-ContentEditorWebPart $spWeb "Сведения СВТ" $content
     
     $listToAdd = "Перечень помещений технологического участка"
     Add-XsltListViewWebPart $spweb  $listToAdd     



      $content ='<div class="ms-rteThemeBackColor-6-3"><hr class="ms-rteElement-Hr"/></div>'+
                '<div class="ms-rteFontSize-3 ms-rteThemeBackColor-6-3">Перечень помещений технологического участка УОС&#160;</div>'+
                '<div class="ms-rteFontSize-1 ms-rteThemeBackColor-6-3"><em></em>&#160;</div>'+
                '<div class="ms-rteFontSize-1 ms-rteThemeBackColor-6-3"><em>Список&#160;редактирует АИБ ТУ УОС при любых&#160;изменениях перечня помещений данного технологического участка</em></div>'+
                '<div class="ms-rteFontSize-1 ms-rteThemeBackColor-6-3"><em></em>&#160;</div>'



     Add-ContentEditorWebPart $spWeb "Сведения помещений" $content 
     
     
     $listToAdd = "Список администраторов технологического участка"
     Add-XsltListViewWebPart $spweb  $listToAdd     
     
    $content ='<div class="ms-rteThemeBackColor-6-3"><hr class="ms-rteElement-Hr"/></div>'+
              '<div class="ms-rteFontSize-3 ms-rteThemeBackColor-6-3"><span>Перечень&#160;администраторов информационной безопасности&#160; </span>технологического участка УОС&#160;</div>'+
              '<div class="ms-rteThemeBackColor-6-3" style="color: #080808"><span class="ms-rteStyle-Normal"><span><span><em></em></span></span></span>&#160;</div>'+
              '<div class="ms-rteThemeBackColor-6-3" style="color: #080808"><span class="ms-rteStyle-Normal ms-rteThemeBackColor-6-3"><span><span>'+
              '<em>Список&#160;редактирует АИБ ТУ УОС при любых&#160;изменениях состава&#160;администраторов <span>данного технологического </span>участка</em></span></span></span></div>'+
              '<div class="ms-rteThemeBackColor-6-3" style="color: #080808"><span class="ms-rteStyle-Normal ms-rteThemeBackColor-6-3"><span><span><em></em></span></span></span>&#160;</div>'

     
     Add-ContentEditorWebPart $spWeb "Сведения администраторов" $content 
     
     
 
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

     Add-ContentEditorWebPart $spWeb "Спрятать правую зону" $content 
		 
}
     

function Costumize-QuickLaunchMenu ($oweb, $CurrentSiteIndex, $spWeb){
        $url = $spWeb.URL
                   
	$qlmenu = @()
	$qlmenu += 'parent,САЙТЫ'
	$qlmenu +=       'child,САЙТЫ,Сайт Информационного модуля,'+$($oWeb.Site + $oWeb.Hierarchy + $oWeb.infoModule)
    	
            for ($i=1; $i -le $oweb.SiteCount ; $i++)
            {
                 if ($i -eq $CurrentSiteIndex)
                 {
                 
                       $url1 =  $webObject.Site + $webObject.Hierarchy + $webObject.NewWebName + $i.ToString().Trim()
                       $webTitle   = $webObject.Title+" № "+ $($i.ToString().PadLeft(2,"0"))

                       $qlmenu +=       'child,САЙТЫ,' + $webTitle +','+$url1
                 }      
           }

	
	
	$qlmenu += 'parent,СПРАВОЧНИКИ'
	$qlmenu +=       'child, СПРАВОЧНИКИ,Пользователи технологического участка,' + $url + '/Lists/TUUsr/'
	$qlmenu +=       'child, СПРАВОЧНИКИ,ТМ идентификаторы,' + $url + '/Lists/TM/'
	$qlmenu += 'parent,ТЕХ. УЧАСТОК'
	$qlmenu +=     'child,ТЕХ. УЧАСТОК,Администраторы технологического участка,' + $url + '/Lists/Admins/'
	$qlmenu +=     'child,ТЕХ. УЧАСТОК, Допуск к СВТ ТУ,' + $url + '/Lists/SVTRestrict/'
	$qlmenu +=     'child,ТЕХ. УЧАСТОК, Перечень СВТ,' + $url + '/SVTs/'
	$qlmenu +=     'child,ТЕХ. УЧАСТОК, Помещения технологического участка,' + $url + '/Lists/Rooms/'
	$qlmenu +=     'child,ТЕХ. УЧАСТОК, Ежедневный отчет,' + $url + '/Lists/DailyReport/'
	
	$qlmenu += 'parent,ОТЧЕТЫ'
	$qlmenu +=     'child,ОТЧЕТЫ,Отчетные формы,' + $url + '/Reports'
	$qlmenu +=     'child,ОТЧЕТЫ,Ежедневные отчеты АИБ в формате MSWORD,' + $url + '/DReports'


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
    $fild += 'FieldRef="'+$fieldid+'" '  # главное поле Lookup
    $fild += ' ShowField="' + $spFieldInternalName +'" />'  
    # Write-Host $fild
    return $fild
}

function Rename-ListFieldTitle(
    $spListObject,
    [string]$NameOfField = "Наименование",
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
    [string]$NameOfField = "Наименование",
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
    # имя списка
    $NameOfListEng     = "TUUsr"
    $NameOfListRus     = "Список пользователей технологического участка"
    $ListDescriptionRus= "Список пользователей технологического участка. Хранится информация о пользователях тех.участков"
    
    # Новое еаименование поля Title
    $TitleNewName      = "Сотрудник Ф.И.О."
    
 
    # создаем список
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # Меняем наименование списка
    $SpList = $spWeb.Lists[$NameOfListEng];
    $Listguid   = $SpList.Id
    # переименовываем список    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus
    
    
    # Меняем наименование поля Title
    Rename-ListFieldTitle $splist $TitleNewName

   
    
    $SpList.Update();
    $spweb.Update();
    $spweb.Dispose();
    
    
    Add-TuUsersSampleData($SpList)
    
    return $Listguid
    
}

function New-TMList( $spweb, $oWeb)
{
    # имя списка
    $NameOfListEng     = "TM"
    $NameOfListRus     = "Список ТМ идентификаторов"
    $ListDescriptionRus= "Список Список ТМ идентификаторов технологического участка. Хранится информация о Список ТМ идентификаторах тех.участка."
    
    # Новое еаименование поля Title
    $TitleNewName      = "№ ТМ-идентификатора"
    
 
    # создаем список
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # Меняем наименование списка
    $SpList = $spWeb.Lists[$NameOfListEng];
    $Listguid   = $SpList.Id
    # переименовываем список    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus
    
    
    # Меняем наименование поля Title
    Rename-ListFieldTitle $splist $TitleNewName

   
    
    $SpList.Update();
    $spweb.Update();
    $spweb.Dispose();
    
    Add-TMSampleData($SpList)
    
    return $Listguid
    
}
function New-AdminsList( $spweb)
{
    # имя списка
    $NameOfListEng     = "Admins"
    $NameOfListRus     = "Список администраторов технологического участка"
    $ListDescriptionRus= "Список администраторов технологического участка. Хранится информация об администраторах тех.участков"
    
    # Новое еаименование поля Title
    $TitleNewName      = "Сотрудник Ф.И.О."
    
 
    # создаем список
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # Меняем наименование списка
    $SpList = $spWeb.Lists[$NameOfListEng];
    $Listguid   = $SpList.Id
    # переименовываем список    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus
    
    
    # Меняем наименование поля Title
    Rename-ListFieldTitle $splist $TitleNewName

    # определяем поля
    $ListFields = @()
    
    $ListFields += '<Field Type="DateTime" DisplayName="Дата ознакомления" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly" Name="DateOfSigned" Description="Дата ознакомления сотрудника с приказом"/>'
    $ListFields += '<Field Type="Choice"   DisplayName="Ознакомился" Required="FALSE" EnforceUniqueValues="FALSE" Format="RadioButtons" FillInChoice="FALSE"  StaticName="IsSigned" Name="IsSigned" ><Default>Нет</Default><CHOICES><CHOICE>Да</CHOICE><CHOICE>Нет</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Choice"   DisplayName="Роль" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" StaticName="AdminRole" Name="AdminRole" ><Default>АИБ</Default><CHOICES><CHOICE>АИБ</CHOICE><CHOICE>Резервный АИБ</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Text"     DisplayName="Должность" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="PersonellPosition" Name="PersonellPosition" />'
    $ListFields += '<Field Type="DateTime" DisplayName="Дата приказа о назначении" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly" StaticName="dpONazn" Name="dpONazn" />'
    $ListFields += '<Field Type="Text"     DisplayName="Номер приказа о назначении" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="npONazn" Name="npONazn" />'
    
       
    # добавляем поля
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
      

      
      
      $viewFields.Add("Роль")
      $viewFields.Add("LinkTitle")
      $viewFields.Add("Номер приказа о назначении")
      $viewFields.Add("Дата приказа о назначении")
      
      $viewRowLimit    = 300
      $viewDefaultView = $true
      $viewPaged       = $true
      $groupBy         = ''
      
      $viewQuery       = '<OrderBy><FieldRef Name="Title" /></OrderBy>'
    
      $newView = $splist.Views.Add($viewTitle, $viewFields, $($groupBy + $viewQuery), $viewRowLimit, $viewPaged, $viewDefaultView)
      $newView.ApplyStyle($spWeb.ViewStyles.StyleById(17)) # С затенением
      $newView.Update()
    

}

function New-RoomList( $spweb, $oWeb)
{
    # имя списка
    $NameOfListEng     = "Rooms"
    $NameOfListRus     = "Перечень помещений технологического участка"
    $ListDescriptionRus= "Перечень помещений технологического участка. Хранится информация об помещения тех.участков"
    
    # Новое еаименование поля Title
    $TitleNewName      = "Номер помещения"
    

    
    # создаем список
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # Меняем наименование списка
    $SpList = $spWeb.Lists[$NameOfListEng];
    
    $Listguid   = $SpList.Id
    
    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus

    
    # Меняем наименование поля Title
    Rename-ListFieldTitle $splist $TitleNewName

    # определяем поля
    $ListFields = @()
    
    
    $ListFields += '<Field Type="DateTime" DisplayName="Дата аттестации" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly"  StaticName="dattDate" Name="dattDate" ColName="dattDate" RowOrdinal="0" Version="1"/>'
    $ListFields += '<Field Type="Text"     DisplayName="Номер акта аттестации" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="tNumAkt" Name="tNumAkt" ColName="tNumAkt" />'
    $ListFields += '<Field Type="DateTime" DisplayName="Наличие списка доступа" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly"  StaticName="dSpisDost" Name="dSpisDost" ColName="dSpisDost" RowOrdinal="0" Version="1" />'
    $ListFields += '<Field Type="Choice"   DisplayName="Система контроля и управления доступом" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrSkud" Name="bhrSkud" ColName="bhrSkud" RowOrdinal="0" Version="1"><Default>Не установлена</Default><CHOICES><CHOICE>Не установлена</CHOICE><CHOICE>Установлена</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Choice"   DisplayName="Кодовый замок" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrLocker" Name="bhrLocker" ColName="bhrLocker" RowOrdinal="0" Version="1"><Default>Не установлен</Default><CHOICES><CHOICE>Не установлен</CHOICE><CHOICE>Установлен</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Choice"   DisplayName="Опечатывание" Description="Опечатывание помещения" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrIsLocked" Name="bhrIsLocked" ColName="bhrIsLocked" RowOrdinal="0" Version="1"><Default>Не выполняется</Default><CHOICES><CHOICE>Не выполняется</CHOICE><CHOICE>Выполняется</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Choice"   DisplayName="Сдача помещения под охрану" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrTypeLocked" Name="bhrTypeLocked" ColName="bhrTypeLocked" RowOrdinal="0" Version="1"><Default>Электронная ключница</Default><CHOICES><CHOICE>Нет</CHOICE><CHOICE>Электронная ключница</CHOICE><CHOICE>На пост охраны</CHOICE></CHOICES></Field>'

    
    # добавляем поля
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
      $viewFields.Add("Дата аттестации")
      $viewFields.Add("Номер акта аттестации")
      $viewFields.Add("Наличие списка доступа")
      $viewFields.Add("Система контроля и управления доступом")
      $viewFields.Add("Кодовый замок")
      
      
      $viewRowLimit    = 300
      $viewDefaultView = $true
      $viewPaged       = $true
      $groupBy         = ''
      
      $viewQuery       = '<OrderBy><FieldRef Name="Title" /></OrderBy>'
    
      $newView = $splist.Views.Add($viewTitle, $viewFields, $($groupBy + $viewQuery), $viewRowLimit, $viewPaged, $viewDefaultView)
      $newView.ApplyStyle($spWeb.ViewStyles.StyleById(17)) # С затенением
      $newView.Update()
    
      
      
    

}
function New-SVTRestrictionList( $spweb, $splist_SVTID, $oWeb)
{
    # имя списка
    $NameOfListEng     = "SVTRestrict"
    $NameOfListRus     = "Допуск к СВТ ТУ"
    $ListDescriptionRus= "Перечень сотрудников допущенных к СВТ. Хранится информация о сотрудниках"
    
    # Новое еаименование поля Title

    $TitleNewName      = "Имя уч. записи в СЗИ от НСД"
    

    
    # создаем список
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # Меняем наименование списка
    $SpList = $spWeb.Lists[$NameOfListEng];
    
    $Listguid   = $SpList.GUID
    
    $Fieldtitle  = $title = $spList.Fields.GetFieldByInternalName("Title")
    $Fieldtitle.Required = $false
    $Fieldtitle.Update()
    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus

    
    # Меняем наименование поля Title
    Rename-ListFieldTitle $splist $TitleNewName

    # определяем поля
    $ListFields = @()
    
    
    
    
    $spFieldLookup = "Инвентарный номер СВТ" 
    $spListParent  = $spweb.Lists["Перечень СВТ"]
    $spFieldParent = "Инвентарный номер"
    # $spFieldOriginal = "tINVN"
    

    $ListFields += Define-LookupField $spListParent $spFieldParent $spFieldLookup 
    # ($spListParent, $spParentLookUpFieldNameRus, $spLookUpFieldNameRus)
    
    $spFieldLookup = "Ф.И.О. допущенного к СВТ"
    $spListParent  = $spweb.Lists["Список пользователей технологического участка"]
    $spFieldParent = "Сотрудник Ф.И.О."
    $spFieldParent = "Title"
    
    $ListFields += Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup 
    
    #$ListFields += '<Field Type="Text" DisplayName="Имя уч. записи в СЗИ от НСД" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="tUserNameSZINSD" Name="tUserNameSZINSD" ColName="tUserNameSZINSD" ><Default/></Field>'

    
    $spFieldLookup = "№ ТМ-идентификатора"
    $spListParent  = $spweb.Lists["Список ТМ идентификаторов"]
    $spFieldParent = "№ ТМ-идентификатора"
    $spFieldParent = "Название"
    $spFieldParent = "Title"
    #Define-LookupFieldTitle

    
    $ListFields += Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup 
    
    
    $ListFields += '<Field Type="Choice" DisplayName="Права доступа" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrAccessRigth" Name="bhrAccessRigth" ColName="bhrAccessRigth" ><Default>Пользователь</Default><CHOICES><CHOICE>Администратор</CHOICE><CHOICE>Рез. администратор</CHOICE><CHOICE>Пользователь</CHOICE></CHOICES></Field>'

    $spFieldLookup = "Назначение СВТ" 
    $spListParent  = $spweb.Lists["Перечень СВТ"]
    $spFieldParent = "Инвентарный номер"
    $spFieldDepended = "АС"
    
        
    $ListFields += Define-DependedLookupField   $spListParent $spFieldParent $spFieldDepended $spFieldLookup  
    # добавляем поля
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
      
      $viewFields.Add("Ф.И.О. допущенного к СВТ")
      $viewFields.Add("LinkTitle")
      $viewFields.Add("№ ТМ-идентификатора")
      $viewFields.Add("Права доступа")
      $viewFields.Add("Назначение СВТ")
      
      $viewRowLimit    = 300
      $viewDefaultView = $true
      $viewPaged       = $true
      $groupBy         = '<GroupBy Collapse="FALSE" GroupLimit="100"><FieldRef Name="'
      $groupBy        += $(Convert-ToInternalName "Инвентарный номер СВТ").Substring(0,32)   # почему-то так
      $groupBy        += '" /></GroupBy>'
      $viewQuery       = '<OrderBy><FieldRef Name="ID" /></OrderBy>'
      $newView = $splist.Views.Add($viewTitle, $viewFields, $($groupBy + $viewQuery), $viewRowLimit, $viewPaged, $viewDefaultView)
      
      $newView.ApplyStyle($spWeb.ViewStyles.StyleById(17)) # С затенением
      $newView.Update()

}

function Add-DocumentSetContentType($spSite,$splist)
{
    $spList.ContentTypesEnabled = $true
    $ctTpAdd = $spSite.RootWeb.ContentTypes["Набор документов"]
    $ctTpRemove=$spSite.RootWeb.ContentTypes["Документ"]
    $spList.ContentTypes.Add($ctTpAdd)
    $SpList.Update()  

}

function Change-ListOrder ($list)
{
	    $CTOrder = "Набор документов","Документ"
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
    $NameOfListRus     = "Перечень СВТ"
    $ListDescriptionRus= "Перечень СВТ. Набор документов для хранения данных об СВТ"


    $spSite = Get-SPSite $spWeb.Site.Url    

    # создаем Библиотеку документов
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "DocumentLibrary")
    
    $SpList = $spWeb.Lists[$NameOfListEng];
          
    Add-DocumentSetContentType $spSite $SpList

    
    $Listguid   = $SpList.Id

    # изменяем порядок типа контента content types
    Change-ListOrder $spList

    Remove-DocLibContentType  $spList "Документ"  

    $Listguid   = $SpList.GUID
    foreach($culture in $spweb.SupportedUiCultures)
    {


      [System.Threading.Thread]::CurrentThread.CurrentUICulture=$culture
      $SpList = $Web.Lists[$NameOfListEng]
      $SpList.Title = $NameOfListRus
      $SpList.Update()

    }
    $ListFields = @()
    


    $ListFields += '<Field Type="Text"   DisplayName="Инвентарный номер" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tINVN" Name="tINVN" ColName="tINVN" />'
    $ListFields += '<Field Type="Text"   DisplayName="Заводской номер" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tSN" Name="tSN" ColName="tSN" />'
    $ListFields += '<Field Type="Text"   DisplayName="Тип ПВМ" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tCompType" Name="tCompType" ColName="tCompType" />'    
    $ListFields += '<Field Type="Text"   DisplayName="АС" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tAS" Name="tAS" ColName="tAS" />'
    $ListFields += '<Field Type="Text"   DisplayName="Ф.И.О. Ответственного" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tFIO" Name="tFIO" ColName="tFIO" />'    
    $ListFields += '<Field Type="Choice" DisplayName="BBK" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" ColName="tBBK" ><Default>Антивирус Касперского 6.0</Default><CHOICES><CHOICE>Антивирус Касперского 6.0</CHOICE><CHOICE>Антивирус Касперского 8.0</CHOICE><CHOICE>Антивирус Касперского 10</CHOICE><CHOICE>Dr.Web</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Choice" DisplayName="Операционная система" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" StaticName="bhrOS" Name="bhrOS" ColName="bhrOS" ><Default>MS DOS</Default><CHOICES><CHOICE>MS DOS</CHOICE><CHOICE>Windows XP</CHOICE><CHOICE>Windows 7</CHOICE><CHOICE>Windows 8</CHOICE><CHOICE>Windows 2003</CHOICE><CHOICE>Windows 2008</CHOICE><CHOICE>Windows 2012</CHOICE></CHOICES></Field>'    
    $ListFields += '<Field Type="Choice" DisplayName="СЗИ от НСД" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrSZINSD" Name="bhrSZINSD" ColName="bhrSZINSD" ><Default>Нет</Default><CHOICES><CHOICE>Нет</CHOICE><CHOICE>Аккорд-Win32</CHOICE><CHOICE>Аккорд-Win64</CHOICE><CHOICE>Аккорд-АМДЗ</CHOICE><CHOICE>SecretNet</CHOICE><CHOICE>Соболь</CHOICE></CHOICES></Field>'
    $ListFields += '<Field Type="Text"   DisplayName="Инв. номер СЗИ" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="tSZIInvN" Name="tSZIInvN" ColName="tSZIInvN" />'
    $ListFields += '<Field Type="Text"   DisplayName="Заводской номер СЗИ" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" StaticName="tSZISN" Name="tSZISN" ColName="tSZISN" />'
    $ListFields += '<Field Type="Text"   DisplayName="Номер помещения" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255"  StaticName="tRoomNumber" Name="tRoomNumber" ColName="tRoomNumber" />' 
    $ListFields += '<Field Type="Calculated" DisplayName="ИД Комп"  ResultType="Text" ><Formula>=1</Formula></Field>'




    
    # добавляем поля
    foreach ($newColumn in $ListFields)
    {
          $disableoutput = $SpList.Fields.AddFieldAsXml($newColumn,$true,[Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView)
    }
    
    $fieldc = $splist.Fields["ИД Комп"]
    $fieldc.Formula = '=СЦЕПИТЬ(" [";[Инв. номер СЗИ];" ";[АС];" ")'
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
      $viewFields.Add("Инвентарный номер")
      $viewFields.Add("Номер помещения")    
      $viewFields.Add("Заводской номер")    
      $viewFields.Add("Тип ПВМ")   
      $viewFields.Add("Ф.И.О. Ответственного")   
      $viewFields.Add("АС")    
      $viewFields.Add("BBK")    
      
      $viewRowLimit    = 300
      $viewDefaultView = $true
      $viewPaged       = $true
      $viewQuery       = '<OrderBy><FieldRef Name="FileLeafRef" /></OrderBy>'
      $newView = $splist.Views.Add($viewTitle, $viewFields,  $viewQuery, $viewRowLimit, $viewPaged, $viewDefaultView)
      
      $newView.ApplyStyle($spWeb.ViewStyles.StyleById(17)) # С затенением
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
      $webTitle       = $oWeb.Title  + " № "+$($CurrentSiteIndex.ToString().PadLeft(2,"0"))
      $WebDescription = $oWeb.Description  + " № "+$($CurrentSiteIndex.ToString().PadLeft(2,"0"))
      
 
      $siteTemplate = "STS#1" # blank Site
      $webUrl=$webSite+$webHierarchy+$newWeb


      # $webUrl
      Remove-SPWeb $webUrl -Confirm:$false -ErrorAction SilentlyContinue 

      write-host $("Веб сайт : " + $webUrl)  -ForegroundColor Green
      write-host "Создание сайта технологического участка..." -nonewline
      # Создание веб-сайта
      $nooutput = new-spweb $webUrl -Template $siteTemplate  -Name $webTitle  -UseParentTopNav -Description $WebDescription # -Description $WebDescription

      $web = get-spweb -Identity $webUrl

      Change-Theme $webSite $web

      $web.Update()
      $web.Dispose()

      write-host "[OK]" -ForegroundColor Yellow
      $web = get-spweb -Identity $webUrl

      write-host "Создание списков и библиотек..." -nonewline
      #создание списков и библиотек документов
      $splistAdminsGuid = New-AdminsList $web $oWeb
     

      $splistUsersGuid = New-UsersList $web

      $splistRoomsGuid = New-RoomList  $web $oWeb


      $splistTMGuid = New-TMList $web 
      $splistSVTsGuid = New-DocumentSet $web $oWeb

      $splistSVTRestrictions = New-SVTRestrictionList $web $splistSVTsGuid $oWeb
      # "Перечень сотрудников допущенных к СВТ. Хранится информация о сотрудниках"
      $splistDailyReportGuid = New-DailyReport $web $oWeb
      
      $splistDReportLibGuid = New-DReportLib  $web $oWeb 
      
      $splistReports = New-ReportDocLib $web
      
      
      write-host "[OK]" -ForegroundColor Yellow

      write-host "Создание панели навигации..." -nonewline
      $web.QuickLaunchEnabled=$true

      Costumize-QuickLaunchMenu $oweb $CurrentSiteIndex $web
      write-host "[OK]" -ForegroundColor Yellow

      write-host "Обновление главной страницы..." -nonewline
      Update-WebParts $web $oWeb
      write-host "[OK]" -ForegroundColor Yellow
      $web.Update()
      write-host "Создан сайт технологического участка ...  "  -ForegroundColor Green
      write-host $($webSite + $webHierarchy + $newWeb)
      write-host
      write-host ===================================
      write-host Проверка Excel... -nonewline
      $ExcelExists = Test-Excel
      if ($ExcelExists)
      {
            write-host "[OK]" -ForegroundColor Yellow
            write-host "Создание отчета технологического участка ...  "  -ForegroundColor Yellow
            $reportName = "Report" + $CurrentSiteIndex.ToString() + ".xlsx"
            $ReportFileName = Create-Excel $reportName $oweb
            
            
            $spReportLib= $web.GetFolder("Reports")
            Write-host $("Помещаем файл "+$ReportFileName+ " в Библиотеку Отчетов.") 
            Add-DocumentToDocLib $spReportLib $ReportFileName "Паспорт ТУ.xlsx"
      }
      else
      {
            write-host "[ERROR]" -ForegroundColor Red
            
            write-host "На данном компьютере не установлен Excel версии 2010 или выше."  -ForegroundColor Yellow
            write-host $("После окончания работы скрипта сохраните каталог ") -ForegroundColor Yellow 
            
            write-host $oweb.iqyFolder -ForegroundColor Cyan 
           
            write-host $("со всеми подкаталогами на компьютере на котором установлен Excel") -ForegroundColor Yellow
            write-host $("и запустите из этого каталога процедуру подготовки сводных отчетов")  -ForegroundColor Yellow
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
    $NameOfListRus     = "Отчеты"
    $ListDescriptionRus= "Отчеты по технологическому участку"


    $spSite = Get-SPSite $spWeb.Site.Url    

    # создаем Библиотеку документов
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

# Инициализация
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
 
                    write-host "Скрипт по созданию Web сайтов технологических участков."
                    write-host "Данный скрипт создает Веб сайты в вашей инфраструктуре."
                    Write-host "Разработка данного скрипта проведена в соответствии с письмом ГУ по ЦФО "
                    Write-host "№ 19-2-18/178560 от 16.12.2016"
                    Write-host 

                    Write-host "-------------------------------------------------------------"
                    Write-host "По всем вопросам связанным с работой данного скрипта, наполнением веб сайтов "
                    write-host "необходимо обращаться в Отделение Калуга ГУ по ЦФО."
                    write-host "Сотрудник ОБиЗИ:Ларин Андрей Юрьевич,т.8-223-1261 - Методическое обеспечение;"
                    write-host "Сотрудник ОИ :Астахов Александр Борисович,т. 8-223-1338 - техническая поддержка."
                    write-host "-------------------------------------------------------------"
                    write-host "Для продолжения нажмите <ENTER>" -nonewline
                    read-host
                    write-host 
                    write-host "ВНИМАНИЕ!!!! Будут созданы следующие веб сайты:" -foreground green
                    write-host "Веб - Сайт Информационного модуля:" -foreground green
                    write-host  $($webSite + $webHierarchy + $infoModule) -foreground yellow
                    write-host 
                    write-host "Веб - Сайты Технологических участков:" -foreground green
                    
                    
                    for ($jj = 1; $jj -le $websiteCount; $jj++)
                    {

                        write-host  $($webSite + $webHierarchy + $newWeb + $jj.Tostring().Trim()) -foreground yellow
                    }

                    write-host "Если такие сайты уже существуют, они будут удалены, все данные на них будут уничтожены!" -foreground green
                    write-host "-------------------------------------------------------------"
                    write-host 
                    write-host "Сейчас будет сформирован " -foreground yellow -nonewline
                    write-host "Пин-код" -foreground cyan -nonewline
                    write-host ", который вы должны будете ввести," -foreground yellow 
                    write-host "Если Вы не уверены, что вы поступаете правильно, просто нажмите Enter," 
                    write-host "скрипт завершит работу, ничего не будет удалено и сайты не будут созданы."
                    write-host 
                    write-host "Ввод верного Пин-Кода будет означать Ваше согласие на создание сайтов."
                    write-host "-------------------------------------------------------------"




                    $randomToken = Get-Random -Maximum 9999 -Minimum 1

                    $Pin = $randomToken.ToString().PadLeft(4, "0")
                    Write-Host $("ПИН-Код :" + $Pin) -foreground cyan
                    Write-Host "введите ПИН-Код:" -nonewline
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
                        Write-Host  "Ничего не было создано. Хорошего дня!"

                    }
        }
        else
        {
                    write-host $("В инициализационном файле " +$($IniFile )+ " указаны параметры") -foreground cyan 
                    write-host $("Site="+$webSite) -foreground Yellow
                    write-host $("Hierarchy="+$webHierarchy) -foreground Yellow
                    write-host $("При попытке обращения к сайту "+$webSite+$webHierarchy+" произошла ошибка.") -foreground Yellow
                    write-host $("Данный сайт не существует или нет прав доступа к этому сайту. ") -foreground cyan 
        }     
        }
        
        else
        {
             write-host $("Обнаружены неверные параметры...")
        }
}
else
{
        Write-host
        Write-Host $("Внимание!  Не был найден инициализационный файл.") -foreground cyan
        Write-host $("Процедура автоматически создала инициализационный файл с именем : ") -foreground cyan 
        Write-host $($IniFile ) -foreground Yellow -nonewline
        Write-host $(" и параметрами по умолчанию.") -foreground cyan 
        Write-host
        Write-host
        Write-Host $("Проверьте параметры данного файла и запустите процедуру еще раз.") -foreground cyan
        
}

Stop-Transcript

