function New-DailyReport( $spweb, $oWeb)
{
    # имя списка
    $NameOfListEng     = "DailyReport"
    $NameOfListRus     = "Ежедневный отчет АИБ-ов технологического участка"
    $ListDescriptionRus= "Ежедневный отчет АИБ-ов технологического участка. Хранится информация о действия АИБ-ов тех.участков"
    
    # Новое еаименование поля Title
    $TitleNewName      = "Участок"
    

    
    # создаем список
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # Меняем наименование списка
    $SpList = $spWeb.Lists[$NameOfListEng];
    
    $Listguid   = $SpList.Id
    
    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus

    $TitleDefault = "Участок «Отчетности отдела платежных систем и расчетов», АИБ – Сергеева Л.Г."
    # Меняем наименование поля Title
    Rename-ListFieldTitle $splist $TitleNewName $TitleDefault

    # определяем поля
    $ListFields = @()
    

    #     Отчет ввел
    $ListFields += '<Field Type="Choice" DisplayName="Отчет ввел" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrReportInp" ColName="bhrReportInp" RowOrdinal="0" Version="2"><Default>Введенская Л.Н.</Default><CHOICES><CHOICE>Введенская Л.Н.</CHOICE><CHOICE>Комонова О.Н.</CHOICE><CHOICE>Сергеева Л.Г.</CHOICE></CHOICES></Field>'

    #     Дата отчета
    $ListFields += '<Field Type="DateTime" DisplayName="Дата отчета" Format="DateOnly" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" StaticName="dReportDate" Name="dReportDate" ColName="dReportDate" RowOrdinal="0" Version="1"><Default>[today]</Default></Field>'
 
    $spFieldLookup = "Сотрудники в первую смену"
    $spListParent  = $spweb.Lists["Список пользователей технологического участка"]
    #$spFieldParent = "Сотрудник Ф.И.О."
    #$spFieldParent = "Название"
    $spFieldParent = "Title"
    $requiried = "TRUE"
    $Notrequiried = "FALSE"
    $MultiChoice = "TRUE"
    #Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup "TRUE" "TRUE"   

    
    #     Сотрудники в первую смену
    
         
    $ListFields +=  Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup $requiried $MultiChoice        
    
    
    
    
    
    #     Сотрудники во вторую смену
    $spFieldLookup = "Сотрудники во вторую смену"
    $ListFields += Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup $Notrequiried $MultiChoice       
    #$ListFields += '<Field Type="LookupMulti" DisplayName="Сотрудники во вторую смену" Required="FALSE" EnforceUniqueValues="FALSE" List="{1154b376-e88b-47e6-a019-d0f5faec4277}" ShowField="Title" Mult="TRUE" Sortable="FALSE" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{a0eaa242-2faa-42cb-bc2f-fc202a0bdb1a}" SourceID="{f5f157e0-3c79-4e42-9ced-bb41e188b8fb}" StaticName="_x041d__x0430__x0020__x0440__x040" Name="_x041d__x0430__x0020__x0440__x040" ColName="int2" RowOrdinal="0" Version="2" Group=""/>'

    #     Отсутствуют на рабочем месте
    $spFieldLookup = "Отсутствуют на рабочем месте"
    $ListFields += Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup $Notrequiried $MultiChoice        

    $spFieldLookup = "Контроль СВТ выполнил АИБ"
    $spListParent  = $spweb.Lists["Перечень СВТ"]    
    $spFieldParent = "ИД Комп"
    $ListFields += Define-LookupFieldMult $spListParent $spFieldParent $spFieldLookup $Notrequiried $MultiChoice
    
    $spFieldLookup = "Контроль СВТ выполнил рез. АИБ"
    $ListFields += Define-LookupFieldMult $spListParent $spFieldParent $spFieldLookup $Notrequiried $MultiChoice
        
   
    #     Контроль СВТ выполнил АИБ
    #$ListFields += '<Field Type="LookupMulti" DisplayName="Контроль СВТ выполнил АИБ" Required="FALSE" EnforceUniqueValues="FALSE" List="{09c23b6f-85a1-42b8-bacd-141af7059dff}" ShowField="_x0421__x0412__x0422__x0020__x04" Mult="TRUE" Sortable="FALSE" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{e847c88a-cb61-46de-accd-00beead6f67a}" SourceID="{f5f157e0-3c79-4e42-9ced-bb41e188b8fb}" StaticName="_x041a__x043e__x043d__x0442__x04" Name="_x041a__x043e__x043d__x0442__x04" ColName="int7" RowOrdinal="0" Version="6" Group="" Description="Инв. номер СВТ"/>'

    #     Контроль СВТ выполнил рез. АИБ
    #$ListFields += '<Field Type="LookupMulti" DisplayName="Контроль СВТ выполнил рез. АИБ" Required="FALSE" EnforceUniqueValues="FALSE" List="{09c23b6f-85a1-42b8-bacd-141af7059dff}" ShowField="_x0421__x0412__x0422__x0020__x04" Mult="TRUE" Sortable="FALSE" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{19b7fe49-f251-4ced-9f5c-293bb5e4b469}" SourceID="{f5f157e0-3c79-4e42-9ced-bb41e188b8fb}" StaticName="_x041a__x043e__x043d__x0442__x040" Name="_x041a__x043e__x043d__x0442__x040" ColName="int8" RowOrdinal="0" Version="5" Group="" Description="Инв. номер СВТ."/>'
     
    #     Опечатывание помещения
    $ListFields += '<Field Type="Choice" DisplayName="Опечатывание помещения" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" StaticName="bhrDfus" ColName="bhrDfus" RowOrdinal="0" Version="1"><Default>Выполнено</Default><CHOICES><CHOICE>Выполнено</CHOICE><CHOICE>Не выполнено</CHOICE></CHOICES></Field>'

    #     Печать  1 на контейнере с ТМ-идентификаторами
    $ListFields += '<Field Type="Choice" DisplayName="Печать  1 на контейнере с ТМ-идентификаторами" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" StaticName="bhrBox1" Name="bhrBox1" ColName="bhrBox1" RowOrdinal="0" Version="1"><Default>Основной</Default><CHOICES><CHOICE>Основной</CHOICE><CHOICE>Резервный</CHOICE></CHOICES></Field>'

    #     Номер 1 печати сотрудника ОБиЗИ
    $ListFields += '<Field Type="Number" DisplayName="Номер 1 печати сотрудника ОБиЗИ" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE"  StaticName="nNomPech1" Name="nNomPech1" ColName="nNomPech1" RowOrdinal="0" Version="1"/>'

    #     Дата опечатывания 1
    $ListFields += '<Field Type="DateTime" DisplayName="Дата опечатывания 1" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly" StaticName="dOpech1" Name="dOpech1" ColName="dOpech1" RowOrdinal="0" Version="1"/>'

    #     Печать  2 на контейнере с ТМ-идентификаторами
    $ListFields += '<Field Type="Choice" DisplayName="Печать  2 на контейнере с ТМ-идентификаторами" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"   StaticName="bhrBox2" Name="bhrBox2" ColName="bhrBox2" RowOrdinal="0" Version="1"><Default>Основной</Default><CHOICES><CHOICE>Основной</CHOICE><CHOICE>Резервный</CHOICE></CHOICES></Field>'

    #     Номер 2 печати сотрудника ОБиЗИ
    $ListFields += '<Field Type="Number" DisplayName="Номер 2 печати сотрудника ОБиЗИ" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" StaticName="nNomPech2" Name="nNomPech2" ColName="nNomPech2" RowOrdinal="0" Version="1"/>'

    #     Дата опечатывания 2
    $ListFields += '<Field Type="DateTime" DisplayName="Дата опечатывания 2" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly"  StaticName="dOpech1" Name="dOpech1" ColName="dOpech1" RowOrdinal="0" Version="1"/>'
     
     
    
    $formula = '=1' 
    #     ОТЧЕТ
    $cf = ""
    $cf += '<Field Type="Calculated" DisplayName="ОТЧЕТ"   ResultType="Text"  >'
    # LCID="1049" EnforceUniqueValues="FALSE" Indexed="FALSE" RowOrdinal="0" Required="FALSE" ReadOnly="TRUE" 
    $cf +=  "<Formula>" + $formula + "</Formula></Field>"

   
    $ListFields += $cf
    
 

    
    # добавляем поля
    foreach ($newColumn in $ListFields)
    {
          $disableoutput = $SpList.Fields.AddFieldAsXml($newColumn,$true,[Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView)
    }
    
    $SpList.Update();
    
    
    # update formula
    $fieldc = $splist.Fields["ОТЧЕТ"]
    $fieldc.Formula = '=ГОД([Дата отчета])&"-"&ЕСЛИ(МЕСЯЦ([Дата отчета])<10;"0";"")&МЕСЯЦ([Дата отчета])&"-"&ЕСЛИ(ДЕНЬ([Дата отчета])<10;"0";"")&ДЕНЬ([Дата отчета])&"-"&[Отчет ввел]'
    $fieldc.Update()
    
    $spweb.Update();
    $spweb.Dispose();
    
    <#
    Create-IQY $spList $oWeb $NameOfListEng
    Add-UsersSampleData $SpList
    Change-RoomsDefaultView  $spweb $spList
    #>
    Change-DailyReportDefaultView $spweb $spList
    
    return $Listguid
    
}

function  Change-DailyReportDefaultView( $spweb, $spList)
{

      
      $viewTitle  = $splist.Title

      $viewFields = New-Object System.Collections.Specialized.StringCollection
      

      
      
      $viewFields.Add("Edit")
      $viewFields.Add("ОТЧЕТ")
      $viewFields.Add("LinkTitle")
      
      $viewRowLimit    = 300
      $viewDefaultView = $true
      $viewPaged       = $true
      $groupBy         = ''
      
      $viewQuery       = '<OrderBy><FieldRef Name="'
      $viewQuery        += $(Convert-ToInternalName "ОТЧЕТ").Substring(0,32)   # почему-то так
      $viewQuery        += '" /></OrderBy>'
    
      $newView = $splist.Views.Add($viewTitle, $viewFields, $($groupBy + $viewQuery), $viewRowLimit, $viewPaged, $viewDefaultView)
      $newView.ApplyStyle($spWeb.ViewStyles.StyleById(17)) # С затенением
      $newView.Update()
    

}
function New-DReportLib( $spweb, $oWeb)
{


    $NameOfListEng     = "DReports"
    $NameOfListRus     = "Ежедневные отчеты АИБ в формате MSWORD"
    $ListDescriptionRus= "Ежедневные отчеты в формате MSWORD. Документы формируются автоматически."


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
    $ListFields = @()
    
    
    

    <#
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
    #>
    
        #     Участок
        Rename-ListFieldTitle $splist "Участок"
          #$ListFields += '<Field ID="{fa564e0f-0c70-4ab9-b863-0177e6ddd247}" Type="Text" Name="Title" ShowInNewForm="FALSE" ShowInFileDlg="FALSE" DisplayName="Участок" Sealed="TRUE" SourceID="http://schemas.microsoft.com/sharepoint/v3" StaticName="Title" ColName="nvarchar7" Description=""/>'

          #     Дата отчета
          #$ListFields += '<Field Type="Text" DisplayName="Дата отчета" />'
          #$ListFields += '<Field Type="Text" DisplayName="Dreport" />' e46368ba-fa3d-4bee-8667-9afad25505d6
          $ListFields += "Text,Dreport,Дата отчета,e46368ba-fa3d-4bee-8667-9afad25505d6"

          #     Все сотрудники работают во вторую смену
          #$ListFields += '<Field Type="Note" DisplayName="Все сотрудники работают во вторую смену"  RichTextMode="Compatible" />'
          #$ListFields += '<Field Type="Note" DisplayName="PersonellInSM2"  RichTextMode="Compatible" />'
          $ListFields += "Note,PersonellInSM2,Все сотрудники работают во вторую смену,40a0e54b-e3b1-4d91-a026-7cb508e65d3c"

          #     Все сотрудники работают в первую смену
          #$ListFields += '<Field Type="Note" DisplayName="Все сотрудники работают в первую смену"  RichTextMode="Compatible" />'
          #$ListFields += '<Field Type="Note" DisplayName="PersonellInSM1"  RichTextMode="Compatible" />'
          $ListFields += "Note,PersonellInSM1,Все сотрудники работают в первую смену,7991c397-adc2-439f-b796-89a1773f635c"

          #     Все сотрудники присутствуют
          #$ListFields += '<Field Type="Note" DisplayName="Все сотрудники присутствуют" />'
          #$ListFields += '<Field Type="Note" DisplayName="PersonellOnWork" />'
          $ListFields += "Note,PersonellOnWork,Все сотрудники присутствуют,e670f38d-1e0f-430e-abdd-0720c316adab"

          #     Контроль СВТ выполнил рез. АИБ
          #$ListFields += '<Field Type="Note" DisplayName="Контроль СВТ выполнил рез. АИБ" />'
          #$ListFields += '<Field Type="Note" DisplayName="ControlByRezAIB" />'
          $ListFields += "Note,ControlByRezAIB,Контроль СВТ выполнил рез. АИБ,c1414425-e607-453e-946e-0e9734147e7d"

          #     Контроль СВТ выполнил АИБ
          #$ListFields += '<Field Type="Note" DisplayName="Контроль СВТ выполнил АИБ" />'
          #$ListFields += '<Field Type="Note" DisplayName="ControlByAIB" />'
          $ListFields += "Note,ControlByAIB,Контроль СВТ выполнил АИБ,83d94c8d-ab37-4da1-8785-afc74ef77af9"

          #     Опечатывание помещения
          #$ListFields += '<Field Type="Text" DisplayName="Опечатывание помещения" />'
          #$ListFields += '<Field Type="Text" DisplayName="RoomStamped" />'
          $ListFields += "Text,RoomStamped,Опечатывание помещения,342830c5-a313-4e4b-8da8-82a49ec7ca49"

          #     Печать 1 на контейнере с ТМ-идентификаторами
          #$ListFields += '<Field Type="Text" DisplayName="Печать 1 на контейнере с ТМ-идентификаторами" />'
          #$ListFields += '<Field Type="Text" DisplayName="StampOnBox1" />'
          $ListFields += "Text,StampOnBox1,Печать 1 на контейнере с ТМ-идентификаторами,51174348-075e-4066-a29b-f33f6a1b0ded"

          #     Номер 1 печати сотрудника ОБиЗИ
          #$ListFields += '<Field Type="Text" DisplayName="Номер 1 печати сотрудника ОБиЗИ" />'
          #$ListFields += '<Field Type="Text" DisplayName="StampNumb1" />'
          $ListFields += "Text,StampNumb1,Номер 1 печати сотрудника ОБиЗИ,aa5b23e2-1dc2-4f81-80d1-84ea4fb3243e"

          #     Дата опечатывания 1
          #$ListFields += '<Field Type="Text" DisplayName="Дата опечатывания 1" />'
          #$ListFields += '<Field Type="Text" DisplayName="DateStamped1" />'
          $ListFields += "Text,DateStamped1,Дата опечатывания 1,2258b0ff-89d8-44dd-95f9-b919e941e6d2"

          #     Печать 2 на контейнере с ТМ-идентификаторами
          #$ListFields += '<Field Type="Text" DisplayName="Печать 2 на контейнере с ТМ-идентификаторами" />'
          #$ListFields += '<Field Type="Text" DisplayName="StampOnBox2" />'
          $ListFields += "Text,StampOnBox2,Печать 2 на контейнере с ТМ-идентификаторами,b6eedd2e-22d4-41fa-a231-8b0a99d26f11"

          #     Номер 2 печати сотрудника ОБиЗИ
          #$ListFields += '<Field Type="Text" DisplayName="Номер 2 печати сотрудника ОБиЗИ" />'
          #$ListFields += '<Field Type="Text" DisplayName="StampNumb2" />'
          $ListFields += "Text,StampNumb2,Номер 2 печати сотрудника ОБиЗИ,ecbfde5d-b9d5-421e-90a9-260ea94a11fd"

          #     Дата опечатывания 2
          #$ListFields += '<Field Type="Text" DisplayName="Дата опечатывания 2" />'
          #$ListFields += '<Field Type="Text" DisplayName="DateStamped2" />'
          $ListFields += "Text,DateStamped2,Дата опечатывания 2,9b22a37e-c06f-4526-b5a7-7d48657e4dfe"

          #     ДС
          #$ListFields += '<Field Type="Text" DisplayName="ДС"  Description="Дата создания документа"/>'
          #$ListFields += '<Field Type="Text" DisplayName="UDocDate"  Description="Дата создания документа"/>'
          $ListFields += "Text,UDocDate,Дата создания документа,63f9fa67-6369-4c7e-9e5e-5c746faaef93"
          
           #$ListFields += "Text,CustomField,CustomField"

          #     СС
          #$ListFields += '<Field Type="Text" DisplayName="СС" Description="Сотрудник который создал документ" />'
          #$ListFields += '<Field Type="Text" DisplayName="UDocPers" Description="Сотрудник который создал документ" />'
          $ListFields += "Text,UDocPers,Сотрудник который создал документ,efc2f184-dc24-4b3c-adf3-b8e97775ab19"

    
    # добавляем поля
    foreach ($newColumn in $ListFields)
    {
          $f  = '<Field '
          $f +='Type="'
          $f += $newColumn.Split(",")[0]
          $f += '" DisplayName="'
          $f += $newColumn.Split(",")[1]
          $f += '" Description="'
          $f += $newColumn.Split(",")[2]
          $f += '" ID="{'
          $f += $newColumn.Split(",")[3]
          $f += '}" '
          $f += ' />'
          $disableoutput = $SpList.Fields.AddFieldAsXml($f,$true,[Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView)
          Rename-ListField $splist $newColumn.Split(",")[1] $newColumn.Split(",")[2]
    }
    
    $templateName = $oWeb.WrkDir + "\Dreport.docx"
    Change-DoclibTemplate  $splist $templateName
    $spweb.Dispose()
    
    
    <#
    Create-IQY $spList $oWeb $NameOfListEng
    Add-SVTsampleData $SpList
    Change-DocumentsetDefaultView $spweb  $splist
    #>
    return $Listguid
        
       
}

