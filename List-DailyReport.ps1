function New-DailyReport( $spweb, $oWeb)
{
    # ��� ������
    $NameOfListEng     = "DailyReport"
    $NameOfListRus     = "���������� ����� ���-�� ���������������� �������"
    $ListDescriptionRus= "���������� ����� ���-�� ���������������� �������. �������� ���������� � �������� ���-�� ���.��������"
    
    # ����� ������������ ���� Title
    $TitleNewName      = "�������"
    

    
    # ������� ������
    
    $spweb.Lists.Add($NameOfListEng, $ListDescriptionRus, "GenericList")
    
        
    # ������ ������������ ������
    $SpList = $spWeb.Lists[$NameOfListEng];
    
    $Listguid   = $SpList.Id
    
    
    Rename-ListNameRus $spWeb $NameOfListEng $NameOfListRus

    $TitleDefault = "������� ����������� ������ ��������� ������ � ��������, ��� � �������� �.�."
    # ������ ������������ ���� Title
    Rename-ListFieldTitle $splist $TitleNewName $TitleDefault

    # ���������� ����
    $ListFields = @()
    

    #     ����� ����
    $ListFields += '<Field Type="Choice" DisplayName="����� ����" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"  StaticName="bhrReportInp" ColName="bhrReportInp" RowOrdinal="0" Version="2"><Default>���������� �.�.</Default><CHOICES><CHOICE>���������� �.�.</CHOICE><CHOICE>�������� �.�.</CHOICE><CHOICE>�������� �.�.</CHOICE></CHOICES></Field>'

    #     ���� ������
    $ListFields += '<Field Type="DateTime" DisplayName="���� ������" Format="DateOnly" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" StaticName="dReportDate" Name="dReportDate" ColName="dReportDate" RowOrdinal="0" Version="1"><Default>[today]</Default></Field>'
 
    $spFieldLookup = "���������� � ������ �����"
    $spListParent  = $spweb.Lists["������ ������������� ���������������� �������"]
    #$spFieldParent = "��������� �.�.�."
    #$spFieldParent = "��������"
    $spFieldParent = "Title"
    $requiried = "TRUE"
    $Notrequiried = "FALSE"
    $MultiChoice = "TRUE"
    #Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup "TRUE" "TRUE"   

    
    #     ���������� � ������ �����
    
         
    $ListFields +=  Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup $requiried $MultiChoice        
    
    
    
    
    
    #     ���������� �� ������ �����
    $spFieldLookup = "���������� �� ������ �����"
    $ListFields += Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup $Notrequiried $MultiChoice       
    #$ListFields += '<Field Type="LookupMulti" DisplayName="���������� �� ������ �����" Required="FALSE" EnforceUniqueValues="FALSE" List="{1154b376-e88b-47e6-a019-d0f5faec4277}" ShowField="Title" Mult="TRUE" Sortable="FALSE" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{a0eaa242-2faa-42cb-bc2f-fc202a0bdb1a}" SourceID="{f5f157e0-3c79-4e42-9ced-bb41e188b8fb}" StaticName="_x041d__x0430__x0020__x0440__x040" Name="_x041d__x0430__x0020__x0440__x040" ColName="int2" RowOrdinal="0" Version="2" Group=""/>'

    #     ����������� �� ������� �����
    $spFieldLookup = "����������� �� ������� �����"
    $ListFields += Define-LookupFieldTitle $spListParent $spFieldParent $spFieldLookup $Notrequiried $MultiChoice        

    $spFieldLookup = "�������� ��� �������� ���"
    $spListParent  = $spweb.Lists["�������� ���"]    
    $spFieldParent = "�� ����"
    $ListFields += Define-LookupFieldMult $spListParent $spFieldParent $spFieldLookup $Notrequiried $MultiChoice
    
    $spFieldLookup = "�������� ��� �������� ���. ���"
    $ListFields += Define-LookupFieldMult $spListParent $spFieldParent $spFieldLookup $Notrequiried $MultiChoice
        
   
    #     �������� ��� �������� ���
    #$ListFields += '<Field Type="LookupMulti" DisplayName="�������� ��� �������� ���" Required="FALSE" EnforceUniqueValues="FALSE" List="{09c23b6f-85a1-42b8-bacd-141af7059dff}" ShowField="_x0421__x0412__x0422__x0020__x04" Mult="TRUE" Sortable="FALSE" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{e847c88a-cb61-46de-accd-00beead6f67a}" SourceID="{f5f157e0-3c79-4e42-9ced-bb41e188b8fb}" StaticName="_x041a__x043e__x043d__x0442__x04" Name="_x041a__x043e__x043d__x0442__x04" ColName="int7" RowOrdinal="0" Version="6" Group="" Description="���. ����� ���"/>'

    #     �������� ��� �������� ���. ���
    #$ListFields += '<Field Type="LookupMulti" DisplayName="�������� ��� �������� ���. ���" Required="FALSE" EnforceUniqueValues="FALSE" List="{09c23b6f-85a1-42b8-bacd-141af7059dff}" ShowField="_x0421__x0412__x0422__x0020__x04" Mult="TRUE" Sortable="FALSE" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{19b7fe49-f251-4ced-9f5c-293bb5e4b469}" SourceID="{f5f157e0-3c79-4e42-9ced-bb41e188b8fb}" StaticName="_x041a__x043e__x043d__x0442__x040" Name="_x041a__x043e__x043d__x0442__x040" ColName="int8" RowOrdinal="0" Version="5" Group="" Description="���. ����� ���."/>'
     
    #     ������������ ���������
    $ListFields += '<Field Type="Choice" DisplayName="������������ ���������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" StaticName="bhrDfus" ColName="bhrDfus" RowOrdinal="0" Version="1"><Default>���������</Default><CHOICES><CHOICE>���������</CHOICE><CHOICE>�� ���������</CHOICE></CHOICES></Field>'

    #     ������  1 �� ���������� � ��-����������������
    $ListFields += '<Field Type="Choice" DisplayName="������  1 �� ���������� � ��-����������������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" StaticName="bhrBox1" Name="bhrBox1" ColName="bhrBox1" RowOrdinal="0" Version="1"><Default>��������</Default><CHOICES><CHOICE>��������</CHOICE><CHOICE>���������</CHOICE></CHOICES></Field>'

    #     ����� 1 ������ ���������� �����
    $ListFields += '<Field Type="Number" DisplayName="����� 1 ������ ���������� �����" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE"  StaticName="nNomPech1" Name="nNomPech1" ColName="nNomPech1" RowOrdinal="0" Version="1"/>'

    #     ���� ������������ 1
    $ListFields += '<Field Type="DateTime" DisplayName="���� ������������ 1" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly" StaticName="dOpech1" Name="dOpech1" ColName="dOpech1" RowOrdinal="0" Version="1"/>'

    #     ������  2 �� ���������� � ��-����������������
    $ListFields += '<Field Type="Choice" DisplayName="������  2 �� ���������� � ��-����������������" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE"   StaticName="bhrBox2" Name="bhrBox2" ColName="bhrBox2" RowOrdinal="0" Version="1"><Default>��������</Default><CHOICES><CHOICE>��������</CHOICE><CHOICE>���������</CHOICE></CHOICES></Field>'

    #     ����� 2 ������ ���������� �����
    $ListFields += '<Field Type="Number" DisplayName="����� 2 ������ ���������� �����" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" StaticName="nNomPech2" Name="nNomPech2" ColName="nNomPech2" RowOrdinal="0" Version="1"/>'

    #     ���� ������������ 2
    $ListFields += '<Field Type="DateTime" DisplayName="���� ������������ 2" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly"  StaticName="dOpech1" Name="dOpech1" ColName="dOpech1" RowOrdinal="0" Version="1"/>'
     
     
    
    $formula = '=1' 
    #     �����
    $cf = ""
    $cf += '<Field Type="Calculated" DisplayName="�����"   ResultType="Text"  >'
    # LCID="1049" EnforceUniqueValues="FALSE" Indexed="FALSE" RowOrdinal="0" Required="FALSE" ReadOnly="TRUE" 
    $cf +=  "<Formula>" + $formula + "</Formula></Field>"

   
    $ListFields += $cf
    
 

    
    # ��������� ����
    foreach ($newColumn in $ListFields)
    {
          $disableoutput = $SpList.Fields.AddFieldAsXml($newColumn,$true,[Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView)
    }
    
    $SpList.Update();
    
    
    # update formula
    $fieldc = $splist.Fields["�����"]
    $fieldc.Formula = '=���([���� ������])&"-"&����(�����([���� ������])<10;"0";"")&�����([���� ������])&"-"&����(����([���� ������])<10;"0";"")&����([���� ������])&"-"&[����� ����]'
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
      $viewFields.Add("�����")
      $viewFields.Add("LinkTitle")
      
      $viewRowLimit    = 300
      $viewDefaultView = $true
      $viewPaged       = $true
      $groupBy         = ''
      
      $viewQuery       = '<OrderBy><FieldRef Name="'
      $viewQuery        += $(Convert-ToInternalName "�����").Substring(0,32)   # ������-�� ���
      $viewQuery        += '" /></OrderBy>'
    
      $newView = $splist.Views.Add($viewTitle, $viewFields, $($groupBy + $viewQuery), $viewRowLimit, $viewPaged, $viewDefaultView)
      $newView.ApplyStyle($spWeb.ViewStyles.StyleById(17)) # � ����������
      $newView.Update()
    

}
function New-DReportLib( $spweb, $oWeb)
{


    $NameOfListEng     = "DReports"
    $NameOfListRus     = "���������� ������ ��� � ������� MSWORD"
    $ListDescriptionRus= "���������� ������ � ������� MSWORD. ��������� ����������� �������������."


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
    $ListFields = @()
    
    
    

    <#
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
    #>
    
        #     �������
        Rename-ListFieldTitle $splist "�������"
          #$ListFields += '<Field ID="{fa564e0f-0c70-4ab9-b863-0177e6ddd247}" Type="Text" Name="Title" ShowInNewForm="FALSE" ShowInFileDlg="FALSE" DisplayName="�������" Sealed="TRUE" SourceID="http://schemas.microsoft.com/sharepoint/v3" StaticName="Title" ColName="nvarchar7" Description=""/>'

          #     ���� ������
          #$ListFields += '<Field Type="Text" DisplayName="���� ������" />'
          #$ListFields += '<Field Type="Text" DisplayName="Dreport" />' e46368ba-fa3d-4bee-8667-9afad25505d6
          $ListFields += "Text,Dreport,���� ������,e46368ba-fa3d-4bee-8667-9afad25505d6"

          #     ��� ���������� �������� �� ������ �����
          #$ListFields += '<Field Type="Note" DisplayName="��� ���������� �������� �� ������ �����"  RichTextMode="Compatible" />'
          #$ListFields += '<Field Type="Note" DisplayName="PersonellInSM2"  RichTextMode="Compatible" />'
          $ListFields += "Note,PersonellInSM2,��� ���������� �������� �� ������ �����,40a0e54b-e3b1-4d91-a026-7cb508e65d3c"

          #     ��� ���������� �������� � ������ �����
          #$ListFields += '<Field Type="Note" DisplayName="��� ���������� �������� � ������ �����"  RichTextMode="Compatible" />'
          #$ListFields += '<Field Type="Note" DisplayName="PersonellInSM1"  RichTextMode="Compatible" />'
          $ListFields += "Note,PersonellInSM1,��� ���������� �������� � ������ �����,7991c397-adc2-439f-b796-89a1773f635c"

          #     ��� ���������� ������������
          #$ListFields += '<Field Type="Note" DisplayName="��� ���������� ������������" />'
          #$ListFields += '<Field Type="Note" DisplayName="PersonellOnWork" />'
          $ListFields += "Note,PersonellOnWork,��� ���������� ������������,e670f38d-1e0f-430e-abdd-0720c316adab"

          #     �������� ��� �������� ���. ���
          #$ListFields += '<Field Type="Note" DisplayName="�������� ��� �������� ���. ���" />'
          #$ListFields += '<Field Type="Note" DisplayName="ControlByRezAIB" />'
          $ListFields += "Note,ControlByRezAIB,�������� ��� �������� ���. ���,c1414425-e607-453e-946e-0e9734147e7d"

          #     �������� ��� �������� ���
          #$ListFields += '<Field Type="Note" DisplayName="�������� ��� �������� ���" />'
          #$ListFields += '<Field Type="Note" DisplayName="ControlByAIB" />'
          $ListFields += "Note,ControlByAIB,�������� ��� �������� ���,83d94c8d-ab37-4da1-8785-afc74ef77af9"

          #     ������������ ���������
          #$ListFields += '<Field Type="Text" DisplayName="������������ ���������" />'
          #$ListFields += '<Field Type="Text" DisplayName="RoomStamped" />'
          $ListFields += "Text,RoomStamped,������������ ���������,342830c5-a313-4e4b-8da8-82a49ec7ca49"

          #     ������ 1 �� ���������� � ��-����������������
          #$ListFields += '<Field Type="Text" DisplayName="������ 1 �� ���������� � ��-����������������" />'
          #$ListFields += '<Field Type="Text" DisplayName="StampOnBox1" />'
          $ListFields += "Text,StampOnBox1,������ 1 �� ���������� � ��-����������������,51174348-075e-4066-a29b-f33f6a1b0ded"

          #     ����� 1 ������ ���������� �����
          #$ListFields += '<Field Type="Text" DisplayName="����� 1 ������ ���������� �����" />'
          #$ListFields += '<Field Type="Text" DisplayName="StampNumb1" />'
          $ListFields += "Text,StampNumb1,����� 1 ������ ���������� �����,aa5b23e2-1dc2-4f81-80d1-84ea4fb3243e"

          #     ���� ������������ 1
          #$ListFields += '<Field Type="Text" DisplayName="���� ������������ 1" />'
          #$ListFields += '<Field Type="Text" DisplayName="DateStamped1" />'
          $ListFields += "Text,DateStamped1,���� ������������ 1,2258b0ff-89d8-44dd-95f9-b919e941e6d2"

          #     ������ 2 �� ���������� � ��-����������������
          #$ListFields += '<Field Type="Text" DisplayName="������ 2 �� ���������� � ��-����������������" />'
          #$ListFields += '<Field Type="Text" DisplayName="StampOnBox2" />'
          $ListFields += "Text,StampOnBox2,������ 2 �� ���������� � ��-����������������,b6eedd2e-22d4-41fa-a231-8b0a99d26f11"

          #     ����� 2 ������ ���������� �����
          #$ListFields += '<Field Type="Text" DisplayName="����� 2 ������ ���������� �����" />'
          #$ListFields += '<Field Type="Text" DisplayName="StampNumb2" />'
          $ListFields += "Text,StampNumb2,����� 2 ������ ���������� �����,ecbfde5d-b9d5-421e-90a9-260ea94a11fd"

          #     ���� ������������ 2
          #$ListFields += '<Field Type="Text" DisplayName="���� ������������ 2" />'
          #$ListFields += '<Field Type="Text" DisplayName="DateStamped2" />'
          $ListFields += "Text,DateStamped2,���� ������������ 2,9b22a37e-c06f-4526-b5a7-7d48657e4dfe"

          #     ��
          #$ListFields += '<Field Type="Text" DisplayName="��"  Description="���� �������� ���������"/>'
          #$ListFields += '<Field Type="Text" DisplayName="UDocDate"  Description="���� �������� ���������"/>'
          $ListFields += "Text,UDocDate,���� �������� ���������,63f9fa67-6369-4c7e-9e5e-5c746faaef93"
          
           #$ListFields += "Text,CustomField,CustomField"

          #     ��
          #$ListFields += '<Field Type="Text" DisplayName="��" Description="��������� ������� ������ ��������" />'
          #$ListFields += '<Field Type="Text" DisplayName="UDocPers" Description="��������� ������� ������ ��������" />'
          $ListFields += "Text,UDocPers,��������� ������� ������ ��������,efc2f184-dc24-4b3c-adf3-b8e97775ab19"

    
    # ��������� ����
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

