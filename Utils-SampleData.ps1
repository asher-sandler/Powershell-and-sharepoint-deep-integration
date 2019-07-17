##############################################
#
#   ������ �������� ������ ��������������� ��������
#
#   ����: Util-SampleData.ps1 
#   �������� : ������� ��� ���������� ��������� ������� 
#              ���������������� �������.
#
#   �����: ������� ��������� ���������, ����� ��������������, ���: 223-1338
#          29AstahovAB@cbr.ru
#
#   ��������� ������, �� ����� ������ �� ���
#   ������� ����������: ������ ������ ����������� � ������� ������ �������. 
#                       �� ������� ��������� ��� ��������.
#   ����: ������ 2017.
#   ������: 17.01.11
#
##############################################

function Add-TuUsersSampleData($splist)
{

$list = @()

$list += "������ �.�."
$list += "������ �.�."
$list += "��������� �.�."
$list += "�������� �.�."
$list += "�������� �.�."
$list += "��������� �.�."
$list += "�������� �.�."
$list += "����������� �.�."

foreach ($el in $list){
            $NewItem = $SpList.Items.Add();
            $NewItem["��������� �.�.�."] = $el
            $NewItem.Update();
        }
}

function Add-UsersSampleData($SpList)
{

$list = @()

  $el = "" | select Room, dAtt,nAkt, sDost, sCont, kZam, Opech, Oxr
  $el.Room = "3475"
  
  
  $el.nAkt = "��-65-24-3-8/9998"
  
  
  $el.sCont = "�����������"
  $el.kZam  = "�� ����������"
  $el.Opech = "�����������"
  $el.Oxr   = "����������� ��������"
  
  $NewItem = $SpList.Items.Add();
  $NewItem["����� ���������"] = $el.Room
  [DateTime]$d = get-Date -Year 2014 -Month 12 -Day 17
  $NewItem["���� ����������"] = $d
  $NewItem["����� ���� ����������"] = $el.nAkt
  [DateTime]$d = Get-Date -Year 2015 -Month 11 -Day 27
  $NewItem["������� ������ �������"] = $d
  $NewItem["������� �������� � ���������� ��������"] = $el.sCont
  $NewItem["������� �����"] = $el.kZam
  $NewItem["������������"] = $el.Opech
  $NewItem["����� ��������� ��� ������"] = $el.Oxr
  $NewItem.Update();
  
  $el = "" | select Room, dAtt,nAkt, sDost, sCont, kZam, Opech, Oxr
  $el.Room = "1141"
  
  
  $el.nAkt = "��-65-24-3-8/9997"
  
  
  $el.sCont = "�����������"
  $el.kZam  = "����������"
  $el.Opech = "�����������"
  $el.Oxr   = "�� ���� ������"
   
  $NewItem = $SpList.Items.Add();
  $NewItem["����� ���������"] = $el.Room
  [DateTime]$d = get-Date -Year 2014 -Month 12 -Day 24
  $NewItem["���� ����������"] = $d

  $NewItem["����� ���� ����������"] = $el.nAkt
  [DateTime]$d = Get-Date -Year 2015 -Month 10 -Day 14
  $NewItem["������� ������ �������"] = $d
  $NewItem["������� �������� � ���������� ��������"] = $el.sCont
  $NewItem["������� �����"] = $el.kZam
  $NewItem["������������"] = $el.Opech
  $NewItem["����� ��������� ��� ������"] = $el.Oxr
  $NewItem.Update();
   
  $el = "" | select Room, dAtt,nAkt, sDost, sCont, kZam, Opech, Oxr
  $el.Room = "2202"
  #[DateTime]$el.dAtt = 
  
  $el.nAkt = "��-65-24-3-8/9996"
  #[DateTime]$el.sDost = 
  
  $el.sCont = "�����������"
  $el.kZam  = "�� ����������"
  $el.Opech = "�����������"
  $el.Oxr   = "���"  
  
  $NewItem = $SpList.Items.Add();
  $NewItem["����� ���������"] = $el.Room
  [DateTime]$d = get-Date -Year 2013 -Month 2 -Day 24
  $NewItem["���� ����������"] = $d
  $NewItem["����� ���� ����������"] = $el.nAkt
  [DateTime]$d = Get-Date -Year 2014 -Month 5 -Day 23
  $NewItem["������� ������ �������"] = $d
  $NewItem["������� �������� � ���������� ��������"] = $el.sCont
  $NewItem["������� �����"] = $el.kZam
  $NewItem["������������"] = $el.Opech
  $NewItem["����� ��������� ��� ������"] = $el.Oxr
  $NewItem.Update();
  
  
}
function  Add-TMSampleData($SpList)
{


$list = @()

$list += "08 000000560426 3B"
$list += "0C 0000001A2531 D6"
$list += "08 000000F334A08 6E"
$list += "08 000000F242A4 6D"
$list += "0C 0000001E5461 56"
$list += "0C 0000001A400E DC"
$list += "0C 0000001A8498 E2"
$list += "0C 0000000A26E8 F2"
$list += "08 000000E33AB4 D2"
$list += "08 000001137F25 B5"
$list += "08 000000F2A1FC 76"
$list += "0C 0000001A6116 13"
$list += "08 000000E35E00 07"
$list += "0C 0000002A8E4A BD"
$list += "000000F36DA54208"
$list += "0A 0000000CFC94 73"
$list += "0A 0000000D3DE1 60"
$list += "0C 00000014FB8D 45"
$list += "0C 000000153322 AE"
$list += "08 000000F30350 68"
$list += "08 000000DFBDC8 F7"

foreach ($el in $list){
            $NewItem = $SpList.Items.Add();
            $NewItem["� ��-��������������"] = $el
            $NewItem.Update();
        }


}
function Add-AdminsSampleData($SpList)
{

       $NewItem = $SpList.Items.Add(); 
       
        $NewItem["��������� �.�.�."] = "�������� �.�."
        [DateTime]$d = Get-Date -Year 2016 -Month 7 -Day 18  # ���� � �����
        $NewItem["���� ������������"]  =  $d 
        $NewItem["�����������"] = "��"#  
        $NewItem["����"] = "��������� ���"# �����   
        $NewItem["���������"]  = "������� �������"# ������������ ����� 
        [DateTime]$d = Get-Date -Year 2015 -Month 9 -Day 30  # ���� � �����   # ���� � ����� 
        $NewItem["���� ������� � ����������"] = $d   
        $NewItem["����� ������� � ����������"]  = "� ���-1-29-65/456"
       
       $NewItem.Update(); 


       $NewItem = $SpList.Items.Add(); 
       
        $NewItem["��������� �.�.�."] = "������� �.�."
        [DateTime]$d = Get-Date -Year 2016 -Month 7 -Day 21  # ���� � �����  
        $NewItem["���� ������������"]  = $d
        $NewItem["�����������"] = "��"#  
        $NewItem["����"] = "���"# �����   
        $NewItem["���������"]  = "������� �������"# ������������ ����� 
        [DateTime]$d =   Get-Date -Year 2015 -Month 11 -Day 19  # ���� � �����   # ���� � �����  
        $NewItem["���� ������� � ����������"] =  $d
        $NewItem["����� ������� � ����������"]  = "� ���-1-29-65/876"
       
       $NewItem.Update(); 


       $NewItem = $SpList.Items.Add(); 
       
        $NewItem["��������� �.�.�."] = "����� �.�."
        [DateTime]$d = Get-Date -Year 2016 -Month 3 -Day 10  # ���� � �����  
        $NewItem["���� ������������"]  = $d
        $NewItem["�����������"] = "��"#  
        $NewItem["����"] = "���"# �����   
        $NewItem["���������"]  = "������� ������ ���������"# ������������ ����� 
        [DateTime]$d = Get-Date -Year 2016 -Month 4 -Day 18 # ���� � �����   # ���� � �����   
        $NewItem["���� ������� � ����������"] =  $d  
        $NewItem["����� ������� � ����������"]  = "� ���-1-29-65/951"
       
       $NewItem.Update(); 


}
function Add-SVTRestrictSampleData($SpList)
{
      $NewItem = $SpList.Items.Add(); 
      
            
      
      $NewItem["����������� ����� ���"]  =  "1;#25736" # �����������   
      $NewItem["�.�.�. ����������� � ���"]  = "7;#�������� �.�."# �����������   
      $NewItem["��� ��. ������ � ��� �� ���"]  =  "��. �������������"# ������������ �����   
      $NewItem["� ��-��������������"]  = "9;#08 000000E33AB4 D2" #�����������   
      $NewItem["����� �������"]  =  "���. �������������"# �����  
      

      $NewItem.Update(); 
      
      $NewItem = $SpList.Items.Add(); 
      
            
      #$NewItem["� �/�"]  = "13"                                   #  ������������ �����   
      $NewItem["����������� ����� ���"]  =  "1;#25736"            # �����������   
      $NewItem["�.�.�. ����������� � ���"]  = "6;#��������� �.�." # �����������   
      $NewItem["��� ��. ������ � ��� �� ���"]  =  "�������������" # ������������ �����   
      $NewItem["� ��-��������������"]  = "12;#0C 0000001A6116 13"       #�����������   
      $NewItem["����� �������"]  =  "������������"# �����  
      # $NewItem["���������� ���"]  =  # ����������� 

      
      $NewItem.Update(); 
           
      
}
function Add-SVTsampleData($SpList)
{

        $ctype = $splist.ContentTypes["����� ����������"]
 
        [HashTable]$docsetpropertes = @{}

        $folder = [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]::Create($splist.RootFolder,"���������_25736",$ctype.ID, $docsetpropertes)
        


        $item = $splist.Items | Where {$_["���"] -eq "���������_25736"}

     
        $Item["BBK"]  = "��������� ����������� 10"#  ����� ����� ����������  
        $Item["��"]  = "�� ��� 2.0; ��� �������������� ���� �� ���"  #  ������������ ����� ����� ����������  
        $Item["��������� �����"]  = "SKAM192345"  #  ������������ ����� ����� ����������  
        $Item["��������� ����� ���"]  = "92116886" #  ������������ ����� ����� ����������  
        $Item["���. ����� ���"]  =  "31313"#  ������������ ����� ����� ����������  
        $Item["����������� �����"]  =  "25736"#  ������������ ����� ����� ����������  
        $Item["��������"]  =   "���������_25736"#  ������������ ����� ����� ����������  
        $Item["����� ���������"]  = "1141" #  ������������ ����� ����� ����������  
        $Item["������������ �������"]  = "Windows 7"  #  ����� ����� ����������  
        $Item["��������"]  = "�� ��� 2.0; ��� �������������� ���� �� ���" #  ������������� ����� ����� ����������  
        $Item["��� �� ���"]  = "������-����"                              #  ����� ����� ����������  
        $Item["��� ���"]  = "Fujitsu-Siemens Esprimo"                     #  ������������ ����� ����� ����������  
        $Item["�.�.�. ��������������"]  = "��������� �.�."                #  ������������ ����� 

        $Item.Update();


        $folder = [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]::Create($splist.RootFolder,"���������_36293",$ctype.ID, $docsetpropertes)

        $item = $splist.Items | Where {$_["���"] -eq "���������_36293"}
         
        $Item["BBK"]  = "��������� ����������� 8"#  ����� ����� ����������  
        $Item["��"]  = "�� ��� 2.0; ��������� ������ �� Oracle + WAS+MQ"  #  ������������ ����� ����� ����������  
        $Item["��������� �����"]  = "CZJ838B2NX"  #  ������������ ����� ����� ����������  
        $Item["��������� ����� ���"]  = "257�-25869" #  ������������ ����� ����� ����������  
        $Item["���. ����� ���"]  =  "31316"#  ������������ ����� ����� ����������  
        $Item["����������� �����"]  =  "36293"#  ������������ ����� ����� ����������  
        $Item["��������"]  =   "���������_36293"#  ������������ ����� ����� ����������  
        $Item["����� ���������"]  = "3475" #  ������������ ����� ����� ����������  
        $Item["������������ �������"]  = "Windows 7"  #  ����� ����� ����������  
        $Item["��������"]  = "�� ��� 2.0; ��������� ������ �� Oracle + WAS+MQ" #  ������������� ����� ����� ����������  
        $Item["��� �� ���"]  = "������-����" #  ����� ����� ����������  
        $Item["��� ���"]  = "HP 7800"#  ������������ ����� ����� ����������  
        $Item["�.�.�. ��������������"]  = "����������� �.�." #  ������������ ����� 

        $Item.Update(); 
       
}
