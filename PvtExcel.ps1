##############################################
#
#   ������ �������� ������ ��������������� ��������
#
#   ����: PvtExcel.ps1 
#   �������� : ������ ������� ������� ������ � ������� xlsx ��� ������� 
#              ���������������� �������.
#
#   �����: ������� ��������� ���������, ����� ��������������, ���: 223-1338
#          29AstahovAB@cbr.ru
#
#   ��������� ������, �� ����� ������ �� ���
#   ������� ����������: ������ ��������� ������� ��������� �� ����������
#                       � ������������� Excel 2010 ��� ����,
#                       �� �������� XLSQRY, ����� ������� ��������
#                       ������������ ������ main.ps1
#   ����: ������ 2017.
#   ������: 17.01.11
#
##############################################
start-transcript "PvtExcel.log"
cls
$0 = $myInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$IniFile = "$dp0\TU.INI"   
. "$dp0\Utils-Excel.ps1"
. "$dp0\Utils-Ini.ps1"


$webObject = Get-WebIniObject $IniFile 
$webObject.WrkDir = $dp0
$webObject.iqyFolder = $dp0

$wrkDir = Get-ChildItem  | ?{$_.PSIsContainer} | Select FullName
$wrkDir = $webObject.iqyFolder + "\" + $webObject.NewWebName

$wdCount = $webObject.SiteCount


for ($jj=1;$jj -le $wdCount; $jj++)
{
      $source = $wrkDir + $jj.ToString() + "\*.iqy"
      $dest  =  $dp0
      $reportName = "Report" + $jj.ToString() + ".xlsx"
      #write-host $("Copy-Item -Path "+$source+" -destination "+$dest+" -Confirm:$false")
      Copy-Item -Path $source -destination $dest -Confirm:$false
      write-host "�������� ������ ���������������� ������� ...  "  -ForegroundColor Yellow
      $ReportFileName = Create-Excel $reportName $webObject
      
      write-host $("����� ���������������� ������� ") 
      write-host $ReportFileName  -ForegroundColor Cyan
      write-host $(" ������.") 
      write-host "��������� ��� � ���������� ���������� "
      write-host ($webObject.Site +  $webObject.Hierarchy + $webObject.NewWebName + $jj.ToString() +"/Reports" ) -ForegroundColor Cyan

}

Stop-Transcript
