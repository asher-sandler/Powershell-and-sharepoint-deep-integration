##############################################
#
#   Модуль создания сайтов технологических участков
#
#   Файл: PvtExcel.ps1 
#   Описание : Модуль создает сводные отчеты в формате xlsx для каждого 
#              технологического участка.
#
#   Автор: Астахов Александр Борисович, отдел информатизации, тел: 223-1338
#          29AstahovAB@cbr.ru
#
#   Отделение Калуга, ГУ Банка России по ЦФО
#   Условия выполнения: Данную процедуру следует запускать на компьютере
#                       с установленным Excel 2010 или выше,
#                       из каталога XLSQRY, после запуска главного
#                       исполняемого модуля main.ps1
#   Дата: Январь 2017.
#   Версия: 17.01.11
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
      write-host "Создание отчета технологического участка ...  "  -ForegroundColor Yellow
      $ReportFileName = Create-Excel $reportName $webObject
      
      write-host $("Отчет технологического участка ") 
      write-host $ReportFileName  -ForegroundColor Cyan
      write-host $(" создан.") 
      write-host "Поместите его в библиотеку документов "
      write-host ($webObject.Site +  $webObject.Hierarchy + $webObject.NewWebName + $jj.ToString() +"/Reports" ) -ForegroundColor Cyan

}

Stop-Transcript
