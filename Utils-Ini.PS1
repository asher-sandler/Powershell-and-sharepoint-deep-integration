##############################################
#
#   Модуль создания сайтов технологических участков
#
#   Файл: Util-Ini.ps1 
#   Описание : Функции для работы с инициализациооным файлом. 
#              технологического участка.
#
#   Автор: Астахов Александр Борисович, отдел информатизации, тел: 223-1338
#          29AstahovAB@cbr.ru
#
#   Отделение Калуга, ГУ Банка России по ЦФО
#   Условия выполнения: Данный модуль запускается в составе других модулей. 
#                       Не следует запускать его отдельно.
#   Дата: Январь 2017.
#   Версия: 17.01.11
#
##############################################

function Test-Ini([string]$tIniFileName,[string]$wrkFolder)
{

      $IsTested = $False
            if (Test-Path $tIniFileName)
            {
                 $IsTested = $True

            }
            else
            {
                 $crlf = [char][int]13+[char][int]10
                 $iniFileshablon =  "; =====    Ini file ===="+$crlf
                 $iniFileshablon += ";"+$crlf+$crlf+$crlf
                 $iniFileshablon += ";Сайт sharepoint не должно быть слеша в конце "+$crlf
                 $iniFileshablon += "Site=http://"+$env:computername+$crlf+$crlf
                 $iniFileshablon += ";Существующая иерархия, где будет создана структура сайтов. Обязательно Слеш в начале и в конце. Параметр может иметь значение, например /otdely/Uchastki/."+$crlf
                 $iniFileshablon += ";Сайт http://site/otdely/Uchastki  должен быть уже создан на момент запуска модуля main.ps1 и пользователь под которым вы зашли на консоль сервера РСИ"+$crlf
                 $iniFileshablon += ";должен иметь права на создание сайтов в данной иерархии."+$crlf
                 $iniFileshablon += ";Сайт технологического участка будет иметь URL http://site/otdely/Uchastki/TUN, где N номер тех. участка "+$crlf
                 $iniFileshablon += "Hierarchy=/"+$crlf+$crlf
                 $iniFileshablon += ";Наименование сайтов технологического участка"+$crlf
                 $iniFileshablon += "NewWebName=TU"+$crlf+$crlf
                 $iniFileshablon += ";Наименование сайта информационного модуля"+$crlf
                 $iniFileshablon += ";Сайт информационного модуля будет иметь URL http://site/otdely/Uchastki/TUInfoModule"+$crlf
                 $iniFileshablon += "InfoModule=TUInfoModule"+$crlf+$crlf
                 $iniFileshablon += ";Отображаемое имя технологического участка"+$crlf
                 $iniFileshablon += "Title=Технологический участок"+$crlf+$crlf
                 $iniFileshablon += ";Описание технологического участка"+$crlf
                 $iniFileshablon += "Description=Информация о Технологическом участке"+$crlf+$crlf
                 $iniFileshablon += ";Количество создаваемых сайтов технологических участков"+$crlf
                 $iniFileshablon += "SiteCount=2"+$crlf+$crlf
                 $iniFileshablon += ";Каталог для хранения временных и служебных файлов сводных отчетов. Если не существует, будет создан автоматически"+$crlf
                 $iniFileshablon += "iqyFolder="+$wrkFolder+"\XLSQRY"+$crlf+$crlf
                 $iniFileshablon += ";========="


                

                  $iniFileshablon | out-file -encoding UTF8 -filepath $tIniFileName # |  Out-Null  
                

                 


            }

      #write-out $IsTested
      return $IsTested



}
function Get-WebIniObject([string]$tIniFileName)
{

$webObject = "" | Select Site, Hierarchy, NewWebName, InfoModule, Title, Description, SiteCount, iqyFolder, WrkDir


$iniFile = Get-Content $tIniFileName

    foreach($line in $iniFile)
    {
        $pLine = $line.Trim().ToUpper()
        
        if ($pLine.Contains("SITE") -and $pLine.Contains("="))
        {
              $qline = $pLine.Split("=")[0].Trim().ToUpper()
              if ($qline -eq "SITE")
              {
                    $aLine = $Line.Split("=")[1].Trim()
                    $webObject.Site        = $aLine
                 
              }
        
        }


        if ($pLine.Contains("HIERARCHY") -and $pLine.Contains("="))
        {
              $qline = $pLine.Split("=")[0].Trim().ToUpper()  
              if ($qline -eq "HIERARCHY")
              {        
                     
                     $aLine = $Line.Split("=")[1].Trim() 
                     $webObject.Hierarchy    = $aLine
                 
              }
                       
        
        }
        if ($pLine.Contains("NEWWEBNAME") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper()                       
               if ($qline -eq "NEWWEBNAME")
               {        
                     
                     $aLine = $Line.Split("=")[1].Trim()
                     $webObject.NewWebName  = $aLine
                  
               }                 
        
        }   
        if ($pLine.Contains("INFOMODULE") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper()         
               if ($qline -eq "INFOMODULE")
               {        
                      
                     $aLine = $Line.Split("=")[1].Trim() 
                     $webObject.InfoModule   = $aLine
                                 
               }
        
        }   
        if ($pLine.Contains("TITLE") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper()   
               if ($qline -eq "TITLE")
               {        
                    
                     $aLine = $Line.Split("=")[1].Trim() 
                     $webObject.Title       = $aLine
                   
               }                 
        
        }        
        if ($pLine.Contains("DESCRIPTION") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper()   
               if ($qline -eq "DESCRIPTION")
               {        
               
                   $aLine = $Line.Split("=")[1].Trim() 
                   $webObject.Description    = $aLine
                  
               }                  
        
        }  
        if ($pLine.Contains("SITECOUNT") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper() 
               if ($qline -eq "SITECOUNT")
               {        
                            
                     $aLine = $Line.Split("=")[1].Trim() 
                     $webObject.SiteCount    = $aLine
                    
               }                  
        
        }    
        if ($pLine.Contains("IQYFOLDER") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper()   
               if ($qline -eq "IQYFOLDER")
               {        
                         
                     $aLine = $Line.Split("=")[1].Trim() 
                     $webObject.iqyFolder  = $aLine
                    
               }                   
        
        } 


    }
    
    return $webObject

}
function Test-WebIniObject([string]$tIniFileName)
{

$webOBool  = "" | Select Site, Hierarchy, NewWebName, InfoModule, Title, Description, SiteCount, iqyFolder

$webOBool.Site         = $false
$webOBool.Hierarchy    = $false
$webOBool.NewWebName   = $false
$webOBool.InfoModule   = $false
$webOBool.Title        = $false
$webOBool.Description  = $false
$webOBool.SiteCount    = $false
$webOBool.iqyFolder    = $false


$iniFile = Get-Content $tIniFileName

    foreach($line in $iniFile)
    {
        $pLine = $line.Trim().ToUpper()
        
        if ($pLine.Contains("SITE") -and $pLine.Contains("="))
        {
              $qline = $pLine.Split("=")[0].Trim().ToUpper()
              if ($qline -eq "SITE")
              {
                    $aLine = $pLine.Split("=")[1].Trim()
                   
                    $webOBool.Site         = $True
              }
        
        }


        if ($pLine.Contains("HIERARCHY") -and $pLine.Contains("="))
        {
              $qline = $pLine.Split("=")[0].Trim().ToUpper()  
              if ($qline -eq "HIERARCHY")
              {        
                     
                     $aLine = $pLine.Split("=")[1].Trim() 
                   
                     $webOBool.Hierarchy    = $True
              }
                       
        
        }
        if ($pLine.Contains("NEWWEBNAME") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper()                       
               if ($qline -eq "NEWWEBNAME")
               {        
                     
                     $aLine = $pLine.Split("=")[1].Trim()
                    
                     $webOBool.NewWebName   = $True 
               }                 
        
        }   
        if ($pLine.Contains("INFOMODULE") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper()         
               if ($qline -eq "INFOMODULE")
               {        
                      
                     $aLine = $pLine.Split("=")[1].Trim() 
                    
                     $webOBool.InfoModule   = $True                   
               }
        
        }   
        if ($pLine.Contains("TITLE") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper()   
               if ($qline -eq "TITLE")
               {        
                    
                     $aLine = $pLine.Split("=")[1].Trim() 
                   
                     $webOBool.Title        = $True  
               }                 
        
        }        
        if ($pLine.Contains("DESCRIPTION") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper()   
               if ($qline -eq "DESCRIPTION")
               {        
               
                   $aLine = $pLine.Split("=")[1].Trim() 
                  
                   $webOBool.Description  = $True 
               }                  
        
        }  
        if ($pLine.Contains("SITECOUNT") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper() 
               if ($qline -eq "SITECOUNT")
               {        
                            
                     $aLine = $pLine.Split("=")[1].Trim() 
                    
                     $webOBool.SiteCount    = $True 
               }                  
        
        }    
        if ($pLine.Contains("IQYFOLDER") -and $pLine.Contains("="))
        {
               $qline = $pLine.Split("=")[0].Trim().ToUpper()   
               if ($qline -eq "IQYFOLDER")
               {        
                         
                     $aLine = $pLine.Split("=")[1].Trim() 
                    
                     $webOBool.iqyFolder    = $True
               }                   
        
        } 


    }

    return $webOBool    

}
function Show-BadIniParams($webOBool, $tIniFileName)
{
    $isGood = $True

    if (!$webOBool.Site)
    {
        $paramName = "Site"
        Show-WhatIsBad $tIniFileName $paramName
        $isGood = $False
        
    } 
    if (!$webOBool.Hierarchy)
    {
        $paramName = "Hierarchy"
        Show-WhatIsBad $tIniFileName $paramName
        $isGood = $False
        
    }         
    if (!$webOBool.NewWebName)
    {
        $paramName = "NewWebName"
        Show-WhatIsBad $tIniFileName $paramName
        $isGood = $False
        
    }    
    if (!$webOBool.InfoModule)
    {
        $paramName = "InfoModule"
        Show-WhatIsBad $tIniFileName $paramName
        $isGood = $False
        
    }     
    if (!$webOBool.Title)
    {
        $paramName = "Title"
        Show-WhatIsBad $tIniFileName $paramName
        $isGood = $False
        
    }     
    if (!$webOBool.Description)
    {
        $paramName = "Description"
        Show-WhatIsBad $tIniFileName $paramName
        $isGood = $False
        
    }    
    if (!$webOBool.SiteCount)
    {
        $paramName = "SiteCount"
        Show-WhatIsBad $tIniFileName $paramName
        $isGood = $False
        
    }

    if (!$webOBool.iqyFolder)
    {
        $paramName = "iqyFolder"
        Show-WhatIsBad $tIniFileName $paramName
        $isGood = $False
        
    }
    
    return $isGood
    
}
function Show-WhatIsBad ($tIniFileName, $paramName)
{
        write-Host $("В INI файле "+$tIniFileName+" не хватает параметра ") -foreground Yellow -nonewline
        write-Host $($paramName) -foreground Cyan -nonewline
        write-Host $(".") -foreground Yellow
        write-Host $("Удалите файл ") -foreground Yellow -nonewline
        write-Host $tIniFileName -foreground Cyan -nonewline 
        write-Host $(" и запустите процедуру заново. Создасться файл с пустыми параметрами.") -foreground Yellow


}
<#
cls
$0 = $myInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$IniFile = "$dp0\TU.INI"    
write-host $inifile
$isIniExists = Test-Ini $IniFile 
$webOBool = Test-WebIniObject  $IniFile
$IsParamFileGood = Show-BadIniParams $webOBool $IniFile
if ($IsParamFileGood)
{
     $webObject = Get-WebIniObject $IniFile 
     $webObject | fl
     
}
else
{
     write-host "Somethimg wrong..."
}

#>







