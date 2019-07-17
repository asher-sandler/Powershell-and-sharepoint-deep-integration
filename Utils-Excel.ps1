##############################################
#
#   Модуль создания сайтов технологических участков
#
#   Файл: Util-Excel.ps1 
#   Описание : Функции для создания сводного отчета. 
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

function Get-Range($Beit)            
{
      $columnL = Convert-ToLetter $Beit.LeftColumn
      $columnR = Convert-ToLetter $Beit.Rightcolumn
      
      $range = $columnL + $Beit.LeftRow.Tostring() + ":"+ $columnR + $Beit.RightRow.Tostring()
                  
  
      return $range

}
function Create-UserPivotTable($wb, $WorkSheet, $Coordinates, $sourceTable, $PivotTableName, $fields, $shNames)
 {
            
            $xlDatabase            = 1
            $xlPivotTableVersion12 = 3
            $xlPivotTableVersion14 = 4
            
            $xlHidden              = 0
            $xlRowField            = 1
            $xlColumnField         = 2
            $xlPageField           = 3
            $xlDataField           = 4
            $xlTabularRow          = 1
            
            $xlDiagonalDown        = 5
            $xlDiagonalUp          = 6
            
            $xlEdgeLeft            = 7
            $xlEdgeTop             = 8
            $xlEdgeBottom          = 9
            $xlEdgeRight           = 10
            $xlInsideVertical      = 11
            
            $xlContinuous          = 1
            $xlThin                = 2
            
            $xlCenter              = -4108
            $xlBottom              = -4107
            
            
            $reference = $shNames[0]+"!R" + $Coordinates.Row.Tostring() +"C" + $Coordinates.Column.Tostring()
            
            #write-host $reference
           
            
            $PivotAdmins = $wb.PivotCaches().Create($xlDataBase, $sourceTable,$xlPivotTableVersion14)
            $PivotAdmins.CreatePivotTable($reference,$PivotTableName)   | Out-Null                            
            Start-Sleep -Milliseconds 500
            
            #[void]$PivotAdmins.Select()
            
            for ($j=0; $j -lt $fields.Count;$j++)
            {     
                   
                   $PivotFields = $WorkSheet.PivotTables($PivotTableName).PivotFields($fields[$j])
                   
                   $PivotFields.Orientation = $xlRowField
                   
		               # clear subtotals

                   for ($k=1; $k -le 12; $k++){
                           $PivotFields.Subtotals($k)=$false
                    }
                   
           
            }
            

            $WorkSheet.PivotTables($PivotTableName).RowAxisLayout($xlTabularRow)
            $WorkSheet.PivotTables($PivotTableName).ColumnGrand = $False
            $WorkSheet.PivotTables($PivotTableName).RowGrand = $False
            $WorkSheet.PivotTables($PivotTableName).ShowTableStyleColumnHeaders=$false
            $WorkSheet.PivotTables($PivotTableName).ShowTableStyleRowHeaders=$false


            $WorkSheet.PivotTables($PivotTableName).ShowDrillIndicators=$false
            #$WorkSheet.PivotTables($PivotTableName).ShowPageMultipleItemLabel=$false
            #$WorkSheet.PivotTables($PivotTableName).ShowTableStyleColumnHeaders=$false
            #$WorkSheet.PivotTables($PivotTableName).ShowTableStyleColumnStripes=$false
            #$WorkSheet.PivotTables($PivotTableName).ShowTableStyleLastColumn=$false
             
 
 
 
            $WorkSheet.PivotTables($PivotTableName).ColumnGrand=$False
            #$WorkSheet.PivotTables($PivotTableName).EnableFieldList=$false
            
            #$mygm =  $WorkSheet.PivotTables($PivotTableName) | gm
            #write-host $mygm
           
            $Beit = Get-TableCoordinates($Pasport.PivotTables($PivotTableName))
            
            $tRange = Get-Range($Beit)
            
            Start-Sleep -Milliseconds 5000
            
      
            $selection = $WorkSheet.Range($tRange)
            $selection.Select() | Out-Null   
            
            $selection.Cells.Borders.Item($xlEdgeLeft).LineStyle        = $xlContinuous
            $selection.Cells.Borders.Item($xlEdgeTop).LineStyle         = $xlContinuous
            $selection.Cells.Borders.Item($xlEdgeBottom).LineStyle      = $xlContinuous
            $selection.Cells.Borders.Item($xlEdgeRight).LineStyle       = $xlContinuous
            $selection.Cells.Borders.Item($xlInsideVertical).LineStyle  = $xlContinuous            
            $selection.Cells.Borders.Item($xlInsideVertical).LineStyle  = $xlContinuous                     
            
            
            $selection.Cells.Borders.Item($xlEdgeLeft).Weight           = $xlThin
            $selection.Cells.Borders.Item($xlEdgeTop).Weight            = $xlThin
            $selection.Cells.Borders.Item($xlEdgeBottom).Weight         = $xlThin
            $selection.Cells.Borders.Item($xlEdgeRight).Weight          = $xlThin
            $selection.Cells.Borders.Item($xlInsideVertical).Weight     = $xlThin     
            $selection.Cells.Borders.Item($xlInsideVertical).Weight     = $xlThin
            
            
            $selection.HorizontalAlignment = $xlCenter 
            $selection.VerticalAlignment   = $xlCenter
            $selection.WrapText            = $True
            $Selection.EntireColumn.AutoFit() | Out-Null
            
            #write-host 1111
            #write-host $Beit
            return $Beit  # coordinates of tables
             
            
                               
 }
function Get-TableCoordinates($PivotTable)
{

            
            $Beit = "" | Select LeftRow,LeftColumn, RightRow, Rightcolumn
            
            $beit.LeftRow     = $PivotTable.RowRange.Row
            $beit.LeftColumn  = $PivotTable.RowRange.Column
            $Beit.RightRow    = $PivotTable.RowRange.Row + $PivotTable.RowRange.Rows.Count - 1
            $Beit.Rightcolumn = $PivotTable.RowRange.Column  + $PivotTable.RowRange.Columns.Count  - 1
         
            
            return $Beit
}

Function Convert-ToLetter($iCol)
{ # As String

   $iAlpha = [int]($iCol / 27)
   $iRemainder = $iCol - ($iAlpha * 26)
   If ($iAlpha -gt 0){
      $ConvertToLetter = [Char]($iAlpha + 64)
   }
   If ($iRemainder -gt 0)
   { # Then
      $ConvertToLetter = $ConvertToLetter + [char]($iRemainder + 64)
   }
   return $ConvertToLetter
}
function Create-DestinationXLSSheets($xlFileName,$wshCount)

{
      
            $excel = New-Object -ComObject Excel.Application

            $xlFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlOpenXMLWorkbook

            $excel.visible = $false
            $excel.DisplayAlerts=$false

            $workbook=$excel.Workbooks.Add()
            $howManyWSToAdd = $wshCount - $workbook.worksheets.count 
            # write-host $howManyWSToAdd,$wshCount,$workbook.worksheets.count 
            
            #read-host
            #if ($howManyWSToAdd -gt 0)
            #{
            #    for($i=1; $i -le $howManyWSToAdd;$i++)               
            #    {
            #        $workbook.workSheets.Add() 
            #    }
            #}
            
            #read-host
            $workbook.SaveAs($xlFileName, $xlFixedFormat)
            $Excel.Workbooks.Close()

            $excel.quit()




            $noOutput = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

            Remove-Variable Excel
            Kill-Excel

}


function Create-XLSQueryFromSPlist($iqyFile,$xlFileName )
{

            $excel = New-Object -ComObject Excel.Application

            $xlFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlOpenXMLWorkbook

            $excel.visible = $false
            $excel.DisplayAlerts=$false
            $workbook=$excel.Workbooks.Open($iqyFile)
            
           
            $workbook.SaveAs($xlFileName, $xlFixedFormat)
            $Excel.Workbooks.Close()

            $excel.quit()


            $noOutput = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

            Remove-Variable Excel
            Kill-Excel            
            

}
function Copy-WorkSheets($sourceXLS,$destXLS,$DestSheet)
{
           $excel = New-Object -ComObject Excel.Application

            $xlFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlOpenXMLWorkbook

            $excel.visible = $false
            $excel.DisplayAlerts=$false
            $wbSource=$excel.Workbooks.Open($sourceXLS)
            $wbDest  =$excel.Workbooks.Open($destXLS)
            
            
            
            $sheetSource = $wbSource.WorkSheets.Item(1)
            $sheetDest   = $wbDest.Sheets.Item($DestSheet)
            
            
            $sheetSource.Copy($sheetDest)
            #read-host
            
            $wbDest.Save()
            
            #$workbook.SaveAs($xlFileName, $xlFixedFormat)
            $Excel.Workbooks.Close()

            $excel.quit()


            $noOutput = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

            Remove-Variable Excel
            Kill-Excel
}

function Create-PivotTable($sourceXLS)
{

           $excel = New-Object -ComObject Excel.Application
           $exvers = $($excel.Version).ToString()
           $rpath = "HKCU:\Software\Microsoft\Office\"+$exvers + "\excel\Security"

            
            
            
      
            New-ItemProperty -Path $rpath  -Name AccessVBOM -Value 1 -Force | Out-Null
            New-ItemProperty -Path $rpath  -Name VBAWarnings -Value 1 -Force | Out-Null
            $xlFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlOpenXMLWorkbook

            $excel.visible = $false
            #$excel.visible = $TRUE
            $excel.DisplayAlerts=$false
            $wb=$excel.Workbooks.Open($sourceXLS)
            $wb.ShowPivotTableFieldList = $true

            $isEnglish = $false
            
            for ($i=1;$i -le $wb.WorkSheets.count; $i++)
            {
		      #write-host 1111
                      #write-host $wb.WorkSheets.Item($i).Name
                      if ($wb.WorkSheets.Item($i).Name -eq "Sheet1")
                      {
                          $isEnglish = $true;
                          break;
                      }
            }
            #write-host 2222
            #write-host $isEnglish
            if ($isEnglish)
	    {
                 $shNames = @("Sheet1","Sheet2","Sheet3")
                 $tablName="Table_"
	    }
            else
            {
                 $shNames = @("Лист1","Лист2","Лист3")
                 $tablName="Таблица_"
	    }

            
          
            $Pasport = $wb.WorkSheets.Item($shNames[0])
            
            $Pasport.Activate()
            
            $Pasport.Range("A3:A3").Cells.Value2="ПАСПОРТ ТЕХНОЛОГИЧЕСКОГО УЧАСТКА"
            $Pasport.Range("A3:A3").Cells.Font.Name = "Times New Roman"
            $Pasport.Range("A3:A3").Cells.Font.Size = 14
            
            $Pasport.Range("A6:A6").Cells.Value2="АДМИНИСТРАТОРЫ ИНФОРМАЦИОННОЙ БЕЗОПАСНОСТИ"
            $Pasport.Range("A6:A6").Cells.Font.Name = "Times New Roman"
            $Pasport.Range("A6:A6").Cells.Font.Size = 14

            $Pasport.Range("A7:A7").Cells.Value2="ТЕХНОЛОГИЧЕСКОГО УЧАСТКА"
            $Pasport.Range("A7:A7").Cells.Font.Name = "Times New Roman"
            $Pasport.Range("A7:A7").Cells.Font.Size = 14
                   
                     
            $pvtTableName = "Admins"
            $Coordinates = "" | Select Row, Column
            $Coordinates.Row = 9; $Coordinates.Column = 2
            $sourceTable = $tablName + $pvtTableName
            $PivotTableName = "PVT_" + $pvtTableName
            
            $fields = @()
            $fields += "Роль"
            $fields += "Сотрудник Ф.И.О."
            $fields += "Должность"
            $fields += "Дата приказа о назначении"
            $fields += "Номер приказа о назначении"
       
            #write-host $pvtTableName                            
            $beit = Create-UserPivotTable $wb $Pasport  $Coordinates  $sourceTable  $PivotTableName  $fields  $shNames
            
            #write-host 2222
            #write-host $beit
            #read-host
            $Coordinates.Row = $beit.RightRow+3; $Coordinates.Column = $beit.LeftColumn
            
            #write-host $beit
            $beit.LeftColumn = 1
            $beit.Rightcolumn = 1
            $beit.RightRow = $beit.RightRow +10
            $beit.LeftRow = $beit.RightRow
            
            #read-host
            $tRange = Get-Range($Beit)  
            
            
            
            $Pasport.Range($tRange).Cells.Value2="ОБЕСПЕЧЕНИЕ ЗАЩИТЫ ПОМЕЩЕНИЙ"
            $Pasport.Range($tRange).Cells.Font.Name = "Times New Roman"
            $Pasport.Range($tRange).Cells.Font.Size = 14
            
            $beit.RightRow = $beit.RightRow + 1
            $beit.LeftRow = $beit.RightRow
 
            $tRange = Get-Range($Beit)  
                        
            $Pasport.Range($tRange).Cells.Value2="В КОТОРЫХ РАСПОЛОЖЕНЫ СРЕДСТВА ВЫЧИСЛИТЕЛЬНОЙ ТЕХНИКИ"
            $Pasport.Range($tRange).Cells.Font.Name = "Times New Roman"
            $Pasport.Range($tRange).Cells.Font.Size = 14
            
            $beit.RightRow = $beit.RightRow + 1
            $beit.LeftRow = $beit.RightRow
                      
            $tRange = Get-Range($Beit)  
                       
            $Pasport.Range($tRange).Cells.Value2=" (СВТ) ТЕХНОЛОГИЧЕСКОГО УЧАСТКА"
            $Pasport.Range($tRange).Cells.Font.Name = "Times New Roman"
            $Pasport.Range($tRange).Cells.Font.Size = 14
             
            $Coordinates.Row = $beit.RightRow+3; $Coordinates.Column = 2
            
            $pvtTableName = "Rooms"
            $sourceTable = $tablName + $pvtTableName
            $PivotTableName = "PVT_" + $pvtTableName
            
            $fields = @()
            $fields += "Номер помещения"
            $fields += "Дата аттестации"
            $fields += "Номер акта аттестации"
            $fields += "Наличие списка доступа"
            $fields += "Система контроля и управления доступом"
            $fields += "Кодовый замок"
       
            #write-host $pvtTableName            
            $beit = Create-UserPivotTable $wb $Pasport  $Coordinates  $sourceTable  $PivotTableName  $fields  $shNames
                          
                          
            #write-host $beit
            $beit.LeftColumn = 1
            $beit.Rightcolumn = 1
            $beit.RightRow = $beit.RightRow +10
            $beit.LeftRow = $beit.RightRow
            
            #read-host
            $tRange = Get-Range($Beit)  
            
            
            
            $Pasport.Range($tRange).Cells.Value2="ПЕРЕЧЕНЬ СВТ НА ТЕХНОЛОГИЧЕСКОМ УЧАСТКЕ"
            $Pasport.Range($tRange).Cells.Font.Name = "Times New Roman"
            $Pasport.Range($tRange).Cells.Font.Size = 14
            
            $beit.RightRow = $beit.RightRow + 1
            $beit.LeftRow = $beit.RightRow


            $Coordinates.Row = $beit.RightRow+3; $Coordinates.Column = 2
            
            $pvtTableName = "SVTs"
            $sourceTable = $tablName + $pvtTableName
            $PivotTableName = "PVT_" + $pvtTableName
            
            $fields = @()
            $fields += "Номер помещения"
            $fields += "Инвентарный номер"
            $fields += "Заводской номер"
            $fields += "Тип ПВМ"
            $fields += "АС"
            $fields += "BBK"
            
            #write-host $pvtTableName               
            $beit = Create-UserPivotTable $wb $Pasport  $Coordinates  $sourceTable  $PivotTableName  $fields  $shNames
                          
                          
            #write-host $beit
            $beit.LeftColumn = 1
            $beit.Rightcolumn = 1
            $beit.RightRow = $beit.RightRow +10
            $beit.LeftRow = $beit.RightRow
            
            #read-host
            $tRange = Get-Range($Beit)  
            
            
            
            $Pasport.Range($tRange).Cells.Value2="СПИСОК РАБОТНИКОВ,"
            $Pasport.Range($tRange).Cells.Font.Name = "Times New Roman"
            $Pasport.Range($tRange).Cells.Font.Size = 14
             
            $beit.RightRow = $beit.RightRow + 1
            $beit.LeftRow = $beit.RightRow
                      
            $tRange = Get-Range($Beit)  
                       
            $Pasport.Range($tRange).Cells.Value2="ДОПУЩЕННЫХ К СВТ ТЕХНОЛОГИЧЕСКОГО УЧАСТКА"
            $Pasport.Range($tRange).Cells.Font.Name = "Times New Roman"
            $Pasport.Range($tRange).Cells.Font.Size = 14
             
            $Coordinates.Row = $beit.RightRow+3; $Coordinates.Column = 2
  
            $pvtTableName = "SVTRestrict"
            $sourceTable = $tablName + $pvtTableName
            $PivotTableName = "PVT_" + $pvtTableName
            
            $fields = @()
            $fields += "Инвентарный номер СВТ"
            $fields += "Назначение СВТ"
            $fields += "Ф.И.О. допущенного к СВТ"
            $fields += "Имя уч. записи в СЗИ от НСД"
            $fields += "Права доступа"
            $fields += "№ ТМ-идентификатора"
            
            #write-host $pvtTableName               
            $beit = Create-UserPivotTable $wb $Pasport  $Coordinates  $sourceTable  $PivotTableName  $fields  $shNames

            #
            $Pasport = $wb.WorkSheets.Item($shNames[0])
            $Pasport.Name = "Паспорт"
              
            $l1 = $wb.WorkSheets.Item($shNames[1])
            $l1.Delete()
            
            $l1 = $wb.WorkSheets.Item($shNames[2])
            $l1.Delete()
            
            
            $SheetAdmins = $wb.WorkSheets.Item("Admins")
            $SheetAdmins.Visible=$false
            
            $SheetRooms = $wb.WorkSheets.Item("Rooms")
            $SheetRooms.Visible=$false
   
            $SheetSVTRestrict = $wb.WorkSheets.Item("SVTRestrict")
            $SheetSVTRestrict.Visible=$false

            $SheetSVTs = $wb.WorkSheets.Item("SVTs")
            $SheetSVTs.Visible=$false
            
            $wb.Save()
            
            #$workbook.SaveAs($xlFileName, $xlFixedFormat)
            $Excel.Workbooks.Close()

            $excel.quit()


            $noOutput = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

            Remove-Variable Excel
     
           Kill-Excel
            
            
            
            
            
}

function Kill-Excel()
{

            Start-Sleep 5

            #read-Host Press a key... 1

            $ExcelCount = @(Get-Process excel -ea 0).Count
            
            #"Excel process 2: {0}" -f $ExcelCount
            if ($ExcelCount -gt 0)
            {
                     $noOutput = Stop-Process -Name Excel
            }
            #read-Host Press a key... 2

            #"Excel process 3: {0}" -f @(Get-Process excel -ea 0).Count

       


}
function Test-Excel()
{
    $isTested = $False

    $exCelApp = Test-Path HKLM:SOFTWARE\Classes\Excel.Application
    if ($exCelApp)
    {
           $excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue 
           $ver = [int]$excel.Version
           if ($ver -ge 14)
           {
              $isTested = $True
           }
           
           Kill-Excel
    }
    return $isTested
}
function Create-IQY($spList, $oWeb, $NameOfList)
{

    
    $crlf=[char][int]13+[char][int]10
    
    $header =  "WEB"+ $crlf+"1"+ $crlf
    $SharePointListView    = $("{"+ $splist.Views[0].Id.ToString()+"}").ToUpper()
    $SharePointListName    = $("{"+$splist.id+"}").ToUpper()
    $RootFolder            = $splist.RootFolder.ServerRelativeUrl
    $RootFolderSlash       = $splist.ParentWebUrl.Replace("/","%2f")
    $SharePointApplication = $splist.ParentWeb.Url+"/_vti_bin"
   
    $iqy  = "WEB"+ $crlf+"1"+ $crlf 
    $iqy += $SharePointApplication +"/owssvr.dll?XMLDATA=1&List=" + $SharePointListName + "&View=" + $SharePointListView + "&RowLimit=0&RootFolder="+$RootFolderSlash + $crlf+ $crlf
    $iqy += "Selection=" + $SharePointListName + "-" + $SharePointListView + $crlf
    $iqy += "EditWebPage="+ $crlf
    $iqy += "Formatting=None"+ $crlf
    $iqy += "PreFormattedTextToColumns=True"+ $crlf
    $iqy += "ConsecutiveDelimitersAsOne=True"+ $crlf
    $iqy += "SingleBlockTextImport=False"+ $crlf
    $iqy += "DisableDateRecognition=False"+ $crlf
    $iqy += "DisableRedirections=False"+ $crlf
    $iqy += "SharePointApplication="+ $SharePointApplication + $crlf
    
    $iqy += "SharePointListView=" + $SharePointListView + $crlf
    $iqy += "SharePointListName=" + $SharePointListName + $crlf
    $iqy += "RootFolder=" + $RootFolder+ $crlf+ $crlf
    
    Check-DirectoryExists $oWeb.iqyFolder
    
    $backupFolder = $oWeb.iqyFolder + "\" + $splist.ParentWebUrl.Split("/")[-1]
    Check-DirectoryExists $backupFolder
    
    $Filename = $oWeb.iqyFolder + "\" + $NameOfList + ".iqy"
    $backupFileName = $backupFolder+ "\" + $NameOfList + ".iqy"
    
    #$xlsIniFile = $oWeb.SiteCount.ToString() + $crlf + $oWeb.NewWebName
    #$xlsIniFileName = $oWeb.iqyFolder + "\xcr.ini"
   
    $iqy | out-file -encoding UTF8 -filepath $Filename
    $iqy | out-file -encoding UTF8 -filepath $backupFileName
    
    #$0 = $myInvocation.MyCommand.Definition
    #$dp0 = [System.IO.Path]::GetDirectoryName($0)
    
    $modules = @()
    
    $modules += "PvtExcel.ps1"
    $modules += "Utils-Ini.ps1"
    $modules += "Utils-Excel.ps1"
    $modules += "Tu.ini"
    $dest    =  $oWeb.iqyFolder + "\"
    foreach ($fil in $modules)
    {
           
           $source  =  $webObject.WrkDir + "\" + $fil
           
           Copy-Item -Path $source -destination $dest -Confirm:$false
    }
    
    #$xlsIniFile | out-file -encoding UTF8 -filepath $xlsIniFileName

   
    

}
function  Check-DirectoryExists($folderName)
{
    
    $IsDirectory = TEST-PATH $($folderName + "/")
    
    if (!$IsDirectory)
    {
         $noout = md $($folderName + "/")
    }
}  
#
#cls
function Create-Excel([string]$xlsFileName="Report.xlsx", $oweb )
{


        $DirItem = Get-Item $oweb.iqyFolder
        $dirXLS  = $DirItem.FullName + "\" 
        #$sourceXLS = "c:\ADMINDIR\UBZI\" + $xlsFileName
        $sourceXLS = $dirXLS + $xlsFileName

        $iqylist = @()

        $iqylist +="Admins"
        $iqylist +="Rooms"
        $iqylist +="SVTRestrict"
        $iqylist +="SVTs"
        WRITE-HOST $("Создаем файл : "+$sourceXLS)
        Kill-Excel


        Create-DestinationXLSSheets $sourceXLS $iqylist.Count


        #$iqyDir = "c:\ADMINDIR\iqy\"
        $iqyDir = $dirXLS
        #xlsDir = "c:\ADMINDIR\UBZI\Excel\"
        $xlsDir = $dirXLS

        foreach ($iqy in $iqylist)
        {
            $iqyFile = $iqyDir + $iqy + ".iqy"
            $xlsFile = $xlsDir + $iqy + ".xlsx"
            
            WRITE-HOST $("Создаем файл доступа к спискам: "+$xlsFile)
            Create-XLSQueryFromSPlist $iqyFile $xlsFile
            

        }

        $destXLS = $sourceXLS
        for($j=0; $j -lt $iqylist.count;$j++)
        {
           $sourceXLSiqy = $xlsDir + $iqylist[$j] + ".xlsx"  
           #WRITE-HOST $("Копируем таблицы доступа к спискам: "+$sourceXLSiqy)
           Copy-WorkSheets $sourceXLSiqy $destXLS $($j+1)
        }


        foreach ($iqy in $iqylist)
        {
              
              $xLSiqy = $xlsDir + $iqy + ".xlsx" 
              #WRITE-HOST $("Удаляем таблицы доступа к спискам: "+$xLSiqy)
              remove-Item $xLSiqy
              $xLSiqy = $xlsDir + $iqy + ".iqy"
              #remove-Item $xLSiqy
        }

        Write-host $("Создаем сводную таблицу.")

        Create-PivotTable $sourceXLS

        WRITE-HOST 
        WRITE-HOST $("Создан файл : "+$sourceXLS)  -ForegroundColor Yellow
        return $sourceXLS
}

