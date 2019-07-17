##############################################
#
#   Модуль создания сайтов технологических участков
#
#   Файл: Util-SampleData.ps1 
#   Описание : Функции для наполнения тестовыми данными 
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

function Add-TuUsersSampleData($splist)
{

$list = @()

$list += "Иванов И.И."
$list += "Петров П.П."
$list += "Кузнецова К.К."
$list += "Семенова С.С."
$list += "Антонова А.А."
$list += "Михайлова М.М."
$list += "Дегунина Д.Д."
$list += "Бескудников Б.Б."

foreach ($el in $list){
            $NewItem = $SpList.Items.Add();
            $NewItem["Сотрудник Ф.И.О."] = $el
            $NewItem.Update();
        }
}

function Add-UsersSampleData($SpList)
{

$list = @()

  $el = "" | select Room, dAtt,nAkt, sDost, sCont, kZam, Opech, Oxr
  $el.Room = "3475"
  
  
  $el.nAkt = "ВН-65-24-3-8/9998"
  
  
  $el.sCont = "Установлена"
  $el.kZam  = "Не установлен"
  $el.Opech = "Выполняется"
  $el.Oxr   = "Электронная ключница"
  
  $NewItem = $SpList.Items.Add();
  $NewItem["Номер помещения"] = $el.Room
  [DateTime]$d = get-Date -Year 2014 -Month 12 -Day 17
  $NewItem["Дата аттестации"] = $d
  $NewItem["Номер акта аттестации"] = $el.nAkt
  [DateTime]$d = Get-Date -Year 2015 -Month 11 -Day 27
  $NewItem["Наличие списка доступа"] = $d
  $NewItem["Система контроля и управления доступом"] = $el.sCont
  $NewItem["Кодовый замок"] = $el.kZam
  $NewItem["Опечатывание"] = $el.Opech
  $NewItem["Сдача помещения под охрану"] = $el.Oxr
  $NewItem.Update();
  
  $el = "" | select Room, dAtt,nAkt, sDost, sCont, kZam, Opech, Oxr
  $el.Room = "1141"
  
  
  $el.nAkt = "ВН-65-24-3-8/9997"
  
  
  $el.sCont = "Установлена"
  $el.kZam  = "Установлен"
  $el.Opech = "Выполняется"
  $el.Oxr   = "На пост охраны"
   
  $NewItem = $SpList.Items.Add();
  $NewItem["Номер помещения"] = $el.Room
  [DateTime]$d = get-Date -Year 2014 -Month 12 -Day 24
  $NewItem["Дата аттестации"] = $d

  $NewItem["Номер акта аттестации"] = $el.nAkt
  [DateTime]$d = Get-Date -Year 2015 -Month 10 -Day 14
  $NewItem["Наличие списка доступа"] = $d
  $NewItem["Система контроля и управления доступом"] = $el.sCont
  $NewItem["Кодовый замок"] = $el.kZam
  $NewItem["Опечатывание"] = $el.Opech
  $NewItem["Сдача помещения под охрану"] = $el.Oxr
  $NewItem.Update();
   
  $el = "" | select Room, dAtt,nAkt, sDost, sCont, kZam, Opech, Oxr
  $el.Room = "2202"
  #[DateTime]$el.dAtt = 
  
  $el.nAkt = "ВН-65-24-3-8/9996"
  #[DateTime]$el.sDost = 
  
  $el.sCont = "Установлена"
  $el.kZam  = "Не установлен"
  $el.Opech = "Выполняется"
  $el.Oxr   = "Нет"  
  
  $NewItem = $SpList.Items.Add();
  $NewItem["Номер помещения"] = $el.Room
  [DateTime]$d = get-Date -Year 2013 -Month 2 -Day 24
  $NewItem["Дата аттестации"] = $d
  $NewItem["Номер акта аттестации"] = $el.nAkt
  [DateTime]$d = Get-Date -Year 2014 -Month 5 -Day 23
  $NewItem["Наличие списка доступа"] = $d
  $NewItem["Система контроля и управления доступом"] = $el.sCont
  $NewItem["Кодовый замок"] = $el.kZam
  $NewItem["Опечатывание"] = $el.Opech
  $NewItem["Сдача помещения под охрану"] = $el.Oxr
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
            $NewItem["№ ТМ-идентификатора"] = $el
            $NewItem.Update();
        }


}
function Add-AdminsSampleData($SpList)
{

       $NewItem = $SpList.Items.Add(); 
       
        $NewItem["Сотрудник Ф.И.О."] = "Сотников С.А."
        [DateTime]$d = Get-Date -Year 2016 -Month 7 -Day 18  # Дата и время
        $NewItem["Дата ознакомления"]  =  $d 
        $NewItem["Ознакомился"] = "Да"#  
        $NewItem["Роль"] = "Резервный АИБ"# Выбор   
        $NewItem["Должность"]  = "ведущий инженер"# Однострочный текст 
        [DateTime]$d = Get-Date -Year 2015 -Month 9 -Day 30  # Дата и время   # Дата и время 
        $NewItem["Дата приказа о назначении"] = $d   
        $NewItem["Номер приказа о назначении"]  = "№ ОДТ-1-29-65/456"
       
       $NewItem.Update(); 


       $NewItem = $SpList.Items.Add(); 
       
        $NewItem["Сотрудник Ф.И.О."] = "Рожкова Т.М."
        [DateTime]$d = Get-Date -Year 2016 -Month 7 -Day 21  # Дата и время  
        $NewItem["Дата ознакомления"]  = $d
        $NewItem["Ознакомился"] = "Да"#  
        $NewItem["Роль"] = "АИБ"# Выбор   
        $NewItem["Должность"]  = "главный инженер"# Однострочный текст 
        [DateTime]$d =   Get-Date -Year 2015 -Month 11 -Day 19  # Дата и время   # Дата и время  
        $NewItem["Дата приказа о назначении"] =  $d
        $NewItem["Номер приказа о назначении"]  = "№ ОДТ-1-29-65/876"
       
       $NewItem.Update(); 


       $NewItem = $SpList.Items.Add(); 
       
        $NewItem["Сотрудник Ф.И.О."] = "Мухин В.В."
        [DateTime]$d = Get-Date -Year 2016 -Month 3 -Day 10  # Дата и время  
        $NewItem["Дата ознакомления"]  = $d
        $NewItem["Ознакомился"] = "Да"#  
        $NewItem["Роль"] = "АИБ"# Выбор   
        $NewItem["Должность"]  = "инженер первой категории"# Однострочный текст 
        [DateTime]$d = Get-Date -Year 2016 -Month 4 -Day 18 # Дата и время   # Дата и время   
        $NewItem["Дата приказа о назначении"] =  $d  
        $NewItem["Номер приказа о назначении"]  = "№ ОДТ-1-29-65/951"
       
       $NewItem.Update(); 


}
function Add-SVTRestrictSampleData($SpList)
{
      $NewItem = $SpList.Items.Add(); 
      
            
      
      $NewItem["Инвентарный номер СВТ"]  =  "1;#25736" # Подстановка   
      $NewItem["Ф.И.О. допущенного к СВТ"]  = "7;#Дегунина Д.Д."# Подстановка   
      $NewItem["Имя уч. записи в СЗИ от НСД"]  =  "Гл. администратор"# Однострочный текст   
      $NewItem["№ ТМ-идентификатора"]  = "9;#08 000000E33AB4 D2" #Подстановка   
      $NewItem["Права доступа"]  =  "Рез. администратор"# Выбор  
      

      $NewItem.Update(); 
      
      $NewItem = $SpList.Items.Add(); 
      
            
      #$NewItem["№ п/п"]  = "13"                                   #  Однострочный текст   
      $NewItem["Инвентарный номер СВТ"]  =  "1;#25736"            # Подстановка   
      $NewItem["Ф.И.О. допущенного к СВТ"]  = "6;#Михайлова М.М." # Подстановка   
      $NewItem["Имя уч. записи в СЗИ от НСД"]  =  "Администратор" # Однострочный текст   
      $NewItem["№ ТМ-идентификатора"]  = "12;#0C 0000001A6116 13"       #Подстановка   
      $NewItem["Права доступа"]  =  "Пользователь"# Выбор  
      # $NewItem["Назначение СВТ"]  =  # Подстановка 

      
      $NewItem.Update(); 
           
      
}
function Add-SVTsampleData($SpList)
{

        $ctype = $splist.ContentTypes["Набор документов"]
 
        [HashTable]$docsetpropertes = @{}

        $folder = [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]::Create($splist.RootFolder,"Документы_25736",$ctype.ID, $docsetpropertes)
        


        $item = $splist.Items | Where {$_["Имя"] -eq "Документы_25736"}

     
        $Item["BBK"]  = "Антивирус Касперского 10"#  Выбор Набор документов  
        $Item["АС"]  = "АС ЭКР 2.0; АРМ администратора СУБД АС ЭКР"  #  Однострочный текст Набор документов  
        $Item["Заводской номер"]  = "SKAM192345"  #  Однострочный текст Набор документов  
        $Item["Заводской номер СЗИ"]  = "92116886" #  Однострочный текст Набор документов  
        $Item["Инв. номер СЗИ"]  =  "31313"#  Однострочный текст Набор документов  
        $Item["Инвентарный номер"]  =  "25736"#  Однострочный текст Набор документов  
        $Item["Название"]  =   "Документы_25736"#  Однострочный текст Набор документов  
        $Item["Номер помещения"]  = "1141" #  Однострочный текст Набор документов  
        $Item["Операционная система"]  = "Windows 7"  #  Выбор Набор документов  
        $Item["Описание"]  = "АС ЭКР 2.0; АРМ администратора СУБД АС ЭКР" #  Многострочный текст Набор документов  
        $Item["СЗИ от НСД"]  = "Аккорд-АМДЗ"                              #  Выбор Набор документов  
        $Item["Тип ПВМ"]  = "Fujitsu-Siemens Esprimo"                     #  Однострочный текст Набор документов  
        $Item["Ф.И.О. Ответственного"]  = "Кузнецова К.К."                #  Однострочный текст 

        $Item.Update();


        $folder = [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]::Create($splist.RootFolder,"Документы_36293",$ctype.ID, $docsetpropertes)

        $item = $splist.Items | Where {$_["Имя"] -eq "Документы_36293"}
         
        $Item["BBK"]  = "Антивирус Касперского 8"#  Выбор Набор документов  
        $Item["АС"]  = "АС ЭКР 2.0; Резервный сервер БД Oracle + WAS+MQ"  #  Однострочный текст Набор документов  
        $Item["Заводской номер"]  = "CZJ838B2NX"  #  Однострочный текст Набор документов  
        $Item["Заводской номер СЗИ"]  = "257Б-25869" #  Однострочный текст Набор документов  
        $Item["Инв. номер СЗИ"]  =  "31316"#  Однострочный текст Набор документов  
        $Item["Инвентарный номер"]  =  "36293"#  Однострочный текст Набор документов  
        $Item["Название"]  =   "Документы_36293"#  Однострочный текст Набор документов  
        $Item["Номер помещения"]  = "3475" #  Однострочный текст Набор документов  
        $Item["Операционная система"]  = "Windows 7"  #  Выбор Набор документов  
        $Item["Описание"]  = "АС ЭКР 2.0; Резервный сервер БД Oracle + WAS+MQ" #  Многострочный текст Набор документов  
        $Item["СЗИ от НСД"]  = "Аккорд-АМДЗ" #  Выбор Набор документов  
        $Item["Тип ПВМ"]  = "HP 7800"#  Однострочный текст Набор документов  
        $Item["Ф.И.О. Ответственного"]  = "Бескудников Б.Б." #  Однострочный текст 

        $Item.Update(); 
       
}
