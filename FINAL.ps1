clear
#Ввод начальной даты
$D1 = Read-Host "Введите начальную дату"
#Ввод конечной даты
$D2 = Read-Host "Введите конечную дату"

#Преобразуем данные начальной и конечной даты в тип DateTime 
$DateStart = New-Object DateTime (($D1[4] + $D1[5] + $D1[6] + $D1[7]), ($D1[2] + $D1[3]), ($D1[0] + $D1[1]), 0, 0, 0)
$DateStop = New-Object DateTime (($D2[4] + $D2[5] + $D2[6] + $D2[7]), ($D2[2] + $D2[3]), ($D2[0] + $D2[1]), 0, 0, 0)

#Цикл на переход +1 день от начальной даты до конечной
for ($i = 0; $DateStart -le $DateStop; $i++)
{
    Write-Host ("Запрашиваемая дата: " + $DateStart)

    #Собираем ссылку сайта и путь папки для скачивания файла
    $SiteAdress = "https://www.atsenergo.ru/nreport?rname=big_nodes_prices_pub&region=eur&rdate=" + $DateStart.Year
    $OutDir = $env:USERPROFILE + "\desktop\Отчет о равновестных ценах\" + $DateStart.Year + "\"

    #Если месяц меньше 10, то к ссылке скачивания (т.к. такой шаблон ссылки) и имени папки прибавляется 0
    if ($DateStart.Month -lt 10) 
    {
        $SiteAdress = $SiteAdress + "0" + $DateStart.Month
        $OutDir = $OutDir + "0" + $DateStart.Month
    } else 
    {
        $SiteAdress = $SiteAdress + $DateStart.Month
        $OutDir = $OutDir + $DateStart.Month
    }
    
    #Если день меньше 10, то к ссылке скачивания (т.к. такой шаблон ссылки) прибавляется 0    
    if ($DateStart.Day -lt 10) 
    {
        $SiteAdress = $SiteAdress + "0" + $DateStart.Day
    } else 
    {
        $SiteAdress = $SiteAdress + $DateStart.Day
    }
    

    Write-Host ("Ссылка на сайт: " + $SiteAdress)

    #Получаем ответ от сервера по запрашиваемому сайту $SiteAdress
    $HttpContent = Invoke-WebRequest -URI $SiteAdress

  
    #Проверяем наличие папки для скачивания, если ее нет, создаем
    if ( -not (Test-Path $OutDir)) {New-Item -Path $OutDir -ItemType "directory" | Out-Null}

    #Делаем выборку по тексту ссылки и присваиваем его переменной NameFile, для имени скачиваемого файла
    $NameFile = $HttpContent.Links | Where-Object {$_.innerText -like "*_eur_big_nodes_prices_pub.xls*"} | fl innerText | Out-String
    #Делаем выборку по тексту ссылки и присваиваем переменной LinkFile ссылку на сам файл
    $LinkFile = $HttpContent.Links | Where-Object {$_.innerText -like "*_eur_big_nodes_prices_pub.xls*"} | fl href | Out-String


    #Если NameFile и LinkFile равно 0, то по текущей дате нет нужной ссылки на скачивание файла
    if (($NameFile -ne 0) -and ($LinkFile -ne 0))
    {
        #Отсекаем все лишнее в ссылке на файл и преобразуем в полный путь
        $LinkFile = $LinkFile.Substring(11)
        $LinkFile = $LinkFile.Remove(37)
        $LinkFile = "https://www.atsenergo.ru/nreport" + $LinkFile
        Write-Host ("Ссылка на файл: " + $LinkFile)

        #Отсекаем все лишнее в имени файла
        $NameFile = $NameFile.Substring(16)
        $NameFile = $NameFile.Remove(37)
    
        #Собираем путь для скачивания с именем файла
        $OutDir = $OutDir + "\" + $NameFile
        Write-Host ("Пусть до файла: " + $OutDir)

        #Загрузка файла
        Invoke-WebRequest $LinkFile -OutFile $OutDir
        Write-Host ("Готово")
    }

    $DateStart = $DateStart.AddDays(1)

}