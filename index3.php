<form method="post" action="" enctype="multipart/form-data">
    <input type="file" name="excel">
    <input type="submit" name="import" value="Импортировать">
</form>

<?php

if (isset($_POST['import']))
{

    $opt=[
        PDO::ATTR_ERRMODE=>PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE=>PDO::FETCH_ASSOC
    ];
    $dbh=new PDO("mysql:host=localhost;dbname=test","root","",$opt );
    //$dbh->query('select * from customers');
    //var_dump($_FILES);
    $name = $_FILES['excel']['name'];
    move_uploaded_file($_FILES['excel']['tmp_name'],$name);
    if (isset($name))
    {
        require_once ('Classes/PHPExcel.php');
        require_once ('Classes/PHPExcel/Writer/Excel5.php');

       // header('Content-Type: text/html; charset=utf-8');
        ini_set('error_reporting', E_ALL);
        ini_set('display_errors', 1);
        ini_set('display_startup_errors', 1);
        if (file_exists($name))
        {




            echo "The file $name exists";

            $pExcel = PHPExcel_IOFactory::load($name);
            $i = 0;

            foreach ($pExcel->getWorksheetIterator() as $worksheet)
            {
                $tables[] = $worksheet->toArray();
                $i++;
            }
            $count = 0;
            $n = 0;

            foreach ($tables as $table){
                foreach ($table as $row)
                {
                    if (!empty($row[0]) and $row[0] !='')
                    {
                        var_dump($row);
                        if ($row[1] !='name') {
                            $str = "insert into test.tovar(`name`, `price`,`img`) value ('{$row[1]}', '{$row[2]}', '{$row[3]}')";

                            $dbh->query($str);
                            //echo $str;
                        }
                    }
                }
            }
        }
    }
}