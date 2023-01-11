<?php

include './vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;

$connection = new PDO("mysql:host=localhost;dbname=phpspreadsheet", "root", "");
$query = "SELECT country.Code, country.Name, country.Continent, country.Region FROM country ORDER BY country.Code DESC LIMIT 50";
$statement = $connection->prepare($query);
$statement->execute();

$result = $statement->fetchAll();
if(isset($_POST["export"])) {

  $file = new Spreadsheet();
  $active_sheet = $file->getActiveSheet();

  $active_sheet->setCellValue('A1', 'Country Name');
  $active_sheet->setCellValue('B1', 'Country Code');
  $active_sheet->setCellValue('C1', 'Continent');
  $active_sheet->setCellValue('D1', 'Region');

  $count = 2;

  foreach($result AS $row) {
    $active_sheet->setCellValue('A' . $count, $row["Name"]);
    $active_sheet->setCellValue('B' . $count, $row["Code"]);
    $active_sheet->setCellValue('C' . $count, $row["Continent"]);
    $active_sheet->setCellValue('D' . $count, $row["Region"]);
    $count = $count + 1;
  }

  $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($file, $_POST["file_type"]);

 
  $file_name = time() . '.' . strtolower($_POST["file_type"]);
    $writer->save($file_name);


  header('Content-Type: application/x-www-form-urlencoded');
  header('Content-Transfer-Encoding: Binary');
  header("Content-disposition: attachment; filename=\"".$file_name."\"");
  readfile($file_name);
  unlink($file_name);
  exit;
}

?>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>EXPORT DATA FROM MYSQL TO EXCEL  AND CSV USING PHPSPREADSHEET</title>
  <!-- bootstrap css -->
  <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
</head>
<body>
<div class="container">
  <div class="mt-5 mb-4">
  </div>
  <div class="row">
    <div class="col-md-10 offset-1">
      <div class="card">
        <div class="card-body">
          <h4 class="card-title">
            <form method="POST">
              <div class="row">
                <div class="col-md-6">World Data</div>
                <div class="col-md-4">
                  <select name="file_type" id="file_type" class="form-control input-sm">
                    <option value="Xlsx">Xlsx</option>
                    <option value="Xls">Xls</option>
                    <option value="Csv">Csv</option>
                  </select>
                </div>
                <div class="col-md-2">
                  <input type="submit" name="export" value="Export" class="btn btn-primary btn-sm">
                </div>
              </div>
            </form>
          </h4>
          <div class="table-responsive mt-3">
            <table class="table table-striped table-bordered">
              <thead>
                <th>Country Name</th>
                <th>Country Code</th>
                <th>Continent</th>
                <th>Region</th>
              </thead>
              <tbody>
                <?php foreach($result AS $row){
                  echo '
                  <tr>
                    <td>'. $row['Name'] .'</td>
                    <td>'. $row['Code'] .'</td>
                    <td>'. $row['Continent'] .'</td>
                    <td>'. $row['Region'] .'</td>
                  </tr>
                  ';
                }
                 ?>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>
  <div class="row">
    <div class="col-md-6 offset-5 mt-5 mb-5">
    </div>
  </div>
</div>


<!-- javascript libraries -->
<script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
</body>
</html>