<?php
ini_set('max_execution_time', 6000);
$datos = array();
$simbolos=["TSLA","FB","V","BRK-B","WMT","AMZN","AAPL","UAA","GOOGL","MSFT","BABA","NFLX","SBUX","MELI","BKNG","EXPE","TRIP","TRVG","FSLR","SPWR","RUN","NVDA","INTC","AMD"];

$contar=0;
foreach ($simbolos as $simbolo){
    $txtCon = file_get_contents("https://finance.yahoo.com/quote/".$simbolo."?p=".$simbolo);

$txtInd = strpos($txtCon, "root.App.main");
$txtExt = substr($txtCon, $txtInd + 15);

$txtInd = strpos($txtExt, "}(this));");
$txtExt = substr($txtExt, 0, $txtInd - 2);

$txtJson = json_decode(utf8_encode($txtExt), true);
$arrDatos = array();
$aux1     = 0;
foreach($txtJson["context"]["dispatcher"]["stores"]["StreamDataStore"]["quoteData"] as $arreglo)
{
   if($arreglo["symbol"]==$simbolos[$contar]){
       array_push($datos,$arreglo["regularMarketPrice"]["fmt"]);
       break;
   }
}
$contar++;
}
?>
<!DOCTYPE html>
<html lang="es">  
  <head>    
    <title>Datos obtenidos de yahoo</title>    
    <meta charset="UTF-8">
    <meta name="title" content="Título de la WEB">
    <meta name="description" content="Descripción de la WEB">    
  </head>  
  <body>    
    <header>
      <h1>Datos obtenidos de yahoo</h1>      
    </header>    
    <section>      
      <article>
        <div>
<table>
  <tr>
    <th>#</th>
    <th>Simbolo</th>
    <th>Valor</th>
  </tr>
<?php
$contar=0;
foreach ($simbolos as $simbolo) {
    echo "<tr><td>".($contar+1)."</td><td>".$simbolo."</td><td>".$datos[$contar]."</td></tr>";
    $contar++;
}
?>
      
</table>
        </div>
      </article>      
    </section>
  </body>  
</html>
<?php
	//Agregamos la librería para leer
	require 'PHPExcel/Classes/PHPExcel/IOFactory.php';
	
	// Creamos un objeto PHPExcel
	$objPHPExcel = new PHPExcel();
	$objReader = PHPExcel_IOFactory::createReader('Excel2007');
	$objPHPExcel = $objReader->load('DatosBolsa.xlsx');
	// Indicamos que se pare en la hoja uno del libro
	$objPHPExcel->setActiveSheetIndex(0);
	
	//Modificamos los valoresde las celdas A2, B2 Y C2
        $celda="";
        $contar=0;
        for($i=27;$i<=50;$i++){
            $celda="B".$i;
	$objPHPExcel->getActiveSheet()->SetCellValue($celda, $datos[$contar]);
        $contar++;
        }

	
	//Guardamos los cambios
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	$objWriter->save("DatosBolsa.xlsx");
        
        echo "<br>El archivo se a actualizado con exito";
?>