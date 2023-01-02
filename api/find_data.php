<?php
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Headers: access");
header("Access-Control-Allow-Methods: GET,POST");
header("Access-Control-Allow-Credentials: true");
header('Content-Type: application/json;charset=utf-8');
require_once __DIR__ . '/../api/web_config.php';

set_time_limit ( 60000 );

use \Psr\Http\Message\ResponseInterface as Response; // ไลบราลี้สำหรับจัดการคำร้องขอ
use \Psr\Http\Message\ServerRequestInterface as Request; // ไลบราลี้สำหรับจัดการคำตอบกลับ

require './vendor/autoload.php'; // ดึงไฟ์ autoload.php เข้ามา
//include_once './class.oracle.php'; // Class Connect Oracle
include_once './util.php'; // ดึงไฟ์ util.php เข้ามา
include_once './web_config.php';
$app = new \Slim\App; // สร้าง object หลักของระบบ

date_default_timezone_set("Asia/Bangkok");

function ConnectDbAll($_sql){
  $DataRows = array();
  $conn = ConnectedDBSO();
  if(!$conn){
    $_err = oci_error();
    echo $_err;
  }else{
    $objParse = oci_parse($conn,$_sql);
    $objEx = oci_execute($objParse);
    if($objEx){
      $objResult = oci_fetch_all($objParse,$DataRows,null,null, OCI_FETCHSTATEMENT_BY_ROW);
    }else{
      echo "Connect Data Base error";
    }
  }
  oci_close($conn);
  return $DataRows;
 
}
function ConnectDbnoAll($_sql){
  $DataRows = array();
  $conn = ConnectedDBSO();
  if(!$conn){
    $_err = oci_error();
    echo $_err;
  }else{
    $objParse = oci_parse($conn,$_sql);
    $objEx = oci_execute($objParse);
    if($objEx){
      
    }else{
      echo "Connect Data Base error";
    }
  }
  oci_close($conn);
  return $DataRows;
  
}

$app->post('/export_data', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    

    
    
  $_sql = "SELECT * FROM OE_RECUT_CHANGE_V WHERE TRUNC(SHIPMENT_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD') and ORG = '".$org."'
  order by SHIPMENT_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/find_item', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    

    
    
  $_sql = "SELECT distinct SO_NO SO_NO,SHIPMENT_DATE  FROM OE_RECUT_CHANGE_V WHERE TRUNC(SHIPMENT_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD') and ORG = '".$org."'
  ORDER BY SHIPMENT_DATE";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });




  $app->post('/find_data', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    $so_no = isset($_REQUEST['so_no']) ? $_REQUEST['so_no'] : '';
    $item_code = isset($_REQUEST['item_code']) ? $_REQUEST['item_code'] : '';

    

    
    
  $_sql = "SELECT * FROM OE_RECUT_CHANGE_V WHERE TRUNC(SHIPMENT_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD') and ORG = '".$org."' and SO_NO = '".$so_no."'
  and ITEM_CODE = '".$item_code."'
  order by SHIPMENT_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/find_distinct_item_code', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    $so_no = isset($_REQUEST['so_no']) ? $_REQUEST['so_no'] : '';
   
    

    
    
  $_sql = "SELECT distinct ITEM_CODE ITEM_CODE,SO_NO FROM OE_RECUT_CHANGE_V WHERE TRUNC(SHIPMENT_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD') and ORG = '".$org."' 
  and SO_NO = '".$so_no."'";
  
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/find_data_item_code', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    $so_no = isset($_REQUEST['so_no']) ? $_REQUEST['so_no'] : '';
    $item_code = isset($_REQUEST['item_code']) ? $_REQUEST['item_code'] : '';
    

    
    
    $_sql = "SELECT * FROM OE_RECUT_CHANGE_V WHERE TRUNC(SHIPMENT_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD') and ORG = '".$org."' and SO_NO = '".$so_no."'
    and ITEM_CODE = '".$item_code."'
    order by SHIPMENT_DATE ASC";
  
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/find_data_na', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    
    
    $_sql = "SELECT * FROM OE_RECUT_CHANGE_V WHERE TRUNC(SHIPMENT_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD') and ORG = '".$org."'
    order by SHIPMENT_DATE ASC";
  
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

$app->run(); //สั่งระบบให้ทำงาน



?>

