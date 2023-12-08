<?php
require __DIR__ . '/../vendor/autoload.php';

use Dotenv\Dotenv;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$env = Dotenv::createImmutable(__DIR__ . '/../');
$env->load();

use MrShan0\PHPFirestore\FirestoreClient;

$firestoreClient = new FirestoreClient($_ENV['FIREBASE_PROJECT_ID'], $_ENV['GCLOUD_FIREBASE_API'], [
    'database' => '(default)',
]);
$collection = $firestoreClient->listDocuments('users');
$collection = json_encode($collection['documents']);

// Decode the JSON data
$data = json_decode($collection, true);
// turn off notices
error_reporting(E_ALL & ~E_NOTICE);
$regroupedData = [];
foreach ($data as $entry) {
    $fields = $entry['fields'];
    $info = $fields['info']['mapValue']['fields'];

    $regroupedData[] = [
        'uId' => $fields['uId']['stringValue'],
        'email' => $fields['email']['stringValue'],
        'role' => $fields['role']['stringValue'],
        'status' => $fields['status']['booleanValue'],
        'created_at' => $fields['created_at']['timestampValue'],
        'info' => [
            'name' => $info['name']['stringValue'],
            'phone' => $info['phone']['stringValue'],
            'age' => $info['age']['integerValue'],
            'birthDate' => $info['birthDate']['timestampValue'],
            'eduLevel' => $info['eduLevel']['stringValue'],
            'shirtSize' => $info['shirtSize']['stringValue'],
            'contacts' => [
                'parentContact' => $info['contacts']['mapValue']['fields']['parentContact']['stringValue'],
                'facebookLink' => $info['contacts']['mapValue']['fields']['facebookLink']['stringValue'],
                'contractEmail' => $info['contacts']['mapValue']['fields']['contractEmail']['stringValue'],
                'lineId' => $info['contacts']['mapValue']['fields']['lineId']['stringValue'],
                'otherContact' => $info['contacts']['mapValue']['fields']['otherContact']['stringValue'],
            ],
        ],
        'assets' => [
            'paymentReceiptSrc' => $fields['assets']['mapValue']['fields']['paymentReceiptSrc']['nullValue'],
            'parentPermissionSrc' => $fields['assets']['mapValue']['fields']['parentPermissionSrc']['nullValue'],
        ],
    ];
}

// Convert the regrouped data to JSON
$jsonData = json_encode($regroupedData, JSON_PRETTY_PRINT);

$spreadsheet = new Spreadsheet();
$activeWorksheet = $spreadsheet->getActiveSheet();
$activeWorksheet->setCellValue('A1', 'รายชื่อผู้สมัครค่ายคอมแคมป์ ครั้งที่ 22');
// set font size A1 to 20 bold
$activeWorksheet->getStyle('A1')->getFont()->setSize(20);

$activeWorksheet->setCellValue('A3', 'ลำดับ');
$activeWorksheet->setCellValue('B3', 'ชื่อ-นามสกุล');
$activeWorksheet->setCellValue('C3', 'อายุ');
$activeWorksheet->setCellValue('D3', 'เบอร์โทรศัพท์');
$activeWorksheet->setCellValue('E3', 'ไลน์');
$activeWorksheet->setCellValue('F3', 'Facebook');
$activeWorksheet->setCellValue('G3', 'อีเมล');
$activeWorksheet->setCellValue('H3', 'ขนาดเสื้อ');
$activeWorksheet->setCellValue('I3', 'ระดับชั้น');
$activeWorksheet->setCellValue('J3', 'ชื่อผู้ปกครอง');
$activeWorksheet->setCellValue('K3', 'เบอร์โทรศัพท์ผู้ปกครอง');
$activeWorksheet->setCellValue('L3', 'อีเมลผู้ปกครอง');
$activeWorksheet->setCellValue('M3', 'อื่นๆ');
$activeWorksheet->setCellValue('N3', 'สถานะการชำระเงิน');
$activeWorksheet->setCellValue('O3', 'สถานะการสมัคร');
$activeWorksheet->setCellValue('P3', 'วันที่สมัคร');

$activeWorksheet->getColumnDimension('A')->setWidth(10);
$activeWorksheet->getColumnDimension('B')->setWidth(30);
$activeWorksheet->getColumnDimension('C')->setWidth(10);
$activeWorksheet->getColumnDimension('D')->setWidth(20);
$activeWorksheet->getColumnDimension('E')->setWidth(20);
$activeWorksheet->getColumnDimension('F')->setWidth(20);
$activeWorksheet->getColumnDimension('G')->setWidth(30);
$activeWorksheet->getColumnDimension('H')->setWidth(10);
$activeWorksheet->getColumnDimension('I')->setWidth(20);
$activeWorksheet->getColumnDimension('J')->setWidth(30);
$activeWorksheet->getColumnDimension('K')->setWidth(20);
$activeWorksheet->getColumnDimension('L')->setWidth(30);
$activeWorksheet->getColumnDimension('M')->setWidth(30);
$activeWorksheet->getColumnDimension('N')->setWidth(20);
$activeWorksheet->getColumnDimension('O')->setWidth(20);
$activeWorksheet->getColumnDimension('P')->setWidth(20);

$startRow = 4;
// insert data from collection
foreach ($regroupedData as $userData) {
    $activeWorksheet->setCellValue("A" .  $startRow, $userData['uId']);
    $activeWorksheet->setCellValue("B" . $startRow, $userData['info']['name']);
    $activeWorksheet->setCellValue("C" . $startRow, $userData['info']['age']);
    $activeWorksheet->setCellValue("D" . $startRow, $userData['info']['phone']);
    $activeWorksheet->setCellValue("E" . $startRow, $userData['info']['contacts']['lineId']);
    $activeWorksheet->setCellValue("F" . $startRow, $userData['info']['contacts']['facebookLink']);
    $activeWorksheet->setCellValue("G" . $startRow, $userData['email']);
    $activeWorksheet->setCellValue("H" . $startRow, $userData['info']['shirtSize']);
    $activeWorksheet->setCellValue("I" . $startRow, $userData['info']['eduLevel']);
    $activeWorksheet->setCellValue("J" . $startRow, $userData['info']['contacts']['parentContact']);
    $activeWorksheet->setCellValue("K" . $startRow, $userData['info']['contacts']['parentContact']);
    $activeWorksheet->setCellValue("L" . $startRow, $userData['info']['contacts']['contractEmail']);
    $activeWorksheet->setCellValue("M" . $startRow, $userData['info']['contacts']['otherContact']);
    $activeWorksheet->setCellValue("N" . $startRow, $userData['status'] ? 'ชำระเงินแล้ว' : 'ยังไม่ชำระเงิน');
    $activeWorksheet->setCellValue("O" . $startRow, $userData['role'] === 'user' ? 'ยังไม่สมัคร' : 'สมัครแล้ว');
    $activeWorksheet->setCellValue("P" . $startRow, $userData['created_at']);
    $startRow++;
}

$activeWorksheet->setCellValue("A" . ($startRow+1), 'ข้อมูล ณ วันที่ ' . date('d/m/Y H:i:s') . ' น.');


$writer = new Xlsx($spreadsheet);
$writer->save('Comcamp22_Exported.xlsx');



// Output the JSON
header('Content-type: application/json');
echo $jsonData;