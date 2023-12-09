<?php
require __DIR__ . '/../vendor/autoload.php';

error_reporting(E_ALL & ~E_WARNING & ~E_NOTICE);

use Dotenv\Dotenv;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use MrShan0\PHPFirestore\FirestoreClient;

// set php time to Asia/Bangkok
date_default_timezone_set('Asia/Bangkok');

class FirebaseDataExporter
{
    private $firestoreClient;

    public function __construct()
    {
        $env = Dotenv::createImmutable(__DIR__ . '/../');
        $env->load();

        $this->firestoreClient = new FirestoreClient($_ENV['FIREBASE_PROJECT_ID'], $_ENV['GCLOUD_FIREBASE_API'], [
            'database' => '(default)',
        ]);
    }

    public function fetchDataFromFirebase()
    {
        $collection = $this->firestoreClient->listDocuments('users');
        $collection = json_encode($collection['documents']);

        $data = json_decode($collection, true);

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
                    'nickname' => $info['nickname']['stringValue'],
                    'phone' => $info['phone']['stringValue'],
                    'age' => $info['age']['integerValue'],
                    'birthDate' => $info['birthDate']['timestampValue'],
                    'eduLevel' => $info['eduLevel']['stringValue'],
                    'shirtSize' => $info['shirtSize']['stringValue'],
                    'foodAllergy' => $info['foodAllergy']['stringValue'],
                    'congenitalDisease' => $info['congenitalDisease']['stringValue'],
                    'drugAllergy' => $info['drugAllergy']['stringValue'],
                    'hasLaptop' => $info['haveLaptop']['booleanValue'],
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

        return $regroupedData;
    }

    public function exportToSpreadsheet($regroupedData, $status = false)
    {
        // if status is true, only export users who have paid
        if ($status) {
            $regroupedData = array_filter($regroupedData, function ($userData) {
                return $userData['status'];
            });
        }

        $spreadsheet = new Spreadsheet();
        $activeWorksheet = $spreadsheet->getActiveSheet();

        $activeWorksheet->setCellValue('A1', 'รายชื่อผู้สมัครค่ายคอมแคมป์ ครั้งที่ 22');
        $activeWorksheet->getStyle('A1')->getFont()->setSize(20);

        $columnHeaders = [
            'A' => 'ลำดับ',
            'B' => 'ชื่อ-นามสกุล',
            'C' => 'อายุ',
            'D' => 'เบอร์โทรศัพท์',
            'E' => 'ไลน์',
            'F' => 'Facebook',
            'G' => 'อีเมล',
            'H' => 'โรคประจำตัว',
            'I' => 'แพ้ยา',
            'J' => 'แพ้อาหาร',
            'K' => 'มีคอมพิวเตอร์ส่วนตัว',
            'L' => 'ขนาดเสื้อ',
            'M' => 'ระดับชั้น',
            'N' => 'ชื่อผู้ปกครอง',
            'O' => 'เบอร์โทรศัพท์ผู้ปกครอง',
            'P' => 'อีเมลผู้ปกครอง',
            'Q' => 'อื่นๆ',
            'R' => 'สถานะการชำระเงิน',
            'S' => 'วันที่สมัคร',
        ];
        $row = 3;
        foreach ($columnHeaders as $column => $header) {
            $activeWorksheet->setCellValue($column . $row, $header);
        }

        $columnWidths = [
            'A' => 10,
            'B' => 30,
            'C' => 10,
            'D' => 20,
            'E' => 20,
            'F' => 20,
            'G' => 30,
            'H' => 30,
            'I' => 30,
            'J' => 30,
            'K' => 10,
            'L' => 10,
            'M' => 20,
            'N' => 30,
            'O' => 20,
            'P' => 30,
            'Q' => 30,
            'R' => 20,
            'S' => 20,
        ];

        foreach ($columnWidths as $column => $width) {
            $activeWorksheet->getColumnDimension($column)->setWidth($width);
        }

        $startRow = 4;
        foreach ($regroupedData as $key => $userData) {
            $key++;
            $activeWorksheet->setCellValue("A" .  $startRow, $key);
            $activeWorksheet->setCellValue("B" . $startRow, $userData['info']['name']);
            $activeWorksheet->setCellValue("C" . $startRow, $userData['info']['age']);
            $activeWorksheet->setCellValue("D" . $startRow, $userData['info']['phone']);
            $activeWorksheet->setCellValue("E" . $startRow, $userData['info']['contacts']['lineId']);
            $activeWorksheet->setCellValue("F" . $startRow, $userData['info']['contacts']['facebookLink']);
            $activeWorksheet->setCellValue("G" . $startRow, $userData['email']);
            $activeWorksheet->setCellValue("H" . $startRow, $userData['info']['congenitalDisease']);
            $activeWorksheet->setCellValue("I" . $startRow, $userData['info']['drugAllergy']);
            $activeWorksheet->setCellValue("J" . $startRow, $userData['info']['foodAllergy']);
            $activeWorksheet->setCellValue("K" . $startRow, $userData['info']['hasLaptop'] ? 'มี' : 'ไม่มี');
            $activeWorksheet->setCellValue("L" . $startRow, $userData['info']['shirtSize']);
            $activeWorksheet->setCellValue("M" . $startRow, $userData['info']['eduLevel']);
            $activeWorksheet->setCellValue("N" . $startRow, $userData['info']['contacts']['parentContact']);
            $activeWorksheet->setCellValue("O" . $startRow, $userData['info']['contacts']['parentContact']);
            $activeWorksheet->setCellValue("P" . $startRow, $userData['info']['contacts']['contractEmail']);
            $activeWorksheet->setCellValue("Q" . $startRow, $userData['info']['contacts']['otherContact']);
            $activeWorksheet->setCellValue("R" . $startRow, $userData['status'] ? 'ชำระเงินแล้ว' : 'ยังไม่ชำระเงิน');
            // firebase datetimestamp to human readable date
            $activeWorksheet->setCellValue("S" . $startRow, $userData['created_at']);
            $startRow++;
        }
        $activeWorksheet->setCellValue("A" . ($startRow + 1), 'ข้อมูล ณ วันที่ ' . date('d/m/Y H:i:s') . ' น.');
        $activeWorksheet->mergeCells("A" . ($startRow + 1) . ":" . "B" . ($startRow + 1));

        $summaryData = [
            'รวมผู้สมัครทั้งหมด' => count($regroupedData) . ' คน',
            'รวมผู้สมัครที่ชำระเงินแล้ว' => count(array_filter($regroupedData, function ($userData) {
                return $userData['status'];
            })) . ' คน',
            'รวมผู้สมัครที่ยังไม่ชำระเงิน' => count(array_filter($regroupedData, function ($userData) {
                return !$userData['status'];
            })) . ' คน'
        ];

        $row = $startRow + 2;
        foreach ($summaryData as $label => $value) {
            $activeWorksheet->setCellValue("A" . $row, $label);
            $activeWorksheet->setCellValue("C" . $row, $value);
            $row++;
        }


        $writer = new Xlsx($spreadsheet);
        $writer->save('Comcamp22_Exported_All.xlsx');
        rename('Comcamp22_Exported_All.xlsx', './storage/Comcamp22_Exported_All.xlsx');

    }

    // export selected fields by array parameters to Spreadsheet
    public function exportSelectedFieldsToSpreadsheet($regroupedData, $fields, $status = true)
    {
        // if status is true, only export users who have paid
        if ($status) {
            $regroupedData = array_filter($regroupedData, function ($userData) {
                return $userData['status'];
            });
        }

        $spreadsheet = new Spreadsheet();
        $activeWorksheet = $spreadsheet->getActiveSheet();

        // exstract data by selected fields and if there is a nested field, extract it too
        $extractedData = [];
        foreach ($regroupedData as $key => $userData) {
            $extractedData[$key] = [];
            // if there is a nested field, extract it too
            if (is_array($userData)) {
                foreach ($userData as $field => $value) {
                    if (is_array($value)) {
                        foreach ($value as $nestedField => $nestedValue) {
                            if (in_array($nestedField, $fields)) {
                                $extractedData[$key][$nestedField] = $nestedValue;
                            }
                        }
                    } else {
                        if (in_array($field, $fields)) {
                            $extractedData[$key][$field] = $value;
                        }
                    }
                }
            }
        }

        // translate all fields to thai
        $thaiFields = [
            'uId' => 'ไอดีผู้สมัคร',
            'email' => 'อีเมล',
            'status' => 'สถานะการชำระเงิน',
            'created_at' => 'วันที่สมัคร',
            'birthDate' => 'วันเกิด',
            'name' => 'ชื่อ-นามสกุล',
            'nickname' => 'ชื่อเล่น',
            'phone' => 'เบอร์โทรศัพท์',
            'age' => 'อายุ',
            'eduLevel' => 'ระดับชั้น',
            'shirtSize' => 'ขนาดเสื้อ',
            'foodAllergy' => 'แพ้อาหาร',
            'congenitalDisease' => 'โรคประจำตัว',
            'drugAllergy' => 'แพ้ยา',
            'hasLaptop' => 'มีคอมพิวเตอร์ส่วนตัว',
            'parentContact' => 'ชื่อผู้ปกครอง',
            'facebookLink' => 'Facebook',
            'contractEmail' => 'อีเมลผู้ปกครอง',
            'lineId' => 'ไลน์',
            'otherContact' => 'อื่นๆ',
        ];

        // set column headers by selected fields
        $columnHeaders = [];
        foreach ($fields as $field) {
            $columnHeaders[] = $thaiFields[$field];
        }

        $alphabet = range('A', 'Z');

        $row = 3;
        foreach ($columnHeaders as $key => $header) {
            $activeWorksheet->setCellValue($alphabet[$key] . $row, $header);
        }

        // set column width automatically
        foreach ($alphabet as $column) {
            $activeWorksheet->getColumnDimension($column)->setAutoSize(true);
        }

        // set data to spreadsheet
        $startRow = 4;
        foreach ($extractedData as $key => $userData) {
            $key++;
            foreach ($userData as $field => $value) {
                $activeWorksheet->setCellValue($alphabet[array_search($field, $fields)] . $startRow, $value);
            }
            $startRow++;
        }

        $activeWorksheet->setCellValue('A1', 'รายชื่อผู้สมัครค่ายคอมแคมป์ ครั้งที่ 22');
        $activeWorksheet->getStyle('A1')->getFont()->setSize(20);

        $writer = new Xlsx($spreadsheet);
        $writer->save('Comcamp22_Exported_Filters.xlsx');
        // move file to storage folder
        rename('Comcamp22_Exported_Filters.xlsx', './storage/Comcamp22_Exported_Filters.xlsx');
    }

    public function exportToJson($regroupedData)
    {
        $jsonData = json_encode($regroupedData, JSON_PRETTY_PRINT);

        return $jsonData;
    }
}