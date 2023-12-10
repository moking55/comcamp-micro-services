<?php

require_once __DIR__ . '/router.php';

// ##################################################
// ##################################################
// ##################################################

// Static GET
// In the URL -> http://localhost
// The output -> Index
get('/', 'src/views/index.php');

get('/comcamp/export', function () {
    require_once __DIR__ . '/src/exporter.php';
    try {
        $exporter = new FirebaseDataExporter();
        $regroupedData = $exporter->fetchDataFromFirebase();
        $serverProtocol = isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] != 'off' ? 'https://' : 'http://';
        $exportPath = $serverProtocol . $_SERVER['HTTP_HOST'] . '/storage/';

        $paidOnly = isset($_GET['paidOnly']) && $_GET['paidOnly'] == 'true';
        if (isset($_GET['fields'])) {
            // convert to array
            $fields = explode(',', $_GET['fields']);
            $exportPath = $exportPath . 'Comcamp22_Exported_Filters.xlsx';
            $exporter->exportSelectedFieldsToSpreadsheet($regroupedData, $fields, $paidOnly);
        } else {
            $exporter->exportToSpreadsheet($regroupedData, $paidOnly);
            $exportPath = $exportPath . 'Comcamp22_Exported_All.xlsx';
        }

        header('Content-Type: application/json');
        echo json_encode([
            'status' => 200,
            'message' => 'OK',
            'url' => $exportPath
        ]);
    } catch (Exception $e) {
        header('Content-Type: application/json');
        echo json_encode([
            'status' => 500,
            'message' => 'Internal Server Error'
        ]);
    }
});

any('/404', 'src/views/error.php');
