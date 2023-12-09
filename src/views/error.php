<?php
header('HTTP/1.0 404 Not Found');
header('content-type: application/json; charset=utf-8');

echo json_encode([
    'status' => 404,
    'message' => 'Not Found'
]);
