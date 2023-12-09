<?php
header('content-type: application/json; charset=utf-8');

// show information about this microservice

echo json_encode([
    'name' => 'Microservice for Comcamp 22',
    'author' => 'Codename_T',
    'github' => 'https://github.com/moking55',
    'version' => '1.0.0',
    'status' => 'running'
]);
