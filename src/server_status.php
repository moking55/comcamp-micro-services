<?php

function checkOnline($domain)
{
    $curlInit = curl_init($domain);
    curl_setopt($curlInit, CURLOPT_CONNECTTIMEOUT, 1);
    curl_setopt($curlInit, CURLOPT_HEADER, true);
    curl_setopt($curlInit, CURLOPT_NOBODY, true);
    curl_setopt($curlInit, CURLOPT_RETURNTRANSFER, true);

    //get answer
    $response = curl_exec($curlInit);

    curl_close($curlInit);
    if ($response) return true;
    return false;
}

$list_sites = [
    "https://comcamp.csmju.com",
    "https://comcamp.codename-t.com",
];

// runnig check
$online = [];
foreach ($list_sites as $site) {
    $online[$site] = checkOnline($site);
}

header('Content-Type: application/json');
echo json_encode($online);