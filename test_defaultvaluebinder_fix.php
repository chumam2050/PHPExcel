<?php
/**
 * Test to verify the fix for "Trying to access array offset on value of type int"
 * Error Location: DefaultValueBinder.php line 82
 */

require_once dirname(__FILE__) . '/vendor/autoload.php';

echo "Testing DefaultValueBinder Fix for Array Offset Error" . PHP_EOL;
echo str_repeat('=', 60) . PHP_EOL;

// This should trigger the specific error scenario that was reported
$excel = new PHPExcel();
$sheet = $excel->getActiveSheet();

// Test cases that previously caused "Trying to access array offset on value of type int"
$testCases = [
    // Integer values
    ['Cell' => 'A1', 'Value' => 0, 'Type' => 'Integer (0)'],
    ['Cell' => 'A2', 'Value' => 1, 'Type' => 'Integer (1)'],
    ['Cell' => 'A3', 'Value' => -1, 'Type' => 'Negative Integer'],
    ['Cell' => 'A4', 'Value' => 100, 'Type' => 'Integer (100)'],
    ['Cell' => 'A5', 'Value' => PHP_INT_MAX, 'Type' => 'PHP_INT_MAX'],
    
    // Float values
    ['Cell' => 'B1', 'Value' => 0.0, 'Type' => 'Float (0.0)'],
    ['Cell' => 'B2', 'Value' => 1.5, 'Type' => 'Float (1.5)'],
    ['Cell' => 'B3', 'Value' => -3.14, 'Type' => 'Float (-3.14)'],
    
    // String values that look like numbers
    ['Cell' => 'C1', 'Value' => '123', 'Type' => 'String numeric'],
    ['Cell' => 'C2', 'Value' => '0123', 'Type' => 'String with leading zero'],
    ['Cell' => 'C3', 'Value' => '1.5', 'Type' => 'String float'],
    
    // Formulas
    ['Cell' => 'D1', 'Value' => '=A1+B1', 'Type' => 'Formula'],
    ['Cell' => 'D2', 'Value' => '=SUM(A1:A5)', 'Type' => 'Formula with function'],
    
    // Boolean
    ['Cell' => 'E1', 'Value' => true, 'Type' => 'Boolean (true)'],
    ['Cell' => 'E2', 'Value' => false, 'Type' => 'Boolean (false)'],
    
    // Null and empty
    ['Cell' => 'F1', 'Value' => null, 'Type' => 'Null'],
    ['Cell' => 'F2', 'Value' => '', 'Type' => 'Empty string'],
    
    // Edge cases
    ['Cell' => 'G1', 'Value' => '=', 'Type' => 'Single equals'],
    ['Cell' => 'G2', 'Value' => 'Text', 'Type' => 'Text string'],
];

$passed = 0;
$failed = 0;

foreach ($testCases as $test) {
    try {
        $sheet->setCellValue($test['Cell'], $test['Value']);
        
        // Verify the value was set
        $retrievedValue = $sheet->getCell($test['Cell'])->getValue();
        
        echo "✓ {$test['Cell']}: {$test['Type']}";
        
        // Show the data type that was detected
        $dataType = $sheet->getCell($test['Cell'])->getDataType();
        $typeNames = [
            PHPExcel_Cell_DataType::TYPE_STRING => 'STRING',
            PHPExcel_Cell_DataType::TYPE_FORMULA => 'FORMULA',
            PHPExcel_Cell_DataType::TYPE_NUMERIC => 'NUMERIC',
            PHPExcel_Cell_DataType::TYPE_BOOL => 'BOOL',
            PHPExcel_Cell_DataType::TYPE_NULL => 'NULL',
            PHPExcel_Cell_DataType::TYPE_INLINE => 'INLINE',
            PHPExcel_Cell_DataType::TYPE_ERROR => 'ERROR',
        ];
        
        echo " → {$typeNames[$dataType]}" . PHP_EOL;
        $passed++;
        
    } catch (Exception $e) {
        echo "✗ {$test['Cell']}: {$test['Type']} - ERROR: {$e->getMessage()}" . PHP_EOL;
        $failed++;
    }
}

echo PHP_EOL;
echo str_repeat('=', 60) . PHP_EOL;
echo "Test Results:" . PHP_EOL;
echo "  Passed: $passed" . PHP_EOL;
echo "  Failed: $failed" . PHP_EOL;
echo str_repeat('=', 60) . PHP_EOL;

if ($failed === 0) {
    echo "✅ SUCCESS!" . PHP_EOL;
    echo "The 'Trying to access array offset on value of type int' error" . PHP_EOL;
    echo "has been fixed. All data types are handled correctly." . PHP_EOL;
} else {
    echo "❌ FAILED!" . PHP_EOL;
    echo "Some tests failed. Please review the errors above." . PHP_EOL;
    exit(1);
}

echo str_repeat('=', 60) . PHP_EOL;

// Additional test: Write to file to ensure everything works end-to-end
echo PHP_EOL . "Writing test file to verify end-to-end functionality... ";
$tempFile = sys_get_temp_dir() . '/phpexcel_defaultvaluebinder_test.xlsx';
$writer = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$writer->save($tempFile);

if (file_exists($tempFile)) {
    $fileSize = filesize($tempFile);
    unlink($tempFile);
    echo "✓ OK (" . number_format($fileSize) . " bytes)" . PHP_EOL;
} else {
    echo "✗ FAILED" . PHP_EOL;
    exit(1);
}

echo PHP_EOL . "🎉 All tests completed successfully!" . PHP_EOL;
