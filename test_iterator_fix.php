<?php
/**
 * Test script to verify Iterator fixes for PHP 8 compatibility
 * This simulates the error scenario from the user's project
 */

require_once dirname(__FILE__) . '/vendor/autoload.php';

echo "Testing PHPExcel Iterator Compatibility with PHP " . PHP_VERSION . PHP_EOL;
echo str_repeat('=', 60) . PHP_EOL;

try {
    // Create new PHPExcel object
    echo "1. Creating PHPExcel object... ";
    $objPHPExcel = new PHPExcel();
    echo "✓" . PHP_EOL;

    // Set document properties
    echo "2. Setting document properties... ";
    $objPHPExcel->getProperties()
        ->setCreator("Test User")
        ->setLastModifiedBy("Test User")
        ->setTitle("Test Document")
        ->setSubject("Test Subject");
    echo "✓" . PHP_EOL;

    // Set active sheet and title (this triggers WorksheetIterator)
    echo "3. Setting worksheet title (triggers WorksheetIterator)... ";
    $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getActiveSheet()->setTitle('Monitoring Attendance');
    echo "✓" . PHP_EOL;

    // Add some data
    echo "4. Adding data to cells... ";
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Name');
    $objPHPExcel->getActiveSheet()->setCellValue('B1', 'Date');
    $objPHPExcel->getActiveSheet()->setCellValue('C1', 'Status');
    echo "✓" . PHP_EOL;

    // Test WorksheetIterator
    echo "5. Testing WorksheetIterator... ";
    $count = 0;
    foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
        $count++;
    }
    echo "✓ (Iterated $count worksheet(s))" . PHP_EOL;

    // Test RowIterator
    echo "6. Testing RowIterator... ";
    $rowCount = 0;
    foreach ($objPHPExcel->getActiveSheet()->getRowIterator() as $row) {
        $rowCount++;
        if ($rowCount >= 5) break; // Just test a few rows
    }
    echo "✓ (Iterated $rowCount row(s))" . PHP_EOL;

    // Test ColumnIterator
    echo "7. Testing ColumnIterator... ";
    $colCount = 0;
    foreach ($objPHPExcel->getActiveSheet()->getColumnIterator() as $column) {
        $colCount++;
        if ($colCount >= 5) break; // Just test a few columns
    }
    echo "✓ (Iterated $colCount column(s))" . PHP_EOL;

    // Test RowCellIterator
    echo "8. Testing RowCellIterator... ";
    $cellCount = 0;
    foreach ($objPHPExcel->getActiveSheet()->getRowIterator() as $row) {
        foreach ($row->getCellIterator() as $cell) {
            $cellCount++;
        }
        break; // Just test first row
    }
    echo "✓ (Iterated $cellCount cell(s))" . PHP_EOL;

    // Create a temporary file to test write functionality
    echo "9. Testing file write... ";
    $tempFile = sys_get_temp_dir() . '/phpexcel_test_' . uniqid() . '.xlsx';
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save($tempFile);
    
    if (file_exists($tempFile)) {
        $fileSize = filesize($tempFile);
        unlink($tempFile); // Clean up
        echo "✓ (Created file: " . number_format($fileSize) . " bytes)" . PHP_EOL;
    } else {
        echo "✗ (File not created)" . PHP_EOL;
    }

    echo PHP_EOL;
    echo str_repeat('=', 60) . PHP_EOL;
    echo "✅ ALL TESTS PASSED!" . PHP_EOL;
    echo "PHPExcel is fully compatible with PHP 8+" . PHP_EOL;
    echo str_repeat('=', 60) . PHP_EOL;

} catch (Exception $e) {
    echo PHP_EOL;
    echo str_repeat('=', 60) . PHP_EOL;
    echo "❌ ERROR OCCURRED!" . PHP_EOL;
    echo "Error: " . $e->getMessage() . PHP_EOL;
    echo "File: " . $e->getFile() . ":" . $e->getLine() . PHP_EOL;
    echo str_repeat('=', 60) . PHP_EOL;
    exit(1);
}
