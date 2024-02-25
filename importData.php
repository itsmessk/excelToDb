<?php
use PhpOffice\PhpSpreadsheet\Shared\Date;

// Load the database configuration file 
include_once 'dbConfig.php';

// Include PhpSpreadsheet library autoloader 
require_once 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

if (isset($_POST['importSubmit'])) {

    // Allowed mime types 
    $excelMimes = array('text/xls', 'text/xlsx', 'application/excel', 'application/vnd.msexcel', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Validate whether selected file is a Excel file 
    if (!empty($_FILES['file']['name']) && in_array($_FILES['file']['type'], $excelMimes)) {

        // If the file is uploaded 
        if (is_uploaded_file($_FILES['file']['tmp_name'])) {
            $reader = new Xlsx();
            $spreadsheet = $reader->load($_FILES['file']['tmp_name']);
            $worksheet = $spreadsheet->getActiveSheet();
            $worksheet_arr = $worksheet->toArray();

            // Remove header row 
            unset($worksheet_arr[0]);

            foreach ($worksheet_arr as $row) {
                $dateString = $row[0];

                // Adjust the createFromFormat to account for two-digit year input.
                // PHP understands "y" as a two-digit year. Use 'd-M-y' for parsing the input date.
                $dateObject = DateTime::createFromFormat('d-M-y', $dateString);

                // Check if date conversion was successful to avoid errors
                if ($dateObject !== false) {
                    // Explicitly handle the year to ensure it's correctly interpreted as 20xx or 19xx
                    // Format the date as 'Y-m-d' for SQL insertion, ensuring a four-digit year
                    $date = $dateObject->format('Y-m-d');
                } else {
                    // Handle error, for example, use a default date or log an error
                    echo "Error parsing date: $dateString";
                    // Set $date to null, a default value, or handle as appropriate
                    $date = null;
                }
                $department = $row[1];
                $section = $row[2];
                $timing = $row[3];
                $duration = $row[4];
                $topic_original = $row[5]; // Assuming this is the column for topics
                if (strtolower($topic_original) == "aptitude") {
                    $topic = 1; // For aptitude
                } elseif (strtolower($topic_original) == "communication") {
                    $topic = 0; // For communication
                } else {
                    $topic = NULL; // For any other topic or value
                }
                $trainer = $row[6];
                $class_delivered = $row[7];
                $attended = $row[8];
                $uniquee = $row[9];

                // Check whether member already exists in the database with the same email 
                $prevQuery = "SELECT id FROM course_details WHERE uniquee = '" . $uniquee . "'";
                $prevResult = $db->query($prevQuery);

                if ($prevResult->num_rows > 0) {
                    // Update course details data in the database 
                    $db->query("UPDATE course_details SET date = '" . $date . "', department = '" . $department . "', section = '" . $section . "', timing = '" . $timing . "', duration = '" . $duration . "', topic = '" . $topic . "', trainer = '" . $trainer . "', class_delivered = '" . $class_delivered . "', attended = '" . $attended . "' WHERE uniquee = '" . $uniquee . "'");
                } else {
                    // Insert course details data into the database 
                    $db->query("INSERT INTO course_details (date, department, section, timing, duration, topic, trainer, class_delivered, attended, uniquee) VALUES ('" . $date . "', '" . $department . "', '" . $section . "', '" . $timing . "', '" . $duration . "', '" . $topic . "', '" . $trainer . "', '" . $class_delivered . "', '" . $attended . "', '" . $uniquee . "')");
                }

            }

            $qstring = '?status=succ';
        } else {
            $qstring = '?status=err';
        }
    } else {
        $qstring = '?status=invalid_file';
    }
}

// Redirect to the listing page 
header("Location: index.php" . $qstring);

?>