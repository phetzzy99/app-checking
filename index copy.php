<!DOCTYPE html>
<html>

<head>
    <title>ID Card Reader</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sweetalert2@11.0.18/dist/sweetalert2.min.css">
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11.0.18/dist/sweetalert2.all.min.js"></script>
    <script src="https://cdn.tailwindcss.com"></script>
</head>

<body class="bg-gray-100">
    <div class="container mx-auto mt-10">
        <div class="bg-white shadow-md rounded px-8 pt-6 pb-8 mb-4">
            <h1 class="text-3xl font-bold mb-4">Upload Excel File</h1>
            <form method="POST" enctype="multipart/form-data" class="mb-4">
                <div class="mb-4">
                    <label class="block text-gray-700 text-sm font-bold mb-2" for="excel_file">Select Excel File:</label>
                    <input class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline" type="file" name="excel_file" accept=".xlsx, .xls" required>
                </div>
                <div>
                    <button class="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded focus:outline-none focus:shadow-outline" type="submit">Upload</button>
                </div>
            </form>
        </div>
        <div id="loading" class="fixed z-50 top-0 left-0 w-screen h-screen flex items-center justify-center bg-gray-500 bg-opacity-50 hidden">
            <div class="spinner-border text-primary" style="width: 3rem; height: 3rem;" role="status">
                <span class="visually-hidden">Loading...</span>
            </div>
        </div>
    </div>

    <script>
        document.querySelector('form').addEventListener('submit', function() {
            document.getElementById('loading').classList.remove('hidden');
        });
    </script>
</body>

</html>


<?php
require 'read_excel.php';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_FILES['excel_file'])) {
        $file_path = $_FILES['excel_file']['tmp_name'];
        $data = read_excel($file_path);

        echo '<div class="bg-white shadow-md rounded px-8 pt-6 pb-8 mb-4">';
        echo '<h2 class="text-2xl font-bold mb-4">Result:</h2>';

        foreach ($data as $row) {
            $id_card = $row[0];
            $name_card = $row[1];
            $api_url = "https://imagesopac.sru.ac.th/v1/api/GetPatronRFine/$id_card";
            $api_token = 'UBWmZBBYYvM6l/vcSSyGYCmS2sirXnRBH0F2vvUFK2ST5NGDu00/v+dsErbOD4m5pgRjkLAd5ZvnWOQImEPlKQ==';

            $headers = [
                'token:' . $api_token
            ];

            $ch = curl_init($api_url);
            curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
            curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
            $response = curl_exec($ch);
            curl_close($ch);

            $result = json_decode($response, true);

            $mismatch = false;

            if ($result) {
                $api_name = $result[0]['BARCODE'];
                if ($id_card !== $api_name) {
                    $mismatch = true;
                }
            }

            echo '<div class="bg-gray-100 rounded-lg p-4 mb-4 ' . ($mismatch ? 'border-l-4 border-red-500' : '') . '">';
            echo '<p class="text-lg font-bold mb-2">รหัสสมาชิก: ' . $id_card . '</p>';
            echo '<p class="text-lg font-bold mb-2">ชื่อ-สกุล: ' . $name_card . '</p>';

            if ($mismatch) {
                echo '<div class="bg-red-100 border-t border-b border-red-500 text-red-700 px-4 py-3 mb-4" role="alert">';
                echo '<p class="font-bold">แจ้งเตือน</p>';
                echo '<p>สมาชิกนี้มีพันธะกับหอสมุด</p>';
                echo '</div>';
            }

            if ($result) {
                foreach ($result as $item) {
                    echo '<div class="mb-2">';
                    echo '<p class="text-gray-700"><strong>Barcode:</strong> ' . $item['BARCODE'] . '</p>';
                    echo '<p class="text-gray-700"><strong>ชื่อเรื่อง:</strong> ' . $item['TITLE'] . '</p>';
                    echo '<p class="text-gray-700"><strong>Call No.:</strong> ' . $item['CALLNO'] . '</p>';
                    echo '<p class="text-gray-700"><strong>วันกำหนดส่ง:</strong> ' . $item['DUEDATE'] . '</p>';
                    echo '<p class="text-gray-700"><strong>วันที่คืน:</strong> ' . $item['RETURNDATE'] . '</p>';
                    echo '<p class="text-gray-700"><strong>ค่าปรับ:</strong> ' . $item['TOTALMONEY'] . '</p>';
                    echo '</div>';
                }
            } else {
                echo '<p class="text-red-500">ไม่พบข้อมูลพันธะกับหอสมุด</p>';
            }

            echo '</div>';
        }

        echo '</div>';

        echo "
        <script>
            Swal.fire({
                icon: 'success',
                title: 'Success',
                text: 'Data has been processed successfully!',
            }).then(() => {
                document.getElementById('loading').classList.add('hidden');
            });
        </script>";
    }
}
?>