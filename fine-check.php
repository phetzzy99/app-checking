<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>ตรวจสอบค้างค่าปรับหนังสือในระบบ</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sweetalert2@11.0.18/dist/sweetalert2.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css">
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11.0.18/dist/sweetalert2.all.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body {
            background-image: url('https://images.unsplash.com/photo-1521587760476-6c12a4b040da?ixlib=rb-4.0.3&ixid=M3wxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8fA%3D%3D&auto=format&fit=crop&w=1470&q=80');
            background-repeat: no-repeat;
            background-size: cover;
            background-position: center;
            min-height: 100vh;
        }

        .container {
            background-color: rgba(255, 255, 255, 0.9);
            border-radius: 10px;
            padding: 30px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
    </style>
</head>

<body>
    <div class="container mx-auto mt-10">
        <div class="bg-white shadow-md rounded px-8 pt-6 pb-8 mb-4">
            <h1 class="text-3xl font-bold mb-4">Upload Excel ตรวจสอบค้างค่าปรับหนังสือในระบบ</h1>
            <form method="POST" enctype="multipart/form-data" class="mb-4">
                <div class="mb-4">
                    <label class="block text-gray-700 text-sm font-bold mb-2" for="excel_file">Select Excel File:</label>
                    <input class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline" type="file" name="excel_file" accept=".xlsx, .xls" required>
                </div>
                <div class="flex items-center justify-between">
                    <button class="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded focus:outline-none focus:shadow-outline" type="submit">Upload</button>
                    <a href="index.html" class="bg-red-500 hover:bg-red-700 text-white font-bold py-2 px-4 rounded focus:outline-none focus:shadow-outline">
                        หน้าแรก
                    </a>
                </div>
            </form>
        </div>
        <div id="loading" class="fixed z-50 top-0 left-0 w-screen h-screen flex items-center justify-center bg-gray-500 bg-opacity-50 hidden">
            <div class="spinner-border text-primary" style="width: 5rem; height: 5rem;" role="status">
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
        $results = [];
        $total_results = 0;
        $mismatch_count = 0;

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

            if ($result) {
                $results[] = [
                    'id_card' => $id_card,
                    'name_card' => $name_card,
                    'items' => $result
                ];

                $total_results++;

                if ($result[0]['BARCODE'] !== $id_card) {
                    $mismatch_count++;
                }
            }
        }

        echo '<div class="bg-white shadow-md rounded px-8 pt-6 pb-8 mb-4">';
        echo '<h2 class="text-2xl font-bold mb-4">Result:</h2>';

        echo '<p class="mb-4">จำนวนผลลัพธ์ทั้งหมด: ' . $total_results . '</p>';
        echo '<p class="mb-4">จำนวนผลลัพธ์ที่มีพันธะกับหอสมุด: ' . $mismatch_count . '</p>';

        echo '<table id="result-table" class="stripe hover">';
        echo '<thead>';
        echo '<tr>';
        echo '<th>ลำดับ</th>';
        echo '<th>รหัสสมาชิก</th>';
        echo '<th>ชื่อ-สกุล</th>';
        echo '<th>รายการหนังสือที่เกินกำหนด</th>';
        echo '<th>สถานะ</th>';
        echo '</tr>';
        echo '</thead>';
        echo '<tbody>';

        $index = 1;

        foreach ($results as $item) {
            $id_card = $item['id_card'];
            $name_card = $item['name_card'];
            $result = $item['items'];

            $mismatch = false;

            if ($result) {
                $api_name = $result[0]['BARCODE'];
                if ($id_card !== $api_name) {
                    $mismatch = true;
                }
            }

            echo '<tr ' . ($mismatch ? 'class="bg-red-100"' : '') . '>';
            echo '<td>' . $index . '</td>';
            echo '<td>' . $id_card . '</td>';
            echo '<td>' . $name_card . '</td>';
            echo '<td>';
            echo '<ul>';
            foreach ($result as $subitem) {
                echo '<li>';
                echo '<strong>Barcode:</strong> ' . $subitem['BARCODE'] . '<br>';
                echo '<strong>ชื่อเรื่อง:</strong> ' . $subitem['TITLE'] . '<br>';
                echo '<strong>Call No.:</strong> ' . $subitem['CALLNO'] . '<br>';
                echo '<strong>วันกำหนดส่ง:</strong> ' . $subitem['DUEDATE'] . '<br>';
                echo '<strong>วันที่คืน:</strong> ' . $subitem['RETURNDATE'] . '<br>';
                echo '<strong>ค่าปรับ:</strong> ' . $subitem['TOTALMONEY'] . ' บาท';
                echo '</li>';
            }
            echo '</ul>';
            echo '</td>';
            echo '<td>' . ($mismatch ? 'มีพันธะกับหอสมุด' : 'ไม่มีพันธะ') . '</td>';
            echo '</tr>';

            $index++;
        }

        echo '</tbody>';
        echo '</table>';

        echo '</div>';

        echo "
        <script>
            $(document).ready(function() {
                $('#result-table').DataTable({
                    lengthMenu: [10, 25, 50, 100],
                    language: {
                        url: 'https://cdn.datatables.net/plug-ins/1.13.4/i18n/th.json'
                    }
                });
                
                document.getElementById('loading').classList.add('hidden');

                Swal.fire({
                    icon: 'success',
                    title: 'Success',
                    text: 'Data has been processed successfully!',
                });
            });
        </script>";
    }
}
?>