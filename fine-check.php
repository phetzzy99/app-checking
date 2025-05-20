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
// เริ่มต้น session ตั้งแต่เริ่มต้นไฟล์
session_start();

require 'read_excel.php';

// ตรวจสอบการ export Excel
if (isset($_GET['export']) && $_GET['export'] === 'excel') {
    // เรียกใช้ไฟล์ export_direct.php และส่งข้อมูลใน session ไป
    if (isset($_SESSION['export_data']) && !empty($_SESSION['export_data'])) {
        require 'export_direct.php';
        $filter = isset($_GET['filter']) ? $_GET['filter'] : 'all';
        // ส่งข้อมูลจาก session ไปยังฟังก์ชัน export
        header("Location: export_direct.php?filter=$filter");
        exit;
    } else {
        echo "<script>
            alert('ไม่พบข้อมูลสำหรับการ Export โปรดอัปโหลดไฟล์ Excel ก่อน');
            window.location.href = 'fine-check.php';
        </script>";
        exit;
    }
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_FILES['excel_file'])) {
        $file_path = $_FILES['excel_file']['tmp_name'];
        $data = read_excel($file_path);
        $results = [];
        $total_results = 0;
        $mismatch_count = 0;
        $no_mismatch_count = 0;

        foreach ($data as $row) {
            // ข้ามแถวแรกถ้าเป็นหัวตาราง (ถ้าจำเป็น)
            if ($row[0] === 'รหัสสมาชิก' || $row[0] === 'ID' || empty($row[0])) {
                continue;
            }

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

            // ตรวจสอบผลลัพธ์จาก API
            $mismatch = false;
            
            if ($result) {
                if ($result[0]['BARCODE'] !== $id_card) {
                    $mismatch = true;
                    $mismatch_count++;
                } else {
                    $no_mismatch_count++;
                }
                
                $results[] = [
                    'id_card' => $id_card,
                    'name_card' => $name_card,
                    'items' => $result,
                    'has_debt' => $mismatch
                ];
                
                $total_results++;
            } else {
                // กรณีไม่มีข้อมูลจาก API ให้เพิ่มเป็นรายการที่ไม่มีพันธะ
                $results[] = [
                    'id_card' => $id_card,
                    'name_card' => $name_card,
                    'items' => [],
                    'has_debt' => false
                ];
                
                $no_mismatch_count++;
                $total_results++;
            }
        }

        // เก็บข้อมูลไว้ใน session เพื่อใช้ในการ export
        $_SESSION['export_data'] = $results;

        echo '<div class="bg-white shadow-md rounded px-8 pt-6 pb-8 mb-4">';
        echo '<h2 class="text-2xl font-bold mb-4">Result:</h2>';

        echo '<p class="mb-4">จำนวนผลลัพธ์ทั้งหมด: ' . $total_results . '</p>';
        echo '<p class="mb-4">จำนวนผลลัพธ์ที่มีพันธะกับหอสมุด: ' . $mismatch_count . '</p>';
        echo '<p class="mb-4">จำนวนผลลัพธ์ที่ไม่มีพันธะกับหอสมุด: ' . $no_mismatch_count . '</p>';

        // เพิ่มปุ่ม Export Excel
        echo '<div class="mb-4 flex space-x-4">';
        echo '<a href="export_direct.php" class="bg-green-500 hover:bg-green-700 text-white font-bold py-2 px-4 rounded focus:outline-none focus:shadow-outline">Export Excel ทั้งหมด</a>';
        echo '<a href="export_direct.php?filter=mismatch" class="bg-red-500 hover:bg-red-700 text-white font-bold py-2 px-4 rounded focus:outline-none focus:shadow-outline">Export Excel เฉพาะรายการที่มีพันธะ</a>';
        echo '<a href="export_direct.php?filter=no_mismatch" class="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded focus:outline-none focus:shadow-outline">Export Excel เฉพาะรายการที่ไม่มีพันธะ</a>';
        echo '</div>';

        // เพิ่มแท็บ
        echo '<div class="mb-4">';
        echo '<ul class="flex border-b">';
        echo '<li class="-mb-px mr-1">';
        echo '<a class="bg-white inline-block border-l border-t border-r rounded-t py-2 px-4 text-blue-700 font-semibold cursor-pointer tab-link active" data-tab="all">ทั้งหมด</a>';
        echo '</li>';
        echo '<li class="mr-1">';
        echo '<a class="bg-white inline-block border-l border-t border-r rounded-t py-2 px-4 text-blue-500 hover:text-blue-800 font-semibold cursor-pointer tab-link" data-tab="has-debt">มีพันธะกับหอสมุด</a>';
        echo '</li>';
        echo '<li class="mr-1">';
        echo '<a class="bg-white inline-block border-l border-t border-r rounded-t py-2 px-4 text-blue-500 hover:text-blue-800 font-semibold cursor-pointer tab-link" data-tab="no-debt">ไม่มีพันธะกับหอสมุด</a>';
        echo '</li>';
        echo '</ul>';
        echo '</div>';

        // ตารางแสดงทั้งหมด
        echo '<div id="tab-all" class="tab-content active">';
        echo '<table id="result-table-all" class="stripe hover w-full">';
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
            $has_debt = $item['has_debt'];

            echo '<tr ' . ($has_debt ? 'class="bg-red-100"' : 'class="bg-green-100"') . '>';
            echo '<td>' . $index . '</td>';
            echo '<td>' . $id_card . '</td>';
            echo '<td>' . $name_card . '</td>';
            echo '<td>';
            if (!empty($result)) {
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
            } else {
                echo 'ไม่มีรายการหนังสือที่เกินกำหนด';
            }
            echo '</td>';
            echo '<td>' . ($has_debt ? 'มีพันธะกับหอสมุด' : 'ไม่มีพันธะ') . '</td>';
            echo '</tr>';

            $index++;
        }

        echo '</tbody>';
        echo '</table>';
        echo '</div>';

        // ตารางแสดงเฉพาะมีพันธะ
        echo '<div id="tab-has-debt" class="tab-content hidden">';
        echo '<table id="result-table-has-debt" class="stripe hover w-full">';
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
            if ($item['has_debt']) {
                $id_card = $item['id_card'];
                $name_card = $item['name_card'];
                $result = $item['items'];

                echo '<tr class="bg-red-100">';
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
                echo '<td>มีพันธะกับหอสมุด</td>';
                echo '</tr>';

                $index++;
            }
        }

        echo '</tbody>';
        echo '</table>';
        echo '</div>';

        // ตารางแสดงเฉพาะไม่มีพันธะ
        echo '<div id="tab-no-debt" class="tab-content hidden">';
        echo '<table id="result-table-no-debt" class="stripe hover w-full">';
        echo '<thead>';
        echo '<tr>';
        echo '<th>ลำดับ</th>';
        echo '<th>รหัสสมาชิก</th>';
        echo '<th>ชื่อ-สกุล</th>';
        echo '<th>สถานะ</th>';
        echo '</tr>';
        echo '</thead>';
        echo '<tbody>';

        $index = 1;

        foreach ($results as $item) {
            if (!$item['has_debt']) {
                $id_card = $item['id_card'];
                $name_card = $item['name_card'];

                echo '<tr class="bg-green-100">';
                echo '<td>' . $index . '</td>';
                echo '<td>' . $id_card . '</td>';
                echo '<td>' . $name_card . '</td>';
                echo '<td>ไม่มีพันธะ</td>';
                echo '</tr>';

                $index++;
            }
        }

        echo '</tbody>';
        echo '</table>';
        echo '</div>';

        echo '</div>';

        echo "
        <script>
            $(document).ready(function() {
                $('#result-table-all').DataTable({
                    lengthMenu: [10, 25, 50, 100],
                    language: {
                        url: 'https://cdn.datatables.net/plug-ins/1.13.4/i18n/th.json'
                    }
                });
                
                $('#result-table-has-debt').DataTable({
                    lengthMenu: [10, 25, 50, 100],
                    language: {
                        url: 'https://cdn.datatables.net/plug-ins/1.13.4/i18n/th.json'
                    }
                });
                
                $('#result-table-no-debt').DataTable({
                    lengthMenu: [10, 25, 50, 100],
                    language: {
                        url: 'https://cdn.datatables.net/plug-ins/1.13.4/i18n/th.json'
                    }
                });
                
                // แท็บคลิก
                $('.tab-link').click(function() {
                    var tabId = $(this).data('tab');
                    
                    // ซ่อนทุกแท็บคอนเท้นท์
                    $('.tab-content').removeClass('active').addClass('hidden');
                    
                    // แสดงแท็บที่คลิก
                    $('#tab-' + tabId).removeClass('hidden').addClass('active');
                    
                    // ปรับสไตล์ของแท็บลิงก์
                    $('.tab-link').removeClass('active text-blue-700').addClass('text-blue-500 hover:text-blue-800');
                    $(this).removeClass('text-blue-500 hover:text-blue-800').addClass('active text-blue-700');
                    
                    // รีเซ็ตตาราง DataTable เพื่อให้แสดงผลถูกต้อง
                    $('#result-table-all').DataTable().columns.adjust();
                    $('#result-table-has-debt').DataTable().columns.adjust();
                    $('#result-table-no-debt').DataTable().columns.adjust();
                });
                
                document.getElementById('loading').classList.add('hidden');

                Swal.fire({
                    icon: 'success',
                    title: 'สำเร็จ',
                    text: 'ประมวลผลข้อมูลเรียบร้อยแล้ว!',
                });
            });
        </script>";
    }
}
?>