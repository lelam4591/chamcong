const express = require('express');
const sql = require('mssql');
const bodyParser = require('body-parser');
const session = require('express-session');
const ExcelJS = require('exceljs');

const app = express();
app.use(express.static('public'));
const port = 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(session({
  secret: 'luakchuong-secret',
  resave: false,
  saveUninitialized: true
}));
// Giao diện chọn kết nối CSDL
app.get('/', (req, res) => {
  res.send(`
    <html>
      <head>
        <title>Kết nối CSDL</title>
        <style>
          body {
            background: #f0f2f5;
            font-family: Arial, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
          }

          .connect-box {
            background: white;
            padding: 30px 40px;
            border-radius: 10px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            width: 350px;
            text-align: center;
          }

          .connect-box h3 {
            margin-bottom: 20px;
            color: #333;
          }

          .connect-box label {
            display: block;
            margin: 10px 0;
            text-align: left;
            font-size: 14px;
          }

          .connect-box input[type="password"] {
            width: 100%;
            padding: 10px;
            margin-top: 10px;
            border-radius: 5px;
            border: 1px solid #ccc;
            font-size: 14px;
          }

          .connect-box button {
            width: 100%;
            padding: 10px;
            margin-top: 20px;
            background: #007BFF;
            color: white;
            border: none;
            border-radius: 5px;
            font-weight: bold;
            font-size: 15px;
            cursor: pointer;
            transition: background 0.3s;
          }

          .connect-box button:hover {
            background: #0056b3;
          }

          .option-group {
            text-align: left;
            margin-bottom: 15px;
            font-size: 14px;
          }

        </style>
      </head>
      <body>
        <div class="connect-box">
          <h3>🔌 Kết nối CSDL</h3>
          <form method="POST" action="/connect">
            <div class="option-group">
              <label><input type="radio" name="option" value="1" checked> Server Chính: pmcc - MITACOSQL</label>
              <label><input type="radio" name="option" value="2"> Server 1.6: sa - MITACO</label>
            </div>
            <input type="password" name="password" placeholder="Nhập mật khẩu..." required>
            <button type="submit">🚀 Kết nối</button>
          </form>
        </div>
      </body>
    </html>
  `);
});


// Xử lý kết nối
app.post('/connect', async (req, res) => {
  const { option, password } = req.body;

  let config;
  if (option === '1') {
    config = {
      user: 'pmcc',
      password,
      server: '103.42.57.125',
      database: 'MITACOSQL',
      options: { encrypt: false, trustServerCertificate: true }
    };
  } else {
    config = {
      user: 'sa',
      password,
      server: '192.168.1.6\\SQL2014',
      database: 'MITACO',
      options: { encrypt: false, trustServerCertificate: true }
    };
  }

  try {
    const pool = await sql.connect(config);
    app.locals.pool = pool;
    res.redirect('/home');
  } catch (err) {
    res.send(`<p style="color:red;">❌ Lỗi kết nối: ${err.message}</p><a href="/">Thử lại</a>`);
  }
});

// Trang chính sau khi kết nối
app.get('/home', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="vi">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Kết nối thành công</title>
        <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Roboto', sans-serif;
            background: #f7f9fc;
            margin: 0;
            padding: 20px;
          }
          h3 {
            color: #333;
            text-align: center;
          }
          .header-actions {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 10px;
            margin-bottom: 20px;
          }
          button {
            background: #14B337;
            border: none;
            color: white;
            padding: 8px 14px;
            border-radius: 6px;
            cursor: pointer;
            transition: background 0.3s;
            font-size: 13px;
          }
          button:hover {
            background: #0D7123;
          }
          form {
            max-width: 800px;
            margin: 0 auto;
          }
          textarea {
  width: 100%;
  max-width: 900px;
  min-width: 300px;
  margin: 0 auto; /* Canh giữa */
  display: block;
  padding: 10px;
  border: 1px solid #ccc;
  border-radius: 6px;
  font-size: 14px;
}

          table {
            border-collapse: collapse;
            margin-top: 20px;
            width: 100%;
            text-align: left;
          }
          table, th, td {
            border: 1px solid #ddd;
          }
          th, td {
            padding: 8px 12px;
          }
          th {
            background: #007BFF;
            color: #fff;
          }
          #noteModal {
            position: fixed;
            top: 10%;
            left: 50%;
            transform: translateX(-50%);
            background: #fff;
            border: 2px solid #ccc;
            padding: 20px;
            z-index: 1000;
            width: 80%;
            max-width: 900px;
            box-shadow: 0 8px 20px rgba(0,0,0,0.2);
          }
          #loading {
            text-align: center;
            color: #D32F2F;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <h3>Kết nối CSDL thành công!</h3>
        <div class="header-actions">
          <button onclick="window.location.href='/disconnect'">🔁 Đổi kết nối CSDL</button>
          <label for="newDate" style="display:none;">Chọn ngày</label>
<input type="date" id="newDate" title="Chọn ngày" placeholder="yyyy-mm-dd">

          <button onclick="replaceDate()">Đổi ngày</button>
        </div>
        <form method="POST" action="/query">
		<div style="display: flex; flex-direction: column; align-items: center; gap: 10px;">
          <textarea name="sql" id="sqlArea" rows="5" placeholder="Nhập câu lệnh SQL tại đây..."></textarea>
          <div style="margin-top: 10px;">
            <button type="submit">Thực hiện</button>
            <button type="button" onclick="abortQuery()">Dừng</button>
            <button type="button" onclick="clearTextarea()">🗑️ Xóa</button>
            <button type="button" onclick="undoTextarea()">↩️ Hoàn tác</button>
          </div>
		</div>
        </form>
        <br>
        <button onclick="showNotes()">📋 Note</button>
        <button onclick="showUpdateForm()">1.Update Giờ Ra</button>
        <button onclick="checkGiaiLao()">2.Check giờ giải lao</button>
		<button onclick="showCheckCong()">3.Check Giờ Công</button>
		<button onclick="showThemGioCong()">➕ Thêm giờ công</button>
		<button onclick="showXoaGioCong()">🗑️ Xóa giờ công</button>
		<button onclick="showUpdateMaVach()">Update Mã Vạch</button>

<div id="updateMaVachForm" style="display:none; margin-top:10px;">
  <form id="formUpdateMaVach">
    <label>Mã NV bắt đầu:</label>
    <input type="number" id="maTu" placeholder="Ví dụ: 12939" required>
    <br><br>
    <label>Mã NV kết thúc:</label>
    <input type="number" id="maDen" placeholder="Ví dụ: 12953" required>
    <br><br>
    <button type="submit">Thực hiện Update</button>
  </form>
</div>


<div id="themGioCongForm" style="display:none; margin-top:10px;">
  <form id="formThemGioCong">
    <label>Mã chấm công:</label>
    <input type="text" id="themMaChamCong" placeholder="Ví dụ: 9141" required>
    <br><br>
    <label>Ngày chấm:</label>
    <input type="date" id="themNgayCham" required>
    <br><br>
    <label>Giờ chấm (HH:mm:ss):</label>
    <input type="time" id="themGioCham" step="1" required>
    <br><br>
    <button type="submit">Thực hiện</button>
  </form>
</div>


<div id="xoaGioCongForm" style="display:none; margin-top:10px;">
  <form id="formXoaGioCong">
    <label>Mã chấm công:</label>
    <input type="text" id="xoaMaChamCong" placeholder="Ví dụ: 9141" required>
    <br><br>
    <label>Ngày chấm:</label>
    <input type="date" id="xoaNgayCham" required>
    <br><br>
    <label>ID:</label>
    <input type="text" id="xoaID" placeholder="Ví dụ: 16099063" required>
    <br><br>
    <button type="submit">Thực hiện Xóa</button>
  </form>
</div>


<div id="checkCongForm" style="display:none; margin-top:10px;">
  <form id="formCheckCong">
    <label>Ngày chấm:</label>
    <label for="checkDate" style="display:none;">Ngày chấm</label>
<input type="date" id="checkDate" required title="Ngày chấm" placeholder="yyyy-mm-dd">

    <br>
    <label>Mã chấm công (cách nhau bởi dấu phẩy):</label>
    <input type="text" id="maChamCong" placeholder="Ví dụ: 9141,12831" required style="width: 100%; padding: 8px;">
    <br><br>
    <button type="submit">Thực hiện</button>
  </form>
</div>
<div id="updateForm" style="display:none; margin-top:10px;">
          <form id="updateGioRaForm">
            <label>Chọn ngày:</label>
            <input type="date" id="updateDate" required>
            <button type="submit">Thực hiện Update</button>
          </form>
        </div>
	
<br>
        <!-- Modal Note -->
        <div id="noteModal" style="display:none;">
          <h3>📌 Các câu SQL hay dùng</h3>
          <textarea style="width:100%; height:300px; padding:3px;" readonly>
KT xóa giờ công giải lao: 
select * from CheckInOut where MaSoMay = 13 and DATEPART(year,ngaycham)=2025 and DATEPART(hour,giocham)=9 order by GioCham desc

KT giờ công:
select * from CheckInOut where machamcong in (9141,12831) and NgayCham = '2025-05-20'

----Update giờ---
update CheckInOut 
set giocham= '2021-12-28 07:16:46.000'
where MaChamCong = 7720 and ngaycham= '2021-12-28'
	  
----Update Mã vạch bắn xe---
update NHANVIEN set GhiChu = CONCAT(1100700,MaChamCong) where MaChamCong between 12939 and 12953

----Làm trống DS NV mới---
update NHANVIEN set nhanvienmoi=0 where nhanvienmoi=1

----Đưa Vào DS NV mới---
update NHANVIEN set NhanVienMoi = 1 where MaNhanVien in (
'S06513',
'S09141')

---chèn giờ công---
INSERT INTO CheckInOut (machamcong, ngaycham, giocham, kieucham , nguoncham, masomay, tenmay)
values ('12831', '2025-05-21 00:00:00.000', '2025-05-21 12:14:01.000',0, 1, 13, N'Máy 13')

INSERT INTO CheckInOut (machamcong, ngaycham, giocham, kieucham , nguoncham, masomay, tenmay)
values ('1120', '2025-05-21 00:00:00.000', '2025-05-21 07:25:02.000',0, 1, 30, N'Máy 30')

INSERT INTO CheckInOut (machamcong, ngaycham, giocham, kieucham , nguoncham, masomay, tenmay)
values ('9141', '2025-05-21 00:00:00.000', '2025-05-21 16:05:37.000',0, 1, 13, N'Máy 13')

----Lọc thẻ từ---
SELECT * FROM nhanvien
WHERE mathe IS NOT NULL AND mathe != '';


----xóa giờ công trong khoảng time ---
delete CheckInOut where MaChamCong =12167 and GioCham between '2019-09-18 07:00:59' and '2019-09-18 07:15:59'
          </textarea>
          <br><br>
          <button onclick="closeNotes()">Đóng</button>
        </div>
        <div id="loading" style="display:none;">⏳ Đang thực hiện truy vấn...</div>
        <!-- Kết quả -->
        <div id="result" style="margin-top: 20px;"></div>
        
<script src="/script.js"></script>
		
      </body>
    </html>
  `);
});


// Truy vấn SQL
app.post('/query', async (req, res) => {
  const query = req.body.sql;
  try {
    const result = await app.locals.pool.request().batch(query);

    // ✅ Nếu là các lệnh như UPDATE/DELETE/INSERT có ảnh hưởng dòng
    const affectedRows = Array.isArray(result.rowsAffected)
      ? result.rowsAffected.reduce((a, b) => a + b, 0)
      : 0;

    if (
      (!result.recordsets || !Array.isArray(result.recordsets) || result.recordsets.length === 0) &&
      affectedRows > 0
    ) {
      return res.send(`<p style="color:green;">✅ Lệnh SQL đã thực hiện thành công (${affectedRows} dòng bị ảnh hưởng).</p>`);
    }

    // ⚠️ Không có dữ liệu trả về, không có dòng nào bị ảnh hưởng
    if (!result.recordsets || !Array.isArray(result.recordsets)) {
      return res.send('<p>⚠️ Lệnh đã thực thi, nhưng không có dữ liệu hoặc dòng nào bị ảnh hưởng.</p>');
    }

    // Format ngày giờ
    const formatDateTime = (d) => {
      if (!(d instanceof Date)) return d;
      return d.toISOString().replace('T', ' ').substring(0, 19);
    };

    // Hiển thị nhiều bảng kết quả nếu có
    const tablesHtml = result.recordsets.map((rows, idx) => {
      if (!rows.length) return `<p>🟡 Kết quả ${idx + 1}: không có dữ liệu.</p>`;

      // Format ngày giờ
      rows = rows.map(row => {
        const { HinhAnh, ...rest } = row;
        for (let key in rest) {
          if (rest[key] instanceof Date) {
            rest[key] = formatDateTime(rest[key]);
          }
        }
        return rest;
      });

      const headers = Object.keys(rows[0]);
      const tableRows = rows.map(row =>
        '<tr>' + headers.map(h => `<td>${row[h] ?? ''}</td>`).join('') + '</tr>'
      ).join('');

      return `
        <h4>Kết quả ${idx + 1}:</h4>
        <table border="1" cellpadding="5" cellspacing="0">
          <thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>
          <tbody>${tableRows}</tbody>
        </table>
        <br>
      `;
    }).join('');

    // Gộp tất cả kết quả vào lastResult để xuất Excel nếu cần
    req.session.lastResult = result.recordsets.flat();

    res.send(`
      <a href="/download-excel">📥 Tải kết quả xuống Excel</a><br><br>
      ${tablesHtml}
    `);
  } catch (err) {
    res.send(`<p style="color:red;">❌ Lỗi: ${err.message}</p>`);
  }
});




// Update Giờ Ra
app.post('/update-gio-ra', async (req, res) => {
  const date = req.body.date;
  const dayOfWeek = new Date(date).getDay(); // 6 = Saturday
  const isSaturday = dayOfWeek === 6;

  const pool = app.locals.pool;
  const log = [];

  try {
    let affectedTotal = 0;

    const run = async (sqlText, desc) => {
      const result = await pool.request().query(sqlText);
      const count = result.rowsAffected[0];
      log.push(`• ${desc}: ${count} dòng`);
      affectedTotal += count;
    };

    await run(`DELETE FROM CheckInOut WHERE MaSoMay = 13 AND GioCham BETWEEN '${date} 09:20:59' AND '${date} 09:50:59'`, 'Xoá máy 13 sáng');
    await run(`DELETE FROM CheckInOut WHERE MaSoMay = 62 AND GioCham BETWEEN '${date} 09:20:59' AND '${date} 09:50:59'`, 'Xoá máy 62 sáng');

    if (!isSaturday) {
      await run(`DELETE FROM CheckInOut WHERE MaSoMay = 13 AND GioCham BETWEEN '${date} 14:20:59' AND '${date} 14:50:59'`, 'Xoá máy 13 chiều');
      await run(`DELETE FROM CheckInOut WHERE MaSoMay = 62 AND GioCham BETWEEN '${date} 14:20:59' AND '${date} 14:50:59'`, 'Xoá máy 62 chiều');
    }

    const updateSQL = `
      UPDATE CheckInOut SET GioCham = '${date} 16:33:59' WHERE id IN (
        SELECT id FROM CheckInOut
        WHERE NgayCham = '${date}' AND GioCham > '${date} 16:45:59'
        AND MaChamCong IN (
          '1002','1003','1006','10123','1014','1015','1039','11288','11596','11810',
          '12613','12727','12800','12831','1305','13691','13771','14042','1451','2548',
          '2583','2674','2805','2896','4012','4318','4533','5649','5982','6094','6175',
          '6188','6458','6513','6835','6945','7324','7719','7892','7935','7983','8957',
          '9141','9220','9958','9960'
        )
      )
    `;
    await run(updateSQL, 'Cập nhật giờ ra');

    res.send(`
      <p style="color:green;">✅ Đã cập nhật Giờ Ra cho ngày <strong>${date}</strong> ${isSaturday ? '(Thứ 7)' : ''}.</p>
      <p><strong>Kết quả:</strong><br>${log.join('<br>')}</p>
    `);
  } catch (err) {
    res.send(`<p style="color:red;">Lỗi khi update: ${err.message}</p>`);
  }
});

// Tải Excel
app.get('/download-excel', async (req, res) => {
  const rows = req.session.lastResult;
  if (!rows || !rows.length) {
    return res.send('<p>Không có dữ liệu để tải.</p>');
  }

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Kết quả');

  sheet.columns = Object.keys(rows[0]).map(key => ({ header: key, key }));
  rows.forEach(row => sheet.addRow(row));

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=ketqua.xlsx');

  await workbook.xlsx.write(res);
  res.end();
});

app.listen(port, () => {
  console.log(`🟢 Server đang chạy tại http://localhost:${port}`);
});

app.get('/disconnect', (req, res) => {
  if (app.locals.pool) {
    app.locals.pool.close(); // đóng kết nối nếu cần
    delete app.locals.pool;
  }
  res.redirect('/');
});
