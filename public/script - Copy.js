          let previousValue = "";
          const textarea = document.getElementById("sqlArea");
          textarea.addEventListener("input", () => {
            if (previousValue === "") {
              previousValue = textarea.value;
            }
          });
          function clearTextarea() {
            previousValue = textarea.value;
            textarea.value = "";
          }
          function undoTextarea() {
            if (previousValue !== "") {
              textarea.value = previousValue;
              previousValue = "";
            } else {
              alert("Không có nội dung để hoàn tác.");
            }
          }
          let controller;
          function abortQuery() {
            if (controller) controller.abort();
            document.getElementById("loading").style.display = "none";
          }
          function showUpdateForm() {
            document.getElementById("updateForm").style.display = "block";
			document.getElementById("checkCongForm").style.display = "none";
			document.getElementById("themGioCongForm").style.display = "none";
			document.getElementById("xoaGioCongForm").style.display = "none";
			document.getElementById("updateMaVachForm").style.display = "none";
          }
          
		  document.querySelector('form').addEventListener('submit', function (e) {
  e.preventDefault();
  controller = new AbortController();
  const sql = e.target.sql.value;
  document.getElementById("loading").style.display = "block";
  fetch('/query', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ sql }),
    signal: controller.signal
  })
  .then(res => res.text())
  .then(html => {
    document.getElementById("loading").style.display = "none";
    document.getElementById("result").innerHTML = html;
  })
  .catch(err => {
    document.getElementById("loading").style.display = "none";
    document.getElementById("result").innerHTML = "<p style='color:red;'>❌ Đã dừng hoặc có lỗi kết nối.</p>";
    console.error(err);
  });
});

          document.getElementById("updateGioRaForm").addEventListener("submit", function(e) {
            e.preventDefault();
            const date = document.getElementById("updateDate").value;
            if (!date) return alert("Vui lòng chọn ngày!");
            fetch('/update-gio-ra', {
              method: 'POST',
              headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
              body: new URLSearchParams({ date })
            })
            .then(res => res.text())
            .then(html => {
              document.getElementById("result").innerHTML = html;
            })
            .catch(err => {
              alert("Lỗi khi thực hiện Update Giờ Ra.");
              console.error(err);
            });
          });
          function showNotes() {
            document.getElementById("noteModal").style.display = "block";
          }
          function closeNotes() {
            document.getElementById("noteModal").style.display = "none";
          }
          function replaceDate() {
            const newDate = document.getElementById("newDate").value;
            if (!newDate.match(/\\d{4}-\\d{2}-\\d{2}/)) {
              alert("Vui lòng chọn ngày hợp lệ (yyyy-mm-dd)");
              return;
            }
            if (!textarea.value) {
              alert("Textarea không có nội dung để thay ngày.");
              return;
            }
            textarea.value = textarea.value.replace(/\\d{4}-\\d{2}-\\d{2}/g, newDate);
          }
		  
		  function showCheckCong() {
  document.getElementById("checkCongForm").style.display = "block";
  document.getElementById("updateForm").style.display = "none";
  document.getElementById("themGioCongForm").style.display = "none";
  document.getElementById("xoaGioCongForm").style.display = "none";
  document.getElementById("updateMaVachForm").style.display = "none";
}

document.getElementById("formCheckCong").addEventListener("submit", function(e) {
  e.preventDefault();
  const date = document.getElementById("checkDate").value;
  const ma = document.getElementById("maChamCong").value.trim();

  if (!date || !ma) return alert("Vui lòng nhập đầy đủ thông tin.");

  const sql = "select * from CheckInOut where machamcong in (" + ma + ") and NgayCham = '" + date + "'";


  document.getElementById("sqlArea").value = sql;

  // Gửi form tự động
  document.querySelector("form").dispatchEvent(new Event("submit"));
});

function checkGiaiLao() {
	document.getElementById("checkCongForm").style.display = "none";
  document.getElementById("updateForm").style.display = "none";
  document.getElementById("themGioCongForm").style.display = "none";
  document.getElementById("xoaGioCongForm").style.display = "none";
  document.getElementById("updateMaVachForm").style.display = "none";
  const sql1 = "select * from CheckInOut where MaSoMay = 13 and DATEPART(year,ngaycham)=2025 and DATEPART(hour,giocham)=9 order by GioCham desc";
  const sql2 = "select * from CheckInOut where MaSoMay = 62 and DATEPART(year,ngaycham)=2025 and DATEPART(hour,giocham)=9 order by GioCham desc";
  const fullSql = sql1 + ";\n" + sql2;
  document.getElementById("sqlArea").value = fullSql;
  document.querySelector("form").dispatchEvent(new Event("submit"));
}

function showThemGioCong() {
  document.getElementById("themGioCongForm").style.display = "block";
document.getElementById("updateForm").style.display = "none";
document.getElementById("checkCongForm").style.display = "none";
document.getElementById("xoaGioCongForm").style.display = "none";
document.getElementById("updateMaVachForm").style.display = "none";

}

document.getElementById("formThemGioCong").addEventListener("submit", function (e) {
  e.preventDefault();

  const ma = document.getElementById("themMaChamCong").value.trim();
  const ngay = document.getElementById("themNgayCham").value;
  const gio = document.getElementById("themGioCham").value;

  if (!ma || !ngay || !gio) {
    alert("Vui lòng nhập đầy đủ thông tin.");
    return;
  }

  const gioChuan = gio.length === 5 ? gio + ':00' : gio;
  const sql = `
INSERT INTO CheckInOut (machamcong, ngaycham, giocham, kieucham , nguoncham, masomay, tenmay)
VALUES ('${ma}', '${ngay} 00:00:00.000', '${ngay} ${gioChuan}.000', 0, 1, 13, N'Máy 13')
  `.trim();

  document.getElementById("loading").style.display = "block";

  fetch('/query', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ sql })
  })
  .then(res => res.text())
  .then(html => {
    document.getElementById("loading").style.display = "none";
    document.getElementById("result").innerHTML = html;
  })
  .catch(err => {
    alert("Lỗi khi thêm giờ công.");
    console.error(err);
  });
});

function showXoaGioCong() {
  document.getElementById("xoaGioCongForm").style.display = "block";
  document.getElementById("updateForm").style.display = "none";
  document.getElementById("checkCongForm").style.display = "none";
  document.getElementById("themGioCongForm").style.display = "none";
  document.getElementById("updateMaVachForm").style.display = "none";
}

document.getElementById("formXoaGioCong").addEventListener("submit", function (e) {
  e.preventDefault();

  const ma = document.getElementById("xoaMaChamCong").value.trim();
  const ngay = document.getElementById("xoaNgayCham").value;
  const id = document.getElementById("xoaID").value.trim();

  if (!ma || !ngay || !id) {
    alert("Vui lòng nhập đầy đủ thông tin.");
    return;
  }

  const sql = `
DELETE FROM CheckInOut 
WHERE machamcong IN (${ma}) AND NgayCham = '${ngay}' AND id = '${id}'
  `.trim();

  document.getElementById("loading").style.display = "block";

  fetch('/query', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ sql })
  })
  .then(res => res.text())
  .then(html => {
    document.getElementById("loading").style.display = "none";
    document.getElementById("result").innerHTML = html;
  })
  .catch(err => {
    alert("Lỗi khi xóa giờ công.");
    console.error(err);
  });
});

function showUpdateMaVach() {
  document.getElementById("updateMaVachForm").style.display = "block";
  document.getElementById("updateForm").style.display = "none";
  document.getElementById("checkCongForm").style.display = "none";
  document.getElementById("themGioCongForm").style.display = "none";
  document.getElementById("xoaGioCongForm").style.display = "none";
}

document.getElementById("formUpdateMaVach").addEventListener("submit", function (e) {
  e.preventDefault();

  const tu = document.getElementById("maTu").value.trim();
  const den = document.getElementById("maDen").value.trim();

  if (!tu || !den) {
    document.getElementById("result").innerHTML = "<p style='color:red;'>❌ Vui lòng nhập đầy đủ mã nhân viên.</p>";
    return;
  }

  const sql = `
UPDATE NHANVIEN 
SET GhiChu = CONCAT(1100700, MaChamCong) 
WHERE MaChamCong BETWEEN ${tu} AND ${den}
  `.trim();

  document.getElementById("loading").style.display = "block";

  fetch('/query', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ sql })
  })
  .then(res => res.text())
  .then(html => {
    document.getElementById("loading").style.display = "none";
    document.getElementById("result").innerHTML = html;
  })
  .catch(err => {
    document.getElementById("loading").style.display = "none";
    document.getElementById("result").innerHTML = "<p style='color:red;'>❌ Lỗi khi gửi truy vấn.</p>";
    console.error(err);
  });
});

