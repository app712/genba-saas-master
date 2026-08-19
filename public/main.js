// ==========================================
// SaaSマスターのGAS URLを設定してください
// ==========================================
const GAS_URL = "https://script.google.com/macros/s/AKfycbzZAaLZChAku2SsLvw1ardx7b2rfttkapgm-WgoISwztQDnfvV1f8iVMbnlQ53ozb_d/exec";

// 画面読み込み時にテナント一覧を取得
window.onload = () => loadTenants();

// テナント一覧取得
async function loadTenants() {
  const tbody = document.getElementById('tenantList');
  try {
    const res = await fetch(GAS_URL);
    const json = await res.json();
    
    if (json.status === "success") {
      tbody.innerHTML = "";
      if (json.companies.length === 0) {
        tbody.innerHTML = "<tr><td colspan='4' class='text-center py-3'>登録されているテナントはありません。</td></tr>";
        return;
      }
      json.companies.forEach(c => {
        tbody.innerHTML += `
          <tr>
            <td class="ps-3 fw-bold text-primary">${c.companyId}</td>
            <td class="fw-bold">${c.companyName}</td>
            <td class="small">${c.createdAt ? new Date(c.createdAt).toLocaleDateString() : '-'}</td>
            <td class="text-center">
              <button onclick="deleteTenant('${c.companyId}')" class="btn btn-outline-danger btn-sm fw-bold">削除</button>
            </td>
          </tr>
        `;
      });
    }
  } catch (e) {
    tbody.innerHTML = "<tr><td colspan='4' class='text-center py-3 text-danger'>データ取得エラー</td></tr>";
  }
}

// テナント新規登録
async function registerTenant() {
  const name = document.getElementById('compName').value;
  const email = document.getElementById('adminEmail').value;
  const pass = document.getElementById('adminPass').value;
  const resDiv = document.getElementById('regResult');

  if (!name || !email || !pass) return alert("全て入力してください。");
  
  resDiv.innerText = "SaaS環境を構築中...\n（ドライブ生成等のため約10〜20秒かかります）";
  resDiv.className = "mt-3 text-secondary";

  try {
    const payload = { action: "registerCompany", companyName: name, adminEmail: email, adminPass: pass };
    const res = await fetch(GAS_URL, { method: "POST", body: JSON.stringify(payload), headers: { 'Content-Type': 'text/plain' } });
    const json = await res.json();
    
    if (json.status === "success") {
      resDiv.innerText = json.message;
      resDiv.className = "mt-3 text-success";
      document.getElementById('compName').value = "";
      document.getElementById('adminEmail').value = "";
      document.getElementById('adminPass').value = "";
      loadTenants(); // 一覧を再取得
    } else {
      resDiv.innerText = "エラー: " + json.message;
      resDiv.className = "mt-3 text-danger";
    }
  } catch (e) {
    resDiv.innerText = "通信エラーが発生しました。";
    resDiv.className = "mt-3 text-danger";
  }
}

// テナント削除
async function deleteTenant(compId) {
  if (!confirm(`本当にテナント [${compId}] を削除しますか？\n※Googleドライブ上のフォルダはゴミ箱に移動されます。`)) return;
  
  try {
    const payload = { action: "deleteCompany", companyId: compId };
    const res = await fetch(GAS_URL, { method: "POST", body: JSON.stringify(payload), headers: { 'Content-Type': 'text/plain' } });
    const json = await res.json();
    
    if (json.status === "success") {
      alert(json.message);
      loadTenants(); // 一覧を再取得
    } else {
      alert("エラー: " + json.message);
    }
  } catch (e) {
    alert("通信エラーが発生しました。");
  }
}