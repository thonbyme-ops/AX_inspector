const files = { a: null, b: null };
let currentMode = "summary";

const MODE_LABELS = {
  summary: { a: "파일 A (예: 검사요청서)", b: "파일 B (예: 검사보고서)" },
  evidence: { a: "파일 A: 노무비 지급 내역서", b: "파일 B: 퇴직공제부금 납부 신고 내역" },
  hrcost: { a: "파일 1: 보험료 납부_단위공사별 (엑셀) 또는 실적정산 (PDF)", b: "파일 2: 퇴직공제부금 납부 신고 내역 (엑셀) 또는 실적정산 (PDF)" },
};
const MODE_ENDPOINT = {
  summary: "/api/compare",
  evidence: "/api/compare_evidence",
  hrcost: "/api/hr_cost/extract",
};
const MODE_BUTTON_LABEL = {
  summary: "비교하기",
  evidence: "비교하기",
  hrcost: "추출하기",
};

function setupModeTabs() {
  document.querySelectorAll(".mode-tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".mode-tab").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      currentMode = btn.dataset.mode;
      ["summary", "evidence", "hrcost", "history"].forEach((m) => {
        document.getElementById(`mode-${m}-desc`).hidden = currentMode !== m;
      });
      document.getElementById("hrcost-hint").hidden = currentMode !== "hrcost";
      document.getElementById("upload-section").hidden = currentMode === "history";
      document.getElementById("history-panel").hidden = currentMode !== "history";
      document.getElementById("result-container").innerHTML = "";
      setStatus("");

      if (currentMode === "history") {
        loadHistoryCompanies();
        return;
      }

      document.getElementById("dz-label-a").textContent = MODE_LABELS[currentMode].a;
      document.getElementById("dz-label-b").textContent = MODE_LABELS[currentMode].b;
      const accept = currentMode === "evidence" ? ".pdf" : ".pdf,.xlsx,.xlsm,.xls";
      document.getElementById("input-a").accept = accept;
      document.getElementById("input-b").accept = accept;
      document.getElementById("compare-btn").textContent = MODE_BUTTON_LABEL[currentMode];
      updateCompareButton();
    });
  });
}

function setupDropzone(slot) {
  const zone = document.getElementById(`dropzone-${slot}`);
  const input = document.getElementById(`input-${slot}`);
  const nameLabel = document.getElementById(`filename-${slot}`);

  const setFile = (file) => {
    if (!file) return;
    files[slot] = file;
    nameLabel.textContent = file.name;
    zone.classList.add("has-file");
    updateCompareButton();
  };

  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", (e) => setFile(e.target.files[0]));

  zone.addEventListener("dragover", (e) => {
    e.preventDefault();
    zone.classList.add("dragover");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("dragover");
    if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
  });
}

function updateCompareButton() {
  const ready = currentMode === "hrcost" ? !!(files.a || files.b) : !!(files.a && files.b);
  document.getElementById("compare-btn").disabled = !ready;
}

function setStatus(msg, isError) {
  const el = document.getElementById("status-msg");
  el.textContent = msg;
  el.classList.toggle("error", !!isError);
}

async function runCompare() {
  const hasPdf = (files.a && files.a.name.toLowerCase().endsWith(".pdf")) ||
    (files.b && files.b.name.toLowerCase().endsWith(".pdf"));
  setStatus(
    currentMode === "hrcost" && hasPdf
      ? "업로드 및 OCR 추출 중... (PDF는 페이지 수에 따라 수 분~1시간 이상 걸릴 수 있습니다)"
      : "업로드 및 추출 중..."
  );
  document.getElementById("compare-btn").disabled = true;

  const form = new FormData();
  if (files.a) form.append("file_a", files.a);
  if (files.b) form.append("file_b", files.b);

  const endpoint = MODE_ENDPOINT[currentMode];
  try {
    const res = await fetch(endpoint, { method: "POST", body: form });
    const data = await res.json();
    if (!res.ok) {
      setStatus(data.error || "비교 중 오류가 발생했습니다.", true);
      updateCompareButton();
      return;
    }
    document.getElementById("result-container").innerHTML = data.html;
    setStatus("완료");
    attachVerifyHandlers();
    attachHrCostSaveHandler();
    setupSortableTables();
    setupTableFilters();
  } catch (err) {
    setStatus("네트워크 오류가 발생했습니다.", true);
  } finally {
    updateCompareButton();
  }
}

function attachHrCostSaveHandler() {
  const resultEl = document.querySelector(".hr-cost-result");
  const btn = document.getElementById("hr-cost-save-btn");
  if (!resultEl || !btn) return;
  const token = resultEl.dataset.token;

  btn.addEventListener("click", async () => {
    btn.disabled = true;
    const msgEl = document.getElementById("hr-cost-save-msg");
    msgEl.textContent = "저장 중...";
    msgEl.style.color = "";
    try {
      const res = await fetch(`/api/hr_cost/save/${token}`, { method: "POST" });
      const data = await res.json();
      if (!res.ok) {
        msgEl.textContent = data.error || "저장 중 오류가 발생했습니다.";
        msgEl.style.color = "#b3261e";
        btn.disabled = false;
        return;
      }
      msgEl.textContent = "DB에 저장되었습니다.";
      msgEl.style.color = "#2e9e56";
    } catch (err) {
      msgEl.textContent = "네트워크 오류가 발생했습니다.";
      msgEl.style.color = "#b3261e";
      btn.disabled = false;
    }
  });
}

function setHistoryStatus(msg, isError) {
  const el = document.getElementById("history-status-msg");
  el.textContent = msg;
  el.style.color = isError ? "#b3261e" : "";
}

async function loadHistoryCompanies() {
  const select = document.getElementById("history-company");
  select.innerHTML = '<option value="">불러오는 중...</option>';
  document.getElementById("history-months").innerHTML = "";
  try {
    const res = await fetch("/api/history/companies");
    const data = await res.json();
    select.innerHTML = '<option value="">선택...</option>';
    (data.companies || []).forEach((c) => {
      const opt = document.createElement("option");
      opt.value = c;
      opt.textContent = c;
      select.appendChild(opt);
    });
  } catch (err) {
    select.innerHTML = '<option value="">불러오기 실패</option>';
  }
  updateHistoryCompareButton();
}

async function loadHistoryMonths(company) {
  const monthsSelect = document.getElementById("history-months");
  monthsSelect.innerHTML = "";
  if (!company) {
    updateHistoryCompareButton();
    return;
  }
  try {
    const res = await fetch(`/api/history/months?company=${encodeURIComponent(company)}`);
    const data = await res.json();
    (data.months || []).forEach((m) => {
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = m;
      monthsSelect.appendChild(opt);
    });
    if (!data.months || !data.months.length) {
      setHistoryStatus("저장된 데이터가 없습니다.");
    } else {
      setHistoryStatus("");
    }
  } catch (err) {
    setHistoryStatus("연월 목록을 불러오지 못했습니다.", true);
  }
  updateHistoryCompareButton();
}

function updateHistoryCompareButton() {
  const company = document.getElementById("history-company").value;
  const months = Array.from(document.getElementById("history-months").selectedOptions).map((o) => o.value);
  document.getElementById("history-compare-btn").disabled = !(company && months.length);
}

async function runHistoryCompare() {
  const company = document.getElementById("history-company").value;
  const months = Array.from(document.getElementById("history-months").selectedOptions).map((o) => o.value);
  setHistoryStatus("조회 중...");
  document.getElementById("history-compare-btn").disabled = true;
  try {
    const res = await fetch("/api/history/compare", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ company, months }),
    });
    const data = await res.json();
    if (!res.ok) {
      setHistoryStatus(data.error || "조회 중 오류가 발생했습니다.", true);
      return;
    }
    document.getElementById("result-container").innerHTML = data.html;
    setHistoryStatus("완료");
    setupSortableTables();
    setupTableFilters();
  } catch (err) {
    setHistoryStatus("네트워크 오류가 발생했습니다.", true);
  } finally {
    updateHistoryCompareButton();
  }
}

function attachVerifyHandlers() {
  const resultEl = document.querySelector(".result");
  if (!resultEl) return;
  const token = resultEl.dataset.token;

  resultEl.querySelectorAll(".verify-buttons button").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const decision = btn.dataset.decision;
      const supervisor = document.getElementById("supervisor-name").value.trim();
      const verifyResultEl = document.getElementById("verify-result");
      if (!supervisor) {
        verifyResultEl.textContent = "검증자 성명을 입력해주세요.";
        verifyResultEl.style.color = "#b3261e";
        return;
      }
      try {
        const res = await fetch("/api/verify", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ token, decision, supervisor }),
        });
        const data = await res.json();
        if (!res.ok) {
          verifyResultEl.textContent = data.description || data.error || "검증 저장에 실패했습니다.";
          verifyResultEl.style.color = "#b3261e";
          return;
        }
        verifyResultEl.textContent = `검증 완료: ${decision === "yes" ? "일치함" : "불일치함"} (검증자: ${supervisor})`;
        verifyResultEl.style.color = "#2e9e56";
      } catch (err) {
        verifyResultEl.textContent = "네트워크 오류가 발생했습니다.";
        verifyResultEl.style.color = "#b3261e";
      }
    });
  });
}

function parseCellValue(text, type) {
  if (type === "number") return parseFloat(text.replace(/,/g, "")) || 0;
  return text.trim();
}

function setupSortableTables() {
  document.querySelectorAll(".sortable-table").forEach((table) => {
    const headers = table.querySelectorAll("thead th.sortable");
    headers.forEach((th, colIndex) => {
      th.addEventListener("click", () => {
        const type = th.dataset.type || "text";
        const asc = th.dataset.sortDir !== "asc";
        headers.forEach((h) => {
          h.classList.remove("sort-asc", "sort-desc");
          delete h.dataset.sortDir;
        });
        th.dataset.sortDir = asc ? "asc" : "desc";
        th.classList.add(asc ? "sort-asc" : "sort-desc");

        const tbody = table.querySelector("tbody");
        const rows = Array.from(tbody.querySelectorAll("tr"));
        rows.sort((r1, r2) => {
          const v1 = parseCellValue(r1.children[colIndex].textContent, type);
          const v2 = parseCellValue(r2.children[colIndex].textContent, type);
          if (v1 < v2) return asc ? -1 : 1;
          if (v1 > v2) return asc ? 1 : -1;
          return 0;
        });
        rows.forEach((row) => tbody.appendChild(row));
      });
    });
  });
}

function setupTableFilters() {
  document.querySelectorAll(".table-filter").forEach((input) => {
    input.addEventListener("input", () => {
      const table = document.getElementById(input.dataset.target);
      if (!table) return;
      const query = input.value.trim().toLowerCase();
      table.querySelectorAll("tbody tr").forEach((row) => {
        const text = row.textContent.toLowerCase();
        row.hidden = query.length > 0 && !text.includes(query);
      });
    });
  });
}

setupDropzone("a");
setupDropzone("b");
setupModeTabs();
document.getElementById("compare-btn").addEventListener("click", runCompare);
document.getElementById("history-company").addEventListener("change", (e) => loadHistoryMonths(e.target.value));
document.getElementById("history-months").addEventListener("change", updateHistoryCompareButton);
document.getElementById("history-compare-btn").addEventListener("click", runHistoryCompare);
