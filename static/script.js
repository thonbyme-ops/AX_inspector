const files = { a: null, b: null };
let currentMode = "summary";

const MODE_LABELS = {
  summary: { a: "파일 A (예: 검사요청서)", b: "파일 B (예: 검사보고서)" },
  evidence: { a: "파일 A: 노무비 지급 내역서", b: "파일 B: 퇴직공제부금 납부 신고 내역" },
  hrcost: { a: "파일 1: 보험료 납부_단위공사별", b: "파일 2: 퇴직공제부금 납부 신고 내역" },
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
      document.getElementById("mode-summary-desc").hidden = currentMode !== "summary";
      document.getElementById("mode-evidence-desc").hidden = currentMode !== "evidence";
      document.getElementById("mode-hrcost-desc").hidden = currentMode !== "hrcost";
      document.getElementById("hrcost-hint").hidden = currentMode !== "hrcost";
      document.getElementById("dz-label-a").textContent = MODE_LABELS[currentMode].a;
      document.getElementById("dz-label-b").textContent = MODE_LABELS[currentMode].b;
      const accept =
        currentMode === "evidence" ? ".pdf" : currentMode === "hrcost" ? ".xlsx,.xlsm,.xls" : ".pdf,.xlsx,.xlsm,.xls";
      document.getElementById("input-a").accept = accept;
      document.getElementById("input-b").accept = accept;
      document.getElementById("compare-btn").textContent = MODE_BUTTON_LABEL[currentMode];
      document.getElementById("result-container").innerHTML = "";
      setStatus("");
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
  setStatus("업로드 및 추출 중...");
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
  } catch (err) {
    setStatus("네트워크 오류가 발생했습니다.", true);
  } finally {
    updateCompareButton();
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

setupDropzone("a");
setupDropzone("b");
setupModeTabs();
document.getElementById("compare-btn").addEventListener("click", runCompare);
