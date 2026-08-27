// 슬롯마다 파일 목록(배열). 대조 모드는 한 슬롯에 여러 개가 들어올 수 있다.
const files = { a: [], b: [] };
let currentMode = "summary";
let currentSubmode = "summary";

const MODE_LABELS = {
  summary: { a: "파일 A (예: 검사요청서)", b: "파일 B (예: 검사보고서)" },
  evidence: { a: "파일 A: 노무비 지급 내역서", b: "파일 B: 퇴직공제부금 납부 신고 내역" },
  hrcost: { a: "파일 1: 보험료 납부_단위공사별 (엑셀) 또는 실적정산 (PDF)", b: "파일 2: 퇴직공제부금 납부 신고 내역 (엑셀) 또는 실적정산 (PDF)" },
  ledger: {
    a: "파일 1: 노무비 지급 명세서 (엑셀)",
    b: "파일 2: 보험료 납부 원장 · 퇴직공제부금 신고 내역 · 공단 발급 PDF (여러 개 선택 가능) — 넣으면 출역x보험료 대조까지",
  },
};
const MODE_ENDPOINT = {
  summary: "/api/compare",
  evidence: "/api/compare_evidence",
  hrcost: "/api/hr_cost/extract",
  ledger: "/api/labor_ledger/extract",
};
const MODE_BUTTON_LABEL = {
  summary: "비교하기",
  evidence: "비교하기",
  hrcost: "추출하기",
  ledger: "정규화하기",
};
// 파일 한 개만 올려도 되는 모드 (비교 모드는 두 개가 다 있어야 한다)
const SINGLE_FILE_MODES = new Set(["hrcost", "ledger"]);

// 모드마다 다루는 서류가 정해져 있는데, 엉뚱한 조합을 넣으면 스캔본 PDF를 수십 분
// OCR한 끝에 빈 결과가 나온다(실제 제보: "기성고 요약표 비교"에 보험료 원장 엑셀과
// 349페이지 실적정산 PDF를 넣고 2분 대기). 파일명으로 미리 알아채 경고한다 --
// 파일명은 바뀔 수 있으니 막지는 않고 확인만 받는다.
const MODE_DOC_HINTS = {
  summary: {
    label: "기성고 요약표 비교",
    patterns: [/검\s*사\s*요\s*청\s*서/, /검\s*사\s*보\s*고\s*서/, /기성고/, /기성\s*검사/],
  },
  evidence: {
    label: "노무비 · 퇴직공제부금 대조",
    patterns: [/노무비\s*지급\s*내역서/, /퇴직공제부금[^]*?신고/],
  },
  hrcost: {
    label: "데이터 추출",
    patterns: [
      /보험료\s*납부/, /단위공사별/, /실적정산/, /납부확인서/, /가입자\s*명부/,
      /결정내역서/, /월별가입자/, /장기요양/, /국민연금/, /건강보험/,
    ],
  },
  ledger: {
    label: "노무비 양식 통일",
    patterns: [/노무비.*(명세|대장)/, /노임\s*대장/, /지급\s*명세/],
  },
};

/** 선택한 파일들이 현재 모드와 맞는지 파일명으로 살펴, 다른 모드 서류로 보이면 안내 문구를 만든다. */
function modeMismatchWarning() {
  const names = ["a", "b"].flatMap((slot) => (files[slot] || []).map((f) => f.name));
  if (!names.length) return null;

  const hint = MODE_DOC_HINTS[currentMode];
  if (!hint) return null;
  // 현재 모드 서류가 하나라도 있으면 사용자가 의도한 조합으로 본다.
  if (names.some((n) => hint.patterns.some((re) => re.test(n)))) return null;

  const suspects = new Map();
  for (const name of names) {
    for (const [mode, other] of Object.entries(MODE_DOC_HINTS)) {
      if (mode === currentMode) continue;
      if (other.patterns.some((re) => re.test(name))) {
        if (!suspects.has(other.label)) suspects.set(other.label, []);
        suspects.get(other.label).push(name);
      }
    }
  }
  if (!suspects.size) return null;  // 아무 패턴에도 안 걸리면(임의 파일명) 조용히 통과

  const detail = [...suspects.entries()]
    .map(([label, files]) => `"${label}" 서류로 보임: ${files.join(", ")}`)
    .join(" / ");
  return `지금 모드는 "${hint.label}"인데, 선택한 파일이 다른 모드 서류 같습니다. ${detail}`;
}

function renderModeWarning() {
  const el = document.getElementById("mode-mismatch-warning");
  if (!el) return;
  const message = modeMismatchWarning();
  el.textContent = message ? `⚠ ${message}` : "";
  el.hidden = !message;
}

function applyModeUI() {
  document.getElementById("dz-label-a").textContent = MODE_LABELS[currentMode].a;
  document.getElementById("dz-label-b").textContent = MODE_LABELS[currentMode].b;
  const accept =
    currentMode === "evidence" ? ".pdf" : currentMode === "ledger" ? ".xlsx,.xlsm,.pdf" : ".pdf,.xlsx,.xlsm,.xls";
  document.getElementById("input-a").accept = accept;
  document.getElementById("input-b").accept = accept;
  // 대조에 쓸 보험료 원장과 퇴직공제 신고 내역을 한 슬롯에 같이 넣을 수 있게 한다.
  document.getElementById("input-b").multiple = currentMode === "ledger";
  document.getElementById("compare-btn").textContent = MODE_BUTTON_LABEL[currentMode];
  updateCompareButton();
  renderModeWarning();
}

const MODAL_TITLES = {
  hrcost: "데이터 추출",
  ledger: "노무비 명세서 양식 통일",
  compare: "신규 기성 검증 및 서류 업로드",
};

function setTopMode(topTab) {
  // topTab: "compare" | "hrcost" | "ledger"
  document.getElementById("mode-compare-desc").hidden = topTab !== "compare";
  document.getElementById("mode-hrcost-desc").hidden = topTab !== "hrcost";
  document.getElementById("mode-ledger-desc").hidden = topTab !== "ledger";
  document.getElementById("compare-submode-section").hidden = topTab !== "compare";
  document.getElementById("hrcost-hint").hidden = !SINGLE_FILE_MODES.has(topTab);
  document.getElementById("result-container").innerHTML = "";
  setStatus("");

  currentMode = topTab === "compare" ? currentSubmode : topTab;
  applyModeUI();
}

function openModal(topTab) {
  const modal = document.getElementById("verify-modal");
  document.getElementById("modal-title").textContent = MODAL_TITLES[topTab] || MODAL_TITLES.compare;
  setTopMode(topTab);
  modal.hidden = false;
}

function closeModal() {
  document.getElementById("verify-modal").hidden = true;
}

function setupModalTriggers() {
  document.getElementById("open-compare-modal-btn").addEventListener("click", () => openModal("compare"));
  document.getElementById("open-hrcost-modal-btn").addEventListener("click", () => openModal("hrcost"));
  document.getElementById("open-ledger-modal-btn").addEventListener("click", () => openModal("ledger"));
  document.getElementById("modal-close-btn").addEventListener("click", closeModal);
  document.getElementById("modal-cancel-btn").addEventListener("click", closeModal);
  document.getElementById("verify-modal").addEventListener("click", (e) => {
    if (e.target.id === "verify-modal") closeModal();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !document.getElementById("verify-modal").hidden) closeModal();
  });
}

function setupSubmodeCards() {
  document.querySelectorAll(".submode-card").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".submode-card").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      currentSubmode = btn.dataset.submode;
      currentMode = currentSubmode;
      document.getElementById("result-container").innerHTML = "";
      setStatus("");
      applyModeUI();
    });
  });
}

function setupDropzone(slot) {
  const zone = document.getElementById(`dropzone-${slot}`);
  const input = document.getElementById(`input-${slot}`);
  const nameLabel = document.getElementById(`filename-${slot}`);

  // 슬롯은 항상 배열로 들고 있는다 -- 노무비x보험료 대조는 보험료 원장과 퇴직공제
  // 신고 내역을 같이 올려야 해서 한 슬롯에 여러 개가 들어올 수 있다(ledger 모드에서만
  // input에 multiple을 켠다). 다른 모드는 1개만 들어와 기존 동작과 같다.
  const setFile = (fileList) => {
    const picked = Array.from(fileList || []).filter(Boolean);
    if (!picked.length) return;
    files[slot] = picked;
    nameLabel.textContent = picked.map((f) => f.name).join(", ");
    zone.classList.add("has-file");
    updateCompareButton();
    renderModeWarning();
  };

  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", (e) => setFile(e.target.files));

  zone.addEventListener("dragover", (e) => {
    e.preventDefault();
    zone.classList.add("dragover");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("dragover");
    if (e.dataTransfer.files.length) setFile(e.dataTransfer.files);
  });
}

const hasFiles = (slot) => !!(files[slot] && files[slot].length);

function updateCompareButton() {
  const ready = SINGLE_FILE_MODES.has(currentMode)
    ? hasFiles("a") || hasFiles("b")
    : hasFiles("a") && hasFiles("b");
  document.getElementById("compare-btn").disabled = !ready;
}

function setStatus(msg, isError) {
  const el = document.getElementById("status-msg");
  el.textContent = msg;
  el.classList.toggle("error", !!isError);
}

const MB = 1024 * 1024;
const maxUploadBytes = () => parseInt(document.body.dataset.maxUploadBytes || "0", 10);
const totalUploadBytes = () =>
  ["a", "b"].reduce((sum, slot) => sum + (files[slot] || []).reduce((s, f) => s + f.size, 0), 0);

// 오래 걸리는 추출(스캔본 OCR은 수십 분)에 경과 시간을 보여준다. 이게 없으면
// 진행 중인지 멈춘 것인지 구분할 수 없어 사용자가 멈춘 것으로 오해한다.
let progressTimer = null;

function startProgress(baseMessage) {
  const startedAt = Date.now();
  const tick = () => {
    const sec = Math.floor((Date.now() - startedAt) / 1000);
    const stamp = sec < 60 ? `${sec}초 경과` : `${Math.floor(sec / 60)}분 ${sec % 60}초 경과`;
    setStatus(`${baseMessage} (${stamp})`);
  };
  tick();
  progressTimer = setInterval(tick, 1000);
}

function stopProgress() {
  if (progressTimer) clearInterval(progressTimer);
  progressTimer = null;
}

async function runCompare() {
  const hasPdf = ["a", "b"].some((slot) =>
    (files[slot] || []).some((f) => f.name.toLowerCase().endsWith(".pdf"))
  );

  // 보내기 전에 용량을 확인한다. 상한을 넘긴 요청은 서버가 본문을 받지 않고 끊어서
  // 브라우저가 응답을 받지 못하고 "업로드 중"에서 매달린다(실측 재현) -- 그래서
  // 아예 보내지 않고 여기서 알려주는 게 유일하게 확실한 방법이다.
  const limit = maxUploadBytes();
  const total = totalUploadBytes();
  if (limit && total > limit) {
    setStatus(
      `선택한 파일 합계가 ${(total / MB).toFixed(1)}MB로 업로드 상한 ` +
        `${Math.round(limit / MB)}MB를 넘습니다. 파일을 나눠서 올려주세요.`,
      true
    );
    return;
  }

  // 모드에 안 맞는 조합이면 한 번 확인받는다 -- 그냥 보내면 스캔본 OCR로 수십 분을
  // 쓴 뒤 빈 결과가 나온다.
  const mismatch = modeMismatchWarning();
  if (mismatch && !window.confirm(`${mismatch}\n\n그래도 이 모드로 진행할까요?`)) {
    setStatus("취소했습니다. 위 안내를 참고해 모드를 바꾸거나 파일을 다시 선택해주세요.", true);
    return;
  }

  const base =
    hasPdf && (currentMode === "hrcost" || currentMode === "ledger" || currentMode === "evidence")
      ? "업로드 및 OCR 추출 중... 스캔본 PDF는 페이지 수에 따라 수 분~1시간 이상 걸립니다"
      : hasPdf
      ? "업로드 및 추출 중... 스캔본 PDF는 OCR이 필요해 오래 걸릴 수 있습니다"
      : "업로드 및 추출 중...";
  startProgress(base);
  document.getElementById("compare-btn").disabled = true;

  const form = new FormData();
  ["a", "b"].forEach((slot) => {
    (files[slot] || []).forEach((file) => form.append(`file_${slot}`, file));
  });

  const endpoint = MODE_ENDPOINT[currentMode];
  try {
    const res = await fetch(endpoint, { method: "POST", body: form });
    // 오류 응답이 JSON이 아닐 수 있다(413 등은 서버 기본 HTML 페이지). 그 경우
    // res.json()이 예외를 던져 실제 원인 대신 "네트워크 오류"만 보였다.
    let data = null;
    try {
      data = await res.json();
    } catch (parseErr) {
      data = null;
    }
    if (!res.ok) {
      stopProgress();
      const fallback =
        res.status === 413
          ? `업로드 용량 상한을 넘었습니다 (보낸 크기 ${(total / MB).toFixed(1)}MB).`
          : `서버 오류가 발생했습니다 (HTTP ${res.status}).`;
      setStatus((data && data.error) || fallback, true);
      updateCompareButton();
      return;
    }
    stopProgress();
    document.getElementById("result-container").innerHTML = data.html;
    setStatus("완료");
    attachVerifyHandlers();
    attachHrCostSaveHandler();
    setupSortableTables();
    setupTableFilters();
    setupHrCostCategoryFilter();
    setupHrCostHistorySection();
  } catch (err) {
    stopProgress();
    // 상한을 넘긴 업로드는 서버가 연결을 끊어 여기로 떨어지기도 한다.
    const hint =
      limit && total > limit * 0.9
        ? ` 선택한 파일 합계가 ${(total / MB).toFixed(1)}MB로 업로드 상한에 가깝습니다 — 파일을 나눠 올려보세요.`
        : "";
    setStatus(`연결이 끊겼습니다. 업로드가 완료되지 못했습니다.${hint}`, true);
  } finally {
    stopProgress();
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

function setupHrCostHistorySection() {
  const resultEl = document.querySelector(".hr-cost-result");
  const section = resultEl && resultEl.querySelector(".history-section");
  if (!section) return;

  const companySelect = section.querySelector(".history-company");
  const monthsSelect = section.querySelector(".history-months");
  const compareBtn = section.querySelector(".history-compare-btn");
  const statusEl = section.querySelector(".history-status-msg");
  const outputEl = section.querySelector(".history-output");

  const setHistoryStatus = (msg, isError) => {
    statusEl.textContent = msg;
    statusEl.style.color = isError ? "#b3261e" : "";
  };

  const updateCompareBtn = () => {
    const months = Array.from(monthsSelect.selectedOptions).map((o) => o.value);
    compareBtn.disabled = !(companySelect.value && months.length);
  };

  const loadMonths = async (company) => {
    monthsSelect.innerHTML = "";
    monthsSelect.disabled = true;
    if (!company) {
      updateCompareBtn();
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
      monthsSelect.disabled = false;
      setHistoryStatus(data.months && data.months.length ? "" : "저장된 데이터가 없습니다.");
    } catch (err) {
      setHistoryStatus("연월 목록을 불러오지 못했습니다.", true);
    }
    updateCompareBtn();
  };

  companySelect.addEventListener("change", (e) => loadMonths(e.target.value));
  monthsSelect.addEventListener("change", updateCompareBtn);

  compareBtn.addEventListener("click", async () => {
    const company = companySelect.value;
    const months = Array.from(monthsSelect.selectedOptions).map((o) => o.value);
    setHistoryStatus("조회 중...");
    compareBtn.disabled = true;
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
      outputEl.innerHTML = data.html;
      setHistoryStatus("완료");
      setupSortableTables();
      setupTableFilters();
    } catch (err) {
      setHistoryStatus("네트워크 오류가 발생했습니다.", true);
    } finally {
      updateCompareBtn();
    }
  });

  (async () => {
    companySelect.innerHTML = '<option value="">불러오는 중...</option>';
    try {
      const res = await fetch("/api/history/companies");
      const data = await res.json();
      companySelect.innerHTML = '<option value="">선택...</option>';
      (data.companies || []).forEach((c) => {
        const opt = document.createElement("option");
        opt.value = c;
        opt.textContent = c;
        companySelect.appendChild(opt);
      });
      let extracted = [];
      try {
        extracted = JSON.parse(resultEl.dataset.companies || "[]");
      } catch (e) {
        extracted = [];
      }
      if (extracted.length === 1 && data.companies && data.companies.includes(extracted[0])) {
        companySelect.value = extracted[0];
        await loadMonths(extracted[0]);
      }
    } catch (err) {
      companySelect.innerHTML = '<option value="">불러오기 실패</option>';
    }
    updateCompareBtn();
  })();
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

function setupHrCostCategoryFilter() {
  const select = document.getElementById("hr-cost-category-select");
  if (!select) return;

  const applyFilter = () => {
    const selected = new Set(Array.from(select.selectedOptions).map((o) => o.value));

    document.querySelectorAll(".hr-cost-table .cat-col").forEach((cell) => {
      cell.classList.toggle("cat-hidden", !selected.has(cell.dataset.cat));
    });

    document.querySelectorAll(".hr-cost-table .sum-cell").forEach((cell) => {
      let sum = 0;
      selected.forEach((cat) => {
        sum += parseInt(cell.dataset[cat] || "0", 10) || 0;
      });
      cell.textContent = sum.toLocaleString("en-US");
    });
  };

  select.addEventListener("change", applyFilter);
  applyFilter();
}

// 텍스트 검색과 등급 탭이 같은 표를 함께 걸러야 해서 한 곳에서 판정한다 --
// 각자 row.hidden을 건드리면 나중에 실행된 쪽이 앞의 조건을 지운다.
function applyRowFilters(table) {
  if (!table) return;
  const input = document.querySelector(`.table-filter[data-target="${table.id}"]`);
  const tabs = document.querySelector(`.tier-tabs[data-target="${table.id}"]`);
  const query = input ? input.value.trim().toLowerCase() : "";
  const active = tabs ? tabs.querySelector(".tier-tab.is-active") : null;
  const tier = active ? active.dataset.tier : "";
  table.querySelectorAll("tbody tr").forEach((row) => {
    const matchesText = !query || row.textContent.toLowerCase().includes(query);
    const matchesTier = !tier || row.dataset.tier === tier;
    row.hidden = !(matchesText && matchesTier);
  });
}

function setupTableFilters() {
  document.querySelectorAll(".table-filter").forEach((input) => {
    input.addEventListener("input", () => {
      applyRowFilters(document.getElementById(input.dataset.target));
    });
  });
  document.querySelectorAll(".tier-tabs").forEach((tabs) => {
    tabs.addEventListener("click", (event) => {
      const button = event.target.closest(".tier-tab");
      if (!button) return;
      tabs.querySelectorAll(".tier-tab").forEach((b) => b.classList.toggle("is-active", b === button));
      applyRowFilters(document.getElementById(tabs.dataset.target));
    });
  });
}

setupDropzone("a");
setupDropzone("b");
setupSubmodeCards();
setupModalTriggers();
applyModeUI();
document.getElementById("compare-btn").addEventListener("click", runCompare);
