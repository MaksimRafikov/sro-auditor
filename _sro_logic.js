
    const INF = Number.POSITIVE_INFINITY;
    const state = {
      membersRows: null,
      contractsRows: null,
      result: null,
      tab: "companies",
    };

    const membersInput = document.getElementById("membersFile");
    const contractsInput = document.getElementById("contractsFile");
    const runBtn = document.getElementById("runBtn");
    const sampleBtn = document.getElementById("sampleBtn");
    const exportBtn = document.getElementById("exportBtn");
    const errorBox = document.getElementById("errorBox");
    const fileHint = document.getElementById("fileHint");
    const results = document.getElementById("results");

    function showError(msg) {
      errorBox.style.display = "block";
      errorBox.textContent = msg;
    }
    function clearError() {
      errorBox.style.display = "none";
      errorBox.textContent = "";
    }

    function updateReady() {
      const ok = !!(state.membersRows && state.contractsRows);
      runBtn.disabled = !ok;
      fileHint.textContent = ok
        ? `Готово: ${state.membersRows.length} строк реестра, ${state.contractsRows.length} договоров.`
        : "Выберите оба файла или нажмите «демо-данные».";
    }

    async function readTable(file) {
      if (typeof XLSX === "undefined") {
        throw new Error(
          "Библиотека Excel не загрузилась. Откройте страницу через локальный сервер (python -m http.server) из папки проекта — нужен файл vendor/xlsx.full.min.js."
        );
      }
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      return XLSX.utils.sheet_to_json(sheet, { defval: "" });
    }

    function normKey(s) {
      return String(s || "")
        .toLowerCase()
        .replace(/ё/g, "е")
        .replace(/[^a-zа-я0-9]+/gi, "");
    }

    function findCol(headers, aliases) {
      const normalized = headers.map((h) => ({ raw: h, n: normKey(h) }));
      for (const alias of aliases) {
        const a = normKey(alias);
        const hit = normalized.find((h) => h.n === a || h.n.includes(a));
        if (hit) return hit.raw;
      }
      return null;
    }

    function normalizeInn(v) {
      const digits = String(v ?? "").replace(/\D/g, "");
      return digits || "";
    }

    function parseMoney(v) {
      if (v === null || v === undefined || v === "") return null;
      if (typeof v === "number" && Number.isFinite(v)) return v;
      let s = String(v).trim();
      if (!s) return null;
      s = s.replace(/\s/g, "").replace(/₽|руб\.?|р\./gi, "");
      if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) {
        s = s.replace(/\./g, "").replace(",", ".");
      } else if (s.includes(",") && s.includes(".")) {
        s = s.replace(/\./g, "").replace(",", ".");
      } else if (s.includes(",")) {
        s = s.replace(",", ".");
      }
      const n = Number(s);
      return Number.isFinite(n) ? n : null;
    }

    function parseLimit(text) {
      if (text === null || text === undefined || text === "") return null;
      const s = String(text).toLowerCase().replace(/ё/g, "е");
      if (/отсут|не\s*указ|нет\b|—|-/.test(s) && !/\d/.test(s)) return null;
      if (/свыше\s*10|10\s*млрд.*и\s*более|более\s*10/.test(s)) return INF;
      const m = s.match(/([\d\s.,]+)\s*(млрд|млн|тыс)?/);
      if (!m) {
        const n = parseMoney(text);
        return n;
      }
      let num = parseMoney(m[1]);
      if (num === null) return null;
      const unit = m[2] || "";
      if (unit.startsWith("млрд")) num *= 1_000_000_000;
      else if (unit.startsWith("млн")) num *= 1_000_000;
      else if (unit.startsWith("тыс")) num *= 1_000;
      return num;
    }

    function normalizeRight(text) {
      const s = String(text || "").toLowerCase().replace(/ё/g, "е");
      if (!s.trim()) return "неизвестно";
      if (/приостан/.test(s)) return "приостановлено";
      if (/исключ|прекращ/.test(s)) return "прекращено";
      if (/действ|актив/.test(s)) return "действует";
      return s.trim();
    }

    function isCompetitive(method, odoFlag) {
      const flag = String(odoFlag ?? "").toLowerCase();
      if (["1", "true", "да", "yes", "y", "+"].includes(flag)) return { competitive: true, unclear: false };
      if (["0", "false", "нет", "no", "n"].includes(flag)) {
        // still may be competitive by method
      }
      const m = String(method ?? "").toLowerCase();
      if (!m.trim()) return { competitive: false, unclear: true };
      if (/44|223|615|конкурс|аукцион|тендер|закупк|торг/.test(m)) {
        return { competitive: true, unclear: false };
      }
      if (/прям|коммерч|без.*конкур/.test(m)) {
        return { competitive: false, unclear: false };
      }
      return { competitive: false, unclear: true };
    }

    function mapMembers(rows) {
      if (!rows.length) throw new Error("Реестр членов СРО пуст.");
      const headers = Object.keys(rows[0]);
      const cols = {
        inn: findCol(headers, ["инн", "inn"]),
        name: findCol(headers, ["контрагент", "наименование", "организация", "член сро", "название"]),
        right: findCol(headers, ["состояние права", "статус права", "право"]),
        vv: findCol(headers, ["уровень вв", "лимит вв", "вв"]),
        odo: findCol(headers, ["уровень одо", "лимит одо", "одо"]),
        oblig: findCol(headers, ["расчет обязательств", "расчёт обязательств", "обязательства"]),
      };
      if (!cols.inn) throw new Error("В реестре членов не найдена колонка ИНН.");
      const byInn = new Map();
      for (const row of rows) {
        const inn = normalizeInn(row[cols.inn]);
        if (!inn) continue;
        byInn.set(inn, {
          inn,
          name: cols.name ? String(row[cols.name] || "").trim() : "",
          right: normalizeRight(cols.right ? row[cols.right] : ""),
          vvText: cols.vv ? String(row[cols.vv] ?? "") : "",
          odoText: cols.odo ? String(row[cols.odo] ?? "") : "",
          vvLimit: cols.vv ? parseLimit(row[cols.vv]) : null,
          odoLimit: cols.odo ? parseLimit(row[cols.odo]) : null,
          registryOblig: cols.oblig ? parseMoney(row[cols.oblig]) : null,
          _cols: cols,
        });
      }
      return { byInn, cols, count: byInn.size };
    }

    function mapContracts(rows) {
      if (!rows.length) throw new Error("Реестр договоров пуст.");
      const headers = Object.keys(rows[0]);
      const cols = {
        inn: findCol(headers, ["инн", "инн подрядчика", "инн участника", "inn"]),
        name: findCol(headers, ["контрагент", "подрядчик", "участник", "наименование", "компания"]),
        amount: findCol(headers, [
          "стоимость принятая сро к учету",
          "стоимость принятая сро к учёту",
          "стоимость к учету",
          "стоимость к учёту",
          "сумма договора",
          "стоимость",
        ]),
        done: findCol(headers, [
          "стоимость принятых работ",
          "исполнено",
          "принято",
          "стоимость исполненных",
        ]),
        method: findCol(headers, ["вид закупки", "способ закупки", "закон", "фз", "тип закупки"]),
        odoFlag: findCol(headers, ["договор одо", "признак одо", "одо"]),
        number: findCol(headers, ["номер договора", "№ договора", "номер"]),
        date: findCol(headers, ["дата договора", "дата заключения", "дата"]),
      };
      if (!cols.inn) throw new Error("В реестре договоров не найдена колонка ИНН.");
      if (!cols.amount) throw new Error("В реестре договоров не найдена колонка стоимости.");

      const list = [];
      for (const row of rows) {
        const inn = normalizeInn(row[cols.inn]);
        if (!inn) continue;
        const amount = parseMoney(row[cols.amount]);
        const done = cols.done ? parseMoney(row[cols.done]) : 0;
        const weirdMoney = amount === null;
        const method = cols.method ? row[cols.method] : "";
        const odoFlag = cols.odoFlag ? row[cols.odoFlag] : "";
        const { competitive, unclear } = isCompetitive(method, odoFlag);
        const executed = done === null ? 0 : done;
        const residual = weirdMoney ? null : Math.max(0, amount - executed);
        list.push({
          inn,
          name: cols.name ? String(row[cols.name] || "").trim() : "",
          number: cols.number ? String(row[cols.number] ?? "") : "",
          date: cols.date ? String(row[cols.date] ?? "") : "",
          amount,
          done: executed,
          residual,
          method: String(method ?? ""),
          competitive,
          unclear,
          weirdMoney,
          assumptionNoDone: !cols.done,
        });
      }
      return { list, cols };
    }

    function diverges(a, b) {
      if (a === null || b === null || a === undefined || b === undefined) return false;
      const diff = Math.abs(a - b);
      const base = Math.max(Math.abs(a), Math.abs(b), 1);
      return diff > 1_000_000 || diff / base > 0.05;
    }

    function analyze(membersRows, contractsRows) {
      const members = mapMembers(membersRows);
      const contracts = mapContracts(contractsRows);
      const byInnContracts = new Map();
      for (const c of contracts.list) {
        if (!byInnContracts.has(c.inn)) byInnContracts.set(c.inn, []);
        byInnContracts.get(c.inn).push(c);
      }

      const companies = [];
      for (const [inn, list] of byInnContracts.entries()) {
        const m = members.byInn.get(inn) || null;
        const comments = [];
        let risk = "НОРМА";

        const maxContract = list.reduce((mx, c) => {
          if (c.amount === null) return mx;
          return Math.max(mx, c.amount);
        }, 0);

        const competitive = list.filter((c) => c.competitive);
        const unclear = list.filter((c) => c.unclear);
        const odoResidual = competitive.reduce((s, c) => s + (c.residual || 0), 0);

        const found = !!m;
        const right = m ? m.right : "не найдено";
        const vvLimit = m ? m.vvLimit : null;
        const odoLimit = m ? m.odoLimit : null;
        const vvText = m ? m.vvText : "";
        const odoText = m ? m.odoText : "";
        const registryOblig = m ? m.registryOblig : null;
        const name = (m && m.name) || list.find((c) => c.name)?.name || "";

        let vvCheck = "н/д";
        let odoCheck = "н/д";

        if (!found) {
          risk = "КРИТИЧНО";
          comments.push("не найдена в реестре СРО");
        }
        if (found && (right === "приостановлено" || right === "прекращено")) {
          risk = "КРИТИЧНО";
          comments.push(`право: ${right}`);
        }
        if (found && vvLimit !== null && Number.isFinite(vvLimit) && maxContract > vvLimit) {
          risk = "КРИТИЧНО";
          vvCheck = "превышение";
          comments.push("превышение ВВ");
        } else if (found && vvLimit === INF) {
          vvCheck = "ок";
        } else if (found && vvLimit !== null && Number.isFinite(vvLimit)) {
          vvCheck = "ок";
        }

        if (competitive.length > 0) {
          if (!found || odoLimit === null) {
            risk = "КРИТИЧНО";
            odoCheck = "нет ОДО";
            comments.push("есть конкурентные договоры, ОДО отсутствует");
          } else if (Number.isFinite(odoLimit) && odoResidual > odoLimit) {
            risk = "КРИТИЧНО";
            odoCheck = "превышение";
            comments.push("превышение ОДО");
          } else {
            odoCheck = "ок";
          }
        } else {
          odoCheck = "не требуется";
        }

        if (unclear.length > 0) {
          if (risk !== "КРИТИЧНО") risk = "РУЧНАЯ ПРОВЕРКА";
          comments.push(`без способа закупки: ${unclear.length}`);
        }
        if (list.some((c) => c.weirdMoney)) {
          if (risk !== "КРИТИЧНО") risk = "РУЧНАЯ ПРОВЕРКА";
          comments.push("странный формат суммы");
        }
        if (list.some((c) => c.assumptionNoDone)) {
          comments.push("исполнение не найдено — остаток = полная стоимость");
        }
        if (found && diverges(odoResidual, registryOblig)) {
          if (risk !== "КРИТИЧНО") risk = "РУЧНАЯ ПРОВЕРКА";
          comments.push("расхождение остатка ОДО с реестром");
        }

        companies.push({
          inn,
          name,
          found,
          right,
          vvText,
          vvLimit,
          maxContract,
          vvCheck,
          odoText,
          odoLimit,
          odoResidual,
          registryOblig,
          odoCheck,
          contractsCount: list.length,
          competitiveCount: competitive.length,
          risk,
          comment: comments.join("; "),
        });
      }

      companies.sort((a, b) => {
        const order = { КРИТИЧНО: 0, "РУЧНАЯ ПРОВЕРКА": 1, НОРМА: 2 };
        return (order[a.risk] - order[b.risk]) || b.maxContract - a.maxContract;
      });

      const summary = {
        contracts: contracts.list.length,
        inns: companies.length,
        found: companies.filter((c) => c.found).length,
        notFound: companies.filter((c) => !c.found).length,
        suspended: companies.filter((c) => c.right === "приостановлено" || c.right === "прекращено").length,
        vvBreach: companies.filter((c) => c.vvCheck === "превышение").length,
        odoBreach: companies.filter((c) => c.odoCheck === "превышение" || c.odoCheck === "нет ОДО").length,
        critical: companies.filter((c) => c.risk === "КРИТИЧНО").length,
        manual: companies.filter((c) => c.risk === "РУЧНАЯ ПРОВЕРКА").length,
        ok: companies.filter((c) => c.risk === "НОРМА").length,
        sumAccepted: contracts.list.reduce((s, c) => s + (c.amount || 0), 0),
        odoResidualTotal: companies.reduce((s, c) => s + (c.odoResidual || 0), 0),
      };

      return {
        summary,
        companies,
        risks: companies.filter((c) => c.risk === "КРИТИЧНО"),
        manual: companies.filter((c) => c.risk === "РУЧНАЯ ПРОВЕРКА"),
        contracts: contracts.list,
        membersMeta: members.cols,
        contractsMeta: contracts.cols,
      };
    }

    function fmtMoney(n) {
      if (n === null || n === undefined || Number.isNaN(n)) return "—";
      if (n === INF) return "без верхнего предела";
      return new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 0 }).format(n) + " ₽";
    }

    function riskPill(risk) {
      const cls = risk === "КРИТИЧНО" ? "crit" : risk === "РУЧНАЯ ПРОВЕРКА" ? "warn" : "ok";
      return `<span class="stamp ${cls}">${risk}</span>`;
    }

    function renderStats(summary) {
      const items = [
        ["Договоров", summary.contracts],
        ["ИНН", summary.inns],
        ["Не найдено", summary.notFound, "crit"],
        ["Право стоп", summary.suspended, "crit"],
        ["Критичные", summary.critical, "crit"],
        ["Ручная", summary.manual, "warn"],
        ["Норма", summary.ok, "ok"],
        ["Остаток ОДО", fmtMoney(summary.odoResidualTotal)],
      ];
      document.getElementById("stats").innerHTML = items
        .map(([l, v, cls]) => `<div class="stat ${cls || ""}"><div class="v">${v}</div><div class="l">${l}</div></div>`)
        .join("");
    }

    function renderTable() {
      const r = state.result;
      if (!r) return;
      let rows = [];
      let head = [];
      if (state.tab === "contracts") {
        head = ["ИНН", "Компания", "№", "К учёту", "Исполнено", "Остаток", "Закупка", "Конкурентный"];
        rows = r.contracts.map((c) => [
          c.inn,
          c.name,
          c.number,
          fmtMoney(c.amount),
          fmtMoney(c.done),
          fmtMoney(c.residual),
          c.method || "—",
          c.unclear ? "неясно" : c.competitive ? "да" : "нет",
        ]);
      } else {
        const src =
          state.tab === "risks" ? r.risks : state.tab === "manual" ? r.manual : r.companies;
        head = [
          "Риск",
          "ИНН",
          "Компания",
          "Право",
          "ВВ",
          "Max договор",
          "ОДО",
          "Остаток ОДО",
          "Комментарий",
        ];
        rows = src.map((c) => [
          riskPill(c.risk),
          c.inn,
          c.name,
          c.right,
          c.vvText || "—",
          fmtMoney(c.maxContract),
          c.odoText || "—",
          fmtMoney(c.odoResidual),
          c.comment || "—",
        ]);
      }
      document.getElementById("thead").innerHTML =
        "<tr>" + head.map((h) => `<th>${h}</th>`).join("") + "</tr>";
      document.getElementById("tbody").innerHTML = rows
        .map((row) => "<tr>" + row.map((cell) => `<td>${cell}</td>`).join("") + "</tr>")
        .join("");
    }

    function run() {
      clearError();
      try {
        state.result = analyze(state.membersRows, state.contractsRows);
        results.classList.remove("hidden");
        exportBtn.classList.remove("hidden");
        renderStats(state.result.summary);
        renderTable();
      } catch (e) {
        showError(e.message || String(e));
      }
    }

    function toCsv(rows) {
      return rows
        .map((r) =>
          r
            .map((v) => {
              const s = String(v ?? "");
              return /["\n,;]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
            })
            .join(";")
        )
        .join("\n");
    }

    function exportCsv() {
      const r = state.result;
      if (!r) return;
      const rows = [
        [
          "ИНН",
          "Компания",
          "Найдена",
          "Право",
          "Уровень ВВ",
          "Лимит ВВ",
          "Max договор",
          "Проверка ВВ",
          "Уровень ОДО",
          "Лимит ОДО",
          "Остаток ОДО",
          "Обязательства реестра",
          "Проверка ОДО",
          "Договоров",
          "Конкурентных",
          "Риск",
          "Комментарий",
        ],
        ...r.companies.map((c) => [
          c.inn,
          c.name,
          c.found ? "да" : "нет",
          c.right,
          c.vvText,
          c.vvLimit === INF ? "inf" : c.vvLimit,
          c.maxContract,
          c.vvCheck,
          c.odoText,
          c.odoLimit === INF ? "inf" : c.odoLimit,
          c.odoResidual,
          c.registryOblig,
          c.odoCheck,
          c.contractsCount,
          c.competitiveCount,
          c.risk,
          c.comment,
        ]),
      ];
      const blob = new Blob(["\uFEFF" + toCsv(rows)], { type: "text/csv;charset=utf-8" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "sro_audit_report.csv";
      a.click();
      URL.revokeObjectURL(a.href);
    }

    const SAMPLE_MEMBERS = [
      {
        Контрагент: "ООО АльфаСтрой",
        ИНН: "7701234567",
        "Состояние права": "Действует",
        "Уровень ВВ": "до 90 млн руб.",
        "Уровень ОДО": "до 90 млн руб.",
        "Расчёт обязательств": 40000000,
      },
      {
        Контрагент: "ООО БетаИнвест",
        ИНН: "7702345678",
        "Состояние права": "Действует",
        "Уровень ВВ": "до 500 млн руб.",
        "Уровень ОДО": "до 500 млн руб.",
        "Расчёт обязательств": 420000000,
      },
      {
        Контрагент: "ООО ГаммаСтрой",
        ИНН: "7703456789",
        "Состояние права": "Приостановлено",
        "Уровень ВВ": "до 500 млн руб.",
        "Уровень ОДО": "до 90 млн руб.",
        "Расчёт обязательств": 10000000,
      },
      {
        Контрагент: "ООО Дельта",
        ИНН: "7704567890",
        "Состояние права": "Действует",
        "Уровень ВВ": "до 90 млн руб.",
        "Уровень ОДО": "",
        "Расчёт обязательств": "",
      },
    ];

    const SAMPLE_CONTRACTS = [
      {
        Контрагент: "ООО АльфаСтрой",
        ИНН: "7701234567",
        "Номер договора": "A-1",
        "Стоимость, принятая СРО к учёту": 70000000,
        "Стоимость принятых работ": 30000000,
        "Вид закупки": "44-ФЗ",
      },
      {
        Контрагент: "ООО БетаИнвест",
        ИНН: "7702345678",
        "Номер договора": "B-1",
        "Стоимость, принятая СРО к учёту": 300000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "223-ФЗ",
      },
      {
        Контрагент: "ООО БетаИнвест",
        ИНН: "7702345678",
        "Номер договора": "B-2",
        "Стоимость, принятая СРО к учёту": 250000000,
        "Стоимость принятых работ": 50000000,
        "Вид закупки": "аукцион",
      },
      {
        Контрагент: "ООО ГаммаСтрой",
        ИНН: "7703456789",
        "Номер договора": "G-1",
        "Стоимость, принятая СРО к учёту": 120000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "прямой",
      },
      {
        Контрагент: "ООО Дельта",
        ИНН: "7704567890",
        "Номер договора": "D-1",
        "Стоимость, принятая СРО к учёту": 50000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "44-ФЗ",
      },
      {
        Контрагент: "ИП Неизвестный",
        ИНН: "500111222333",
        "Номер договора": "X-1",
        "Стоимость, принятая СРО к учёту": 15000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "",
      },
    ];

    membersInput.addEventListener("change", async (e) => {
      clearError();
      try {
        state.membersRows = await readTable(e.target.files[0]);
        updateReady();
      } catch (err) {
        showError("Не удалось прочитать реестр членов: " + err.message);
      }
    });

    contractsInput.addEventListener("change", async (e) => {
      clearError();
      try {
        state.contractsRows = await readTable(e.target.files[0]);
        updateReady();
      } catch (err) {
        showError("Не удалось прочитать договоры: " + err.message);
      }
    });

    runBtn.addEventListener("click", run);
    exportBtn.addEventListener("click", exportCsv);
    sampleBtn.addEventListener("click", () => {
      state.membersRows = SAMPLE_MEMBERS;
      state.contractsRows = SAMPLE_CONTRACTS;
      membersInput.value = "";
      contractsInput.value = "";
      updateReady();
      run();
    });

    document.querySelectorAll(".tabs button").forEach((btn) => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".tabs button").forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");
        state.tab = btn.dataset.tab;
        renderTable();
      });
    });
  