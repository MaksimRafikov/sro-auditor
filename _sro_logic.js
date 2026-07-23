
    const INF = Number.POSITIVE_INFINITY;
    const state = {
      membersRows: null,
      contractsRows: null,
      result: null,
      tab: "companies",
      filter: null,
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
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
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

    function startOfDay(d) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }

    function parseDate(v) {
      if (v === null || v === undefined || v === "") return null;
      if (v instanceof Date && !Number.isNaN(v.getTime())) return startOfDay(v);
      if (typeof v === "number" && Number.isFinite(v)) {
        // Excel serial (1900 date system)
        const epoch = Date.UTC(1899, 11, 30);
        const d = new Date(epoch + Math.round(v) * 86400000);
        return Number.isNaN(d.getTime()) ? null : startOfDay(d);
      }
      const s = String(v).trim();
      if (!s) return null;
      const m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{2,4})$/);
      if (m) {
        let y = Number(m[3]);
        if (y < 100) y += 2000;
        const d = new Date(y, Number(m[2]) - 1, Number(m[1]));
        return Number.isNaN(d.getTime()) ? null : startOfDay(d);
      }
      const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (iso) {
        const d = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
        return Number.isNaN(d.getTime()) ? null : startOfDay(d);
      }
      const parsed = new Date(s);
      return Number.isNaN(parsed.getTime()) ? null : startOfDay(parsed);
    }

    function fmtDate(d) {
      if (!d) return "—";
      const dd = String(d.getDate()).padStart(2, "0");
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      return `${dd}.${mm}.${d.getFullYear()}`;
    }

    /** Классификация способа закупки: 44 / 223 / 615 / direct / other_comp / unclear */
    function classifyMethod(method, odoFlag) {
      const flag = String(odoFlag ?? "").toLowerCase().trim();
      const flagYes = ["1", "true", "да", "yes", "y", "+"].includes(flag);
      const m = String(method ?? "").toLowerCase().replace(/ё/g, "е");

      if (/615/.test(m)) return { type: "615", competitive: true, unclear: false };
      if (/44/.test(m)) return { type: "44", competitive: true, unclear: false };
      if (/223/.test(m)) return { type: "223", competitive: true, unclear: false };
      if (/конкурс|аукцион|тендер|закупк|торг/.test(m)) {
        return { type: "other_comp", competitive: true, unclear: false };
      }
      if (flagYes) return { type: "other_comp", competitive: true, unclear: false };
      if (/прям|коммерч|без.*конкур/.test(m)) {
        return { type: "direct", competitive: false, unclear: false };
      }
      if (!m.trim()) return { type: "unclear", competitive: false, unclear: true };
      return { type: "unclear", competitive: false, unclear: true };
    }

    function dateInSuspension(contractDate, suspFrom, suspTo) {
      if (!contractDate) return false;
      if (!suspFrom && !suspTo) return false;
      if (suspFrom && contractDate < suspFrom) return false;
      if (suspTo && contractDate > suspTo) return false;
      return true;
    }

    function collectHeaders(rows) {
      const keys = new Set();
      for (const row of rows) {
        Object.keys(row).forEach((k) => keys.add(k));
      }
      return [...keys];
    }

    function mapMembers(rows) {
      if (!rows.length) throw new Error("Реестр членов СРО пуст.");
      const headers = collectHeaders(rows);
      const cols = {
        inn: findCol(headers, ["инн", "inn"]),
        name: findCol(headers, ["контрагент", "наименование", "организация", "член сро", "название"]),
        right: findCol(headers, ["состояние права", "статус права", "право"]),
        vv: findCol(headers, ["уровень вв", "лимит вв", "вв"]),
        odo: findCol(headers, ["уровень одо", "лимит одо", "одо"]),
        oblig: findCol(headers, ["расчет обязательств", "расчёт обязательств", "обязательства"]),
        suspFrom: findCol(headers, [
          "дата приостановления",
          "дата начала приостановки",
          "приостановлено с",
          "дата приостановки",
          "начало приостановки",
        ]),
        suspTo: findCol(headers, [
          "дата возобновления",
          "дата окончания приостановки",
          "приостановлено по",
          "окончание приостановки",
          "возобновлено",
        ]),
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
          suspFrom: cols.suspFrom ? parseDate(row[cols.suspFrom]) : null,
          suspTo: cols.suspTo ? parseDate(row[cols.suspTo]) : null,
          _cols: cols,
        });
      }
      return { byInn, cols, count: byInn.size };
    }

    function mapContracts(rows) {
      if (!rows.length) throw new Error("Реестр договоров пуст.");
      const headers = collectHeaders(rows);
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
        const { type, competitive, unclear } = classifyMethod(method, odoFlag);
        const executed = done === null ? 0 : done;
        const residual = weirdMoney ? null : Math.max(0, amount - executed);
        const dateRaw = cols.date ? row[cols.date] : "";
        const dateObj = cols.date ? parseDate(dateRaw) : null;
        list.push({
          inn,
          name: cols.name ? String(row[cols.name] || "").trim() : "",
          number: cols.number ? String(row[cols.number] ?? "") : "",
          date: dateObj ? fmtDate(dateObj) : dateRaw ? String(dateRaw) : "",
          dateObj,
          amount,
          done: executed,
          residual,
          method: String(method ?? ""),
          methodType: type,
          competitive,
          unclear,
          weirdMoney,
          assumptionNoDone: !cols.done,
          inSuspensionPeriod: false,
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
        const flags = {
          odoExceed: false,
          noOdo: false,
          vvExceed: false,
          suspended: false,
          odoMismatch: false,
          notFound: false,
        };

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
        const suspFrom = m ? m.suspFrom : null;
        const suspTo = m ? m.suspTo : null;
        const name = (m && m.name) || list.find((c) => c.name)?.name || "";
        const rightStopped = right === "приостановлено" || right === "прекращено";

        let vvCheck = "н/д";
        let odoCheck = "н/д";

        // Договоры в период приостановки
        let contractsInSuspension = [];
        if (m && (suspFrom || suspTo)) {
          for (const c of list) {
            if (dateInSuspension(c.dateObj, suspFrom, suspTo)) {
              c.inSuspensionPeriod = true;
              contractsInSuspension.push(c);
            }
          }
        }

        if (!found) {
          risk = "КРИТИЧНО";
          flags.notFound = true;
          comments.push("не найдена в реестре СРО");
        }

        // Приостановление: договоры в периоде — приоритет; иначе статус права
        if (contractsInSuspension.length > 0) {
          risk = "КРИТИЧНО";
          flags.suspended = true;
          const period =
            (suspFrom ? fmtDate(suspFrom) : "…") + " — " + (suspTo ? fmtDate(suspTo) : "…");
          comments.push(
            `договор(ы) в период приостановки (${period}): ${contractsInSuspension.length}`
          );
        } else if (found && rightStopped) {
          risk = "КРИТИЧНО";
          flags.suspended = true;
          comments.push(`право: ${right}`);
        }

        if (found && vvLimit !== null && Number.isFinite(vvLimit) && maxContract > vvLimit) {
          risk = "КРИТИЧНО";
          flags.vvExceed = true;
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
            flags.noOdo = true;
            odoCheck = "нет ОДО";
            if (rightStopped) {
              comments.push("нет ОДО при договорах ОДО (акцент: право приостановлено)");
            } else {
              comments.push("есть конкурентные договоры, ОДО отсутствует");
            }
          } else if (Number.isFinite(odoLimit) && odoResidual > odoLimit) {
            risk = "КРИТИЧНО";
            flags.odoExceed = true;
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
        // Сравниваем остаток ОДО с реестром только если есть конкурентные договоры
        if (found && competitive.length > 0 && diverges(odoResidual, registryOblig)) {
          if (risk !== "КРИТИЧНО") risk = "РУЧНАЯ ПРОВЕРКА";
          flags.odoMismatch = true;
          comments.push("расхождение остатка ОДО с реестром (пересчёт ЕРЧ)");
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
          contractsInSuspensionCount: contractsInSuspension.length,
          suspFrom,
          suspTo,
          risk,
          flags,
          comment: comments.join("; "),
        });
      }

      companies.sort((a, b) => {
        const order = { КРИТИЧНО: 0, "РУЧНАЯ ПРОВЕРКА": 1, НОРМА: 2 };
        return order[a.risk] - order[b.risk] || b.maxContract - a.maxContract;
      });

      const byType = (t) => contracts.list.filter((c) => c.methodType === t).length;
      const summary = {
        contracts: contracts.list.length,
        byFz44: byType("44"),
        byFz223: byType("223"),
        byFz615: byType("615"),
        byDirect: byType("direct"),
        byOtherComp: byType("other_comp"),
        byUnclear: byType("unclear"),
        inns: companies.length,
        found: companies.filter((c) => c.found).length,
        notFound: companies.filter((c) => c.flags.notFound).length,
        odoExceed: companies.filter((c) => c.flags.odoExceed).length,
        noOdo: companies.filter((c) => c.flags.noOdo).length,
        vvExceed: companies.filter((c) => c.flags.vvExceed).length,
        suspended: companies.filter((c) => c.flags.suspended).length,
        odoMismatch: companies.filter((c) => c.flags.odoMismatch).length,
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

    function companiesByFilter(r, filter) {
      if (!filter) return r.companies;
      if (filter === "risks") return r.risks;
      if (filter === "manual") return r.manual;
      if (filter === "odoExceed") return r.companies.filter((c) => c.flags.odoExceed);
      if (filter === "noOdo") return r.companies.filter((c) => c.flags.noOdo);
      if (filter === "vvExceed") return r.companies.filter((c) => c.flags.vvExceed);
      if (filter === "suspended") return r.companies.filter((c) => c.flags.suspended);
      if (filter === "odoMismatch") return r.companies.filter((c) => c.flags.odoMismatch);
      return r.companies;
    }

    function renderStats(summary) {
      const s = summary;
      document.getElementById("stats").innerHTML = `
        <div class="summary-grid">
          <div class="summary-block">
            <h3>Договоры</h3>
            <div class="summary-row main"><span>Количество договоров</span><strong>${s.contracts}</strong></div>
            <div class="summary-row sub"><span>из них по 44-ФЗ</span><strong>${s.byFz44}</strong></div>
            <div class="summary-row sub"><span>по 223-ФЗ</span><strong>${s.byFz223}</strong></div>
            <div class="summary-row sub"><span>по 615-ФЗ</span><strong>${s.byFz615}</strong></div>
            <div class="summary-row sub"><span>прямые</span><strong>${s.byDirect}</strong></div>
            ${
              s.byOtherComp
                ? `<div class="summary-row sub muted"><span>прочие конкурентные</span><strong>${s.byOtherComp}</strong></div>`
                : ""
            }
            ${
              s.byUnclear
                ? `<div class="summary-row sub muted"><span>без вида закупки</span><strong>${s.byUnclear}</strong></div>`
                : ""
            }
          </div>
          <div class="summary-block">
            <h3>Члены СРО</h3>
            <button type="button" class="summary-row clickable crit" data-filter="odoExceed">
              <span>1. Превышен уровень ОДО</span><strong>${s.odoExceed}</strong>
            </button>
            <button type="button" class="summary-row clickable crit" data-filter="noOdo">
              <span>2. Нет ОДО (есть договоры ОДО)</span><strong>${s.noOdo}</strong>
            </button>
            <button type="button" class="summary-row clickable crit" data-filter="vvExceed">
              <span>3. Превышен уровень ВВ</span><strong>${s.vvExceed}</strong>
            </button>
            <button type="button" class="summary-row clickable crit" data-filter="suspended">
              <span>4. Приостановлен</span><strong>${s.suspended}</strong>
            </button>
            <button type="button" class="summary-row clickable warn" data-filter="odoMismatch">
              <span>5. Ручная проверка (остаток ОДО ≠ реестр)</span><strong>${s.odoMismatch}</strong>
            </button>
            <div class="summary-row main odo-total"><span>Остаток ОДО</span><strong>${fmtMoney(s.odoResidualTotal)}</strong></div>
          </div>
        </div>`;

      document.querySelectorAll("#stats [data-filter]").forEach((btn) => {
        btn.addEventListener("click", () => {
          const f = btn.dataset.filter;
          state.filter = state.filter === f ? null : f;
          state.tab = "companies";
          document.querySelectorAll(".tabs button").forEach((b) => {
            b.classList.toggle("active", b.dataset.tab === "companies");
            b.setAttribute("aria-selected", b.dataset.tab === "companies" ? "true" : "false");
          });
          document.querySelectorAll("#stats [data-filter]").forEach((b) => {
            b.classList.toggle("active", state.filter === b.dataset.filter);
          });
          renderTable();
        });
      });
    }

    function renderTable() {
      const r = state.result;
      if (!r) return;
      let rows = [];
      let head = [];
      if (state.tab === "contracts") {
        head = [
          "ИНН",
          "Компания",
          "№",
          "Дата",
          "К учёту",
          "Исполнено",
          "Остаток",
          "Закупка",
          "Тип",
          "Приостановка",
        ];
        rows = r.contracts.map((c) => [
          c.inn,
          c.name,
          c.number,
          c.date || "—",
          fmtMoney(c.amount),
          fmtMoney(c.done),
          fmtMoney(c.residual),
          c.method || "—",
          c.methodType === "44"
            ? "44-ФЗ"
            : c.methodType === "223"
              ? "223-ФЗ"
              : c.methodType === "615"
                ? "615-ФЗ"
                : c.methodType === "direct"
                  ? "прямой"
                  : c.methodType === "other_comp"
                    ? "конкур."
                    : "неясно",
          c.inSuspensionPeriod ? "в периоде" : "—",
        ]);
      } else {
        const src =
          state.tab === "risks"
            ? r.risks
            : state.tab === "manual"
              ? r.manual
              : companiesByFilter(r, state.filter);
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
        state.filter = null;
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
          "В периоде приостановки",
          "Превышен ОДО",
          "Нет ОДО",
          "Превышен ВВ",
          "Приостановлен",
          "Расхождение ОДО",
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
          c.contractsInSuspensionCount,
          c.flags.odoExceed ? "да" : "нет",
          c.flags.noOdo ? "да" : "нет",
          c.flags.vvExceed ? "да" : "нет",
          c.flags.suspended ? "да" : "нет",
          c.flags.odoMismatch ? "да" : "нет",
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
        "Дата приостановления": "01.03.2025",
        "Дата возобновления": "",
      },
      {
        Контрагент: "СТС Кузнецов",
        ИНН: "7704567890",
        "Состояние права": "Приостановлено",
        "Уровень ВВ": "до 90 млн руб.",
        "Уровень ОДО": "",
        "Расчёт обязательств": "",
        "Дата приостановления": "15.01.2025",
        "Дата возобновления": "",
      },
      {
        Контрагент: "ООО Еpsilon",
        ИНН: "7705678901",
        "Состояние права": "Действует",
        "Уровень ВВ": "до 90 млн руб.",
        "Уровень ОДО": "до 90 млн руб.",
        "Расчёт обязательств": 10000000,
      },
    ];

    const SAMPLE_CONTRACTS = [
      {
        Контрагент: "ООО АльфаСтрой",
        ИНН: "7701234567",
        "Номер договора": "A-1",
        "Дата заключения": "10.02.2025",
        "Стоимость, принятая СРО к учёту": 70000000,
        "Стоимость принятых работ": 30000000,
        "Вид закупки": "44-ФЗ",
      },
      {
        Контрагент: "ООО БетаИнвест",
        ИНН: "7702345678",
        "Номер договора": "B-1",
        "Дата заключения": "05.01.2025",
        "Стоимость, принятая СРО к учёту": 300000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "223-ФЗ",
      },
      {
        Контрагент: "ООО БетаИнвест",
        ИНН: "7702345678",
        "Номер договора": "B-2",
        "Дата заключения": "20.02.2025",
        "Стоимость, принятая СРО к учёту": 250000000,
        "Стоимость принятых работ": 50000000,
        "Вид закупки": "аукцион",
      },
      {
        Контрагент: "ООО ГаммаСтрой",
        ИНН: "7703456789",
        "Номер договора": "G-1",
        "Дата заключения": "20.04.2025",
        "Стоимость, принятая СРО к учёту": 120000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "прямой",
      },
      {
        Контрагент: "СТС Кузнецов",
        ИНН: "7704567890",
        "Номер договора": "K-1",
        "Дата заключения": "01.02.2025",
        "Стоимость, принятая СРО к учёту": 50000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "44-ФЗ",
      },
      {
        Контрагент: "ООО Еpsilon",
        ИНН: "7705678901",
        "Номер договора": "E-1",
        "Дата заключения": "12.03.2025",
        "Стоимость, принятая СРО к учёту": 40000000,
        "Стоимость принятых работ": 5000000,
        "Вид закупки": "615-ФЗ",
      },
      {
        Контрагент: "ИП Неизвестный",
        ИНН: "500111222333",
        "Номер договора": "X-1",
        "Дата заключения": "01.01.2025",
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
        if (state.tab !== "companies") state.filter = null;
        document.querySelectorAll("#stats [data-filter]").forEach((b) => b.classList.remove("active"));
        btn.setAttribute("aria-selected", "true");
        document.querySelectorAll(".tabs button").forEach((b) => {
          if (b !== btn) b.setAttribute("aria-selected", "false");
        });
        renderTable();
      });
    });
