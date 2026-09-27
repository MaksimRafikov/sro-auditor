
    const INF = Number.POSITIVE_INFINITY;
    const state = {
      membersRows: null,
      contractsRows: null,
      membersMeta: null,
      membersSource: null,
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
      if (ok) {
        fileHint.textContent = `Готово: ${state.membersRows.length} членов реестра, ${state.contractsRows.length} договоров.`;
      } else if (state.membersRows) {
        fileHint.textContent = `Реестр членов готов (${state.membersRows.length}). Загрузите договоры.`;
      } else if (state.contractsRows) {
        fileHint.textContent = "Договоры загружены. Обновите реестр членов в шаге 1.";
      } else {
        fileHint.textContent = "Обновите реестр членов и загрузите договоры.";
      }
    }

    async function readTable(file) {
      if (typeof XLSX === "undefined") {
        throw new Error(
          "Библиотека Excel не загрузилась. Откройте страницу через локальный сервер из папки проекта — нужен файл vendor/xlsx-js-style.min.js."
        );
      }
      const buf = await file.arrayBuffer();
      // В CSV нет типов, и SheetJS угадывает их по-американски: ИНН 0278105091
      // становится числом без ведущего нуля, 04.12.2020 — 4 декабря,
      // а «1234,56» — 123456. Читаем CSV текстом и разбираем своими парсерами.
      const isCsv = /\.csv$/i.test(file.name) || file.type === "text/csv";
      const wb = isCsv
        ? XLSX.read(buf, { type: "array", raw: true })
        : XLSX.read(buf, { type: "array", cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
    }

    function normKey(s) {
      return String(s || "")
        .toLowerCase()
        .replace(/ё/g, "е")
        .replace(/[^a-zа-я0-9]+/gi, "");
    }

    /**
     * Подбор колонки по алиасам.
     * 1) точное совпадение нормализованного имени (по порядку алиасов);
     * 2) includes — максимальная доля aliasLen/headerLen (и более ранний алиас при равенстве),
     *    чтобы «контрагент» не цеплял «…обязательствах контрагента».
     * Короткие алиасы (< 4 символов после нормализации) — только exact.
     */
    function findCol(headers, aliases) {
      const normalized = headers.map((h) => ({ raw: h, n: normKey(h) }));
      for (const alias of aliases) {
        const a = normKey(alias);
        if (!a) continue;
        const exact = normalized.find((h) => h.n === a);
        if (exact) return exact.raw;
      }
      let best = null;
      let bestScore = -1;
      let bestAliasOrder = Infinity;
      aliases.forEach((alias, aliasIdx) => {
        const a = normKey(alias);
        if (!a || a.length < 4) return;
        for (const h of normalized) {
          if (!h.n.includes(a)) continue;
          const score = a.length / h.n.length;
          if (score > bestScore || (Math.abs(score - bestScore) < 1e-9 && aliasIdx < bestAliasOrder)) {
            bestScore = score;
            best = h.raw;
            bestAliasOrder = aliasIdx;
          }
        }
      });
      return bestScore >= 0.35 ? best : null;
    }

    function isYesFlag(v) {
      const s = String(v ?? "")
        .toLowerCase()
        .replace(/ё/g, "е")
        .trim();
      return ["1", "true", "да", "yes", "y", "+"].includes(s);
    }

    function normalizeInn(v) {
      const digits = String(v ?? "").replace(/\D/g, "");
      if (!digits) return "";
      // Excel/CSV часто отдают ИНН числом и съедают ведущий ноль (регионы 01–09).
      // ИНН ЮЛ — 10 цифр, ИП — 12; код региона 00 не существует, ноль всегда один.
      if (digits.length === 9 || digits.length === 11) return "0" + digits;
      return digits;
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

    function dateInPeriods(contractDate, periods) {
      if (!contractDate || !periods || !periods.length) return false;
      return periods.some((p) => dateInSuspension(contractDate, p.from, p.to));
    }

    /** «30.08.2017 — 13.09.2017; 26.04.2018 — открыто» → несколько периодов. */
    function parseSuspensionText(text) {
      const raw = String(text ?? "").trim();
      if (!raw) return [];
      const periods = [];
      for (const chunk of raw.split(/\s*;\s*/)) {
        const found = [];
        const re = /(\d{1,2}[./]\d{1,2}[./]\d{2,4})/g;
        let match;
        while ((match = re.exec(chunk))) {
          const d = parseDate(match[1]);
          if (d) found.push(d);
        }
        if (!found.length) continue;
        const open = /открыт/i.test(chunk);
        periods.push({ from: found[0], to: open ? null : found[1] || null });
      }
      return periods;
    }

    function sameDay(a, b) {
      if (!a && !b) return true;
      if (!a || !b) return false;
      return a.getTime() === b.getTime();
    }

    function mergePeriods(list) {
      const out = [];
      for (const p of list) {
        if (!p || !p.from) continue;
        const dup = out.some((q) => sameDay(q.from, p.from) && sameDay(q.to, p.to));
        if (!dup) out.push({ from: p.from, to: p.to || null });
      }
      out.sort((a, b) => a.from - b.from);
      return out;
    }

    function formatPeriods(periods) {
      return periods
        .map((p) => `${fmtDate(p.from)} — ${p.to ? fmtDate(p.to) : "открыто"}`)
        .join("; ");
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
        name: findCol(headers, [
          "контрагент",
          "наименование",
          "организация",
          "член сро",
          "сокращенное наименование",
          "название",
        ]),
        right: findCol(headers, ["состояние права", "статус права", "право"]),
        vv: findCol(headers, ["уровень вв", "лимит вв", "уровень ответственности вв", "вв"]),
        odo: findCol(headers, ["уровень одо", "лимит одо", "уровень ответственности одо"]),
        oblig: findCol(headers, [
          "расчет обязательств",
          "расчёт обязательств",
          "расчет размера обязательств",
          "обязательства",
        ]),
        obligDate: findCol(headers, [
          "дата расчета размера обязательств",
          "дата расчёта размера обязательств",
          "дата расчета обязательств",
          "дата расчёта обязательств",
        ]),
        regDate: findCol(headers, [
          "дата регистрации в реестре сро",
          "дата регистрации в реестре",
          "дата вступления в сро",
          "дата вступления",
          "дата приема в члены",
        ]),
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
        suspPeriods: findCol(headers, [
          "периоды приостановок",
          "периоды приостановки",
          "история приостановок",
        ]),
      };
      if (!cols.inn) throw new Error("В реестре членов не найдена колонка ИНН.");
      if (cols.oblig && normKey(cols.oblig).includes("дата")) cols.oblig = null;
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
          obligDate: cols.obligDate ? parseDate(row[cols.obligDate]) : null,
          regDate: cols.regDate ? parseDate(row[cols.regDate]) : null,
          suspFrom: cols.suspFrom ? parseDate(row[cols.suspFrom]) : null,
          suspTo: cols.suspTo ? parseDate(row[cols.suspTo]) : null,
          periods: mergePeriods([
            ...(cols.suspFrom && parseDate(row[cols.suspFrom])
              ? [{ from: parseDate(row[cols.suspFrom]), to: cols.suspTo ? parseDate(row[cols.suspTo]) : null }]
              : []),
            ...(cols.suspPeriods ? parseSuspensionText(row[cols.suspPeriods]) : []),
          ]),
          _cols: cols,
        });
      }
      return { byInn, cols, count: byInn.size };
    }

    function mapContracts(rows) {
      if (!rows.length) throw new Error("Реестр договоров пуст.");
      const headers = collectHeaders(rows);
      const cols = {
        inn: findCol(headers, [
          "инн члена сро",
          "инн члена",
          "инн подрядчика",
          "инн участника",
          "инн поставщика",
          "инн",
          "inn",
        ]),
        name: findCol(headers, [
          "член сро поставщик",
          "член сро",
          "подрядчик",
          "контрагент",
          "участник",
          "наименование",
          "компания",
        ]),
        amount: findCol(headers, [
          "стоимость принятая сро к учету",
          "стоимость принятая сро к учёту",
          "стоимость к учету",
          "стоимость к учёту",
          "стоимость работ по договору",
          "сумма договора",
          "стоимость работ",
          "стоимость",
        ]),
        done: findCol(headers, [
          "сумма принятых работ по договору",
          "стоимость принятых работ по договору",
          "сумма принятых работ",
          "стоимость принятых работ",
          "принятых работ",
          "стоимость исполненных работ",
          "стоимость исполненных",
          "исполнено",
          "принято",
        ]),
        method: findCol(headers, ["вид закупки", "способ закупки", "тип закупки", "закон"]),
        odoFlag: findCol(headers, ["договор одо", "признак одо"]),
        number: findCol(headers, ["номер договора", "№ договора"]),
        date: findCol(headers, [
          "дата заключения договора",
          "дата заключения",
          "дата договора",
          "дата",
        ]),
        exclude: findCol(headers, [
          "не учитывать договор в обязательствах контрагента",
          "не учитывать договор в обязательствах",
          "не учитывать договор",
          "не использовать",
        ]),
      };
      if (!cols.inn) throw new Error("В реестре договоров не найдена колонка ИНН.");
      if (!cols.amount) throw new Error("В реестре договоров не найдена колонка стоимости.");

      const list = [];
      for (const row of rows) {
        const inn = normalizeInn(row[cols.inn]);
        if (!inn) continue;
        const excluded = cols.exclude ? isYesFlag(row[cols.exclude]) : false;
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
          excluded,
          inSuspensionPeriod: false,
          beforeMembership: false,
          membershipRegDate: null,
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
          vvExceedSusp: false,
          odoExceedSusp: false,
          odoMismatch: false,
          dataIncomplete: false,
          notFound: false,
        };

        const counted = list.filter((c) => !c.excluded);
        const excludedCount = list.length - counted.length;
        const found = !!m;
        const regDate = m ? m.regDate : null;
        for (const c of list) {
          const before = !!(regDate && c.dateObj && c.dateObj < regDate);
          c.beforeMembership = before;
          c.membershipRegDate = before ? regDate : null;
        }
        const beforeMembershipList = list.filter((c) => c.beforeMembership);
        // В лимиты входят договоры с даты регистрации. Более ранние остаются в отчёте как основание:
        // на дату заключения членства в реестре не было.
        const inScope = counted.filter((c) => !c.beforeMembership);

        const maxContract = inScope.reduce((mx, c) => {
          if (c.amount === null) return mx;
          return Math.max(mx, c.amount);
        }, 0);

        const competitive = inScope.filter((c) => c.competitive);
        const unclear = inScope.filter((c) => c.unclear);
        const odoResidual = competitive.reduce((s, c) => s + (c.residual || 0), 0);

        const right = m ? m.right : "не найдено";
        const vvLimit = m ? m.vvLimit : null;
        const odoLimit = m ? m.odoLimit : null;
        const vvText = m ? m.vvText : "";
        const odoText = m ? m.odoText : "";
        const registryOblig = m ? m.registryOblig : null;
        const obligDate = m ? m.obligDate : null;
        const periods = m ? m.periods : [];
        const name = (m && m.name) || list.find((c) => c.name)?.name || "";
        const rightStopped = right === "приостановлено" || right === "прекращено";
        const lawContracts = inScope.filter(
          (c) => c.methodType === "44" || c.methodType === "223" || c.methodType === "615"
        );

        let vvCheck = "н/д";
        let odoCheck = "н/д";
        let vvCheckSusp = "н/д";
        let odoCheckSusp = "н/д";

        let contractsInSuspension = [];
        if (periods.length) {
          for (const c of inScope) {
            if (dateInPeriods(c.dateObj, periods)) {
              c.inSuspensionPeriod = true;
              contractsInSuspension.push(c);
            }
          }
        }
        const suspCompetitive = contractsInSuspension.filter((c) => c.competitive);
        const maxInSuspension = contractsInSuspension.reduce((mx, c) => {
          if (c.amount === null) return mx;
          return Math.max(mx, c.amount);
        }, 0);
        const odoResidualSusp = suspCompetitive.reduce((s, c) => s + (c.residual || 0), 0);

        if (!found) {
          risk = "КРИТИЧНО";
          flags.notFound = true;
          comments.push("не найдена в реестре СРО");
        }

        // 4. Договоры, заключённые внутри любого периода приостановки.
        // Статус права без таких договоров остаётся критичным, но в счётчик 4 не входит.
        if (contractsInSuspension.length > 0) {
          risk = "КРИТИЧНО";
          flags.suspended = true;
          comments.push(
            `договоры в период приостановки (${formatPeriods(periods)}): ${contractsInSuspension.length}`
          );
          if (found && vvLimit !== null && Number.isFinite(vvLimit)) {
            if (maxInSuspension > vvLimit) {
              flags.vvExceedSusp = true;
              vvCheckSusp = "превышение";
              comments.push("превышение ВВ в период приостановки");
            } else {
              vvCheckSusp = "ок";
            }
          } else if (found && vvLimit === INF) {
            vvCheckSusp = "ок";
          }
          if (suspCompetitive.length > 0 && found && odoLimit !== null && Number.isFinite(odoLimit)) {
            if (odoResidualSusp > odoLimit) {
              flags.odoExceedSusp = true;
              odoCheckSusp = "превышение";
              comments.push("превышение ОДО в период приостановки");
            } else {
              odoCheckSusp = "ок";
            }
          } else if (suspCompetitive.length > 0 && found && odoLimit === INF) {
            odoCheckSusp = "ок";
          }
        } else if (found && rightStopped) {
          risk = "КРИТИЧНО";
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
            if (lawContracts.length > 0) {
              risk = "КРИТИЧНО";
              flags.noOdo = true;
              odoCheck = "нет ОДО";
              comments.push("нет уровня ОДО при договорах 44/223/615");
            } else {
              if (risk !== "КРИТИЧНО") risk = "РУЧНАЯ ПРОВЕРКА";
              flags.dataIncomplete = true;
              comments.push("конкурентные договоры не 44/223/615, уровень ОДО не указан");
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
          flags.dataIncomplete = true;
          comments.push(`без способа закупки: ${unclear.length}`);
        }
        if (inScope.some((c) => c.weirdMoney) || beforeMembershipList.some((c) => c.weirdMoney)) {
          if (risk !== "КРИТИЧНО") risk = "РУЧНАЯ ПРОВЕРКА";
          flags.dataIncomplete = true;
          comments.push("странный формат суммы");
        }
        if (inScope.some((c) => c.assumptionNoDone) || beforeMembershipList.some((c) => c.assumptionNoDone)) {
          if (risk !== "КРИТИЧНО") risk = "РУЧНАЯ ПРОВЕРКА";
          flags.dataIncomplete = true;
          comments.push("исполнение не найдено — остаток = полная стоимость");
        }
        if (excludedCount > 0) {
          comments.push(`исключено из обязательств: ${excludedCount}`);
        }
        if (beforeMembershipList.length > 0) {
          const nums = beforeMembershipList.map((c) => c.number).filter(Boolean);
          const shown = nums.slice(0, 8).join(", ");
          const more = nums.length > 8 ? ` и ещё ${nums.length - 8}` : "";
          comments.push(
            `до регистрации в реестре СРО (${fmtDate(regDate)}): ${beforeMembershipList.length} дог.` +
              (shown ? ` № ${shown}${more}` : "") +
              " — на дату заключения членства нет, в ВВ/ОДО не входят"
          );
        }
        // Сверка ЕРЧ: не раньше даты регистрации и не позже даты расчёта.
        let odoResidualErch = null;
        if (found && competitive.length > 0 && registryOblig !== null) {
          const undated = competitive.filter((c) => !c.dateObj);
          if (obligDate && undated.length === 0) {
            const asOf = competitive.filter((c) => c.dateObj <= obligDate);
            const afterCount = competitive.length - asOf.length;
            odoResidualErch = asOf.reduce((s, c) => s + (c.residual || 0), 0);
            if (diverges(odoResidualErch, registryOblig)) {
              if (risk !== "КРИТИЧНО") risk = "РУЧНАЯ ПРОВЕРКА";
              flags.odoMismatch = true;
              comments.push(
                `расхождение остатка ОДО с реестром на ${fmtDate(obligDate)} (пересчёт ЕРЧ)`
              );
            } else if (afterCount > 0 && diverges(odoResidual, registryOblig)) {
              comments.push(
                `сверка ЕРЧ на ${fmtDate(obligDate)} сошлась; договоры после этой даты в расчёт реестра не входят`
              );
            }
          } else if (diverges(odoResidual, registryOblig)) {
            if (risk !== "КРИТИЧНО") risk = "РУЧНАЯ ПРОВЕРКА";
            flags.odoMismatch = true;
            comments.push(
              obligDate && undated.length
                ? "расхождение остатка ОДО с реестром (пересчёт ЕРЧ); есть договоры без даты, срез на дату расчёта не применён"
                : "расхождение остатка ОДО с реестром (пересчёт ЕРЧ)"
            );
          }
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
          obligDate,
          regDate,
          periods,
          periodsText: periods.length ? formatPeriods(periods) : "",
          odoResidualErch,
          odoCheck,
          contractsCount: counted.length,
          competitiveCount: competitive.length,
          excludedCount,
          beforeMembershipCount: beforeMembershipList.length,
          contractsInSuspensionCount: contractsInSuspension.length,
          maxInSuspension: contractsInSuspension.length ? maxInSuspension : null,
          vvCheckSusp,
          odoResidualSusp: contractsInSuspension.length ? odoResidualSusp : null,
          odoCheckSusp,
          risk,
          flags,
          comment: comments.join("; "),
        });
      }

      companies.sort((a, b) => {
        const order = { КРИТИЧНО: 0, "РУЧНАЯ ПРОВЕРКА": 1, НОРМА: 2 };
        return order[a.risk] - order[b.risk] || b.maxContract - a.maxContract;
      });

      const byType = (t) =>
        contracts.list.filter((c) => !c.excluded && c.methodType === t).length;
      const summary = {
        contracts: contracts.list.filter((c) => !c.excluded).length,
        contractsRaw: contracts.list.length,
        excludedContracts: contracts.list.filter((c) => c.excluded).length,
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
        vvExceedSusp: companies.filter((c) => c.flags.vvExceedSusp).length,
        odoExceedSusp: companies.filter((c) => c.flags.odoExceedSusp).length,
        odoMismatch: companies.filter((c) => c.flags.odoMismatch).length,
        beforeMembershipCompanies: companies.filter((c) => c.beforeMembershipCount > 0).length,
        beforeMembershipContracts: contracts.list.filter((c) => c.beforeMembership).length,
        critical: companies.filter((c) => c.risk === "КРИТИЧНО").length,
        manual: companies.filter((c) => c.risk === "РУЧНАЯ ПРОВЕРКА").length,
        manualErch: companies.filter(
          (c) => c.risk === "РУЧНАЯ ПРОВЕРКА" && c.flags.odoMismatch
        ).length,
        manualIncomplete: companies.filter(
          (c) => c.risk === "РУЧНАЯ ПРОВЕРКА" && c.flags.dataIncomplete
        ).length,
        ok: companies.filter((c) => c.risk === "НОРМА").length,
        sumAccepted: contracts.list
          .filter((c) => !c.excluded)
          .reduce((s, c) => s + (c.amount || 0), 0),
        odoResidualTotal: companies.reduce((s, c) => s + (c.odoResidual || 0), 0),
        assumptionNoDone: !contracts.cols.done,
        mapped: {
          amount: contracts.cols.amount,
          done: contracts.cols.done,
          date: contracts.cols.date,
          number: contracts.cols.number,
          name: contracts.cols.name,
          method: contracts.cols.method,
          exclude: contracts.cols.exclude,
          memberVv: members.cols.vv,
          memberOdo: members.cols.odo,
          memberOblig: members.cols.oblig,
          memberObligDate: members.cols.obligDate,
          memberRegDate: members.cols.regDate,
        },
      };

      const risks = companies.filter((c) => c.risk === "КРИТИЧНО");
      return {
        summary,
        companies,
        risks,
        actionQueue: buildActionQueue(companies),
        manual: companies.filter((c) => c.risk === "РУЧНАЯ ПРОВЕРКА"),
        manualErch: companies.filter(
          (c) => c.risk === "РУЧНАЯ ПРОВЕРКА" && c.flags.odoMismatch
        ),
        manualIncomplete: companies.filter(
          (c) => c.risk === "РУЧНАЯ ПРОВЕРКА" && c.flags.dataIncomplete
        ),
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

    /** Тяжесть для очереди «К исполнению»: 0 приостановка/право → 1 ВВ/ОДО → 2 нет в реестре. */
    function actionSeverity(c) {
      const f = c.flags || {};
      const rightIssue =
        f.suspended || c.right === "приостановлено" || c.right === "прекращено";
      if (rightIssue) return 0;
      if (f.vvExceed || f.odoExceed || f.noOdo || f.vvExceedSusp || f.odoExceedSusp) return 1;
      if (f.notFound) return 2;
      return 3;
    }

    function actionRiskType(c) {
      const f = c.flags || {};
      const parts = [];
      if (f.suspended) parts.push("Приостановка");
      else if (c.right === "приостановлено" || c.right === "прекращено") {
        parts.push(c.right === "прекращено" ? "Право прекращено" : "Право приостановлено");
      }
      if (f.vvExceed || f.vvExceedSusp) parts.push("Превышение ВВ");
      if (f.odoExceed || f.odoExceedSusp) parts.push("Превышение ОДО");
      if (f.noOdo) parts.push("Нет ОДО");
      if (f.notFound) parts.push("Нет в реестре");
      return parts.length ? parts.join("; ") : "Критично";
    }

    function actionMetric(c) {
      return `max ${fmtMoney(c.maxContract)} · ОДО ${fmtMoney(c.odoResidual)}`;
    }

    function actionCheckHint(c) {
      const f = c.flags || {};
      const hints = [];
      if (f.suspended) {
        hints.push(
          `сверить ${c.contractsInSuspensionCount || 0} дог. с периодами приостановки`
        );
      } else if (c.right === "приостановлено" || c.right === "прекращено") {
        hints.push(`подтвердить статус права («${c.right}») в реестре`);
      }
      if (f.vvExceed || f.vvExceedSusp) {
        hints.push(
          `сверить max договор ${fmtMoney(c.maxContract)} с лимитом ВВ ${fmtMoney(c.vvLimit)}`
        );
      }
      if (f.odoExceed || f.odoExceedSusp) {
        hints.push(
          `сверить остаток ОДО ${fmtMoney(c.odoResidual)} с лимитом ${fmtMoney(c.odoLimit)}`
        );
      }
      if (f.noOdo) hints.push("проверить, почему нет уровня ОДО при 44/223/615");
      if (f.notFound) hints.push("найти ИНН в реестре членов / уточнить принадлежность к СРО");
      if (!hints.length && c.comment) return c.comment;
      return hints.join("; ") || "разобрать критичный риск";
    }

    function letterDraftFor(c) {
      const name = c.name || "—";
      const type =
        c.risk === "КРИТИЧНО"
          ? actionRiskType(c)
          : c.comment || c.risk || "без формулировки";
      const vvLabel = (c.vvText || fmtMoney(c.vvLimit) || "н/д").replace(/\.\s*$/, "");
      const odoLabel = (c.odoText || fmtMoney(c.odoLimit) || "н/д").replace(/\.\s*$/, "");
      let lead;
      if (c.risk === "КРИТИЧНО") {
        lead = `По результатам сверки договоров члена СРО ИНН ${c.inn} (${name}) выявлено критическое отклонение: ${type}.`;
      } else if (c.risk === "РУЧНАЯ ПРОВЕРКА") {
        lead = `По результатам сверки договоров члена СРО ИНН ${c.inn} (${name}) требуется ручная проверка: ${type}.`;
      } else {
        lead = `Справка по члену СРО ИНН ${c.inn} (${name}): статус «${c.risk || "НОРМА"}».`;
      }
      const bits = [lead];
      if (c.vvLimit != null || c.maxContract) {
        bits.push(
          `Максимальный один договор: ${fmtMoney(c.maxContract)} при уровне ВВ ${vvLabel}.`
        );
      }
      if (c.odoLimit != null || c.odoResidual) {
        bits.push(
          `Остаток обязательств (ОДО): ${fmtMoney(c.odoResidual)} при уровне ОДО ${odoLabel}.`
        );
      }
      if (c.periodsText) bits.push(`Периоды приостановки: ${c.periodsText}.`);
      if (c.comment && c.risk === "КРИТИЧНО") bits.push(`Детали: ${c.comment}.`);
      else if (c.comment && c.risk !== "РУЧНАЯ ПРОВЕРКА") bits.push(`Детали: ${c.comment}.`);
      bits.push("Рекомендуется провести проверку и зафиксировать решение в протоколе.");
      return bits.join(" ");
    }

    function actionLetterDraft(c) {
      return letterDraftFor(c);
    }

    function buildActionQueue(companies) {
      return companies
        .filter((c) => c.risk === "КРИТИЧНО")
        .map((c) => ({
          ...c,
          queueSeverity: actionSeverity(c),
          riskTypeShort: actionRiskType(c),
          metricShort: actionMetric(c),
          checkHint: actionCheckHint(c),
          letterDraft: actionLetterDraft(c),
          sortAmount: Math.max(c.maxContract || 0, c.odoResidual || 0),
        }))
        .sort(
          (a, b) =>
            a.queueSeverity - b.queueSeverity ||
            b.sortAmount - a.sortAmount ||
            String(a.inn).localeCompare(String(b.inn), "ru")
        );
    }

    function companiesByFilter(r, filter) {
      if (!filter) return r.companies;
      if (filter === "risks") return r.risks;
      if (filter === "manual") return r.manual;
      if (filter === "manualErch") return r.manualErch;
      if (filter === "manualIncomplete") return r.manualIncomplete;
      if (filter === "odoExceed") return r.companies.filter((c) => c.flags.odoExceed);
      if (filter === "noOdo") return r.companies.filter((c) => c.flags.noOdo);
      if (filter === "vvExceed") return r.companies.filter((c) => c.flags.vvExceed);
      if (filter === "suspended") return r.companies.filter((c) => c.flags.suspended);
      if (filter === "vvExceedSusp") return r.companies.filter((c) => c.flags.vvExceedSusp);
      if (filter === "odoExceedSusp") return r.companies.filter((c) => c.flags.odoExceedSusp);
      if (filter === "odoMismatch") return r.companies.filter((c) => c.flags.odoMismatch);
      if (filter === "beforeMembership") return r.companies.filter((c) => c.beforeMembershipCount > 0);
      return r.companies;
    }

    function renderMapBanner(summary) {
      const box = document.getElementById("mapWarn");
      if (!box) return;
      const m = summary.mapped || {};
      const lines = [];
      if (summary.assumptionNoDone) {
        lines.push(
          "Колонка исполнения не найдена — остаток ОДО = полная стоимость договора. Проверьте, что в файле есть «Сумма/стоимость принятых работ»."
        );
      }
      if (!m.date) {
        lines.push("Дата договора не распознана — проверка приостановки по датам недоступна.");
      }
      if (m.memberOblig && !m.memberObligDate) {
        lines.push(
          "Дата расчёта обязательств не найдена — сверка ЕРЧ сравнивает остаток по всем договорам с даты регистрации."
        );
      }
      if (!m.memberRegDate) {
        lines.push(
          "Дата регистрации в реестре СРО не найдена — договоры до вступления из лимитов не выводятся."
        );
      }
      const mapBits = [
        m.amount ? `стоимость: «${m.amount}»` : null,
        m.done ? `исполнено: «${m.done}»` : "исполнено: не найдено",
        m.date ? `дата: «${m.date}»` : null,
      ].filter(Boolean);
      if (mapBits.length) {
        lines.push("Колонки договоров: " + mapBits.join(" · "));
      }
      if (summary.excludedContracts > 0) {
        lines.push(`Исключено из обязательств по флагу: ${summary.excludedContracts} дог.`);
      }
      if (summary.beforeMembershipContracts > 0) {
        lines.push(
          `До регистрации в реестре: ${summary.beforeMembershipContracts} дог. На дату заключения членства нет, в лимиты ВВ/ОДО не входят.`
        );
      }
      if (!lines.length) {
        box.style.display = "none";
        box.textContent = "";
        return;
      }
      box.style.display = "block";
      box.className = summary.assumptionNoDone ? "map-warn danger" : "map-warn";
      box.textContent = lines.join(" ");
    }

    function renderStats(summary) {
      const s = summary;
      renderMapBanner(s);
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
            ${
              s.excludedContracts
                ? `<div class="summary-row sub muted"><span>исключены из обязательств</span><strong>${s.excludedContracts}</strong></div>`
                : ""
            }
          </div>
          <div class="summary-block">
            <h3>Члены СРО</h3>
            <button type="button" class="summary-row clickable crit" data-filter="noOdo">
              <span>1. Нет уровня ОДО (есть 44, 223, 615)</span><strong>${s.noOdo}</strong>
            </button>
            <button type="button" class="summary-row clickable crit" data-filter="vvExceed">
              <span>2. Превышен уровень ВВ</span><strong>${s.vvExceed}</strong>
            </button>
            <button type="button" class="summary-row clickable crit" data-filter="odoExceed">
              <span>3. Превышен уровень ОДО</span><strong>${s.odoExceed}</strong>
            </button>
            <button type="button" class="summary-row clickable crit" data-filter="suspended">
              <span>4. Договоры во время приостановки</span><strong>${s.suspended}</strong>
            </button>
            <button type="button" class="summary-row clickable crit nest" data-filter="vvExceedSusp">
              <span>4.1. Превышен ВВ в период приостановки</span><strong>${s.vvExceedSusp}</strong>
            </button>
            <button type="button" class="summary-row clickable crit nest" data-filter="odoExceedSusp">
              <span>4.2. Превышен ОДО во время приостановки</span><strong>${s.odoExceedSusp}</strong>
            </button>
            <button type="button" class="summary-row clickable warn" data-filter="odoMismatch">
              <span>6. Сверка ЕРЧ</span><strong>${s.odoMismatch}</strong>
            </button>
            <button type="button" class="summary-row clickable warn" data-filter="manualErch">
              <span>Ручная · ЕРЧ</span><strong>${s.manualErch}</strong>
            </button>
            <button type="button" class="summary-row clickable warn" data-filter="manualIncomplete">
              <span>Ручная · данные неполные</span><strong>${s.manualIncomplete}</strong>
            </button>
            <button type="button" class="summary-row clickable" data-filter="beforeMembership">
              <span>До регистрации в реестре</span><strong>${s.beforeMembershipCompanies}</strong>
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

    function escHtml(s) {
      return String(s ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    }

    function companyByInn(inn) {
      const r = state.result;
      if (!r || !inn) return null;
      return (r.companies || []).find((c) => String(c.inn) === String(inn)) || null;
    }

    function methodTypeLabel(t) {
      if (t === "44") return "44-ФЗ";
      if (t === "223") return "223-ФЗ";
      if (t === "615") return "615-ФЗ";
      if (t === "direct") return "прямой";
      if (t === "other_comp") return "конкур.";
      return "неясно";
    }

    function renderTable() {
      const r = state.result;
      if (!r) return;
      let rows = [];
      let head = [];
      let wrapCols = new Set();
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
          "Исключён",
          "До регистрации",
          "Приостановка",
        ];
        rows = r.contracts.map((c) => ({
          inn: c.inn,
          cells: [
            c.inn,
            c.name,
            c.number,
            c.date || "—",
            fmtMoney(c.amount),
            fmtMoney(c.done),
            c.excluded ? "—" : fmtMoney(c.residual),
            c.method || "—",
            methodTypeLabel(c.methodType),
            c.excluded ? "да" : "—",
            c.beforeMembership ? "да" : "—",
            c.inSuspensionPeriod ? "в периоде" : "—",
          ],
        }));
      } else if (state.tab === "queue") {
        const src = r.actionQueue || [];
        head = ["ИНН", "Компания", "Тип риска", "Max / остаток ОДО", "Что проверить"];
        wrapCols = new Set([2, 4]);
        rows = src.map((c) => ({
          inn: c.inn,
          cells: [
            c.inn,
            c.name || "—",
            c.riskTypeShort,
            c.metricShort,
            c.checkHint,
          ],
        }));
      } else {
        const src =
          state.tab === "risks"
            ? r.risks
            : state.tab === "manualErch"
              ? r.manualErch
              : state.tab === "manualIncomplete"
                ? r.manualIncomplete
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
        wrapCols = new Set([8]);
        rows = src.map((c) => ({
          inn: c.inn,
          cells: [
            riskPill(c.risk),
            c.inn,
            c.name,
            c.right,
            c.vvText || "—",
            fmtMoney(c.maxContract),
            c.odoText || "—",
            fmtMoney(c.odoResidual),
            c.comment || "—",
          ],
        }));
      }
      document.getElementById("thead").innerHTML =
        "<tr>" + head.map((h) => `<th>${h}</th>`).join("") + "</tr>";
      document.getElementById("tbody").innerHTML = rows
        .map((row) => {
          const innAttr = row.inn ? ` data-inn="${escHtml(row.inn)}"` : "";
          const openable = row.inn ? ' class="row-open" tabindex="0" role="button"' : "";
          const label = row.inn
            ? ` aria-label="Открыть карточку ИНН ${escHtml(row.inn)}"`
            : "";
          return (
            `<tr${innAttr}${openable}${label}>` +
            row.cells
              .map(
                (cell, i) =>
                  `<td${wrapCols.has(i) ? ' class="cell-wrap"' : ""}>${cell}</td>`
              )
              .join("") +
            "</tr>"
          );
        })
        .join("");
      updateTabCounts(r);
    }

    function updateTabCounts(r) {
      const counts = {
        queue: (r.actionQueue || []).length,
        risks: (r.risks || []).length,
        manualErch: (r.manualErch || []).length,
        manualIncomplete: (r.manualIncomplete || []).length,
      };
      document.querySelectorAll(".tabs button[data-tab]").forEach((btn) => {
        const base = btn.dataset.label || btn.textContent;
        const n = counts[btn.dataset.tab];
        btn.textContent = n != null ? `${base} ${n}` : base;
      });
    }

    const innDrawer = {
      overlay: document.getElementById("innDrawerOverlay"),
      panel: document.getElementById("innDrawer"),
      body: document.getElementById("innDrawerBody"),
      title: document.getElementById("innDrawerTitle"),
      meta: document.getElementById("innDrawerMeta"),
      stamp: document.getElementById("innDrawerStamp"),
      copyBtn: document.getElementById("innCopyLetter"),
      closeBtn: document.getElementById("innDrawerClose"),
      copyStatus: document.getElementById("innCopyStatus"),
      lastFocus: null,
      letter: "",
      inn: null,
    };

    function rightChipClass(right) {
      const s = String(right || "").toLowerCase();
      if (s.includes("приостанов") || s.includes("прекращ") || s.includes("не найден")) {
        return "inn-chip inn-chip--crit";
      }
      if (s.includes("действ")) return "inn-chip inn-chip--ok";
      return "inn-chip";
    }

    function checkChip(label, value) {
      const v = String(value || "н/д");
      let cls = "inn-chip";
      if (v === "превышение") cls += " inn-chip--crit";
      else if (v === "ок") cls += " inn-chip--ok";
      return `<span class="${cls}">${escHtml(label)}: ${escHtml(v)}</span>`;
    }

    function renderInnCard(inn) {
      const c = companyByInn(inn);
      if (!c || !innDrawer.body) return;
      innDrawer.inn = String(c.inn);
      innDrawer.letter = letterDraftFor(c);
      if (innDrawer.title) {
        innDrawer.title.textContent = c.name || `ИНН ${c.inn}`;
      }
      if (innDrawer.meta) {
        innDrawer.meta.innerHTML = `<span class="inn-mono">ИНН ${escHtml(c.inn)}</span>`;
      }
      if (innDrawer.stamp) {
        innDrawer.stamp.innerHTML = riskPill(c.risk);
      }
      if (innDrawer.copyStatus) {
        innDrawer.copyStatus.textContent = "";
        innDrawer.copyStatus.classList.add("hidden");
      }

      const contracts = (state.result.contracts || []).filter(
        (x) => String(x.inn) === String(c.inn)
      );
      const periodsHtml = c.periodsText
        ? `<ul class="inn-periods">${c.periodsText
            .split("; ")
            .map((p) => `<li>${escHtml(p)}</li>`)
            .join("")}</ul>`
        : `<p class="inn-empty">Периодов приостановки нет</p>`;

      const contractRows = contracts.length
        ? contracts
            .map((x) => {
              const flags = [];
              if (x.excluded) flags.push("искл.");
              if (x.beforeMembership) flags.push("до рег.");
              if (x.inSuspensionPeriod) flags.push("приостановка");
              return `<tr>
                <td>${escHtml(x.number || "—")}</td>
                <td class="num">${escHtml(x.date || "—")}</td>
                <td class="num">${escHtml(fmtMoney(x.amount))}</td>
                <td class="num">${x.excluded ? "—" : escHtml(fmtMoney(x.residual))}</td>
                <td>${escHtml(methodTypeLabel(x.methodType))}</td>
                <td class="cell-wrap">${flags.length ? escHtml(flags.join(", ")) : "—"}</td>
              </tr>`;
            })
            .join("")
        : `<tr><td colspan="6" class="inn-empty">Договоров по ИНН нет</td></tr>`;

      innDrawer.body.innerHTML = `
        <section class="inn-block" aria-labelledby="innLimitsTitle">
          <h3 id="innLimitsTitle" class="inn-block-title">Лимиты и факт</h3>
          <dl class="inn-grid">
            <div><dt>Право</dt><dd><span class="${rightChipClass(c.right)}">${escHtml(c.right || "—")}</span></dd></div>
            <div><dt>Уровень ВВ</dt><dd class="inn-mono">${escHtml(c.vvText || fmtMoney(c.vvLimit) || "—")}</dd></div>
            <div><dt>Max договор</dt><dd class="inn-mono">${escHtml(fmtMoney(c.maxContract))}</dd></div>
            <div><dt>Проверка ВВ</dt><dd>${checkChip("ВВ", c.vvCheck)}</dd></div>
            <div><dt>Уровень ОДО</dt><dd class="inn-mono">${escHtml(c.odoText || fmtMoney(c.odoLimit) || "—")}</dd></div>
            <div><dt>Остаток ОДО</dt><dd class="inn-mono">${escHtml(fmtMoney(c.odoResidual))}</dd></div>
            <div><dt>Проверка ОДО</dt><dd>${checkChip("ОДО", c.odoCheck)}</dd></div>
            <div><dt>Реестр обязательств</dt><dd class="inn-mono">${escHtml(fmtMoney(c.registryOblig))}</dd></div>
          </dl>
        </section>
        <section class="inn-block" aria-labelledby="innSuspTitle">
          <h3 id="innSuspTitle" class="inn-block-title">Приостановки</h3>
          ${periodsHtml}
          <p class="inn-aside">
            Договоров в периоде: <strong class="inn-mono">${c.contractsInSuspensionCount || 0}</strong>
            · ВВ в приостановке: ${checkChip("ВВ", c.vvCheckSusp)}
            · ОДО в приостановке: ${checkChip("ОДО", c.odoCheckSusp)}
          </p>
        </section>
        <section class="inn-block" aria-labelledby="innContractsTitle">
          <h3 id="innContractsTitle" class="inn-block-title">
            Договоры ИНН
            <span class="inn-count">${contracts.length}</span>
          </h3>
          <div class="inn-table-wrap">
            <table class="inn-table">
              <thead>
                <tr>
                  <th>№</th>
                  <th>Дата</th>
                  <th>К учёту</th>
                  <th>Остаток</th>
                  <th>Тип</th>
                  <th>Метки</th>
                </tr>
              </thead>
              <tbody>${contractRows}</tbody>
            </table>
          </div>
        </section>
        ${
          c.comment
            ? `<section class="inn-block">
                <h3 class="inn-block-title">Комментарий проверки</h3>
                <p class="inn-comment">${escHtml(c.comment)}</p>
              </section>`
            : ""
        }
        <section class="inn-block inn-letter-preview" aria-labelledby="innLetterTitle">
          <h3 id="innLetterTitle" class="inn-block-title">Формулировка для письма</h3>
          <p class="inn-letter-text">${escHtml(innDrawer.letter)}</p>
        </section>
      `;
    }

    function openInnCard(inn) {
      if (!innDrawer.overlay || !companyByInn(inn)) return;
      innDrawer.lastFocus = document.activeElement;
      renderInnCard(inn);
      innDrawer.overlay.classList.remove("hidden");
      innDrawer.overlay.removeAttribute("hidden");
      document.body.classList.add("inn-drawer-open");
      const focusEl = innDrawer.closeBtn || innDrawer.panel;
      if (focusEl && typeof focusEl.focus === "function") focusEl.focus();
    }

    function closeInnCard() {
      if (!innDrawer.overlay || innDrawer.overlay.classList.contains("hidden")) return;
      innDrawer.overlay.classList.add("hidden");
      innDrawer.overlay.setAttribute("hidden", "");
      document.body.classList.remove("inn-drawer-open");
      innDrawer.inn = null;
      innDrawer.letter = "";
      if (innDrawer.lastFocus && typeof innDrawer.lastFocus.focus === "function") {
        innDrawer.lastFocus.focus();
      }
      innDrawer.lastFocus = null;
    }

    async function copyInnLetter() {
      if (!innDrawer.letter) return;
      const done = () => {
        if (!innDrawer.copyStatus) return;
        innDrawer.copyStatus.textContent = "Скопировано";
        innDrawer.copyStatus.classList.remove("hidden");
      };
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(innDrawer.letter);
          done();
          return;
        }
      } catch (_) {
        /* fallback below */
      }
      const ta = document.createElement("textarea");
      ta.value = innDrawer.letter;
      ta.setAttribute("readonly", "");
      ta.style.position = "fixed";
      ta.style.left = "-9999px";
      document.body.appendChild(ta);
      ta.select();
      try {
        document.execCommand("copy");
        done();
      } finally {
        document.body.removeChild(ta);
      }
    }

    function run() {
      clearError();
      try {
        closeInnCard();
        state.filter = null;
        state.tab = "queue";
        state.result = analyze(state.membersRows, state.contractsRows);
        results.classList.remove("hidden");
        exportBtn.classList.remove("hidden");
        document.querySelectorAll(".tabs button").forEach((b) => {
          const on = b.dataset.tab === "queue";
          b.classList.toggle("active", on);
          b.setAttribute("aria-selected", on ? "true" : "false");
        });
        renderStats(state.result.summary);
        renderTable();
      } catch (e) {
        showError(e.message || String(e));
      }
    }

    const XL = {
      header: {
        font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11, name: "Calibri" },
        fill: { patternType: "solid", fgColor: { rgb: "1E293B" } },
        alignment: { vertical: "center", wrapText: true, horizontal: "left" },
        border: {
          top: { style: "thin", color: { rgb: "CBD5E1" } },
          bottom: { style: "thin", color: { rgb: "CBD5E1" } },
          left: { style: "thin", color: { rgb: "CBD5E1" } },
          right: { style: "thin", color: { rgb: "CBD5E1" } },
        },
      },
      cell: {
        font: { sz: 10, name: "Calibri" },
        alignment: { vertical: "center", wrapText: true },
        border: {
          top: { style: "thin", color: { rgb: "E2E8F0" } },
          bottom: { style: "thin", color: { rgb: "E2E8F0" } },
          left: { style: "thin", color: { rgb: "E2E8F0" } },
          right: { style: "thin", color: { rgb: "E2E8F0" } },
        },
      },
      label: {
        font: { bold: true, sz: 10, name: "Calibri", color: { rgb: "334155" } },
        fill: { patternType: "solid", fgColor: { rgb: "F1F5F9" } },
        alignment: { vertical: "center" },
      },
      money: "#,##0",
      risk: {
        КРИТИЧНО: {
          font: { bold: true, sz: 10, color: { rgb: "B91C1C" }, name: "Calibri" },
          fill: { patternType: "solid", fgColor: { rgb: "FEE2E2" } },
          alignment: { vertical: "center", horizontal: "center" },
        },
        "РУЧНАЯ ПРОВЕРКА": {
          font: { bold: true, sz: 10, color: { rgb: "B45309" }, name: "Calibri" },
          fill: { patternType: "solid", fgColor: { rgb: "FEF3C7" } },
          alignment: { vertical: "center", horizontal: "center" },
        },
        НОРМА: {
          font: { bold: true, sz: 10, color: { rgb: "047857" }, name: "Calibri" },
          fill: { patternType: "solid", fgColor: { rgb: "D1FAE5" } },
          alignment: { vertical: "center", horizontal: "center" },
        },
      },
      rowFill: {
        КРИТИЧНО: "FEF2F2",
        "РУЧНАЯ ПРОВЕРКА": "FFFBEB",
        НОРМА: "FFFFFF",
      },
    };

    function excelLimit(v) {
      if (v === null || v === undefined) return "";
      if (v === INF) return "без предела";
      return v;
    }

    function styleSheet(ws, opts) {
      const { moneyCols = [], riskCol = -1, headerRows = 1, widths = [] } = opts;
      if (!ws["!ref"]) return ws;
      const range = XLSX.utils.decode_range(ws["!ref"]);
      ws["!cols"] = widths.map((wch) => ({ wch }));
      ws["!rows"] = [{ hpt: 24 }];
      for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws[addr]) continue;
          const isHeader = R < headerRows;
          let base = isHeader ? { ...XL.header } : { ...XL.cell };
          if (!isHeader && riskCol >= 0) {
            const riskAddr = XLSX.utils.encode_cell({ r: R, c: riskCol });
            const riskVal = ws[riskAddr] && ws[riskAddr].v;
            const fillRgb = XL.rowFill[riskVal];
            if (fillRgb && fillRgb !== "FFFFFF") {
              base = {
                ...base,
                fill: { patternType: "solid", fgColor: { rgb: fillRgb } },
              };
            }
            if (C === riskCol && XL.risk[riskVal]) {
              base = { ...XL.risk[riskVal], border: XL.cell.border };
            }
          }
          if (!isHeader && moneyCols.includes(C) && typeof ws[addr].v === "number") {
            ws[addr].z = XL.money;
            base = {
              ...base,
              alignment: { ...base.alignment, horizontal: "right" },
            };
          }
          if (isHeader && opts.summaryLabels && C === 0) {
            base = { ...XL.label, border: XL.cell.border };
          }
          ws[addr].s = base;
        }
      }
      if (opts.autofilter && range.e.r >= headerRows) {
        ws["!autofilter"] = { ref: ws["!ref"] };
      }
      ws["!freeze"] = { xSplit: 0, ySplit: headerRows, topLeftCell: "A" + (headerRows + 1), state: "frozen" };
      return ws;
    }

    function companyRows(list) {
      return list.map((c) => [
        c.inn,
        c.name,
        c.found ? "да" : "нет",
        c.right,
        c.vvText || "",
        excelLimit(c.vvLimit),
        c.maxContract || 0,
        c.vvCheck,
        c.odoText || "",
        excelLimit(c.odoLimit),
        c.odoResidual || 0,
        c.registryOblig == null ? "" : c.registryOblig,
        c.obligDate ? fmtDate(c.obligDate) : "",
        c.odoResidualErch == null ? "" : c.odoResidualErch,
        c.odoCheck,
        c.contractsCount,
        c.competitiveCount,
        c.excludedCount || 0,
        c.contractsInSuspensionCount || 0,
        c.maxInSuspension == null ? "" : c.maxInSuspension,
        c.vvCheckSusp || "",
        c.odoResidualSusp == null ? "" : c.odoResidualSusp,
        c.odoCheckSusp || "",
        c.regDate ? fmtDate(c.regDate) : "",
        c.beforeMembershipCount || 0,
        c.risk,
        c.comment || "",
      ]);
    }

    const COMPANY_HEADER = [
      "ИНН",
      "Компания",
      "Найдена в реестре",
      "Состояние права",
      "Уровень ВВ",
      "Лимит ВВ, ₽",
      "Макс. один договор, ₽",
      "Проверка ВВ",
      "Уровень ОДО",
      "Лимит ОДО, ₽",
      "Остаток ОДО по договорам, ₽",
      "Расчёт обязательств из реестра, ₽",
      "Дата расчёта обязательств",
      "Остаток ОДО на дату расчёта, ₽",
      "Проверка ОДО",
      "Кол-во договоров",
      "Кол-во конкурентных",
      "Исключено из обязательств",
      "Договоров в периоде приостановки",
      "Макс. договор в приостановке, ₽",
      "ВВ в приостановке",
      "Остаток ОДО в приостановке, ₽",
      "ОДО в приостановке",
      "Дата регистрации в реестре",
      "Договоров до регистрации",
      "Итоговый риск",
      "Комментарий",
    ];

    const COMPANY_WIDTHS = [12, 28, 12, 16, 18, 14, 16, 12, 18, 14, 18, 18, 14, 18, 14, 10, 10, 12, 12, 16, 14, 18, 14, 16, 14, 16, 42];
    const COMPANY_MONEY = [5, 6, 9, 10, 11, 13, 19, 21];
    const COMPANY_RISK_COL = 25;

    function exportExcel() {
      const r = state.result;
      if (!r) return;
      if (typeof XLSX === "undefined") {
        showError("Библиотека Excel не загрузилась.");
        return;
      }

      const s = r.summary;
      const today = new Date();
      const dateStr =
        String(today.getDate()).padStart(2, "0") +
        "." +
        String(today.getMonth() + 1).padStart(2, "0") +
        "." +
        today.getFullYear();

      const meta = state.membersMeta;
      const summaryAoA = [
        ["Показатель", "Значение"],
        ["Дата проверки", dateStr],
        [
          "Источник реестра членов",
          meta
            ? `НОСТРОЙ, СРО ${meta.sro_id}`
            : state.membersSource === "file"
              ? "файл реестра членов"
              : "демо-данные",
        ],
        ["Дата выгрузки реестра членов", meta ? fmtExportedAt(meta.exported_at) : ""],
        ["Членов в реестре", meta ? meta.stats.members : state.membersRows.length],
        ["Всего договоров", s.contracts],
        ["из них по 44-ФЗ", s.byFz44],
        ["по 223-ФЗ", s.byFz223],
        ["по 615-ФЗ", s.byFz615],
        ["прямые", s.byDirect],
        ["прочие конкурентные", s.byOtherComp],
        ["без вида закупки", s.byUnclear],
        ["Уникальных ИНН", s.inns],
        ["Найдено в реестре", s.found],
        ["Не найдено", s.notFound],
        ["1. Нет уровня ОДО (есть 44, 223, 615)", s.noOdo],
        ["2. Превышен уровень ВВ", s.vvExceed],
        ["3. Превышен уровень ОДО", s.odoExceed],
        ["4. Договоры во время приостановки", s.suspended],
        ["4.1. Превышен ВВ в период приостановки", s.vvExceedSusp],
        ["4.2. Превышен ОДО во время приостановки", s.odoExceedSusp],
        ["6. Сверка ЕРЧ", s.odoMismatch],
        ["Договоры до регистрации в реестре", s.beforeMembershipContracts],
        ["Компаний с договорами до регистрации", s.beforeMembershipCompanies],
        ["Критичные (всего)", s.critical],
        ["Ручная · ЕРЧ", s.manualErch],
        ["Ручная · данные неполные", s.manualIncomplete],
        ["Ручная проверка (всего)", s.manual],
        ["Норма", s.ok],
        ["Сумма, принятая СРО к учёту, ₽", s.sumAccepted],
        ["Остаток ОДО (расчёт), ₽", s.odoResidualTotal],
      ];

      const contractsAoA = [
        [
          "ИНН",
          "Компания",
          "Номер договора",
          "Дата",
          "Стоимость к учёту, ₽",
          "Исполнено, ₽",
          "Остаток, ₽",
          "Способ закупки",
          "Тип",
          "Конкурентный",
          "Участвует в ОДО",
          "Исключён из обязательств",
          "До регистрации в реестре",
          "В периоде приостановки",
          "Комментарий",
        ],
        ...r.contracts.map((c) => {
          const typeLabel =
            c.methodType === "44"
              ? "44-ФЗ"
              : c.methodType === "223"
                ? "223-ФЗ"
                : c.methodType === "615"
                  ? "615-ФЗ"
                  : c.methodType === "direct"
                    ? "прямой"
                    : c.methodType === "other_comp"
                      ? "прочий конкурентный"
                      : "неясно";
          const notes = [];
          if (c.weirdMoney) notes.push("странный формат суммы");
          if (c.assumptionNoDone) notes.push("исполнение не найдено");
          if (c.excluded) notes.push("не учитывать в обязательствах");
          if (c.beforeMembership) {
            notes.push(
              `заключён до регистрации в реестре СРО ${fmtDate(c.membershipRegDate)} — на дату заключения членства нет`
            );
          }
          if (c.inSuspensionPeriod) notes.push("дата в периоде приостановки");
          return [
            c.inn,
            c.name,
            c.number,
            c.date || "",
            c.amount == null ? "" : c.amount,
            c.done,
            c.excluded || c.residual == null ? "" : c.residual,
            c.method || "",
            typeLabel,
            c.unclear ? "неясно" : c.competitive ? "да" : "нет",
            c.excluded || c.beforeMembership ? "нет" : c.competitive ? "да" : "нет",
            c.excluded ? "да" : "нет",
            c.beforeMembership ? "да" : "нет",
            c.inSuspensionPeriod ? "да" : "нет",
            notes.join("; "),
          ];
        }),
      ];

      const companiesAoA = [COMPANY_HEADER, ...companyRows(r.companies)];
      const risksAoA = [COMPANY_HEADER, ...companyRows(r.risks)];
      const manualErchAoA = [COMPANY_HEADER, ...companyRows(r.manualErch || [])];
      const manualIncompleteAoA = [
        COMPANY_HEADER,
        ...companyRows(r.manualIncomplete || []),
      ];
      const queue = r.actionQueue || buildActionQueue(r.companies);
      const queueAoA = [
        [
          "ИНН",
          "Компания",
          "Тип риска",
          "Макс. один договор, ₽",
          "Остаток ОДО, ₽",
          "Что проверить",
          "Черновик для письма / протокола",
        ],
        ...queue.map((c) => [
          c.inn,
          c.name || "",
          c.riskTypeShort,
          c.maxContract || 0,
          c.odoResidual || 0,
          c.checkHint,
          c.letterDraft,
        ]),
      ];

      const wsSummary = XLSX.utils.aoa_to_sheet(summaryAoA);
      wsSummary["!cols"] = [{ wch: 48 }, { wch: 22 }];
      wsSummary["!rows"] = [{ hpt: 24 }];
      const sumRange = XLSX.utils.decode_range(wsSummary["!ref"]);
      for (let R = sumRange.s.r; R <= sumRange.e.r; R++) {
        const a = XLSX.utils.encode_cell({ r: R, c: 0 });
        const b = XLSX.utils.encode_cell({ r: R, c: 1 });
        if (R === 0) {
          if (wsSummary[a]) wsSummary[a].s = XL.header;
          if (wsSummary[b]) wsSummary[b].s = XL.header;
          continue;
        }
        if (wsSummary[a]) wsSummary[a].s = { ...XL.label, border: XL.cell.border };
        if (wsSummary[b]) {
          const isMoney =
            typeof wsSummary[b].v === "number" &&
            /₽|остаток ОДО|сумма/i.test(String(wsSummary[a]?.v || ""));
          if (isMoney) wsSummary[b].z = XL.money;
          wsSummary[b].s = {
            ...XL.cell,
            font: { bold: true, sz: 10, name: "Calibri" },
            alignment: {
              vertical: "center",
              horizontal: typeof wsSummary[b].v === "number" ? "right" : "left",
            },
          };
        }
      }

      const wsCompanies = styleSheet(XLSX.utils.aoa_to_sheet(companiesAoA), {
        headerRows: 1,
        widths: COMPANY_WIDTHS,
        moneyCols: COMPANY_MONEY,
        riskCol: COMPANY_RISK_COL,
        autofilter: true,
      });
      const wsContracts = styleSheet(XLSX.utils.aoa_to_sheet(contractsAoA), {
        headerRows: 1,
        widths: [12, 28, 14, 12, 16, 14, 14, 16, 14, 12, 14, 12, 16, 12, 52],
        moneyCols: [4, 5, 6],
        autofilter: true,
      });
      const wsRisks = styleSheet(XLSX.utils.aoa_to_sheet(risksAoA), {
        headerRows: 1,
        widths: COMPANY_WIDTHS,
        moneyCols: COMPANY_MONEY,
        riskCol: COMPANY_RISK_COL,
        autofilter: true,
      });
      const wsManualErch = styleSheet(XLSX.utils.aoa_to_sheet(manualErchAoA), {
        headerRows: 1,
        widths: COMPANY_WIDTHS,
        moneyCols: COMPANY_MONEY,
        riskCol: COMPANY_RISK_COL,
        autofilter: true,
      });
      const wsManualIncomplete = styleSheet(
        XLSX.utils.aoa_to_sheet(manualIncompleteAoA),
        {
          headerRows: 1,
          widths: COMPANY_WIDTHS,
          moneyCols: COMPANY_MONEY,
          riskCol: COMPANY_RISK_COL,
          autofilter: true,
        }
      );
      const wsQueue = styleSheet(XLSX.utils.aoa_to_sheet(queueAoA), {
        headerRows: 1,
        widths: [12, 28, 22, 16, 16, 42, 56],
        moneyCols: [3, 4],
        autofilter: true,
      });

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, wsSummary, "Сводка");
      XLSX.utils.book_append_sheet(wb, wsQueue, "К исполнению");
      XLSX.utils.book_append_sheet(wb, wsCompanies, "Проверка компаний");
      XLSX.utils.book_append_sheet(wb, wsContracts, "Договоры с расчётом");
      XLSX.utils.book_append_sheet(wb, wsRisks, "Риски");
      XLSX.utils.book_append_sheet(wb, wsManualErch, "ЕРЧ");
      XLSX.utils.book_append_sheet(wb, wsManualIncomplete, "Данные неполные");

      const fname = `СРО_сверка_договоры_лимиты_${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}.xlsx`;
      XLSX.writeFile(wb, fname);
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
        "Расчёт обязательств": 300000000,
        "Дата расчёта размера обязательств": "31.01.2025",
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
        Контрагент: "ООО ДельтаМост",
        ИНН: "7706789012",
        "Состояние права": "Приостановлено",
        "Уровень ВВ": "до 90 млн руб.",
        "Уровень ОДО": "до 90 млн руб.",
        "Расчёт обязательств": 240000000,
        "Периоды приостановок": "01.11.2024 — 15.12.2024; 01.03.2025 — открыто",
      },
      {
        Контрагент: "ООО Еpsilon",
        ИНН: "7705678901",
        "Состояние права": "Действует",
        "Уровень ВВ": "до 90 млн руб.",
        "Уровень ОДО": "до 90 млн руб.",
        "Расчёт обязательств": 10000000,
        "Дата расчёта размера обязательств": "01.04.2025",
        "Дата регистрации в реестре СРО": "01.01.2025",
      },
      {
        Контрагент: "ООО ЗетаСервис",
        ИНН: "7707890123",
        "Состояние права": "Действует",
        "Уровень ВВ": "до 90 млн руб.",
        "Уровень ОДО": "до 90 млн руб.",
        "Расчёт обязательств": 0,
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
        Контрагент: "ООО ДельтаМост",
        ИНН: "7706789012",
        "Номер договора": "D-0",
        "Дата заключения": "01.02.2025",
        "Стоимость, принятая СРО к учёту": 10000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "прямой",
      },
      {
        Контрагент: "ООО ДельтаМост",
        ИНН: "7706789012",
        "Номер договора": "D-1",
        "Дата заключения": "10.11.2024",
        "Стоимость, принятая СРО к учёту": 40000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "44-ФЗ",
      },
      {
        Контрагент: "ООО ДельтаМост",
        ИНН: "7706789012",
        "Номер договора": "D-2",
        "Дата заключения": "20.04.2025",
        "Стоимость, принятая СРО к учёту": 200000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "223-ФЗ",
      },
      {
        Контрагент: "ООО Еpsilon",
        ИНН: "7705678901",
        "Номер договора": "E-0",
        "Дата заключения": "15.06.2024",
        "Стоимость, принятая СРО к учёту": 120000000,
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
      {
        Контрагент: "ООО ЗетаСервис",
        ИНН: "7707890123",
        "Номер договора": "Z-1",
        "Дата заключения": "10.02.2025",
        "Стоимость, принятая СРО к учёту": 8000000,
        "Стоимость принятых работ": 0,
        "Вид закупки": "",
      },
    ];

    membersInput.addEventListener("change", async (e) => {
      clearError();
      try {
        state.membersRows = await readTable(e.target.files[0]);
        state.membersMeta = null;
        state.membersSource = "file";
        onMembersFromFile(e.target.files[0], state.membersRows.length);
        updateReady();
      } catch (err) {
        showError("Не удалось прочитать реестр членов: " + err.message);
      }
    });

    contractsInput.addEventListener("change", async (e) => {
      clearError();
      const file = e.target.files[0];
      try {
        state.contractsRows = await readTable(file);
        markFilled("contracts", file.name);
        updateReady();
      } catch (err) {
        showError("Не удалось прочитать договоры: " + err.message);
      }
    });

    runBtn.addEventListener("click", run);
    exportBtn.addEventListener("click", exportExcel);
    sampleBtn.addEventListener("click", () => {
      state.membersRows = SAMPLE_MEMBERS;
      state.contractsRows = SAMPLE_CONTRACTS;
      state.membersMeta = null;
      state.membersSource = "sample";
      membersInput.value = "";
      contractsInput.value = "";
      markFilled("members", null);
      markFilled("contracts", "демо · договоры");
      onMembersFromSample(SAMPLE_MEMBERS.length);
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

    const resultsTable = document.querySelector("#results .table-wrap table");
    if (resultsTable) {
      resultsTable.addEventListener("click", (e) => {
        const tr = e.target.closest("tr[data-inn]");
        if (!tr || !resultsTable.contains(tr)) return;
        openInnCard(tr.dataset.inn);
      });
      resultsTable.addEventListener("keydown", (e) => {
        if (e.key !== "Enter" && e.key !== " ") return;
        const tr = e.target.closest("tr[data-inn]");
        if (!tr || e.target !== tr) return;
        e.preventDefault();
        openInnCard(tr.dataset.inn);
      });
    }

    if (innDrawer.closeBtn) innDrawer.closeBtn.addEventListener("click", closeInnCard);
    if (innDrawer.copyBtn) innDrawer.copyBtn.addEventListener("click", copyInnLetter);
    if (innDrawer.overlay) {
      innDrawer.overlay.addEventListener("click", (e) => {
        if (e.target === innDrawer.overlay) closeInnCard();
      });
    }
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") closeInnCard();
    });

    /* --- Шаг 1: реестр членов из НОСТРОЙ через локальный helper ------------- */

    const HELPER_ORIGIN = "http://127.0.0.1:8765";
    const SRO_ID_KEY = "sro-auditor:sroId";
    const POLL_MS = 700;
    const HELPER_CMD = "python tools/sro_server.py";

    const sroInput = document.getElementById("sroId");
    const fetchBtn = document.getElementById("fetchBtn");
    const sroTitle = document.getElementById("sroTitle");
    const registryStatus = document.getElementById("registryStatus");
    const registryStatusText = document.getElementById("registryStatusText");
    const registryProgress = document.getElementById("registryProgress");
    const registryBar = document.getElementById("registryBar");
    const registryMeta = document.getElementById("registryMeta");
    const registryActions = document.getElementById("registryActions");
    const registryError = document.getElementById("registryError");
    const membersFallback = document.getElementById("membersFallback");
    const membersPreview = document.getElementById("membersPreview");
    const membersPreviewBtn = document.getElementById("membersPreviewBtn");
    const membersExcelBtn = document.getElementById("membersExcelBtn");
    const membersPreviewCount = document.getElementById("membersPreviewCount");
    const membersPreviewFilter = document.getElementById("membersPreviewFilter");
    const membersPreviewHead = document.getElementById("membersPreviewHead");
    const membersPreviewBody = document.getElementById("membersPreviewBody");
    const membersPreviewEmpty = document.getElementById("membersPreviewEmpty");

    const MEMBERS_PREVIEW_COLS = [
      ["ИНН", "inn"],
      ["Наименование", "name"],
      ["Состояние права", "right"],
      ["Уровень ВВ", "vv"],
      ["Уровень ОДО", "odo"],
      ["Дата регистрации", "reg"],
      ["Приостановки", "susp"],
      ["Записей по ИНН", "dup"],
    ];

    let membersPreviewOpen = false;

    const fileCards = {
      members: { card: document.getElementById("membersCard"), name: document.getElementById("membersName") },
      contracts: { card: document.getElementById("contractsCard"), name: document.getElementById("contractsName") },
    };

    /** Отметить слот файла: name = имя файла или null, чтобы сбросить слот. */
    function markFilled(kind, name) {
      const { card, name: nameEl } = fileCards[kind];
      card.classList.toggle("filled", !!name);
      if (name) card.dataset.ready = "true";
      else delete card.dataset.ready;
      nameEl.textContent = name || "файл не выбран";
    }

    // file:// умеет дотянуться до helper, https (GitHub Pages) — нет (mixed content).
    const apiBase =
      location.protocol === "http:" ? "" : location.protocol === "file:" ? HELPER_ORIGIN : null;

    let pollTimer = null;

    function esc(s) {
      return String(s ?? "").replace(/[&<>"]/g, (c) => `&${{ "&": "amp", "<": "lt", ">": "gt", '"': "quot" }[c]};`);
    }

    function setRegStatus(kind, html) {
      registryStatus.dataset.state = kind;
      registryStatusText.innerHTML = html;
    }

    function showRegError(msg) {
      registryError.style.display = "block";
      registryError.textContent = msg;
    }
    function clearRegError() {
      registryError.style.display = "none";
      registryError.textContent = "";
    }

    function setProgress(done, total) {
      registryProgress.classList.remove("hidden");
      if (total > 0) {
        registryProgress.dataset.indeterminate = "false";
        registryBar.style.width = Math.round((done / total) * 100) + "%";
        registryProgress.setAttribute("aria-valuenow", String(Math.round((done / total) * 100)));
      } else {
        registryProgress.dataset.indeterminate = "true";
        registryBar.style.width = "";
        registryProgress.removeAttribute("aria-valuenow");
      }
    }

    function hideProgress() {
      registryProgress.classList.add("hidden");
      registryProgress.dataset.indeterminate = "false";
      registryBar.style.width = "0";
    }

    function fmtExportedAt(iso) {
      const d = new Date(iso);
      if (Number.isNaN(d.getTime())) return String(iso || "—");
      const time = `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
      return `${fmtDate(startOfDay(d))} ${time}`;
    }

    function sroDisplayName(payload) {
      if (!payload) return { full: "", short: "" };
      const full = String(payload.sro_name || "").trim();
      const short = String(payload.sro_short_name || "").trim();
      return { full, short };
    }

    function renderSroTitle(payload) {
      const { full, short } = sroDisplayName(payload);
      if (!full && !short) {
        sroTitle.classList.add("hidden");
        sroTitle.textContent = "";
        return;
      }
      const primary = full || short;
      // Короткое — только если реально короче полного (у части СРО «short» длиннее).
      const secondary =
        full && short && short !== full && short.length < full.length ? short : "";
      sroTitle.innerHTML =
        esc(primary) +
        (secondary ? `<span class="sro-short">${esc(secondary)}</span>` : "");
      sroTitle.classList.remove("hidden");
    }

    function clearSroTitle() {
      sroTitle.classList.add("hidden");
      sroTitle.textContent = "";
    }

    function rightClass(right) {
      const s = String(right || "").toLowerCase();
      if (s.includes("приостанов")) return "right-suspended";
      if (s.includes("прекращ")) return "right-terminated";
      return "";
    }

    function membersPreviewRows(rows) {
      return (rows || []).map((r) => ({
        inn: String(r["ИНН"] ?? ""),
        name: String(r["Наименование"] ?? ""),
        right: String(r["Состояние права"] ?? ""),
        vv: String(r["Уровень ВВ"] ?? ""),
        odo: String(r["Уровень ОДО"] ?? ""),
        reg: String(r["Дата регистрации в реестре СРО"] ?? ""),
        susp: String(r["Периоды приостановок"] ?? ""),
        dup: r["Записей по ИНН"] > 1 ? String(r["Записей по ИНН"]) : "",
      }));
    }

    function setMembersPreviewOpen(open) {
      membersPreviewOpen = !!open;
      membersPreview.classList.toggle("hidden", !membersPreviewOpen);
      membersPreviewBtn.setAttribute("aria-expanded", membersPreviewOpen ? "true" : "false");
      membersPreviewBtn.textContent = membersPreviewOpen ? "Скрыть реестр" : "Показать реестр";
      if (membersPreviewOpen) {
        renderMembersPreview();
        membersPreviewFilter.focus();
      }
    }

    function hideMembersPreview() {
      membersPreviewOpen = false;
      membersPreview.classList.add("hidden");
      registryActions.classList.add("hidden");
      membersPreviewBtn.setAttribute("aria-expanded", "false");
      membersPreviewBtn.textContent = "Показать реестр";
      membersPreviewFilter.value = "";
      membersPreviewHead.innerHTML = "";
      membersPreviewBody.innerHTML = "";
      membersPreviewEmpty.classList.add("hidden");
    }

    function renderMembersPreview() {
      const rows = membersPreviewRows(state.membersRows);
      const q = String(membersPreviewFilter.value || "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ");
      const filtered = !q
        ? rows
        : rows.filter((r) =>
            [r.inn, r.name, r.right, r.vv, r.odo, r.reg, r.susp].join(" ").toLowerCase().includes(q)
          );

      membersPreviewHead.innerHTML =
        "<tr>" +
        MEMBERS_PREVIEW_COLS.map(([label, key]) => {
          const cls = key === "inn" || key === "dup" ? ' class="num"' : "";
          return `<th${cls}>${esc(label)}</th>`;
        }).join("") +
        "</tr>";

      if (!filtered.length) {
        membersPreviewBody.innerHTML = "";
        membersPreviewEmpty.classList.remove("hidden");
      } else {
        membersPreviewEmpty.classList.add("hidden");
        membersPreviewBody.innerHTML = filtered
          .map((r) => {
            const cells = MEMBERS_PREVIEW_COLS.map(([, key]) => {
              const val = r[key] || "—";
              if (key === "inn" || key === "dup") {
                return `<td class="num">${esc(val)}</td>`;
              }
              if (key === "name" || key === "susp") {
                return `<td class="wrap">${esc(val)}</td>`;
              }
              if (key === "right") {
                const cls = rightClass(r.right);
                return `<td class="${cls}">${esc(val)}</td>`;
              }
              return `<td>${esc(val)}</td>`;
            }).join("");
            return `<tr>${cells}</tr>`;
          })
          .join("");
      }

      const total = rows.length;
      const shown = filtered.length;
      membersPreviewCount.textContent =
        q && shown !== total ? `${shown} из ${total}` : `${total} строк`;
    }

    function showMembersPreviewControls(rowsCount) {
      registryActions.classList.remove("hidden");
      membersPreviewCount.textContent = `${rowsCount} строк`;
      membersExcelBtn.disabled = !rowsCount;
      if (membersPreviewOpen) renderMembersPreview();
    }

    const MEMBERS_EXPORT_COLS = [
      "ИНН",
      "Наименование",
      "Полное наименование",
      "Состояние права",
      "Уровень ВВ",
      "Уровень ОДО",
      "Предел ВВ, ₽",
      "Предел ОДО, ₽",
      "Расчёт обязательств",
      "Дата расчёта размера обязательств",
      "Дата регистрации в реестре СРО",
      "Периоды приостановок",
      "Периоды ограничения по конкурентным",
      "Реестровый номер",
      "Тип члена",
      "Регион",
      "Формулировка ВВ (НОСТРОЙ)",
      "Формулировка ОДО (НОСТРОЙ)",
      "Взнос КФ ВВ",
      "Взнос КФ ОДО",
      "Соответствие требованиям",
      "Дата прекращения членства",
      "Основание прекращения",
      "Решений по праву",
      "История решений",
      "Обновлено в НОСТРОЙ",
      "Ссылка НОСТРОЙ",
      "Записей по ИНН",
    ];

    const MEMBERS_EXPORT_MONEY = new Set([
      "Расчёт обязательств",
      "Взнос КФ ВВ",
      "Взнос КФ ОДО",
    ]);

    function membersExportColumns(rows) {
      if (!rows || !rows.length) return MEMBERS_EXPORT_COLS.slice();
      const present = new Set();
      rows.forEach((r) => Object.keys(r).forEach((k) => present.add(k)));
      const ordered = MEMBERS_EXPORT_COLS.filter((k) => present.has(k));
      const extras = [...present].filter((k) => !MEMBERS_EXPORT_COLS.includes(k));
      // Контрагент из демо/ручных файлов — в начало рядом с наименованием.
      if (extras.includes("Контрагент") && !ordered.includes("Наименование")) {
        ordered.unshift("Контрагент");
        extras.splice(extras.indexOf("Контрагент"), 1);
      }
      return ordered.concat(extras);
    }

    function exportMembersExcel() {
      const rows = state.membersRows;
      if (!rows || !rows.length) {
        showRegError("Нет загруженного реестра членов для выгрузки.");
        return;
      }
      if (typeof XLSX === "undefined") {
        showRegError("Библиотека Excel не загрузилась.");
        return;
      }
      clearRegError();

      const cols = membersExportColumns(rows);
      const moneyCols = cols
        .map((c, i) => (MEMBERS_EXPORT_MONEY.has(c) ? i : -1))
        .filter((i) => i >= 0);
      const aoa = [cols].concat(
        rows.map((r) =>
          cols.map((c) => {
            const v = r[c];
            if (v == null || v === "") return "";
            if (MEMBERS_EXPORT_MONEY.has(c) && typeof v === "number") return v;
            if (typeof v === "number" || typeof v === "boolean") return v;
            return String(v);
          })
        )
      );

      const widths = cols.map((c) => {
        if (c === "ИНН" || c === "Записей по ИНН") return 14;
        if (c.includes("наименование") || c === "Наименование" || c === "Контрагент") return 36;
        if (c.includes("Период") || c.includes("История") || c.includes("Основание")) return 28;
        if (c.includes("Ссылка")) return 32;
        if (c.includes("Формулировка")) return 28;
        return 16;
      });

      const ws = styleSheet(XLSX.utils.aoa_to_sheet(aoa), {
        headerRows: 1,
        widths,
        moneyCols,
        autofilter: true,
      });

      const meta = state.membersMeta;
      const summaryAoA = [
        ["Показатель", "Значение"],
        ["СРО", meta ? meta.sro_id : ""],
        ["Название СРО", meta ? meta.sro_name || meta.sro_short_name || "" : ""],
        ["Источник", meta ? meta.source || "НОСТРОЙ" : state.membersSource === "file" ? "файл" : "демо"],
        ["Дата выгрузки", meta ? fmtExportedAt(meta.exported_at) : ""],
        ["Членов в файле", rows.length],
      ];
      if (meta && meta.stats) {
        const s = meta.stats;
        summaryAoA.push(
          ["С приостановками", s.with_suspensions ?? ""],
          ["Право действует", s.right_active ?? ""],
          ["Право приостановлено", s.right_suspended ?? ""],
          ["Право прекращено", s.right_terminated ?? ""],
          ["Без уровня ОДО", s.no_odo_level ?? ""],
          ["ИНН с 2+ записями", s.duplicate_inn ?? ""]
        );
      }
      const wsSummary = XLSX.utils.aoa_to_sheet(summaryAoA);
      const sumRange = XLSX.utils.decode_range(wsSummary["!ref"]);
      for (let R = sumRange.s.r; R <= sumRange.e.r; R++) {
        const a = XLSX.utils.encode_cell({ r: R, c: 0 });
        const b = XLSX.utils.encode_cell({ r: R, c: 1 });
        if (wsSummary[a]) {
          wsSummary[a].s = R === 0 ? { ...XL.header } : { ...XL.label, border: XL.cell.border };
        }
        if (wsSummary[b]) {
          wsSummary[b].s = R === 0 ? { ...XL.header } : { ...XL.cell };
        }
      }
      wsSummary["!cols"] = [{ wch: 28 }, { wch: 56 }];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, wsSummary, "Сводка");
      XLSX.utils.book_append_sheet(wb, ws, "Реестр членов");

      const today = new Date();
      const ymd =
        `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-` +
        String(today.getDate()).padStart(2, "0");
      const sroPart = meta && meta.sro_id != null ? `_СРО_${meta.sro_id}` : "";
      XLSX.writeFile(wb, `реестр_членов${sroPart}_${ymd}.xlsx`);
    }

    function renderRegistryMeta(payload) {
      const s = payload.stats || {};
      const items = [
        ["Членов загружено", s.members, ""],
        ["Дата выгрузки", fmtExportedAt(payload.exported_at), ""],
        ["С приостановками", s.with_suspensions, s.with_suspensions ? "warn" : ""],
        ["Право приостановлено", s.right_suspended, s.right_suspended ? "crit" : ""],
        ["Право прекращено", s.right_terminated, ""],
        ["Без уровня ОДО", s.no_odo_level, s.no_odo_level ? "warn" : ""],
      ];
      if (s.duplicate_inn) items.push(["ИНН с 2+ записями", s.duplicate_inn, "warn"]);
      if (s.cards_failed) items.push(["Карточек не скачано", s.cards_failed, "crit"]);
      if (s.unparsed_levels) items.push(["Уровень не распознан", s.unparsed_levels, "crit"]);

      registryMeta.innerHTML = items
        .map(
          ([label, value, cls]) =>
            `<div><dt>${esc(label)}</dt><dd class="${cls}">${esc(value ?? "—")}</dd></div>`
        )
        .join("");
      registryMeta.classList.remove("hidden");
    }

    function applyRegistry(payload) {
      state.membersRows = payload.members;
      state.membersMeta = payload;
      state.membersSource = "registry";
      membersInput.value = "";
      markFilled("members", null);
      renderSroTitle(payload);
      renderRegistryMeta(payload);
      showMembersPreviewControls(payload.members.length);
      updateReady();
    }

    function onMembersFromFile(file, rows) {
      hideProgress();
      registryMeta.classList.add("hidden");
      clearSroTitle();
      markFilled("members", file.name);
      setRegStatus("ok", `Реестр членов из файла <code>${esc(file.name)}</code>: ${rows} строк.`);
      showMembersPreviewControls(rows);
      if (membersPreviewOpen) renderMembersPreview();
    }

    function onMembersFromSample(rows) {
      hideProgress();
      registryMeta.classList.add("hidden");
      clearSroTitle();
      setRegStatus("ok", `Демо-данные: ${rows} членов. НОСТРОЙ не запрашивался.`);
      showMembersPreviewControls(rows);
      if (membersPreviewOpen) renderMembersPreview();
    }

    function sroIdValue() {
      const digits = String(sroInput.value || "").replace(/\D/g, "");
      return digits;
    }

    async function api(path, init) {
      const resp = await fetch(apiBase + path, init);
      return resp;
    }

    async function helperAlive() {
      if (apiBase === null) return false;
      try {
        const resp = await api("/api/health");
        if (!resp.ok) return false;
        const body = await resp.json();
        return body.app === "sro-auditor-helper";
      } catch {
        return false;
      }
    }

    function helperMissingStatus() {
      fetchBtn.disabled = true;
      hideProgress();
      registryMeta.classList.add("hidden");
      hideMembersPreview();
      clearSroTitle();
      setRegStatus(
        "warn",
        `Локальный helper не запущен. В папке проекта: <code>${esc(HELPER_CMD)}</code> — ` +
          "и откройте чекер по адресу из консоли. Либо загрузите реестр членов вручную."
      );
      membersFallback.open = true;
    }

    async function loadCache(sroId, opts) {
      const quiet = !!(opts && opts.quiet);
      try {
        const resp = await api(`/api/nostroy/${sroId}`);
        if (resp.status === 404) {
          registryMeta.classList.add("hidden");
          hideMembersPreview();
          clearSroTitle();
          setRegStatus(
            "warn",
            `Локального кэша по СРО ${esc(sroId)} нет. Нажмите «Обновить с НОСТРОЙ».`
          );
          return false;
        }
        if (!resp.ok) throw new Error(`helper ответил ${resp.status}`);
        const payload = await resp.json();
        applyRegistry(payload);
        const { full, short } = sroDisplayName(payload);
        const label =
          short && (!full || short.length < full.length)
            ? short
            : full && full.length <= 96
              ? full
              : short || "";
        const nameBit = label ? ` · ${esc(label)}` : "";
        setRegStatus(
          "ok",
          `СРО ${esc(payload.sro_id)}${nameBit}: ${payload.members.length} членов из локального кэша.`
        );
        return true;
      } catch (err) {
        if (!quiet) showRegError("Не удалось прочитать кэш: " + err.message);
        return false;
      }
    }

    function pollProgress(sroId) {
      clearTimeout(pollTimer);
      pollTimer = setTimeout(async () => {
        let job;
        try {
          const resp = await api(`/api/nostroy/${sroId}/progress`);
          job = await resp.json();
        } catch {
          finishRefresh();
          setRegStatus("error", "Helper перестал отвечать. Проверьте окно с сервером.");
          membersFallback.open = true;
          return;
        }
        if (job.state === "running") {
          setProgress(job.done || 0, job.total || 0);
          setRegStatus("work", esc(job.message || "выгрузка…"));
          pollProgress(sroId);
          return;
        }
        finishRefresh();
        if (job.state === "error") {
          setRegStatus("error", `Выгрузка не удалась: ${esc(job.error || job.message)}`);
          showRegError(
            "Выгрузка не удалась. Повторите позже или загрузите реестр членов вручную " +
              "(файл cache/nostroy/<номер>/members.csv остаётся от прошлой выгрузки)."
          );
          membersFallback.open = true;
          return;
        }
        await loadCache(sroId, {});
      }, POLL_MS);
    }

    function finishRefresh() {
      clearTimeout(pollTimer);
      hideProgress();
      fetchBtn.disabled = false;
      fetchBtn.textContent = "Обновить с НОСТРОЙ";
    }

    async function refreshFromNostroy() {
      const sroId = sroIdValue();
      clearRegError();
      if (!sroId) {
        setRegStatus("error", "Укажите номер СРО — только цифры, как в адресе реестра.");
        sroInput.focus();
        return;
      }
      localStorage.setItem(SRO_ID_KEY, sroId);
      fetchBtn.disabled = true;
      fetchBtn.textContent = "Обновление…";
      setProgress(0, 0);
      setRegStatus("work", "подключение к НОСТРОЙ");
      try {
        const resp = await api(`/api/nostroy/${sroId}/refresh`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ force: false }),
        });
        const body = await resp.json();
        if (!resp.ok && resp.status !== 409) throw new Error(body.error || `helper ответил ${resp.status}`);
        if (resp.status === 409) {
          setRegStatus("work", esc(body.message));
        }
        pollProgress(sroId);
      } catch (err) {
        finishRefresh();
        setRegStatus("error", "Helper не отвечает. Запущен ли " + `<code>${esc(HELPER_CMD)}</code>?`);
        showRegError(err.message);
        membersFallback.open = true;
      }
    }

    fetchBtn.addEventListener("click", refreshFromNostroy);
    membersPreviewBtn.addEventListener("click", () => {
      if (!state.membersRows || !state.membersRows.length) return;
      setMembersPreviewOpen(!membersPreviewOpen);
    });
    membersExcelBtn.addEventListener("click", exportMembersExcel);
    membersPreviewFilter.addEventListener("input", () => {
      if (membersPreviewOpen) renderMembersPreview();
    });
    sroInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        refreshFromNostroy();
      }
    });
    sroInput.addEventListener("change", async () => {
      const sroId = sroIdValue();
      if (!sroId || fetchBtn.disabled) return;
      localStorage.setItem(SRO_ID_KEY, sroId);
      if (state.membersSource === "registry") {
        state.membersRows = null;
        state.membersMeta = null;
        state.membersSource = null;
        registryMeta.classList.add("hidden");
        hideMembersPreview();
        clearSroTitle();
        updateReady();
      }
      await loadCache(sroId, { quiet: true });
    });

    (async function initRegistry() {
      const saved = localStorage.getItem(SRO_ID_KEY);
      if (saved) sroInput.value = saved;
      if (!(await helperAlive())) {
        helperMissingStatus();
        return;
      }
      hideProgress();
      await loadCache(sroIdValue(), { quiet: true });
    })();
