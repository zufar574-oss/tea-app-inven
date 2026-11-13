/* data.js — загрузка и фильтрация базы МойСклад.
 * Фаза 2: parseWorkbook + filterRows (7 правил SPEC).
 * Фаза 3: search().
 * Фаза 5: buildAct() — формирование строк акта для экспорта в МойСклад. */
(function () {
  "use strict";

  // Белый список колонок (точные имена, SPEC). Все 14 обязательны.
  var COL = {
    uuid: "UUID",
    name: "Наименование",
    article: "Артикул",
    code: "Код",
    unit: "Единица измерения",
    ean13: "Штрихкод EAN13",
    ean8: "Штрихкод EAN8",
    code128: "Штрихкод Code128",
    upc: "Штрихкод UPC",
    gtin: "Штрихкод GTIN",
    weighted: "Весовой товар",
    type: "Тип",
    archived: "Архивный",
    groups: "Группы"
  };

  var ALL_COLS = [
    COL.uuid, COL.name, COL.article, COL.code, COL.unit,
    COL.ean13, COL.ean8, COL.code128, COL.upc, COL.gtin,
    COL.weighted, COL.type, COL.archived, COL.groups
  ];

  // Колонки акта — ПОЛНЫЙ шаблон МойСклада (36 колонок, порядок как в Пример.xlsx).
  // МойСклад при «Искать: Автоматически» узнаёт файл именно по этой структуре:
  // видит 36 знакомых колонок с родными именами и автосопоставляет их.
  // Файл с меньшим числом колонок МойСклад не узнаёт и ставит «Не загружать».
  // Заполняем те колонки, что у нас есть из базы; остальные оставляем пустыми.
  var ACT_COLS = [
    "UUID",
    "Тип",
    "Группы",
    "Код",
    "Наименование",
    "Внешний код",
    "Артикул",
    "Единица измерения",
    "Вес",
    "Объем",
    "Весовой товар",
    "Разливной товар",
    "Цена продажи",
    "Валюта (Цена продажи)",
    "Закупочная цена",
    "Валюта (Закупочная цена)",
    "Неснижаемый остаток",
    "Штрихкод EAN13",
    "Штрихкод EAN8",
    "Штрихкод Code128",
    "Штрихкод UPC",
    "Описание",
    "Минимальная цена",
    "Валюта (Минимальная цена)",
    "Страна",
    "НДС",
    "Поставщик",
    "Количество",
    "Изображение",
    "ГТД",
    "Код товара модификации",
    "Характеристика:эко",
    "Характеристика:Цвет",
    "Упаковка:мешок",
    "Алкогольная продукция",
    "Содержит акцизную марку"
  ];

  // Поле количества зависит от режима: в документе Приемка оно называется
  // «Количество», в документе Инвентаризация — «Фактический остаток» (имя
  // поля в экране импорта МойСклада). С родным именем поля МойСклад
  // сопоставляет колонку файла автоматически, без ручного выбора.
  var QTY_COL = {
    reception: "Количество",
    inventory: "Фактический остаток"
  };

  function asText(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim();
  }

  /* Определить ед. изм. по правилам SPEC:
   *   г/шт → как есть; кг → null (исключить); пусто → пересчёт по весовому
   *   (да→г, нет→шт, иное→null); прочее → null.
   * Возвращает "г" | "шт" | null. */
  function resolveUnit(ei, weighted) {
    if (ei === "г" || ei === "шт") return ei;
    if (ei === "кг") return null; // исключаем целиком
    if (ei === "") {
      if (weighted === "да") return "г";
      if (weighted === "нет") return "шт";
      return null; // весовой не да/нет — исключаем
    }
    return null; // любое иное значение (включая *)
  }

  window.TeaApp.data = {

    COL: COL,
    ALL_COLS: ALL_COLS,

    /* Прочитать .xlsx из ArrayBuffer.
     * Возвращает { ok:true, rows } или { ok:false, error }.
     * error: "not_excel" — файл не читается как Excel
     *        "empty"     — нет строк данных
     *        "no_column" — не хватает колонки белого списка (missing — имя) */
    parseWorkbook: function (arrayBuffer) {
      var wb;
      try {
        wb = XLSX.read(arrayBuffer, { type: "array" });
      } catch (e) {
        return { ok: false, error: "not_excel" };
      }
      if (!wb.SheetNames || !wb.SheetNames.length) return { ok: false, error: "not_excel" };

      var ws = wb.Sheets[wb.SheetNames[0]];
      var rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!rows.length) return { ok: false, error: "empty" };

      // Проверка всех 14 колонок белого списка
      var headers = Object.keys(rows[0]);
      for (var i = 0; i < ALL_COLS.length; i++) {
        if (headers.indexOf(ALL_COLS[i]) === -1) {
          return { ok: false, error: "no_column", missing: ALL_COLS[i] };
        }
      }
      return { ok: true, rows: rows };
    },

    /* Применить фильтры SPEC к строкам. Возвращает { items, report }.
     * items — массив товаров для базы; report — счётчики для отчёта. */
    filterRows: function (rows) {
      var items = [];
      var report = {
        loaded: 0,
        servicesSkipped: 0,
        withoutUuid: 0,
        kgExcluded: 0,
        archivedHidden: 0
      };

      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var uuid = asText(r[COL.uuid]);
        if (!uuid) { report.withoutUuid++; continue; }

        var type = asText(r[COL.type]);
        if (type !== "Товар") { report.servicesSkipped++; continue; }

        var ei = asText(r[COL.unit]);
        var weighted = asText(r[COL.weighted]);
        var unit = resolveUnit(ei, weighted);
        if (ei === "кг") { report.kgExcluded++; continue; } // кг считаем отдельно
        if (unit === null) continue; // иные/непонятные — просто исключаем

        var archived = asText(r[COL.archived]) === "да";
        if (archived) report.archivedHidden++;

        items.push({
          uuid: uuid,
          name: asText(r[COL.name]),
          article: asText(r[COL.article]),
          code: asText(r[COL.code]), // текст, ведущие нули важны
          unit: unit,
          barcodes: {
            ean13: asText(r[COL.ean13]),
            ean8: asText(r[COL.ean8]),
            code128: asText(r[COL.code128]),
            upc: asText(r[COL.upc]),
            gtin: asText(r[COL.gtin])
          },
          weighted: weighted,
          archived: archived,
          groups: asText(r[COL.groups])
        });
        report.loaded++;
      }

      return { items: items, report: report };
    },

    /* Форматирование отчёта в строку для модала. */
    formatReport: function (r) {
      return "Загружено " + r.loaded + " товаров" +
             " · услуг пропущено " + r.servicesSkipped +
             " · без UUID — " + r.withoutUuid +
             " · «кг» исключено " + r.kgExcluded +
             " · архивных скрыто " + r.archivedHidden;
    },

    // Кэш базы для поиска. Устанавливается из main.js после загрузки/восстановления.
    _base: null,
    setBase: function (base) { this._base = base || null; },
    getBase: function () { return this._base; },

    /* Нормализованное имя новой позиции — ключ для суммирования одинаковых.
     * trim + схлопывание внутренних пробелов + lowercase. Чувствительность к
     * регистру убираем: «Подарок» и «подарок» — одна позиция. */
    normalizeNewName: function (name) {
      return (name || "").trim().replace(/\s+/g, " ").toLowerCase();
    },

    /* Поиск по базе. SPEC п.2.
     * query — строка; opts.showArchived — показывать архивные.
     * Возвращает массив товаров, отсортированный:
     *   1) точные совпадения по штрихкоду (EAN13→EAN8→Code128→UPC→GTIN)
     *   2) текстовые совпадения (наименование/артикул/код/штрихкод содержит)
     *      по алфавиту наименования.
     * При < 2 символов — пустой массив (дропдаун не показываем). */
    search: function (query, opts) {
      opts = opts || {};
      var showArchived = !!opts.showArchived;
      var q = (query || "").trim().toLowerCase();
      if (q.length < 2) return [];

      var items = this._base && this._base.items ? this._base.items : [];
      var pool = [];
      for (var i = 0; i < items.length; i++) {
        if (!showArchived && items[i].archived) continue;
        pool.push(items[i]);
      }

      var BC_ORDER = ["ean13", "ean8", "code128", "upc", "gtin"];
      var barcodeHits = []; // {item, priority}
      var textHits = [];

      for (var j = 0; j < pool.length; j++) {
        var it = pool[j];
        var bc = it.barcodes || {};

        // Точное совпадение по штрихкоду — высший приоритет
        var bcMatched = false;
        for (var b = 0; b < BC_ORDER.length; b++) {
          var v = bc[BC_ORDER[b]];
          if (v && v.toLowerCase() === q) {
            barcodeHits.push({ item: it, priority: b });
            bcMatched = true;
            break;
          }
        }
        if (bcMatched) continue;

        // Текстовое вхождение: наименование, артикул, код, любой штрихкод
        var inName = it.name.toLowerCase().indexOf(q) !== -1;
        var inCode = it.code.toLowerCase().indexOf(q) !== -1;
        var inArticle = it.article.toLowerCase().indexOf(q) !== -1;
        var inAnyBc = false;
        if (!inName && !inCode && !inArticle) {
          for (var c = 0; c < BC_ORDER.length; c++) {
            var v2 = bc[BC_ORDER[c]];
            if (v2 && v2.toLowerCase().indexOf(q) !== -1) { inAnyBc = true; break; }
          }
        }
        if (inName || inCode || inArticle || inAnyBc) {
          textHits.push(it);
        }
      }

      // Штрихкод-совпадения — по приоритету; текстовые — по алфавиту
      barcodeHits.sort(function (a, b) { return a.priority - b.priority; });
      textHits.sort(function (a, b) {
        return a.name.localeCompare(b.name, "ru");
      });

      var result = [];
      for (var k = 0; k < barcodeHits.length; k++) result.push(barcodeHits[k].item);
      for (var m = 0; m < textHits.length; m++) result.push(textHits[m]);
      return result;
    },

    ACT_COLS: ACT_COLS,

    /* Имя файла акта: <Режим>_<DD-MM-YYYY>_<HH-MM>.xlsx.
     * Время через дефис (двоеточие недопустимо в именах на Windows). */
    formatActFileName: function (mode) {
      var prefix = mode === "inventory" ? "Инвентарка" : "Приемка";
      var d = new Date();
      function pad(n) { return (n < 10 ? "0" : "") + n; }
      var date = pad(d.getDate()) + "-" + pad(d.getMonth() + 1) + "-" + d.getFullYear();
      var time = pad(d.getHours()) + "-" + pad(d.getMinutes());
      return prefix + "_" + date + "_" + time + ".xlsx";
    },

    /* Сформировать строки акта из списка выбранного.
     * mode = "reception" | "inventory".
     *   reception: позиции с qty ≤ 0 в акт НЕ попадают (контрольный №6).
     *   inventory: попадают все, включая нули (инвентаризация фиксирует остаток).
     * Каждая строка содержит ВСЕ 36 колонок шаблона МойСклада (незаполненные — ""),
     * чтобы МойСклад узнал файл как «родной» и автосопоставил колонки.
     * Поле количества пишется под режим: приёмка — «Количество», инвентарка —
     * «Фактический остаток» (QTY_COL). Список колонок возвращается в cols.
     * Новые позиции: UUID пустой (МойСклад создаст), Тип «Товар», ед. изм. «шт». */
    buildAct: function (mode, list) {
      list = list || [];
      var qtyCol = QTY_COL[mode] || QTY_COL.reception;
      var rows = [];
      for (var i = 0; i < list.length; i++) {
        var it = list[i];
        if (mode === "reception" && (!it.qty || it.qty <= 0)) continue;

        var unit = it.unit || "";
        if (it.isNew && !(unit && unit.trim())) unit = "шт";

        var bc = it.barcodes || {};
        var row = {
          "UUID": it.uuid || "",
          "Тип": "Товар",
          "Группы": it.groups || "",
          "Код": it.code || "",
          "Наименование": it.name || "",
          "Внешний код": "",
          "Артикул": it.article || "",
          "Единица измерения": unit,
          "Вес": "",
          "Объем": "",
          "Весовой товар": it.weighted || "",
          "Разливной товар": "",
          "Цена продажи": "",
          "Валюта (Цена продажи)": "",
          "Закупочная цена": "",
          "Валюта (Закупочная цена)": "",
          "Неснижаемый остаток": "",
          "Штрихкод EAN13": bc.ean13 || "",
          "Штрихкод EAN8": bc.ean8 || "",
          "Штрихкод Code128": bc.code128 || "",
          "Штрихкод UPC": bc.upc || "",
          "Описание": "",
          "Минимальная цена": "",
          "Валюта (Минимальная цена)": "",
          "Страна": "",
          "НДС": "",
          "Поставщик": "",
          "Изображение": "",
          "ГТД": "",
          "Код товара модификации": "",
          "Характеристика:эко": "",
          "Характеристика:Цвет": "",
          "Упаковка:мешок": "",
          "Алкогольная продукция": "",
          "Содержит акцизную марку": ""
        };
        row[qtyCol] = it.qty || 0;
        rows.push(row);
      }
      // Колонки акта: 36 колонок шаблона, поле количества — под режим.
      var cols = ACT_COLS.slice();
      cols[ACT_COLS.indexOf("Количество")] = qtyCol;
      return { rows: rows, cols: cols, fileName: this.formatActFileName(mode) };
    },

    /* Перенос позиций приёмки в стартовый список инвентарки (SPEC раздел 5).
     * Берём только позиции с UUID (новые — без UUID — отбрасываем),
     * количества обнуляем (инвентарка считает фактический остаток заново).
     * Сохраняем все поля товара базы (groups/weighted/barcodes/article/code/unit),
     * чтобы акт сверки выгружался корректно. */
    buildInventoryFromReception: function (list) {
      list = list || [];
      var inv = [];
      for (var i = 0; i < list.length; i++) {
        var it = list[i];
        if (it.isNew) continue;          // новые позиции не переносим
        if (!it.uuid) continue;          // без UUID — не сверишь
        inv.push({
          uuid: it.uuid,
          name: it.name || "",
          code: it.code || "",
          article: it.article || "",
          unit: it.unit || "",
          groups: it.groups || "",
          weighted: it.weighted || "",
          barcodes: it.barcodes || {},
          qty: 0,                        // обнуляем — считаем факт заново
          isNew: false
        });
      }
      return inv;
    }
  };
})();
