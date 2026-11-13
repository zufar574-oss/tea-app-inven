/* main.js — связка модулей, обработчики, поток работы.
 * Фаза 1: заставка → Старт.
 * Фаза 2: загрузка базы.
 * Фаза 3: переход в режимы, поиск, тумблер «Архивные».
 * Фаза 4: окно количества, список выбранных, суммирование, очистка. */
(function () {
  "use strict";

  var SPLASH_MS = 1000;

  // Короткие ссылки на модули
  var storage = null, data = null, ui = null;

  var loading = false; // защита от повторного выбора файла во время загрузки

  var currentMode = null;        // "reception" | "inventory"
  var currentList = [];          // выбранные товары текущего режима
  var searchDebounce = null;     // таймер дебаунса поиска
  var DEBOUNCE_MS = 150;

  function $(id) { return document.getElementById(id); }

  /* ---------- Заставка ---------- */
  function showStart() {
    if ($("splash")) $("splash").classList.remove("is-active");
    if ($("start")) $("start").classList.add("is-active");
    initStart();
  }

  /* ---------- Оверлей / Модал ---------- */
  function showOverlay(text) {
    if (text) $("overlayText").textContent = text;
    $("overlay").classList.add("is-active");
  }
  function hideOverlay() {
    $("overlay").classList.remove("is-active");
  }

  function showModal(opts) {
    $("modalTitle").textContent = opts.title || "";
    $("modalText").textContent = opts.text || "";
    var actionsEl = $("modalActions");
    actionsEl.innerHTML = "";
    var acts = opts.actions || [{ label: "Понятно", primary: true }];
    acts.forEach(function (a) {
      var b = document.createElement("button");
      b.type = "button";
      b.className = "btn " + (a.primary ? "btn--primary" : "btn--secondary");
      b.textContent = a.label;
      b.addEventListener("click", function () {
        hideModal();
        if (typeof a.onClick === "function") a.onClick();
      });
      actionsEl.appendChild(b);
    });
    $("modal").classList.add("is-active");
  }
  function hideModal() {
    $("modal").classList.remove("is-active");
  }

  /* ---------- Форматирование даты «Обновлена: DD.MM HH:MM» ---------- */
  function formatDate(ts) {
    var d = new Date(ts);
    function pad(n) { return (n < 10 ? "0" : "") + n; }
    return pad(d.getDate()) + "." + pad(d.getMonth() + 1) + " " +
           pad(d.getHours()) + ":" + pad(d.getMinutes());
  }

  /* ---------- Обновление строки базы и кнопок на Старте ---------- */
  function applyBase(base) {
    // Делаем базу доступной для поиска
    data.setBase(base);

    var baseEl = $("startBase");
    if (base && base.items && base.items.length) {
      baseEl.classList.add("start__base--ok");
      baseEl.querySelector(".start__base-line").textContent =
        "База: " + base.items.length + " товара";
      baseEl.querySelector(".start__base-date").textContent =
        "Обновлена: " + formatDate(base.savedAt);
      $("btnReception").disabled = false;
      $("btnInventory").disabled = false;
    } else {
      baseEl.classList.remove("start__base--ok");
      baseEl.querySelector(".start__base-line").textContent = "База не загружена";
      baseEl.querySelector(".start__base-date").textContent = "";
      $("btnReception").disabled = true;
      $("btnInventory").disabled = true;
    }
  }

  /* ---------- Инициализация Старта: восстановить базу и настройки ---------- */
  function initStart() {
    var st = storage.getBaseState();
    if (st.corrupt) {
      showModal({
        title: "База повреждена",
        text: "Загрузите выгрузку МойСклад заново кнопкой «Обновить базу».",
        actions: [{ label: "Понятно", primary: true }]
      });
    }
    if (st.base) {
      applyBase(st.base);
    }
    // Восстановление тумблера «Архивные»
    var showArchived = storage.getSetting("showArchived", false);
    $("toggleArchived").checked = !!showArchived;
  }

  function showArchivedOn() {
    return !!$("toggleArchived").checked;
  }

  /* ---------- Загрузка базы из файла ---------- */
  function handleFile(file) {
    if (loading) return;

    // Проверка расширения .xlsx
    var name = (file.name || "").toLowerCase();
    if (!name.endsWith(".xlsx")) {
      showModal({
        title: "Не тот файл",
        text: "Это не Excel-файл. Нужен .xlsx из МойСклад.",
        actions: [{ label: "Понятно", primary: true }]
      });
      return;
    }

    loading = true;
    showOverlay("Загружаем базу…");

    var reader = new FileReader();
    reader.onload = function () {
      // Небольшая задержка — дать оверлею отрисоваться до тяжёлого парсинга
      window.setTimeout(function () {
        try {
          var result = data.parseWorkbook(reader.result);
          if (!result.ok) {
            hideOverlay();
            loading = false;
            showParseError(result);
            return;
          }
          var filtered = data.filterRows(result.rows);
          var base = {
            version: 1,
            savedAt: Date.now(),
            items: filtered.items
          };
          var saved = storage.saveBase(base);
          hideOverlay();
          loading = false;
          if (!saved) {
            showModal({
              title: "Не удалось сохранить",
              text: "База слишком велика для хранения в браузере. Обратитесь к разработчику.",
              actions: [{ label: "Понятно", primary: true }]
            });
            return;
          }
          applyBase(base);
          showModal({
            title: "База загружена",
            text: data.formatReport(filtered.report),
            actions: [{ label: "Готово", primary: true }]
          });
        } catch (e) {
          hideOverlay();
          loading = false;
          showModal({
            title: "Ошибка",
            text: "Не удалось прочитать файл. Попробуйте другую выгрузку МойСклад.",
            actions: [{ label: "Понятно", primary: true }]
          });
        }
      }, 50);
    };
    reader.onerror = function () {
      hideOverlay();
      loading = false;
      showModal({
        title: "Ошибка",
        text: "Не удалось открыть файл.",
        actions: [{ label: "Понятно", primary: true }]
      });
    };
    reader.readAsArrayBuffer(file);
  }

  function showParseError(result) {
    var text;
    if (result.error === "not_excel") {
      text = "Это не Excel-файл. Нужен .xlsx из МойСклад.";
    } else if (result.error === "empty") {
      text = "Файл пустой. Похоже, это не выгрузка МойСклад.";
    } else if (result.error === "no_column") {
      text = "В файле нет колонки «" + result.missing + "». " +
             "Похоже, это не выгрузка МойСклад.";
    } else {
      text = "Не удалось прочитать файл.";
    }
    showModal({
      title: result.error === "no_column" ? "Не та выгрузка" : "Ошибка файла",
      text: text,
      actions: [{ label: "Понятно", primary: true }]
    });
  }

  /* ---------- Режимы: переход на рабочий экран ---------- */
  function enterMode(mode) {
    currentMode = mode;

    if (mode === "inventory") {
      // SPEC раздел 5: инвентарка «начинает с нуля (пустой список).
      // Если есть незавершённая инвентарка — спросить Продолжить? Да/Нет».
      var saved = storage.loadList("inventory");
      if (saved.length > 0) {
        showModal({
          title: "Продолжить?",
          text: "Найдена незавершённая инвентаризация (" + saved.length + " поз.). Продолжить?",
          actions: [
            { label: "Да", primary: true, onClick: function () { finishEnterMode("inventory", saved); } },
            { label: "Нет", primary: false, onClick: function () {
              currentList = [];
              storage.clearList("inventory");
              finishEnterMode("inventory", []);
            } }
          ]
        });
        return;
      }
      finishEnterMode("inventory", []);
      return;
    }

    // Приёмка: всегда возобновляет сохранённый список (SPEC раздел 5).
    finishEnterMode("reception", storage.loadList("reception"));
  }

  function finishEnterMode(mode, list) {
    currentMode = mode;
    ui.showWork(mode);
    currentList = list;
    renderList();
    refocusSearch();
  }

  function backToStart() {
    // SPEC Состояние 5: при непустом списке — подтверждение (список не теряется)
    if (currentList.length > 0) {
      showModal({
        title: "Выйти?",
        text: "Список не выгружен. Выйти? (список сохранится)",
        actions: [
          { label: "Остаться", primary: true },
          { label: "Выйти", primary: false, onClick: doBackToStart }
        ]
      });
      return;
    }
    doBackToStart();
  }

  function doBackToStart() {
    currentMode = null;
    currentList = [];
    ui.clearSearchResults();
    ui.showStart();
  }

  /* ---------- Сохранение и рендер списка ---------- */
  function persistList() {
    if (!currentMode) return;
    storage.saveList(currentMode, currentList);
  }

  function renderList() {
    ui.renderWorkList(currentList, {
      onEdit: onEditRow,
      onRemove: onRemoveRow
    });
  }

  function refocusSearch() {
    var inp = $("searchInput");
    if (inp) inp.focus();
  }

  /* ---------- Поиск ---------- */
  function runSearch() {
    var q = $("searchInput").value;
    if (q.trim().length < 2) {
      ui.clearSearchResults();
      return;
    }
    var items = data.search(q, { showArchived: showArchivedOn() });
    ui.renderSearchResults(items, q, onPickItem, onCreateNew);
  }

  function onSearchInput() {
    if (searchDebounce) window.clearTimeout(searchDebounce);
    searchDebounce = window.setTimeout(runSearch, DEBOUNCE_MS);
  }

  function onSearchKey(e) {
    if (e.key !== "Enter") return;
    // Enter (в т.ч. от сканера): если ровно одно совпадение — выбрать его
    var q = $("searchInput").value;
    if (q.trim().length < 2) return;
    var items = data.search(q, { showArchived: showArchivedOn() });
    if (items.length === 1) {
      onPickItem(items[0]);
    }
  }

  /* ---------- Выбор товара из базы → окно количества ---------- */
  function onPickItem(item) {
    ui.clearSearchResults();
    $("searchInput").value = "";
    ui.openQtyModal({
      title: item.name,
      sub: (item.code ? "Код " + item.code : "") +
           (item.article ? (item.code ? " · " : "") + "Арт. " + item.article : "") +
           " · ед.: " + item.unit,
      onConfirm: function (qty) { addItem(item, qty); }
    });
  }

  /* ---------- Создать новую позицию → окно новой позиции ---------- */
  function onCreateNew() {
    var q = $("searchInput").value;
    ui.clearSearchResults();
    $("searchInput").value = "";
    ui.openNewPosModal({
      initialName: q && q.trim().length >= 2 ? q.trim() : "",
      onConfirm: function (name, qty, unit) { addNewItem(name, qty, unit); }
    });
  }

  /* ---------- Добавить товар из базы (с суммированием) ---------- */
  function addItem(item, qty) {
    var existing = null;
    for (var i = 0; i < currentList.length; i++) {
      if (!currentList[i].isNew && currentList[i].uuid === item.uuid) {
        existing = currentList[i];
        break;
      }
    }
    if (existing) {
      existing.qty += qty; // суммируем в ту же строку (контрольный пример №5)
      ui.toast("+" + qty + " → " + item.name +
               " (итого " + existing.qty + ")");
    } else {
      currentList.unshift({
        uuid: item.uuid,
        name: item.name,
        code: item.code || "",
        article: item.article || "",
        unit: item.unit || "",
        groups: item.groups || "",
        weighted: item.weighted || "",
        barcodes: item.barcodes || {},
        qty: qty,
        isNew: false
      });
    }
    persistList();
    renderList();
    refocusSearch();
  }

  /* ---------- Добавить новую позицию (с суммированием по имени) ---------- */
  function addNewItem(name, qty, unit) {
    var key = data.normalizeNewName(name);
    var existing = null;
    for (var i = 0; i < currentList.length; i++) {
      if (currentList[i].isNew &&
          data.normalizeNewName(currentList[i].name) === key) {
        existing = currentList[i];
        break;
      }
    }
    if (existing) {
      existing.qty += qty;
      // Обновляем ед. изм. на последнюю указанную (для согласованности)
      if (unit) existing.unit = unit;
      ui.toast("+" + qty + " → " + existing.name + " (итого " + existing.qty + ")");
    } else {
      currentList.unshift({
        uuid: "",
        name: name,
        code: "",
        unit: unit || "шт",
        qty: qty,
        isNew: true
      });
    }
    persistList();
    renderList();
    refocusSearch();
  }

  /* ---------- Правка количества (клик по строке) ---------- */
  function onEditRow(item, index) {
    ui.openQtyModal({
      title: item.name,
      sub: (item.code ? "Код " + item.code : "") +
           (item.unit ? (item.code ? " · " : "") + "ед.: " + item.unit : "") +
           (item.isNew ? " · новая" : ""),
      initialQty: item.qty,
      onConfirm: function (qty) { editItem(index, qty); }
    });
  }

  function editItem(index, qty) {
    if (index < 0 || index >= currentList.length) return;
    currentList[index].qty = qty;
    persistList();
    renderList();
    refocusSearch();
  }

  /* ---------- Удаление строки ---------- */
  function onRemoveRow(index) {
    if (index < 0 || index >= currentList.length) return;
    currentList.splice(index, 1);
    persistList();
    renderList();
  }

  /* ---------- Очистить список (с подтверждением) ---------- */
  function clearListConfirm() {
    if (!currentList.length) return;
    showModal({
      title: "Очистить список?",
      text: "В списке " + currentList.length + " " +
            posWord(currentList.length) + ". Точно очистить?",
      actions: [
        { label: "Нет", primary: true },
        { label: "Да", primary: false, onClick: function () {
          currentList = [];
          persistList();
          renderList();
        } }
      ]
    });
  }

  // Согласование слова «позиция»: 1 позиция, 2 позиции, 5 позиций
  function posWord(n) {
    var n10 = n % 10, n100 = n % 100;
    if (n10 === 1 && n100 !== 11) return "позиция";
    if (n10 >= 2 && n10 <= 4 && (n100 < 10 || n100 >= 20)) return "позиции";
    return "позиций";
  }

  /* ---------- Экспорт акта в .xlsx (Фаза 5) ---------- */
  var lastExport = null; // { fileName, blob, url, mode }

  function canShareFiles() {
    if (!navigator.canShare) return false;
    try { return navigator.canShare({ files: [new File(["test"], "t.xlsx")] }); }
    catch (e) { return false; }
  }

  function buildBlob() {
    var act = data.buildAct(currentMode, currentList);
    // Колонки — из buildAct: для инвентарки поле количества называется
    // «Фактический остаток» (шаблон импорта МойСклада), для приёмки — «Количество».
    var cols = act.cols || data.ACT_COLS;
    var ws = XLSX.utils.json_to_sheet(act.rows, { header: cols });

    // Ширины колонок — по содержимому (шапка + данные), с разумными пределами.
    // На распознавание МойСкладом ширины не влияют, но файл выглядит читаемо.
    var widths = cols.map(function (c) {
      var max = c.length;
      act.rows.forEach(function (r) {
        var v = r[c] == null ? "" : String(r[c]);
        if (v.length > max) max = v.length;
      });
      return { wch: Math.min(Math.max(max + 2, 8), 40) };
    });
    ws["!cols"] = widths;

    // Штрихкодам задаём текстовый формат («@»), чтобы ведущие нули не терялись
    // при открытии в Excel (как в шаблоне МойСклада, где UPC имеет формат «@»).
    var BC_COLS = ["Штрихкод EAN13", "Штрихкод EAN8", "Штрихкод Code128", "Штрихкод UPC"];
    BC_COLS.forEach(function (cname) {
      var ci = cols.indexOf(cname);
      if (ci < 0) return;
      for (var r = 0; r < act.rows.length; r++) {
        var addr = XLSX.utils.encode_cell({ r: r + 1, c: ci }); // +1: строка 0 = шапка
        if (ws[addr]) ws[addr].z = "@";
      }
    });

    var wb = XLSX.utils.book_new();
    // Лист «Sheet0» — как в шаблоне МойСклада (Пример.xlsx).
    // МойСклад ожидает именно это имя листа при автораспознавании.
    XLSX.utils.book_append_sheet(wb, ws, "Sheet0");

    // bookSST: true — обязательная опция. Без неё SheetJS пишет строки как
    // t="str" (кэш формулы) — это невалидный тип для ячеек с данными, и парсер
    // МойСклада (Apache POI) видит такие ячейки пустыми: шапка и данные не
    // отображаются, все колонки встают в «Не загружать».
    // С bookSST SheetJS создаёт xl/sharedStrings.xml и пишет строки как t="s"
    // со ссылкой по индексу — ровно так же, как сам МойСклад в своём шаблоне.
    var arr = XLSX.write(wb, { bookType: "xlsx", type: "array", bookSST: true });
    return { fileName: act.fileName, blob: new Blob([arr], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }) };
  }

  function downloadBlob(blob, fileName) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    // Не отзываем URL сразу — нужен для «Поделиться» и повторного скачивания.
    // Отзываем при уходе с экрана «Готово» (onBackToList/onNewAct).
    return url;
  }

  function exportAct() {
    if (!currentMode || currentList.length === 0) return;
    try {
      var built = buildBlob();
    } catch (e) {
      showModal({
        title: "Ошибка экспорта",
        text: "Не удалось сформировать файл. Попробуйте ещё раз.",
        actions: [{ label: "Понятно", primary: true }]
      });
      return;
    }
    // Снимаем прошлый URL, если был
    if (lastExport && lastExport.url) URL.revokeObjectURL(lastExport.url);

    var url = downloadBlob(built.blob, built.fileName);
    lastExport = { fileName: built.fileName, blob: built.blob, url: url, mode: currentMode };

    ui.showDone({
      fileName: built.fileName,
      canShare: canShareFiles(),
      onDownload: onDownloadAgain,
      onShare: onShare,
      // «Сверить остатки» доступен только после приёмки (SPEC раздел 5).
      onCheck: currentMode === "reception" ? onCheckRemainders : null,
      onBackToList: onBackToList,
      onNewAct: onNewAct
    });
  }

  function onDownloadAgain() {
    if (!lastExport) return;
    // Новый URL для повторного скачивания (старый оставляем для «Поделиться»)
    downloadBlob(lastExport.blob, lastExport.fileName);
  }

  function onShare() {
    if (!lastExport || !navigator.share) return;
    var file = new File([lastExport.blob], lastExport.fileName,
      { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    navigator.share({ files: [file], title: "Акт", text: lastExport.fileName })
      .catch(function () { /* отмена — молча */ });
  }

  function onBackToList() {
    // Вернуться к списку (сохранён), экспорт не чистит список
    if (lastExport && lastExport.url) { URL.revokeObjectURL(lastExport.url); lastExport = null; }
    ui.showWork(currentMode);
    renderList();
    refocusSearch();
  }

  function onNewAct() {
    // Подтверждение очистки списка
    showModal({
      title: "Новый акт?",
      text: "Текущий список будет очищен. Продолжить?",
      actions: [
        { label: "Нет", primary: true },
        { label: "Да", primary: false, onClick: function () {
          if (lastExport && lastExport.url) { URL.revokeObjectURL(lastExport.url); lastExport = null; }
          currentList = [];
          persistList();
          ui.showWork(currentMode);
          renderList();
          refocusSearch();
        } }
      ]
    });
  }

  /* «Сверить остатки» — после экспорта приёмки (SPEC раздел 5).
   * Да → инвентарка: позиции приёмки с UUID (новые отброшены), количества 0.
   * Нет → остаёмся в приёмке (список приёмки не трогаем). */
  function onCheckRemainders() {
    var recepList = storage.loadList("reception");
    var withUuid = recepList.filter(function (it) { return !it.isNew && it.uuid; }).length;
    var dropped = recepList.length - withUuid;
    var detail = "Перенос в инвентарку: " + withUuid + " поз." +
                 (dropped > 0 ? " (новые отброшены: " + dropped + ")" : "") +
                 ". Количество обнулится. Продолжить?";
    showModal({
      title: "Сверить остатки?",
      text: detail,
      actions: [
        { label: "Да", primary: true, onClick: startInventoryFromReception },
        { label: "Нет", primary: false }
      ]
    });
  }

  function startInventoryFromReception() {
    var recepList = storage.loadList("reception");
    var inv = data.buildInventoryFromReception(recepList);
    // Отзываем URL прошлого экспорта
    if (lastExport && lastExport.url) { URL.revokeObjectURL(lastExport.url); lastExport = null; }
    // Меняем режим, сохраняем стартовый список инвентарки
    currentMode = "inventory";
    currentList = inv;
    persistList();
    ui.showWork("inventory");
    renderList();
    refocusSearch();
  }

  /* ---------- Обработчики ---------- */
  function bindEvents() {
    $("btnUpdateBase").addEventListener("click", function () {
      $("fileInput").value = "";
      $("fileInput").click();
    });

    $("fileInput").addEventListener("change", function () {
      if (this.files && this.files[0]) {
        handleFile(this.files[0]);
      }
    });

    // Переход в режимы
    $("btnReception").addEventListener("click", function () {
      if (!this.disabled) enterMode("reception");
    });
    $("btnInventory").addEventListener("click", function () {
      if (!this.disabled) enterMode("inventory");
    });

    // Назад с рабочего экрана
    $("btnWorkBack").addEventListener("click", backToStart);

    // Очистить список
    $("btnClearList").addEventListener("click", clearListConfirm);

    // Экспорт акта
    $("btnExport").addEventListener("click", exportAct);

    // Поиск
    $("searchInput").addEventListener("input", onSearchInput);
    $("searchInput").addEventListener("keydown", onSearchKey);

    // Тумблер «Архивные» — сохраняем состояние
    $("toggleArchived").addEventListener("change", function () {
      storage.setSetting("showArchived", this.checked);
    });
  }

  /* ---------- Старт ---------- */
  function init() {
    storage = window.TeaApp.storage;
    data = window.TeaApp.data;
    ui = window.TeaApp.ui;
    if (!storage || !data || !ui) {
      console.error("TeaApp modules not ready");
      return;
    }
    ui.init(); // подключить кнопки модалов количества/новой позиции
    bindEvents();
    window.setTimeout(showStart, SPLASH_MS);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
