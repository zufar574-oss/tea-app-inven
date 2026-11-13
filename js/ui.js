/* ui.js — отрисовка экранов, выпадающий список поиска, toast,
 * окно количества, окно новой позиции, список выбранных товаров, экран «Готово».
 * Фаза 3: навигация, поиск, toast.
 * Фаза 4: окно количества, окно новой позиции, список выбранных.
 * Фаза 5: активация экспорта, экран «Готово», ед. изм. для новой позиции. */
(function () {
  "use strict";

  function $(id) { return document.getElementById(id); }

  /* ---------- Навигация между экранами ---------- */
  function setActiveScreen(name) {
    var screens = ["start", "work", "done"];
    for (var i = 0; i < screens.length; i++) {
      var el = $(screens[i]);
      if (el) el.classList.toggle("is-active", screens[i] === name);
    }
  }

  function showStart() { setActiveScreen("start"); }

  function showWork(mode) {
    setActiveScreen("work");
    var title = mode === "inventory" ? "ИНВЕНТАРКА" : "ПРИЁМКА";
    $("workTitle").textContent = title;
    // Очистка поиска и фокус
    $("searchInput").value = "";
    clearSearchResults();
    renderWorkList([], null); // чистый экран; main.js дозагрузит сохранённый список
    window.setTimeout(function () {
      var inp = $("searchInput");
      if (inp) inp.focus();
    }, 50);
  }

  /* ---------- Подпись элемента: код или артикул ---------- */
  function itemSubLine(it) {
    var parts = [];
    if (it.code) parts.push("Код " + it.code);
    if (it.article) parts.push("Арт. " + it.article);
    return parts.join(" · ");
  }

  /* ---------- Рендер результатов поиска (дропдаун) ---------- */
  function renderSearchResults(items, query, onPick, onCreate) {
    var box = $("searchDropdown");
    box.innerHTML = "";

    var rows = document.createDocumentFragment();

    items.forEach(function (it) {
      var row = document.createElement("button");
      row.type = "button";
      row.className = "dd-row" + (it.archived ? " dd-row--archived" : "");
      var left = document.createElement("span");
      left.className = "dd-row__name";
      left.textContent = it.name + (it.archived ? " (архив)" : "");
      var sub = document.createElement("span");
      sub.className = "dd-row__sub";
      sub.textContent = itemSubLine(it);
      var unit = document.createElement("span");
      unit.className = "dd-row__unit";
      unit.textContent = it.unit;
      row.appendChild(left);
      row.appendChild(sub);
      row.appendChild(unit);
      row.addEventListener("click", function () { onPick(it); });
      rows.appendChild(row);
    });

    if (query && query.trim().length >= 2) {
      var create = document.createElement("button");
      create.type = "button";
      create.className = "dd-row dd-row--create";
      create.textContent = "+ Создать новую позицию";
      create.addEventListener("click", function () { onCreate(); });
      rows.appendChild(create);
    }

    box.appendChild(rows);
    box.classList.toggle("is-active", box.children.length > 0);
  }

  function clearSearchResults() {
    var box = $("searchDropdown");
    box.innerHTML = "";
    box.classList.remove("is-active");
  }

  /* ---------- Список выбранных товаров ----------
   * items — массив { uuid, name, code, unit, qty, isNew }.
   * handlers = { onEdit(item, index), onRemove(index) }. */
  function renderWorkList(items, handlers) {
    items = items || [];
    handlers = handlers || {};
    var body = $("workList");
    body.innerHTML = "";

    if (!items.length) {
      var ph = document.createElement("p");
      ph.className = "work__placeholder";
      ph.textContent = "Найдите товар по названию, коду или отсканируйте штрихкод";
      body.appendChild(ph);
    } else {
      items.forEach(function (it, idx) {
        var row = document.createElement("div");
        row.className = "work-row" + (it.isNew ? " work-row--new" : "");
        var main = document.createElement("div");
        main.className = "work-row__main";
        var name = document.createElement("span");
        name.className = "work-row__name";
        name.textContent = it.name;
        var sub = document.createElement("span");
        sub.className = "work-row__sub";
        sub.textContent = it.code ? ("Код " + it.code) : "";
        main.appendChild(name);
        main.appendChild(sub);

        var qty = document.createElement("span");
        qty.className = "work-row__qty";
        qty.textContent = it.qty + (it.unit ? " " + it.unit : "");

        var rm = document.createElement("button");
        rm.type = "button";
        rm.className = "work-row__remove";
        rm.setAttribute("aria-label", "Удалить");
        rm.textContent = "✕";
        var removeIdx = idx;
        rm.addEventListener("click", function (e) {
          e.stopPropagation();
          if (typeof handlers.onRemove === "function") handlers.onRemove(removeIdx);
        });

        row.appendChild(main);
        row.appendChild(qty);
        row.appendChild(rm);
        var editIdx = idx, editItem = it;
        row.addEventListener("click", function () {
          if (typeof handlers.onEdit === "function") handlers.onEdit(editItem, editIdx);
        });
        body.appendChild(row);
      });
    }

    // Счётчик и кнопки футера
    $("workCounter").textContent = items.length + " поз.";
    $("btnClearList").disabled = items.length === 0;
    $("btnExport").disabled = items.length === 0; // Ф5: экспорт активен при непустом списке
  }

  /* ---------- Toast ---------- */
  var toastTimer = null;
  function toast(msg) {
    var el = $("toast");
    el.textContent = msg;
    el.classList.add("is-active");
    if (toastTimer) window.clearTimeout(toastTimer);
    toastTimer = window.setTimeout(function () {
      el.classList.remove("is-active");
    }, 2000);
  }

  /* ---------- Окно количества ---------- */
  var qtyConfirm = null;

  function setupQtyModal() {
    $("qtyCancel").addEventListener("click", closeQtyModal);
    $("qtyAdd").addEventListener("click", function () {
      var raw = $("qtyInput").value || "";
      var digits = raw.replace(/\D/g, "");
      var qty = digits === "" ? 0 : parseInt(digits, 10);
      var h = qtyConfirm;
      closeQtyModal();
      if (typeof h === "function") h(qty);
    });
    // Только цифры, дроби и минусы отсекаются на лету
    $("qtyInput").addEventListener("input", function () {
      this.value = this.value.replace(/\D/g, "");
    });
    $("qtyInput").addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); $("qtyAdd").click(); }
      if (e.key === "Escape") { closeQtyModal(); }
    });
  }

  function openQtyModal(opts) {
    opts = opts || {};
    $("qtyTitle").textContent = opts.title || "";
    $("qtySub").textContent = opts.sub || "";
    $("qtyInput").value = (opts.initialQty != null && opts.initialQty !== "")
      ? String(opts.initialQty) : "";
    qtyConfirm = opts.onConfirm || null;
    $("qtyModal").classList.add("is-active");
    window.setTimeout(function () {
      var inp = $("qtyInput");
      if (inp) { inp.focus(); inp.select(); }
    }, 30);
  }

  function closeQtyModal() {
    $("qtyModal").classList.remove("is-active");
    qtyConfirm = null;
  }

  /* ---------- Окно новой позиции ---------- */
  var newPosConfirm = null;

  function setupNewPosModal() {
    $("newPosCancel").addEventListener("click", closeNewPosModal);
    $("newPosAdd").addEventListener("click", function () {
      var name = ($("newPosName").value || "").trim();
      if (!name) {
        $("newPosName").focus();
        return;
      }
      var digits = ($("newPosQty").value || "").replace(/\D/g, "");
      var qty = digits === "" ? 0 : parseInt(digits, 10);
      // Ед. изм. нужна для импорта в МойСклад; по умолчанию «шт»
      var unit = ($("newPosUnit").value || "").trim() || "шт";
      var h = newPosConfirm;
      closeNewPosModal();
      if (typeof h === "function") h(name, qty, unit);
    });
    $("newPosQty").addEventListener("input", function () {
      this.value = this.value.replace(/\D/g, "");
    });
    $("newPosName").addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); $("newPosQty").focus(); }
    });
    $("newPosQty").addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); $("newPosUnit").focus(); }
    });
    $("newPosUnit").addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); $("newPosAdd").click(); }
      if (e.key === "Escape") { closeNewPosModal(); }
    });
  }

  function openNewPosModal(opts) {
    opts = opts || {};
    $("newPosName").value = opts.initialName || "";
    $("newPosQty").value = (opts.initialQty != null && opts.initialQty !== "")
      ? String(opts.initialQty) : "";
    $("newPosUnit").value = opts.initialUnit || "шт";
    newPosConfirm = opts.onConfirm || null;
    $("newPosModal").classList.add("is-active");
    window.setTimeout(function () { $("newPosName").focus(); }, 30);
  }

  function closeNewPosModal() {
    $("newPosModal").classList.remove("is-active");
    newPosConfirm = null;
  }

  /* ---------- Экран «Готово» (после экспорта) ----------
   * opts = { fileName, canShare, onDownload, onShare, onBackToList, onNewAct, onCheck }.
   * «Поделиться» — только при canShare. «Сверить остатки» (onCheck) — только для приёмки.
   * Инструкции по импорту в МойСклад на экране не показываются (убраны по
   * запросу пользователя) — полные шаги импорта живут в README, п. 4. */
  function showDone(opts) {
    opts = opts || {};
    $("doneFileName").textContent = opts.fileName || "";
    $("doneFileName").title = opts.fileName || "";

    // Кнопка «Поделиться» — только при поддержке Web Share с файлами
    $("btnDoneShare").style.display = opts.canShare ? "" : "none";
    // Кнопка «Сверить остатки» — только для режима приёмки (onCheck задан)
    $("btnDoneCheck").style.display = opts.onCheck ? "" : "none";

    // Перевешиваем обработчики (один клик — один листенер через клон)
    function rebind(id, fn) {
      var el = $(id);
      var clone = el.cloneNode(true);
      el.parentNode.replaceChild(clone, el);
      clone.addEventListener("click", fn);
    }
    rebind("btnDoneDownload", opts.onDownload || function () {});
    rebind("btnDoneShare", opts.onShare || function () {});
    rebind("btnDoneCheck", opts.onCheck || function () {});
    rebind("btnDoneBack", opts.onBackToList || function () {});
    rebind("btnDoneNew", opts.onNewAct || function () {});

    setActiveScreen("done");
  }

  /* ---------- Инициализация модалов (один раз) ---------- */
  function init() {
    setupQtyModal();
    setupNewPosModal();
  }

  window.TeaApp.ui = {
    init: init,
    showStart: showStart,
    showWork: showWork,
    showDone: showDone,
    renderSearchResults: renderSearchResults,
    clearSearchResults: clearSearchResults,
    renderWorkList: renderWorkList,
    toast: toast,
    openQtyModal: openQtyModal,
    closeQtyModal: closeQtyModal,
    openNewPosModal: openNewPosModal,
    closeNewPosModal: closeNewPosModal
  };
})();
