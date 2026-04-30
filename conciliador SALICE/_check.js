
    const $ = (id) => document.getElementById(id);
    const nf = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    // V3.7: estado de última corrida + validaciones manuales (promover Dudosos a Validados)
    let lastData = null;
    let forcedValidations = [];

    // Persistencia local de configuración
    const STORAGE_KEY = 'conciliador_cfg_v1';

    function readConfigFromUI() {
      return {
        margin: Number($("margin").value),
        t_days_sus: Number($("t_days_sus").value),
        max_opts: Number($("max_opts").value),
        day_weight_before: Number($("day_weight_before").value),
        day_weight_after: Number($("day_weight_after").value),
        peso_valid: Number($("peso_valid").value),
        peso_dudoso: Number($("peso_dudoso").value),
        mp_penalty: Number($("mp_penalty").value),
        cuit_mismatch_penalty: Number($("cuit_mismatch_penalty").value),
        alt_delta: Number($("alt_delta").value),
        pen_salice_galicia: Number($("pen_salice_galicia").value),
        pen_alarcon_bbva: Number($("pen_alarcon_bbva").value),
        show_peso: !!$("show_peso").checked,
        show_cuit: !!$("show_cuit").checked,
        persist_cfg: !!$("persist_cfg").checked,
      };
    }

    function applyConfigToUI(cfg) {
      if (!cfg || typeof cfg !== 'object') return;
      const setIf = (id, key) => {
        if (cfg[key] === undefined || cfg[key] === null) return;
        const el = $(id);
        if (!el) return;
        el.value = String(cfg[key]);
      };
      setIf('margin', 'margin');
      setIf('t_days_sus', 't_days_sus');
      setIf('max_opts', 'max_opts');
      setIf('day_weight_before', 'day_weight_before');
      setIf('day_weight_after', 'day_weight_after');
      setIf('peso_valid', 'peso_valid');
      setIf('peso_dudoso', 'peso_dudoso');
      setIf('mp_penalty', 'mp_penalty');
      setIf('cuit_mismatch_penalty', 'cuit_mismatch_penalty');
      setIf('alt_delta', 'alt_delta');
      setIf('pen_salice_galicia', 'pen_salice_galicia');
      setIf('pen_alarcon_bbva', 'pen_alarcon_bbva');
      if (cfg.show_peso !== undefined) $("show_peso").checked = !!cfg.show_peso;
      if (cfg.show_cuit !== undefined) $("show_cuit").checked = !!cfg.show_cuit;
      if (cfg.persist_cfg !== undefined) $("persist_cfg").checked = !!cfg.persist_cfg;
    }

    function setCfgStatus(msg) {
      const el = $("cfg_save_status");
      if (!el) return;
      el.textContent = msg || '';
      if (msg) {
        window.clearTimeout(setCfgStatus._t);
        setCfgStatus._t = window.setTimeout(() => { el.textContent = ''; }, 1800);
      }
    }

    function saveConfigToLocal() {
      const shouldPersist = $("persist_cfg").checked;
      if (!shouldPersist) {
        localStorage.removeItem(STORAGE_KEY);
        setCfgStatus('');
        return;
      }
      const cfg = readConfigFromUI();
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
        setCfgStatus('Guardado.');
      } catch (e) {
        setCfgStatus('No se pudo guardar.');
      }
    }

    function loadConfigFromLocal() {
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return;
        const cfg = JSON.parse(raw);
        applyConfigToUI(cfg);
      } catch (_) {
        // Si hay basura en storage, lo limpiamos.
        localStorage.removeItem(STORAGE_KEY);
      }
    }

    // Upload UI labels + drag & drop
    function setFileName(inputId, labelId) {
      const inp = $(inputId);
      const lab = $(labelId);
      if (!inp || !lab) return;
      const name = (inp.files && inp.files[0]) ? inp.files[0].name : "Ningún archivo seleccionado";
      lab.textContent = name;
      lab.classList.toggle('muted', !(inp.files && inp.files[0]));
    }
    ['excel','pdf_salice','pdf_alarcon'].forEach(id => {
      const map = { excel:'excel_name', pdf_salice:'salice_name', pdf_alarcon:'alarcon_name' };
      const el = $(id);
      if (el) el.addEventListener('change', () => setFileName(id, map[id]));
      // init
      setFileName(id, map[id]);
    });

    // Enable drag & drop onto the visible drop areas.
    // Note: browsers don't allow setting input.files directly in all cases,
    // so we use a DataTransfer to assign the FileList safely.
    function bindDropArea(inputId, nameId) {
      const inp = $(inputId);
      const nameEl = $(nameId);
      if (!inp || !nameEl) return;
      // The drop target is the <label class="drop" for this input.
      const dropEl = document.querySelector(`label.drop[for="${inputId}"]`);
      if (!dropEl) return;

      const prevent = (e) => { e.preventDefault(); e.stopPropagation(); };
      dropEl.addEventListener('dragenter', (e) => { prevent(e); dropEl.classList.add('dragover'); });
      dropEl.addEventListener('dragover', (e) => { prevent(e); dropEl.classList.add('dragover'); });
      dropEl.addEventListener('dragleave', (e) => { prevent(e); dropEl.classList.remove('dragover'); });
      dropEl.addEventListener('drop', (e) => {
        prevent(e);
        dropEl.classList.remove('dragover');
        const files = e.dataTransfer?.files;
        if (!files || files.length === 0) return;

        const f0 = files[0];
        if (inputId === 'excel') {
          const ok = /\.(xlsx|xls)$/i.test(f0.name);
          if (!ok) return;
        } else {
          const ok = /\.(pdf)$/i.test(f0.name);
          if (!ok) return;
        }

        const dt = new DataTransfer();
        dt.items.add(f0);
        inp.files = dt.files;
        inp.dispatchEvent(new Event('change', { bubbles: true }));
      });
    }

    bindDropArea('excel', 'excel_name');
    bindDropArea('pdf_salice', 'salice_name');
    bindDropArea('pdf_alarcon', 'alarcon_name');

    function renderTable(el, rows, opts = null) {
      el.innerHTML = "";
      if (!rows || rows.length === 0) {
        el.innerHTML = "<tr><td class='muted'>Sin filas</td></tr>";
        return;
      }
      const cols = Array.from(new Set(rows.flatMap(r => Object.keys(r))));
      const actionHeader = (opts && opts.action) ? (opts.actionHeader || 'Acción') : null;
      if (actionHeader) cols.push(actionHeader);
      const thead = document.createElement('thead');
      const trh = document.createElement('tr');
      cols.forEach(c => {
        const th = document.createElement('th');
        th.textContent = c;
        trh.appendChild(th);
      });
      thead.appendChild(trh);
      el.appendChild(thead);
      const tbody = document.createElement('tbody');
      rows.forEach((r, idx) => {
        const tr = document.createElement('tr');
        cols.forEach(c => {
          const td = document.createElement('td');
          if (actionHeader && c === actionHeader) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn-mini';
            btn.textContent = opts.actionLabel || 'Validar';
            btn.addEventListener('click', () => {
              if (typeof opts.onAction === 'function') opts.onAction(idx, r);
            });
            td.appendChild(btn);
            tr.appendChild(td);
            return;
          }
          const raw = r[c];
          if (raw === null || raw === undefined) {
            td.textContent = "";
          } else {
            const isMoneyCol = /Importe|Dif importe|Peso/i.test(c);
            const isNumber = typeof raw === 'number' || (!!raw && !isNaN(raw) && raw.trim && raw.trim() !== '' && !isNaN(Number(raw)));
            if (isMoneyCol && isNumber) {
              td.textContent = nf.format(Number(raw));
            } else {
              td.textContent = String(raw);
            }
          }
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      el.appendChild(tbody);
    }

    function dudosoKey(row) {
      return {
        fila_excel: row ? row['Fila Excel'] : undefined,
        nro_recibo: row ? row['Nro recibo'] : undefined,
        medio_pago: row ? row['Medio de pago'] : undefined,
      };
    }

    function promoteDudoso(idx) {
      if (!lastData || !Array.isArray(lastData.dudosos) || !Array.isArray(lastData.validados)) return;
      const row = lastData.dudosos[idx];
      if (!row) return;

      // Guardar override para exportación
      const k = dudosoKey(row);
      if (k.fila_excel !== undefined && k.fila_excel !== null &&
          k.nro_recibo !== undefined && k.nro_recibo !== null &&
          k.medio_pago !== undefined && k.medio_pago !== null) {
        const sig = `${k.fila_excel}||${k.nro_recibo}||${k.medio_pago}`;
        const exists = forcedValidations.some(x => `${x.fila_excel}||${x.nro_recibo}||${x.medio_pago}` === sig);
        if (!exists) forcedValidations.push(k);
      }

      // Mover en UI
      lastData.dudosos.splice(idx, 1);
      lastData.validados.push(row);

      $("cnt_ok").textContent = `(${lastData.validados.length})`;
      $("cnt_warn").textContent = `(${lastData.dudosos.length})`;

      renderTable($("tbl_ok"), lastData.validados);
      renderTable($("tbl_warn"), lastData.dudosos, {
        action: true,
        actionHeader: 'Acción',
        actionLabel: 'Validar',
        onAction: (i) => promoteDudoso(i),
      });
    }

    async function run() {
      $("error").textContent = "";
      $("status").textContent = "Procesando...";
      $("run").disabled = true;
      try {
        const excel = $("excel").files[0];
        const pdfSalice = $("pdf_salice").files[0];
        const pdfAlarcon = $("pdf_alarcon").files[0];
        if (!excel || (!pdfSalice && !pdfAlarcon)) {
          throw new Error("Tenés que seleccionar un Excel y al menos 1 PDF (SALICE o Alarcón)." );
        }

        const margin = encodeURIComponent($("margin").value);
        const tDaysSus = encodeURIComponent($("t_days_sus").value);
        const maxOpts = encodeURIComponent($("max_opts").value);
        const dayWeightBefore = encodeURIComponent($("day_weight_before").value);
        const dayWeightAfter = encodeURIComponent($("day_weight_after").value);
        const pesoValid = encodeURIComponent($("peso_valid").value);
        const pesoDudoso = encodeURIComponent($("peso_dudoso").value);
        const mpPenalty = encodeURIComponent($("mp_penalty").value);
        const cuitMismatchPenalty = encodeURIComponent($("cuit_mismatch_penalty").value);
        const penSaliceGalicia = encodeURIComponent($("pen_salice_galicia").value);
        const penAlarconBbva = encodeURIComponent($("pen_alarcon_bbva").value);
        const altDelta = encodeURIComponent($("alt_delta").value);

        const url = `/compare?margin_days=${margin}`
          + `&tolerance_days_suspect=${tDaysSus}`
          + `&max_options=${maxOpts}`
          + `&day_weight_bank_before=${dayWeightBefore}`
          + `&day_weight_bank_after=${dayWeightAfter}`
          + `&valid_max_peso=${pesoValid}`
          + `&dudoso_max_peso=${pesoDudoso}`
          + `&mp_mismatch_penalty=${mpPenalty}`
          + `&penalty_salice_to_galicia=${penSaliceGalicia}`
          + `&penalty_alarcon_to_bbva=${penAlarconBbva}`
          + `&cliente_cuit_mismatch_penalty=${cuitMismatchPenalty}`
          + `&alternatives_cost_delta=${altDelta}`
          + `&show_peso=${encodeURIComponent($("show_peso").checked ? 1 : 0)}`
          + `&show_cuit=${encodeURIComponent($("show_cuit").checked ? 1 : 0)}`;

        const form = new FormData();
        form.append('excel', excel);
        if (pdfSalice) form.append('pdf_salice', pdfSalice);
        if (pdfAlarcon) form.append('pdf_alarcon', pdfAlarcon);

        const res = await fetch(url, { method: 'POST', body: form });
        const text = await res.text();
        if (!res.ok) throw new Error(text);

        const data = JSON.parse(text);

        lastData = data;
        forcedValidations = [];

        $("meta").textContent = JSON.stringify(data.meta, null, 2);
        $("cnt_ok").textContent = `(${data.validados.length})`;
        $("cnt_warn").textContent = `(${data.dudosos.length})`;
        $("cnt_bad").textContent = `(${data.no_encontrados.length})`;

        renderTable($("tbl_ok"), data.validados);
        renderTable($("tbl_warn"), data.dudosos, {
          action: true,
          actionHeader: 'Acción',
          actionLabel: 'Validar',
          onAction: (i) => promoteDudoso(i),
        });
        renderTable($("tbl_bad"), data.no_encontrados);

        $("dl_xlsx").disabled = false;
        $("dl_user_xlsx").disabled = false;
        $("status").textContent = "Listo.";
      } catch (e) {
        $("error").textContent = e?.message || String(e);
        $("status").textContent = "";
        $("dl_xlsx").disabled = true;
        $("dl_user_xlsx").disabled = true;
      } finally {
        $("run").disabled = false;
      }
    }

    async function download(format) {
      $("error").textContent = "";
      try {
        // Mostrar estado de procesamiento también durante la exportación
        $("status").textContent = "Procesando...";
        $("run").disabled = true;
        $("dl_user_xlsx").disabled = true;
        $("dl_xlsx").disabled = true;

        const excel = $("excel").files[0];
        const pdfSalice = $("pdf_salice").files[0];
        const pdfAlarcon = $("pdf_alarcon").files[0];
        if (!excel || (!pdfSalice && !pdfAlarcon)) {
          throw new Error("Tenés que seleccionar un Excel y al menos 1 PDF." );
        }

        const margin = encodeURIComponent($("margin").value);
        const tDaysSus = encodeURIComponent($("t_days_sus").value);
        const maxOpts = encodeURIComponent($("max_opts").value);
        const dayWeightBefore = encodeURIComponent($("day_weight_before").value);
        const dayWeightAfter = encodeURIComponent($("day_weight_after").value);
        const pesoValid = encodeURIComponent($("peso_valid").value);
        const pesoDudoso = encodeURIComponent($("peso_dudoso").value);
        const mpPenalty = encodeURIComponent($("mp_penalty").value);
        const cuitMismatchPenalty = encodeURIComponent($("cuit_mismatch_penalty").value);
        const penSaliceGalicia = encodeURIComponent($("pen_salice_galicia").value);
        const penAlarconBbva = encodeURIComponent($("pen_alarcon_bbva").value);
        const altDelta = encodeURIComponent($("alt_delta").value);

        const url = `/export?format=${format}&margin_days=${margin}`
          + `&tolerance_days_suspect=${tDaysSus}`
          + `&max_options=${maxOpts}`
          + `&day_weight_bank_before=${dayWeightBefore}`
          + `&day_weight_bank_after=${dayWeightAfter}`
          + `&valid_max_peso=${pesoValid}`
          + `&dudoso_max_peso=${pesoDudoso}`
          + `&mp_mismatch_penalty=${mpPenalty}`
          + `&penalty_salice_to_galicia=${penSaliceGalicia}`
          + `&penalty_alarcon_to_bbva=${penAlarconBbva}`
          + `&cliente_cuit_mismatch_penalty=${cuitMismatchPenalty}`
          + `&alternatives_cost_delta=${altDelta}`
          + `&show_peso=${encodeURIComponent($("show_peso").checked ? 1 : 0)}`
          + `&show_cuit=${encodeURIComponent($("show_cuit").checked ? 1 : 0)}`;

        const form = new FormData();
        form.append('excel', excel);
        if (forcedValidations && forcedValidations.length) {
          form.append('force_validations', JSON.stringify(forcedValidations));
        }
        if (pdfSalice) form.append('pdf_salice', pdfSalice);
        if (pdfAlarcon) form.append('pdf_alarcon', pdfAlarcon);

        const res = await fetch(url, { method: 'POST', body: form });
        if (!res.ok) throw new Error(await res.text());

        const blob = await res.blob();
        const a = document.createElement('a');
        const ext = (format === 'xlsx' || format === 'devxlsx') ? 'xlsx' : 'zip';
        a.href = URL.createObjectURL(blob);
        if (format === 'xlsx') {
          const base = (excel && excel.name) ? excel.name.replace(/\.[^/.]+$/, '') : 'ingresos';
          a.download = `${base}_conciliado.${ext}`;
        }
        else if (format === 'devxlsx') a.download = `resultado_conciliacion_dev.${ext}`;
        else a.download = `resultado_conciliacion.${ext}`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        $("status").textContent = "Listo.";
      } catch (e) {
        $("error").textContent = e?.message || String(e);
        $("status").textContent = "";
      } finally {
        $("run").disabled = false;
        // Rehabilitar según si ya hay resultado cargado o no
        const canDownload = !$("dl_xlsx").disabled || !$("dl_user_xlsx").disabled;
        // Si antes de descargar estaban habilitados, los volvemos a habilitar.
        // (Si no hay resultados, seguirán deshabilitados por el flujo normal.)
        if ($("meta").textContent && $("meta").textContent.trim().length > 0) {
          $("dl_xlsx").disabled = false;
          $("dl_user_xlsx").disabled = false;
        }
      }
    }

    $("run").addEventListener('click', run);
    $("dl_user_xlsx").addEventListener('click', () => download('xlsx'));
    $("dl_xlsx").addEventListener('click', () => download('devxlsx'));

    // Config modal
    function setDefaults() {
      $("margin").value = 5;
      $("t_days_sus").value = 7;
      $("max_opts").value = 4;
      $("day_weight_before").value = 40;
      $("day_weight_after").value = 50;
      $("peso_valid").value = 150;
      $("peso_dudoso").value = 3500;
      $("mp_penalty").value = 35;
      $("cuit_mismatch_penalty").value = 90;
      $("alt_delta").value = 50;
      $("pen_salice_galicia").value = 45;
      $("pen_alarcon_bbva").value = 45;
      $("show_peso").checked = false;
      $("show_cuit").checked = false;
      $("persist_cfg").checked = false;
    }
    setDefaults();
    loadConfigFromLocal();

    // Config modal (se abre desde el menú)
    // (Mantenemos compatibilidad por si en algún momento vuelve a existir el botón)
    const legacyCfgBtn = $("btn_cfg");
    if (legacyCfgBtn) legacyCfgBtn.addEventListener('click', () => { $("cfg_modal").style.display = 'block'; });
    $("btn_cfg_close").addEventListener('click', () => { $("cfg_modal").style.display = 'none'; });
    $("cfg_modal").addEventListener('click', (e) => { if (e.target.id === 'cfg_modal') $("cfg_modal").style.display = 'none'; });
    $("btn_cfg_reset").addEventListener('click', () => {
      setDefaults();
      localStorage.removeItem(STORAGE_KEY);
      setCfgStatus('Defaults restaurados.');
    });
    $("btn_cfg_save").addEventListener('click', () => { saveConfigToLocal(); });
    $("btn_cfg_apply").addEventListener('click', () => {
      // Si está activado, guardamos antes de cerrar. Si no, limpiamos.
      if ($("persist_cfg").checked) saveConfigToLocal();
      else localStorage.removeItem(STORAGE_KEY);
      $("cfg_modal").style.display = 'none';
    });

    // Pop-out menu
    function closeMenu() { $("menu_dd").style.display = 'none'; }
    function toggleMenu() {
      $("menu_dd").style.display = ($("menu_dd").style.display === 'block') ? 'none' : 'block';
    }
    $("menu_btn").addEventListener('click', (e) => { e.stopPropagation(); toggleMenu(); });
    document.addEventListener('click', () => closeMenu());
    $("menu_dd").addEventListener('click', (e) => e.stopPropagation());

    $("menu_cfg").addEventListener('click', () => { closeMenu(); setCfgStatus(''); $("cfg_modal").style.display = 'block'; });
    $("menu_meta").addEventListener('click', () => { closeMenu(); $("meta_modal").style.display = 'block'; });

    // Meta modal
    $("btn_meta_close").addEventListener('click', () => { $("meta_modal").style.display = 'none'; });
    $("meta_modal").addEventListener('click', (e) => { if (e.target.id === 'meta_modal') $("meta_modal").style.display = 'none'; });
  
