(function () {
  // --- DOM refs ---
  const bgFile = document.getElementById("bgFile"); //template form
  const csvFile = document.getElementById("csvFile");
  const btnParse = document.getElementById("btnParse");
  const btnGenerate = document.getElementById("btnGenerate");
  const btnExportZip = document.getElementById("btnExportZip");
  const listEl = document.getElementById("list");
  const canvas = document.getElementById("preview");
  const ctx = canvas.getContext("2d");
  const dateFrom = document.getElementById("dateFrom");
  const dateTo = document.getElementById("dateTo");
  //export parameters
  const wMobile = document.getElementById("wMobile"); //width
  const jpegQ = document.getElementById("jpegQ"); //quality of jpeg

  // --- state ---
  let bgImage = null; // HTMLImageElement
  let tasks = []; // parsed tasks (raw)
  let joined = []; // main + assistant merged
  let filtered = []; // filtered by date
  let renderedItems = new Map(); // filename -> Blob (jpeg)

  // === Coordinates for PNG 1240×1754 ===
  const MAP = {
    baseW: 1240,
    baseH: 1754,
    fields: [
      { key: "Person", x: 230, y: 325, size: 56, align: "left", maxWidth: 900 }, // "Имя"
      {
        key: "Assistant",
        x: 405,
        y: 455,
        size: 56,
        align: "left",
        maxWidth: 900,
      }, // "Напарник"
      { key: "Date", x: 260, y: 580, size: 56, align: "left", maxWidth: 520 }, // "Дата"
      {
        key: "Assignment",
        x: 445,
        y: 705,
        size: 56,
        align: "left",
        maxWidth: 750,
        wrap: true,
        lineHeight: 1.2,
      }, // "Задание"
    ],
    checkboxes: [
      { key: "SchoolMain", x: 130, y: 975, size: 50 }, // "В главном зале"
    ],
  };

  // === Date shift ===
  const defaultDateOffsetDays = 1; // monday -> tuesday
  const specialOffsets = {
    /* "YYYY-MM-DD": days */
  };

  function startOfISOWeek(d) {
    const day = (d.getDay() + 6) % 7; // 0..6, где 0 — monday
    const res = new Date(d);
    res.setDate(res.getDate() - day);
    res.setHours(0, 0, 0, 0);
    return res;
  }

  // === Helpers ===
  function enableParseIfReady() {
    btnParse.disabled = !(bgFile.files.length && csvFile.files.length);
  }
  bgFile.addEventListener("change", enableParseIfReady);
  csvFile.addEventListener("change", enableParseIfReady);

  function parseDate(s) {
    if (!s) return null;
    s = String(s).trim();

    let y, m, d;
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      // 2025-05-05
      [y, m, d] = s.split("-").map(Number);
    } else if (/^\d{4}\.\d{2}\.\d{2}$/.test(s)) {
      // 2025.05.05
      [y, m, d] = s.split(".").map(Number);
    } else if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) {
      // 05.05.2025
      [d, m, y] = s.split(".").map(Number);
    } else {
      // fallback — local date from string
      const tmp = new Date(s);
      if (isNaN(+tmp)) return null;
      y = tmp.getFullYear();
      m = tmp.getMonth() + 1;
      d = tmp.getDate();
    }

    const date = new Date(y, m - 1, d);
    date.setHours(0, 0, 0, 0);
    return date;
  }

  function fmtDate(date) {
    if (!date) return "";
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function fmtDateHuman(date) {
    if (!date) return "";
    return date.toLocaleDateString("ru-RU", { day: "numeric", month: "long" });
  }

  function wrapText(text, maxWidth, font) {
    ctx.font = font;
    const words = String(text || "").split(/\s+/);
    const lines = [];
    let line = "";
    for (const w of words) {
      const test = line ? line + " " + w : w;
      const wpx = ctx.measureText(test).width;
      if (wpx <= maxWidth || !line) {
        line = test;
      } else {
        lines.push(line);
        line = w;
      }
    }
    if (line) lines.push(line);
    return lines;
  }

  function drawCheck(x, y, s) {
    ctx.lineWidth = Math.max(2, Math.round(s / 10));
    ctx.strokeStyle = "#000";
    ctx.beginPath();
    ctx.moveTo(x + s * 0.15, y + s * 0.55);
    ctx.lineTo(x + s * 0.45, y + s * 0.85);
    ctx.lineTo(x + s * 0.9, y + s * 0.2);
    ctx.stroke();
  }

  async function loadBgImage(file) {
    if (!file || !(file instanceof Blob) || file.size === 0) {
      throw new Error("loadBgImage: invalid file");
    }
    const url = URL.createObjectURL(file);
    try {
      const img = new Image();
      img.decoding = "async";
      await new Promise((resolve, reject) => {
        img.onload = () => resolve();
        img.onerror = () => reject(new Error("Failed to load image"));
        img.src = url;
      });
      bgImage = img;
      return img;
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  function drawRecordToCanvas(rec) {
    if (!bgImage) return;
    canvas.width = bgImage.naturalWidth;
    canvas.height = bgImage.naturalHeight;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(bgImage, 0, 0, canvas.width, canvas.height);

    const scaleX = canvas.width / MAP.baseW;
    const scaleY = canvas.height / MAP.baseH;
    const fontFamily = "Inter, Noto Sans, system-ui, Arial";

    for (const f of MAP.fields) {
      let val = rec[f.key] ?? "";
      if (f.key === "Date" && rec.DateDisplay) {
        val = rec.DateDisplay;
      }
      const x = Math.round(f.x * scaleX);
      const y = Math.round(f.y * scaleY);
      const size = Math.round((f.size * (scaleX + scaleY)) / 2);
      const maxW = f.maxWidth ? Math.round(f.maxWidth * scaleX) : undefined;

      ctx.fillStyle = "#000";
      ctx.font = `bold ${size}px ${fontFamily}`;
      ctx.textAlign = f.align || "left";
      ctx.textBaseline = "top";

      if (f.wrap && maxW) {
        const lines = wrapText(val, maxW, ctx.font);
        const lh = Math.round(size * (f.lineHeight || 1.15));
        let yy = y;
        for (const line of lines) {
          ctx.fillText(line, f.align === "right" ? x + maxW : x, yy, maxW);
          yy += lh;
        }
      } else {
        ctx.fillText(
          String(val || ""),
          f.align === "right" && maxW ? x + maxW : x,
          y,
          maxW
        );
      }
    }

    if (String(rec.School).trim() === "1") {
      const cb = MAP.checkboxes[0];
      const x = Math.round(cb.x * scaleX);
      const y = Math.round(cb.y * scaleY);
      const s = Math.round((cb.size * (scaleX + scaleY)) / 2);
      drawCheck(x, y, s);
    }
  }

  async function canvasToJpegBlob(targetW) {
    if (!targetW || targetW >= canvas.width) {
      const b = await new Promise((res) =>
        canvas.toBlob(res, "image/jpeg", Number(jpegQ.value || 0.9))
      );
      return b;
    }
    const t = document.createElement("canvas");
    const ratio = targetW / canvas.width;
    t.width = targetW;
    t.height = Math.round(canvas.height * ratio);
    const tctx = t.getContext("2d");
    tctx.drawImage(canvas, 0, 0, t.width, t.height);
    const b = await new Promise((res) =>
      t.toBlob(res, "image/jpeg", Number(jpegQ.value || 0.9))
    );
    return b;
  }

  function nameSafe(s) {
    return String(s || "")
      .replace(/[^A-Za-zА-Яа-яЁё_\- ]+/g, "") // only latin or cyrillic
      .trim()
      .replace(/\s+/g, "_"); // spaces → _
  }
  function detectMinMaxDates(rows) {
    let min = null,
      max = null;
    for (const r of rows) {
      const date = parseDate(r.Date);
      if (!date) continue;
      if (!min || date < min) min = date;
      if (!max || date > max) max = date;
    }
    return { min, max };
  }

  function renderList() {
    listEl.innerHTML = "";
    filtered.forEach((r) => {
      const div = document.createElement("div");
      div.className = "item";
      const left = document.createElement("div");
      left.className = "meta";
      const title = document.createElement("div");
      title.textContent = `${r.Date} · ${r.Person}`;
      const sub = document.createElement("small");
      sub.textContent = `${
        r.Assistant ? "Напарник: " + r.Assistant + " · " : ""
      }${r.Assignment || ""}`;
      left.appendChild(title);
      left.appendChild(sub);
      const btn = document.createElement("button");
      btn.textContent = "Показать";
      btn.addEventListener("click", () => drawRecordToCanvas(r));
      div.appendChild(left);
      div.appendChild(btn);
      listEl.appendChild(div);
    });
  }

  async function parseCsv(file) {
    const csvText = await file.text();
    const parsed = Papa.parse(csvText, { header: true, skipEmptyLines: true });
    return parsed.data
      .map((r) => ({
        Date: (r.Date || "").trim(),
        Person: (r.Person || "").trim(),
        PartType: (r.PartType || "").trim(),
        Assignment: (r.Assignment || "").trim(),
        School: String(r.School || "").trim(),
      }))
      .filter((r) => r.Date && r.Person);
  }

  function joinTasks(raw) {
    const allowed = new Set([
      "BibleReading",
      "Apply1",
      "Apply2",
      "Apply3",
      "Apply4",
    ]);
    const mains = new Map();
    const assistants = new Map();
    for (const r of raw) {
      const base = String(r.PartType || "")
        .replace(/Assistant/i, "")
        .trim();
      if (!allowed.has(base)) continue; //
      const key = `${r.Date}__${base}`;
      if (/Assistant/i.test(r.PartType || "")) {
        assistants.set(key, r);
      } else {
        mains.set(key, r);
      }
    }

    const res = [];
    for (const [key, main] of mains) {
      const a = assistants.get(key);

      // Date shift: default + check for special list
      let date = parseDate(main.Date);
      if (date) {
        const weekMonday = fmtDate(date);
        const offset =
          (specialOffsets[weekMonday] ?? defaultDateOffsetDays) | 0;
        date.setDate(date.getDate() + offset);
      }

      res.push({
        Date: fmtDate(date),
        DateDisplay: fmtDateHuman(date),
        Person: main.Person,
        Assistant: a ? a.Person : "",
        Assignment: main.Assignment,
        School: main.School,
      });
    }
    return res;
  }

  function applyFilter() {
    const from = parseDate(dateFrom.value);
    const to = parseDate(dateTo.value);
    filtered = joined
      .filter((r) => {
        const date = parseDate(r.Date);
        if (!date) return false;
        const okFrom = from
          ? date >=
            new Date(from.getFullYear(), from.getMonth(), from.getDate())
          : true;
        const okTo = to
          ? date <=
            new Date(
              to.getFullYear(),
              to.getMonth(),
              to.getDate(),
              23,
              59,
              59,
              999
            )
          : true;
        return okFrom && okTo;
      })
      .sort(
        (a, b) =>
          parseDate(a.Date) - parseDate(b.Date) ||
          a.Person.localeCompare(b.Person, "ru")
      );
    renderList();
  }

  // --- events ---
  btnParse.addEventListener("click", async () => {
    if (!(bgFile.files[0] && csvFile.files[0])) return;
    await loadBgImage(bgFile.files[0]);

    tasks = await parseCsv(csvFile.files[0]);
    joined = joinTasks(tasks);

    const { min, max } = detectMinMaxDates(joined);
    if (min) dateFrom.value = fmtDate(min);
    if (max) dateTo.value = fmtDate(max);

    applyFilter();
    btnGenerate.disabled = filtered.length === 0;
    btnExportZip.disabled = true;
  });

  dateFrom.addEventListener("change", () => {
    applyFilter();
    btnGenerate.disabled = filtered.length === 0;
  });
  dateTo.addEventListener("change", () => {
    applyFilter();
    btnGenerate.disabled = filtered.length === 0;
  });

  btnGenerate.addEventListener("click", async () => {
    if (!filtered.length) return;
    renderedItems.clear();
    drawRecordToCanvas(filtered[0]); // preview first

    const targetW = Number(wMobile.value || 1280);
    for (const rec of filtered) {
      drawRecordToCanvas(rec);
      const blob = await canvasToJpegBlob(targetW);
      const fname = `${rec.Date}_${nameSafe(rec.Person)}.jpg`;
      renderedItems.set(fname, blob);
    }
    btnExportZip.disabled = renderedItems.size === 0;
  });

  btnExportZip.addEventListener("click", async () => {
    if (!renderedItems.size) return;
    console.log({ renderedItems });
    const zip = new JSZip();
    for (const [name, blob] of renderedItems) {
      zip.file(name, blob);
    }
    const content = await zip.generateAsync({ type: "blob" });

    const url = URL.createObjectURL(content);
    const a = document.createElement("a");
    a.href = url;
    a.download = `S89_${fmtDate(new Date())}.zip`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  // --- sanity tests (run in console) ---
  // 1) parseDate formats
  console.assert(
    fmtDate(parseDate("2025-05-05")) === "2025-05-05",
    "parseDate ISO failed"
  );
  console.assert(
    fmtDate(parseDate("2025.05.05")) === "2025-05-05",
    "parseDate yyyy.mm.dd failed"
  );
  console.assert(
    fmtDate(parseDate("05.05.2025")) === "2025-05-05",
    "parseDate dd.mm.yyyy failed"
  );
})();
