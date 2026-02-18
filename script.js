(() => {
  "use strict";

  // ── DOM References ──
  const form = document.getElementById("eventForm");
  const outputSection = document.getElementById("outputSection");

  // ── Snippet state (preserved across style-option changes) ──
  let currentIcsContent = "";
  const allDayCheckbox = document.getElementById("allDay");
  const startTimeGroup = document.getElementById("startTimeGroup");
  const endTimeGroup = document.getElementById("endTimeGroup");
  const timezoneGroup = document.getElementById("timezoneGroup");
  const timezoneSelect = document.getElementById("timezone");
  const startDateInput = document.getElementById("startDate");
  const endDateInput = document.getElementById("endDate");
  const startTimeInput = document.getElementById("startTime");
  const endTimeInput = document.getElementById("endTime");
  const toast = document.getElementById("toast");

  // ── Initialise ──
  populateTimezones();
  setDefaults();
  bindEvents();

  // ── Timezone List ──
  function populateTimezones() {
    let zones;
    try {
      zones = Intl.supportedValuesOf("timeZone");
    } catch {
      zones = [
        "America/New_York", "America/Chicago", "America/Denver",
        "America/Los_Angeles", "America/Anchorage", "Pacific/Honolulu",
        "Europe/London", "Europe/Paris", "Europe/Berlin",
        "Asia/Tokyo", "Asia/Shanghai", "Asia/Kolkata",
        "Australia/Sydney", "Pacific/Auckland", "UTC"
      ];
    }

    const localZone = Intl.DateTimeFormat().resolvedOptions().timeZone;

    zones.forEach((tz) => {
      const opt = document.createElement("option");
      opt.value = tz;
      opt.textContent = tz.replace(/_/g, " ");
      if (tz === localZone) opt.selected = true;
      timezoneSelect.appendChild(opt);
    });
  }

  function setDefaults() {
    const now = new Date();
    const nextHour = new Date(now);
    nextHour.setHours(now.getHours() + 1, 0, 0, 0);
    const endTime = new Date(nextHour);
    endTime.setHours(endTime.getHours() + 1);

    startDateInput.value = formatLocalDate(nextHour);
    endDateInput.value = formatLocalDate(endTime);
    startTimeInput.value = formatLocalTime(nextHour);
    endTimeInput.value = formatLocalTime(endTime);
  }

  // ── Events ──
  function bindEvents() {
    allDayCheckbox.addEventListener("change", toggleAllDay);
    form.addEventListener("submit", handleSubmit);

    startDateInput.addEventListener("change", () => {
      if (endDateInput.value < startDateInput.value) {
        endDateInput.value = startDateInput.value;
      }
    });

    document.querySelectorAll(".btn-copy").forEach((btn) => {
      btn.addEventListener("click", handleCopyButton);
    });

    document.querySelectorAll(".btn-copy-url").forEach((btn) => {
      btn.addEventListener("click", handleCopyUrl);
    });

    document.querySelectorAll('input[name^="snippet"]').forEach((input) => {
      input.addEventListener("change", () => {
        if (!outputSection.hidden) refreshSnippet();
      });
    });
  }

  function toggleAllDay() {
    const isAllDay = allDayCheckbox.checked;
    startTimeGroup.style.display = isAllDay ? "none" : "";
    endTimeGroup.style.display = isAllDay ? "none" : "";
    timezoneGroup.style.display = isAllDay ? "none" : "";
    startTimeInput.required = !isAllDay;
    endTimeInput.required = !isAllDay;
  }

  function handleSubmit(e) {
    e.preventDefault();

    const event = readForm();
    if (!event) return;

    const googleUrl = buildGoogleUrl(event);
    const outlookUrl = buildOutlookUrl(event);
    const yahooUrl = buildYahooUrl(event);
    const icsContent = buildIcs(event);

    // Set link hrefs
    document.getElementById("googleLink").href = googleUrl;
    document.getElementById("outlookLink").href = outlookUrl;
    document.getElementById("yahooLink").href = yahooUrl;

    // Set embed URLs
    document.getElementById("googleUrl").value = googleUrl;
    document.getElementById("outlookUrl").value = outlookUrl;
    document.getElementById("yahooUrl").value = yahooUrl;

    // ICS download
    document.getElementById("icsDownload").onclick = () => downloadIcs(icsContent, event.title);

    // ICS copy
    document.getElementById("icsLinkCopy").onclick = () => {
      copyToClipboard(icsContent);
      showToast("ICS content copied!");
    };

    // HTML Snippet
    currentIcsContent = icsContent;
    refreshSnippet();

    // Show output
    outputSection.hidden = false;
    outputSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  // ── Read Form ──
  function readForm() {
    const title = document.getElementById("title").value.trim();
    const description = document.getElementById("description").value.trim();
    const location = document.getElementById("location").value.trim();
    const allDay = allDayCheckbox.checked;
    const timezone = timezoneSelect.value;

    const startDate = startDateInput.value;
    const endDate = endDateInput.value;
    const startTime = startTimeInput.value || "00:00";
    const endTime = endTimeInput.value || "23:59";

    if (!title || !startDate || !endDate) return null;

    return { title, description, location, allDay, timezone, startDate, endDate, startTime, endTime };
  }

  // ── Google Calendar ──
  // https://www.google.com/calendar/render?action=TEMPLATE&text=...&dates=START/END&details=...&location=...
  function buildGoogleUrl(ev) {
    const params = new URLSearchParams();
    params.set("action", "TEMPLATE");
    params.set("text", ev.title);

    if (ev.allDay) {
      const start = ev.startDate.replace(/-/g, "");
      // Google all-day end date is exclusive, so add one day
      const endExclusive = addDays(ev.endDate, 1).replace(/-/g, "");
      params.set("dates", `${start}/${endExclusive}`);
    } else {
      const start = toUtcString(ev.startDate, ev.startTime, ev.timezone);
      const end = toUtcString(ev.endDate, ev.endTime, ev.timezone);
      params.set("dates", `${start}/${end}`);
    }

    const details = descriptionWithMap(ev.description, ev.location);
    if (details) params.set("details", details);
    if (ev.location) params.set("location", ev.location);

    return `https://www.google.com/calendar/render?${params.toString()}`;
  }

  // ── Office 365 / Outlook ──
  // Uses the deeplink/compose endpoint (outlook.office.com)
  function buildOutlookUrl(ev) {
    const base = "https://outlook.office.com/calendar/deeplink/compose";
    const params = new URLSearchParams();
    params.set("path", "/calendar/action/compose");
    params.set("rru", "addevent");
    params.set("subject", ev.title);

    if (ev.allDay) {
      params.set("startdt", ev.startDate);
      params.set("enddt", addDays(ev.endDate, 1));
      params.set("allday", "true");
    } else {
      // Outlook deeplink expects UTC ISO dates with Z suffix
      params.set("startdt", toIsoUtc(ev.startDate, ev.startTime, ev.timezone));
      params.set("enddt", toIsoUtc(ev.endDate, ev.endTime, ev.timezone));
    }

    const body = descriptionWithMap(ev.description, ev.location);
    if (body) params.set("body", body);
    if (ev.location) params.set("location", ev.location);

    return `${base}?${params.toString()}`;
  }

  // ── Yahoo Calendar ──
  function buildYahooUrl(ev) {
    const params = new URLSearchParams();
    params.set("v", "60");
    params.set("title", ev.title);

    if (ev.allDay) {
      const start = ev.startDate.replace(/-/g, "");
      const end = addDays(ev.endDate, 1).replace(/-/g, "");
      params.set("st", start);
      params.set("et", end);
      params.set("dur", "allday");
    } else {
      params.set("st", toUtcString(ev.startDate, ev.startTime, ev.timezone));
      params.set("et", toUtcString(ev.endDate, ev.endTime, ev.timezone));
    }

    const desc = descriptionWithMap(ev.description, ev.location);
    if (desc) params.set("desc", desc);
    if (ev.location) params.set("in_loc", ev.location);

    return `https://calendar.yahoo.com/?${params.toString()}`;
  }

  // ── ICS File ──
  function buildIcs(ev) {
    const uid = `${Date.now()}-${Math.random().toString(36).slice(2)}@addtocalendar`;
    const now = formatIcsUtcNow();

    let dtStart, dtEnd;
    if (ev.allDay) {
      dtStart = `DTSTART;VALUE=DATE:${ev.startDate.replace(/-/g, "")}`;
      dtEnd = `DTEND;VALUE=DATE:${addDays(ev.endDate, 1).replace(/-/g, "")}`;
    } else {
      dtStart = `DTSTART:${toUtcString(ev.startDate, ev.startTime, ev.timezone)}`;
      dtEnd = `DTEND:${toUtcString(ev.endDate, ev.endTime, ev.timezone)}`;
    }

    const lines = [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//Add To Calendar Creator//EN",
      "CALSCALE:GREGORIAN",
      "METHOD:PUBLISH",
      "BEGIN:VEVENT",
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      dtStart,
      dtEnd,
      `SUMMARY:${escapeIcs(ev.title)}`,
    ];

    const icsDesc = descriptionWithMap(ev.description, ev.location);
    if (icsDesc) lines.push(`DESCRIPTION:${escapeIcs(icsDesc)}`);
    if (ev.location) lines.push(`LOCATION:${escapeIcs(ev.location)}`);

    lines.push("END:VEVENT", "END:VCALENDAR");

    return lines.join("\r\n");
  }

  function downloadIcs(content, title) {
    const blob = new Blob([content], { type: "text/calendar;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${slugify(title)}.ics`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast("ICS file downloaded!");
  }

  // ── Date / Time Helpers ──
  function formatLocalDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  function formatLocalTime(date) {
    const h = String(date.getHours()).padStart(2, "0");
    const m = String(date.getMinutes()).padStart(2, "0");
    return `${h}:${m}`;
  }

  function addDays(dateStr, n) {
    const d = new Date(dateStr + "T00:00:00");
    d.setDate(d.getDate() + n);
    return formatLocalDate(d);
  }

  /**
   * Convert a local date + time + timezone to UTC string: YYYYMMDDTHHmmssZ
   */
  function toUtcString(dateStr, timeStr, timezone) {
    const dtString = `${dateStr}T${timeStr}:00`;

    // Build a Date that represents the wall-clock time in the given timezone
    // by computing the UTC offset for that timezone at the given instant.
    const localDate = new Date(dtString);

    // Get the offset between the local system tz and the target tz
    const formatter = new Intl.DateTimeFormat("en-US", {
      timeZone: timezone,
      year: "numeric", month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit", second: "2-digit",
      hour12: false,
    });

    // Parse target-tz representation of an arbitrary reference point
    // and use it to derive the actual UTC instant.
    const parts = formatter.formatToParts(localDate);
    const get = (type) => {
      const p = parts.find((p) => p.type === type);
      return p ? p.value : "00";
    };

    const tzDate = new Date(
      `${get("year")}-${get("month")}-${get("day")}T${get("hour")}:${get("minute")}:${get("second")}`
    );

    // Difference between what JS thinks and what the target tz shows
    const diffMs = localDate.getTime() - tzDate.getTime();
    const utcMs = localDate.getTime() + diffMs;
    const utc = new Date(utcMs);

    const y = utc.getUTCFullYear();
    const mo = String(utc.getUTCMonth() + 1).padStart(2, "0");
    const d = String(utc.getUTCDate()).padStart(2, "0");
    const h = String(utc.getUTCHours()).padStart(2, "0");
    const mi = String(utc.getUTCMinutes()).padStart(2, "0");
    const s = String(utc.getUTCSeconds()).padStart(2, "0");

    return `${y}${mo}${d}T${h}${mi}${s}Z`;
  }

  function toIsoLocal(dateStr, timeStr) {
    return `${dateStr}T${timeStr}:00`;
  }

  /**
   * Convert local date + time + timezone to ISO 8601 UTC string: YYYY-MM-DDTHH:mm:ssZ
   * Used by the Outlook deeplink endpoint.
   */
  function toIsoUtc(dateStr, timeStr, timezone) {
    const compact = toUtcString(dateStr, timeStr, timezone); // YYYYMMDDTHHmmssZ
    const y = compact.slice(0, 4);
    const mo = compact.slice(4, 6);
    const d = compact.slice(6, 8);
    const h = compact.slice(9, 11);
    const mi = compact.slice(11, 13);
    const s = compact.slice(13, 15);
    return `${y}-${mo}-${d}T${h}:${mi}:${s}Z`;
  }

  function formatIcsUtcNow() {
    const d = new Date();
    return (
      d.getUTCFullYear().toString() +
      String(d.getUTCMonth() + 1).padStart(2, "0") +
      String(d.getUTCDate()).padStart(2, "0") +
      "T" +
      String(d.getUTCHours()).padStart(2, "0") +
      String(d.getUTCMinutes()).padStart(2, "0") +
      String(d.getUTCSeconds()).padStart(2, "0") +
      "Z"
    );
  }

  // ── Location / Maps Helpers ──
  function isUrl(str) {
    return /^https?:\/\//i.test(str);
  }

  function buildMapsUrl(location) {
    return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(location)}`;
  }

  /**
   * If the location is a physical address (not a URL), append a Google Maps
   * link to the description so calendar recipients can tap through to a map.
   */
  function descriptionWithMap(description, location) {
    if (!location || isUrl(location)) return description;
    const mapsLink = buildMapsUrl(location);
    const mapLine = `Map: ${mapsLink}`;
    return description ? `${description}\n\n${mapLine}` : mapLine;
  }

  // ── Snippet Refresh ──
  function getSnippetOpts() {
    const radio = (name) =>
      document.querySelector(`input[name="${name}"]:checked`).value;
    const look = radio("snippetLook");
    return {
      addTo:     radio("snippetText")  === "addto",
      fullWidth: radio("snippetWidth") === "full",
      outline:   look === "outline",
      icons:     look === "icons",
    };
  }

  function refreshSnippet() {
    const googleUrl  = document.getElementById("googleLink").href;
    const outlookUrl = document.getElementById("outlookLink").href;
    const yahooUrl   = document.getElementById("yahooLink").href;
    const snippet = buildHtmlSnippet(
      googleUrl, outlookUrl, yahooUrl, currentIcsContent, getSnippetOpts()
    );
    document.getElementById("htmlSnippet").value = snippet;
    document.getElementById("snippetPreview").innerHTML = snippet;
    document.getElementById("copyHtmlSnippet").onclick = () => {
      copyToClipboard(snippet);
      showToast("HTML snippet copied!");
    };
  }

  // ── HTML Snippet Builder ──
  /**
   * Escape a string for use inside an HTML attribute value (double-quoted).
   */
  function escHtmlAttr(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  // ── Icon Snippet Builder ──
  /**
   * Shared helper: 45×45 rounded-square icon with a centred brand logo.
   * @param {string} bg    – background fill colour
   * @param {string} paths – inner SVG elements (paths, rects, …)
   * @param {number} vw    – source viewBox width
   * @param {number} vh    – source viewBox height
   */
  function svgBrandIcon(bg, paths, vw, vh) {
    const size  = 20;
    const scale = size / Math.max(vw, vh);
    const tx    = (45 - vw * scale) / 2;
    const ty    = (45 - vh * scale) / 2;
    const svg =
      `<svg xmlns="http://www.w3.org/2000/svg" width="45" height="45" viewBox="0 0 45 45">` +
      `<rect width="45" height="45" rx="8" fill="${bg}"/>` +
      `<g transform="translate(${tx.toFixed(2)},${ty.toFixed(2)}) scale(${scale.toFixed(4)})">` +
      paths +
      `</g>` +
      `</svg>`;
    return "data:image/svg+xml;base64," + btoa(svg);
  }

  /** Google Calendar icon: white background + 4-colour Google G logo. */
  function googleIconUri() {
    return svgBrandIcon("#ffffff",
      `<path fill="#4285f4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/>` +
      `<path fill="#34a853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/>` +
      `<path fill="#fbbc05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l3.66-2.84z"/>` +
      `<path fill="#ea4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/>`,
      24, 24);
  }

  /** Office 365 icon: white background + Microsoft 4-colour tile logo. */
  function outlookIconUri() {
    return svgBrandIcon("#ffffff",
      `<rect x="0"    y="0"    width="9.5" height="9.5" fill="#f25022"/>` +
      `<rect x="11.5" y="0"    width="9.5" height="9.5" fill="#7fba00"/>` +
      `<rect x="0"    y="11.5" width="9.5" height="9.5" fill="#00a4ef"/>` +
      `<rect x="11.5" y="11.5" width="9.5" height="9.5" fill="#ffb900"/>`,
      21, 21);
  }

  /** Apple / ICS icon: dark background + white Apple logo silhouette. */
  function appleIconUri() {
    return svgBrandIcon("#1c1c1e",
      `<path fill="#ffffff" d="M12.152 6.896c-.948 0-2.415-1.078-3.96-1.04-2.04.027-3.91 1.183-4.961 3.014-2.117 3.675-.546 9.103 1.519 12.09 1.013 1.459 2.208 3.09 3.792 3.039 1.52-.065 2.09-.987 3.935-.987 1.831 0 2.35.987 3.96.948 1.637-.026 2.676-1.48 3.676-2.948 1.156-1.688 1.636-3.325 1.662-3.415-.039-.013-3.182-1.221-3.22-4.857-.026-3.04 2.48-4.494 2.597-4.559-1.429-2.09-3.623-2.324-4.39-2.376-2-.156-3.675 1.09-4.61 1.09zM15.53 3.83c.843-1.012 1.4-2.427 1.245-3.83-1.207.052-2.662.805-3.532 1.818-.78.896-1.454 2.338-1.273 3.714 1.338.104 2.715-.688 3.559-1.701"/>`,
      24, 24);
  }

  /** Yahoo icon: purple background + bold "Y!" text. */
  function yahooIconUri() {
    const svg =
      `<svg xmlns="http://www.w3.org/2000/svg" width="45" height="45" viewBox="0 0 45 45">` +
      `<rect width="45" height="45" rx="8" fill="#6001d2"/>` +
      `<text x="22.5" y="31" font-family="Arial,sans-serif" font-size="20" font-weight="bold" fill="#fff" text-anchor="middle">Y!</text>` +
      `</svg>`;
    return "data:image/svg+xml;base64," + btoa(svg);
  }

  /**
   * Build the email-safe icon-style snippet.
   * Mirrors the AddEvent.com pattern: optional <p> title, then
   * <p style="font-size:0"> containing inline <a><img></a> links.
   */
  function buildIconSnippet(googleUrl, outlookUrl, yahooUrl, icsContent, opts) {
    const { addTo } = opts;
    const icsHref = "data:text/calendar;charset=utf-8," + encodeURIComponent(icsContent);

    const imgStyle = "width:45px;height:45px;display:inline;margin:0 4px;";
    const aStyle   = "display:inline;";

    const icons = [
      { href: googleUrl,  src: googleIconUri(),  alt: "Google Calendar", download: false },
      { href: outlookUrl, src: outlookIconUri(), alt: "Office 365",      download: false },
      { href: yahooUrl,   src: yahooIconUri(),   alt: "Yahoo Calendar",  download: false },
      { href: icsHref,    src: appleIconUri(),   alt: "Apple / ICS",     download: true  },
    ];

    const iconLinks = icons.map(({ href, src, alt, download }) => {
      const extra = download ? ` download="event.ics"` : ` target="_blank" rel="noopener"`;
      return (
        `<a href="${escHtmlAttr(href)}"${extra} title="${alt}" style="${aStyle}">` +
        `<img src="${src}" alt="${alt}" width="45" height="45" border="0" style="${imgStyle}" />` +
        `</a>`
      );
    }).join("\n  ");

    const titleP = addTo
      ? `<p style="margin:0 0 10px 0;text-align:center;font-size:16px;font-weight:bold;font-family:sans-serif;color:#000000;">Add to your calendar</p>\n`
      : "";

    return (
      titleP +
      `<p style="margin:0;text-align:center;font-size:0;">\n` +
      `  ${iconLinks}\n` +
      `</p>`
    );
  }

  /**
   * Build a self-contained HTML snippet with inline-styled "Add to Calendar"
   * buttons for all four providers. The ICS button uses a data: URI so it
   * works without a server.
   *
   * opts: { addTo: bool, fullWidth: bool, outline: bool, icons: bool }
   */
  function buildHtmlSnippet(googleUrl, outlookUrl, yahooUrl, icsContent, opts = {}) {
    if (opts.icons) {
      return buildIconSnippet(googleUrl, outlookUrl, yahooUrl, icsContent, opts);
    }

    const { addTo = true, fullWidth = false, outline = false } = opts;

    const icsHref =
      "data:text/calendar;charset=utf-8," + encodeURIComponent(icsContent);

    const prefix = addTo ? "Add to " : "";
    const entries = [
      { href: googleUrl,  color: "#4285f4", label: prefix + "Google Calendar",  download: false },
      { href: outlookUrl, color: "#0078d4", label: prefix + "Office\u00a0365",  download: false },
      { href: yahooUrl,   color: "#6001d2", label: prefix + "Yahoo Calendar",   download: false },
      { href: icsHref,    color: "#1c1c1e", label: prefix + "Apple\u00a0/\u00a0ICS", download: true },
    ];

    const base =
      "text-decoration:none;border-radius:6px;font-size:14px;" +
      "font-weight:600;font-family:sans-serif;line-height:1.2;text-align:center;" +
      (fullWidth
        ? "display:block;width:100%;padding:10px 20px;box-sizing:border-box;"
        : "display:inline-block;padding:10px 20px;");

    function btnStyle(color) {
      return outline
        ? `${base}color:${color};background:#ffffff;border:2px solid ${color};`
        : `${base}color:#ffffff;background:${color};border:none;`;
    }

    const wrapStyle = fullWidth
      ? "display:flex;flex-direction:column;gap:8px;"
      : "display:flex;flex-wrap:wrap;gap:8px;";

    const buttons = entries.map(({ href, color, label, download }) => {
      const extra = download
        ? ' download="event.ics"'
        : ' target="_blank" rel="noopener"';
      return `<a href="${escHtmlAttr(href)}"${extra} style="${btnStyle(color)}">${label}</a>`;
    }).join("\n  ");

    return `<div style="${wrapStyle}">\n  ${buttons}\n</div>`;
  }

  // ── String Helpers ──
  function escapeIcs(str) {
    return str
      .replace(/\\/g, "\\\\")
      .replace(/;/g, "\\;")
      .replace(/,/g, "\\,")
      .replace(/\n/g, "\\n");
  }

  function slugify(str) {
    return str
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/(^-|-$)/g, "")
      .slice(0, 50) || "event";
  }

  // ── Clipboard ──
  function copyToClipboard(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text);
    } else {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
    }
  }

  function handleCopyButton(e) {
    const btn = e.currentTarget;
    const targetId = btn.dataset.target;
    const anchor = document.getElementById(targetId);
    if (anchor && anchor.href) {
      copyToClipboard(anchor.href);
      showToast("Link copied to clipboard!");
    }
  }

  function handleCopyUrl(e) {
    const btn = e.currentTarget;
    const inputId = btn.dataset.input;
    const input = document.getElementById(inputId);
    if (input) {
      copyToClipboard(input.value);
      showToast("URL copied to clipboard!");
    }
  }

  // ── Toast ──
  let toastTimer;
  function showToast(msg) {
    toast.textContent = msg;
    toast.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toast.classList.remove("show"), 2400);
  }
})();
