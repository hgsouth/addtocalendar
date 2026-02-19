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

    document.getElementById("snippetContent")?.addEventListener("change", () => {
      if (!outputSection.hidden) refreshSnippet();
    });
    document.getElementById("snippetAccent")?.addEventListener("input", () => {
      if (!outputSection.hidden) refreshSnippet();
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

    const googleUrl      = buildGoogleUrl(event);
    const outlookLiveUrl = buildOutlookLiveUrl(event);
    const outlookUrl     = buildOutlookUrl(event);
    const yahooUrl       = buildYahooUrl(event);
    const icsContent     = buildIcs(event);

    // Set link hrefs
    document.getElementById("googleLink").href      = googleUrl;
    document.getElementById("outlookLiveLink").href  = outlookLiveUrl;
    document.getElementById("outlookLink").href      = outlookUrl;
    document.getElementById("yahooLink").href        = yahooUrl;

    // Set embed URLs
    document.getElementById("googleUrl").value      = googleUrl;
    document.getElementById("outlookLiveUrl").value  = outlookLiveUrl;
    document.getElementById("outlookUrl").value      = outlookUrl;
    document.getElementById("yahooUrl").value        = yahooUrl;

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

  // ── Outlook URL helpers ──
  // Shared builder — base differs for personal vs work/school accounts.
  function _buildOutlookComposeUrl(base, ev) {
    const params = new URLSearchParams();
    params.set("rru", "addevent");
    params.set("path", "/calendar/action/compose");
    params.set("subject", ev.title);

    if (ev.allDay) {
      params.set("startdt", ev.startDate);
      params.set("enddt", addDays(ev.endDate, 1));
      params.set("allday", "true");
    } else {
      params.set("startdt", toIsoUtc(ev.startDate, ev.startTime, ev.timezone));
      params.set("enddt", toIsoUtc(ev.endDate, ev.endTime, ev.timezone));
    }

    const body = descriptionWithMap(ev.description, ev.location);
    if (body) params.set("body", body);
    if (ev.location) params.set("location", ev.location);

    return `${base}?${params.toString()}`;
  }

  /** Outlook.com — personal Microsoft accounts */
  function buildOutlookLiveUrl(ev) {
    return _buildOutlookComposeUrl(
      "https://outlook.live.com/calendar/0/action/compose", ev);
  }

  /** Office 365 / Microsoft 365 — work & school accounts */
  function buildOutlookUrl(ev) {
    return _buildOutlookComposeUrl(
      "https://outlook.office.com/calendar/0/action/compose", ev);
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
    const radio  = (name) => document.querySelector(`input[name="${name}"]:checked`)?.value;
    const select = (id)   => document.getElementById(id)?.value;
    return {
      content:     select("snippetContent") ?? "icon-text",
      outline:     radio("snippetStyle")    === "outline",
      shape:       radio("snippetShape")    ?? "rounded",
      size:        radio("snippetSize")     ?? "36",
      align:       radio("snippetAlign")    ?? "center",
      colorMode:   radio("snippetColor")    ?? "native",
      accentColor: document.getElementById("snippetAccent")?.value ?? "#0f766e",
      addTo:       radio("snippetText")     === "addto",
      fullWidth:   radio("snippetWidth")    === "full",
    };
  }

  function refreshSnippet() {
    const googleUrl      = document.getElementById("googleLink").href;
    const outlookLiveUrl = document.getElementById("outlookLiveLink").href;
    const outlookUrl     = document.getElementById("outlookLink").href;
    const yahooUrl       = document.getElementById("yahooLink").href;
    const snippet = buildHtmlSnippet(
      googleUrl, outlookLiveUrl, outlookUrl, yahooUrl, currentIcsContent, getSnippetOpts()
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
  function svgBrandIcon(bg, paths, vw, vh, { rx = 8, fillOverride = null } = {}) {
    const size  = 20;
    const scale = size / Math.max(vw, vh);
    const tx    = (45 - vw * scale) / 2;
    const ty    = (45 - vh * scale) / 2;
    const rendered = fillOverride
      ? paths.replace(/fill="[^"]*"/g, `fill="${fillOverride}"`)
      : paths;
    const svg =
      `<svg xmlns="http://www.w3.org/2000/svg" width="45" height="45" viewBox="0 0 45 45">` +
      `<rect width="45" height="45" rx="${rx}" fill="${bg}"/>` +
      `<g transform="translate(${tx.toFixed(2)},${ty.toFixed(2)}) scale(${scale.toFixed(4)})">` +
      rendered +
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

  /** Apple / ICS icon: no background + black Apple logo silhouette. */
  function appleIconUri() {
    return svgBrandIcon("transparent",
      `<path fill="#000000" d="M12.152 6.896c-.948 0-2.415-1.078-3.96-1.04-2.04.027-3.91 1.183-4.961 3.014-2.117 3.675-.546 9.103 1.519 12.09 1.013 1.459 2.208 3.09 3.792 3.039 1.52-.065 2.09-.987 3.935-.987 1.831 0 2.35.987 3.96.948 1.637-.026 2.676-1.48 3.676-2.948 1.156-1.688 1.636-3.325 1.662-3.415-.039-.013-3.182-1.221-3.22-4.857-.026-3.04 2.48-4.494 2.597-4.559-1.429-2.09-3.623-2.324-4.39-2.376-2-.156-3.675 1.09-4.61 1.09zM15.53 3.83c.843-1.012 1.4-2.427 1.245-3.83-1.207.052-2.662.805-3.532 1.818-.78.896-1.454 2.338-1.273 3.714 1.338.104 2.715-.688 3.559-1.701"/>`,
      24, 24);
  }

  // ── Path data reused across icon modes ──

  const _GOOGLE_PATHS =
    `<path fill="#4285f4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/>` +
    `<path fill="#34a853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/>` +
    `<path fill="#fbbc05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l3.66-2.84z"/>` +
    `<path fill="#ea4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/>`;

  // Outlook.com icon — viewBox 0 0 103.17322 104.31332.
  // style="fill:..." converted to fill="...", Inkscape attrs stripped.
  const _OUTLOOK_PATHS =
    `<path fill="#0072c6" d="m64.566509,22.116383v20.404273l7.130526,4.489881c0.188058,0.05485,0.595516,0.05877,0.783574,0L103.16929,26.320259c0,-2.44867,-2.28412,-4.203876,-3.573094,-4.203876H64.566509z"/>` +
    `<path fill="#0072c6" d="m64.566509,50.13308l6.507584,4.470291c0.916782,0.673874,2.021622,0,2.021622,0-1.100922,0.673874,30.077495,-20.035993,30.077495,-20.035993v37.501863c0,4.082422,-2.61322,5.794531,-5.551621,5.794531H64.562591V50.13308z"/>` +
    `<g transform="matrix(3.9178712,0,0,3.9178712,-13.481403,-41.384473)">` +
    `<path fill="#0072c6" d="m11.321,20.958c-0.566,0-1.017,0.266-1.35,0.797-0.333,0.531-0.5,1.234-0.5,2.109,0,0.888,0.167,1.59,0.5,2.106,0.333,0.517,0.77,0.774,1.31,0.774,0.557,0,0.999,-0.251,1.325,-0.753,0.326,-0.502,0.49,-1.199,0.49,-2.09,0,-0.929,-0.158,-1.652,-0.475,-2.169-0.317,-0.516-0.75,-0.774-1.3,-0.774z"/>` +
    `<path fill="#0072c6" d="m3.441,13.563v20.375l15.5,3.25V10.563l-15.5,3zm10.372,13.632c-0.655,0.862-1.509,1.294-2.563,1.294-1.027,0-1.863,-0.418-2.51,-1.253C8.094,26.4,7.77,25.312,7.77,23.97c0,-1.417,0.328,-2.563,0.985,-3.438,0.657,-0.875,1.527,-1.313,2.61,-1.313,1.023,0,1.851,0.418,2.482,1.256,0.632,0.838,0.948,1.942,0.948,3.313,0.001,1.409,-0.327,2.545,-0.982,3.407z"/>` +
    `</g>`;

  // Microsoft 365 icon — normalized to a 38.94 × 46.913 viewBox.
  // Original path coordinates offset by (-469.87, -633.186).
  const _M365_PATHS =
    `<path fill="#C84B3C" fill-rule="evenodd" d="m0,37.844,0-28.52,25.229-9.3238,13.711,4.3877,0,38.392-13.711,4.133-25.229-9.0691,25.229,3.0361,0-33.201-16.454,3.8392,0,22.487z"/>`;

  const _APPLE_PATHS =
    `<path fill="#000000" d="M12.152 6.896c-.948 0-2.415-1.078-3.96-1.04-2.04.027-3.91 1.183-4.961 3.014-2.117 3.675-.546 9.103 1.519 12.09 1.013 1.459 2.208 3.09 3.792 3.039 1.52-.065 2.09-.987 3.935-.987 1.831 0 2.35.987 3.96.948 1.637-.026 2.676-1.48 3.676-2.948 1.156-1.688 1.636-3.325 1.662-3.415-.039-.013-3.182-1.221-3.22-4.857-.026-3.04 2.48-4.494 2.597-4.559-1.429-2.09-3.623-2.324-4.39-2.376-2-.156-3.675 1.09-4.61 1.09zM15.53 3.83c.843-1.012 1.4-2.427 1.245-3.83-1.207.052-2.662.805-3.532 1.818-.78.896-1.454 2.338-1.273 3.714 1.338.104 2.715-.688 3.559-1.701"/>`;

  /**
   * Compact icon SVG at `size × size`, no background rect.
   * If `fillOverride` is given, all fill attributes are replaced with that colour.
   */
  function svgInlineIcon(paths, vw, vh, size, fillOverride) {
    const scale    = size / Math.max(vw, vh);
    const tx       = (size - vw * scale) / 2;
    const ty       = (size - vh * scale) / 2;
    const rendered = fillOverride
      ? paths.replace(/fill="[^"]*"/g, `fill="${fillOverride}"`)
      : paths;
    return "data:image/svg+xml;base64," + btoa(
      `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">` +
      `<g transform="translate(${tx.toFixed(2)},${ty.toFixed(2)}) scale(${scale.toFixed(4)})">` +
      rendered + `</g></svg>`
    );
  }

  /** Google G: multi-colour paths, or all-white when `white` is true. */
  function googleInlineIconUri(size, white) {
    return svgInlineIcon(_GOOGLE_PATHS, 24, 24, size, white ? "#ffffff" : null);
  }
  /** Microsoft 4-colour tiles (Outlook.com), or all-white. */
  function outlookInlineIconUri(size, white) {
    return svgInlineIcon(_OUTLOOK_PATHS, 103.17322, 104.31332, size, white ? "#ffffff" : null);
  }
  /** Microsoft 365 icon, or all-white. */
  function m365InlineIconUri(size, white) {
    return svgInlineIcon(_M365_PATHS, 38.94, 46.913, size, white ? "#ffffff" : null);
  }
  /** Yahoo "Y!" text, white or native purple. */
  function yahooInlineIconUri(size, white) {
    const color = white ? "#ffffff" : "#6001d2";
    const cy    = (size * 0.78).toFixed(1);
    const fs    = (size * 0.72).toFixed(1);
    return "data:image/svg+xml;base64," + btoa(
      `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">` +
      `<text x="${size / 2}" y="${cy}" font-family="Arial,sans-serif" font-size="${fs}" ` +
      `font-weight="bold" fill="${color}" text-anchor="middle">Y!</text></svg>`
    );
  }
  /** Apple logo silhouette, white or black. */
  function appleInlineIconUri(size, white) {
    return svgInlineIcon(_APPLE_PATHS, 24, 24, size, white ? "#ffffff" : "#000000");
  }

  /** Yahoo icon: no background, purple "Y!" text. */
  function yahooIconUri() {
    const svg =
      `<svg xmlns="http://www.w3.org/2000/svg" width="45" height="45" viewBox="0 0 45 45">` +
      `<text x="22.5" y="31" font-family="Arial,sans-serif" font-size="20" font-weight="bold" fill="#6001d2" text-anchor="middle">Y!</text>` +
      `</svg>`;
    return "data:image/svg+xml;base64," + btoa(svg);
  }

  /**
   * Build the email-safe icon-style snippet.
   * Mirrors the AddEvent.com pattern: optional <p> title, then
   * <p style="font-size:0"> containing inline <a><img></a> links.
   */
  function buildIconSnippet(googleUrl, outlookLiveUrl, outlookUrl, yahooUrl, icsContent, opts) {
    const { addTo, align = "center", shape = "rounded", size = "36",
            colorMode = "native", accentColor = "#0f766e" } = opts;
    const icsHref = "data:text/calendar;charset=utf-8," + encodeURIComponent(icsContent);
    const ta = align === "left" ? "left" : "center";

    const ICON_PX = { "28": 32, "32": 38, "36": 45, "40": 50, "44": 56 };
    const SVG_RX  = { square: 0, rounded: 8, pill: 22.5 };
    const px     = ICON_PX[size] ?? 45;
    const radius = _RADIUS_MAP[shape] ?? "8px";
    const rx     = SVG_RX[shape] ?? 8;

    const useCustom = colorMode === "custom";
    const border    = useCustom ? "none" : "1px solid #e0e0e0";

    const imgStyle = `width:${px}px;height:${px}px;display:inline;margin:0 4px;border:${border};border-radius:${radius};transition:box-shadow .15s;`;
    const aStyle   = "display:inline;text-decoration:none;";

    // Yahoo is text-based, so needs its own SVG builder
    function yahooIcon(bg, textColor) {
      const bgRect = bg === "transparent" ? "" : `<rect width="45" height="45" rx="${rx}" fill="${bg}"/>`;
      return "data:image/svg+xml;base64," + btoa(
        `<svg xmlns="http://www.w3.org/2000/svg" width="45" height="45" viewBox="0 0 45 45">` +
        bgRect +
        `<text x="22.5" y="31" font-family="Arial,sans-serif" font-size="20" font-weight="bold" fill="${textColor}" text-anchor="middle">Y!</text>` +
        `</svg>`
      );
    }

    const io = { rx };                                      // shared icon opts
    const iw = { rx, fillOverride: "#ffffff" };             // white fill icon opts

    const icons = useCustom
      ? [
          { href: googleUrl,      src: svgBrandIcon(accentColor, _GOOGLE_PATHS,  24,    24,     iw),  alt: "Google Calendar", download: false },
          { href: outlookLiveUrl,  src: svgBrandIcon(accentColor, _OUTLOOK_PATHS, 103.17322, 104.31332, iw),  alt: "Outlook.com",     download: false },
          { href: outlookUrl,      src: svgBrandIcon(accentColor, _M365_PATHS,   38.94, 46.913, iw),  alt: "Office 365",      download: false },
          { href: yahooUrl,        src: yahooIcon(accentColor, "#ffffff"),                              alt: "Yahoo Calendar",  download: false },
          { href: icsHref,         src: svgBrandIcon(accentColor, _APPLE_PATHS,   24,    24,     iw),  alt: "Apple / ICS",     download: true  },
        ]
      : [
          { href: googleUrl,      src: svgBrandIcon("#ffffff", _GOOGLE_PATHS,  24,    24,     io),  alt: "Google Calendar", download: false },
          { href: outlookLiveUrl,  src: svgBrandIcon("#ffffff", _OUTLOOK_PATHS, 103.17322, 104.31332, io),  alt: "Outlook.com",     download: false },
          { href: outlookUrl,      src: svgBrandIcon("#ffffff", _M365_PATHS,   38.94, 46.913, io),  alt: "Office 365",      download: false },
          { href: yahooUrl,        src: yahooIcon("#ffffff", "#6001d2"),                              alt: "Yahoo Calendar",  download: false },
          { href: icsHref,         src: svgBrandIcon("#ffffff", _APPLE_PATHS,   24,    24,     io),  alt: "Apple / ICS",     download: true  },
        ];

    const iconLinks = icons.map(({ href, src, alt, download }) => {
      const extra = download ? ` download="event.ics"` : ` target="_blank" rel="noopener"`;
      return (
        `<a href="${escHtmlAttr(href)}"${extra} title="${alt}" style="${aStyle}">` +
        `<img src="${src}" alt="${alt}" width="${px}" height="${px}" border="0" class="atc-icon" style="${imgStyle}" />` +
        `</a>`
      );
    }).join("\n  ");

    const titleP = addTo
      ? `<p style="margin:0 0 10px 0;text-align:${ta};font-size:16px;font-weight:bold;font-family:sans-serif;color:#000000;">Add to your calendar</p>\n`
      : "";

    const styleBlock = `<style>.atc-icon:hover{box-shadow:0 3px 10px rgba(0,0,0,.15)}</style>\n`;
    return (
      styleBlock +
      titleP +
      `<p style="margin:0;text-align:${ta};font-size:0;">\n` +
      `  ${iconLinks}\n` +
      `</p>`
    );
  }

  const _RADIUS_MAP = { square: "0px", rounded: "8px", pill: "999px" };
  const _SIZE_MAP = {
    "28": { h: "28px", fs: "12px", lp: "0 10px", gap: "7px",  icon: 14 },
    "32": { h: "32px", fs: "13px", lp: "0 12px", gap: "8px",  icon: 16 },
    "36": { h: "36px", fs: "14px", lp: "0 14px", gap: "9px",  icon: 18 },
    "40": { h: "40px", fs: "15px", lp: "0 16px", gap: "10px", icon: 20 },
    "44": { h: "44px", fs: "16px", lp: "0 18px", gap: "11px", icon: 22 },
  };

  /**
   * Build a combined icon + text button snippet.
   * Each provider gets one <a> with a small inline icon and a label.
   */
  function buildIconTextSnippet(googleUrl, outlookLiveUrl, outlookUrl, yahooUrl, icsContent, opts) {
    const {
      addTo       = true,
      fullWidth   = false,
      outline     = false,
      shape       = "rounded",
      size        = "36",
      align       = "center",
      colorMode   = "native",
      accentColor = "#0f766e",
    } = opts;

    const icsHref        = "data:text/calendar;charset=utf-8," + encodeURIComponent(icsContent);
    const radius         = _RADIUS_MAP[shape] ?? "8px";
    const { h, fs, lp, gap, icon } = _SIZE_MAP[size] ?? _SIZE_MAP["36"];
    const useWhiteIcons  = !outline;
    const prefix         = addTo ? "Add to " : "";
    const justifyContent = align === "left" ? "flex-start" : "center";

    const entries = [
      { href: googleUrl,      color: "#4285f4", label: prefix + "Google Calendar",        iconUri: googleInlineIconUri(icon, useWhiteIcons),  download: false },
      { href: outlookLiveUrl,  color: "#0072c6", label: prefix + "Outlook.com",            iconUri: outlookInlineIconUri(icon, useWhiteIcons), download: false },
      { href: outlookUrl,      color: "#C84B3C", label: prefix + "Office\u00a0365",        iconUri: m365InlineIconUri(icon, useWhiteIcons),    download: false },
      { href: yahooUrl,        color: "#6001d2", label: prefix + "Yahoo Calendar",         iconUri: yahooInlineIconUri(icon, useWhiteIcons),   download: false },
      { href: icsHref,         color: "#1c1c1e", label: prefix + "Apple\u00a0/\u00a0ICS", iconUri: appleInlineIconUri(icon, useWhiteIcons),   download: true  },
    ];

    const wrapStyle = fullWidth
      ? `display:flex;flex-direction:column;gap:8px;align-items:${align === "left" ? "flex-start" : "center"};`
      : `display:flex;flex-wrap:wrap;gap:8px;justify-content:${justifyContent};`;

    const buttons = entries.map(({ href, color, label, iconUri, download }) => {
      const bg    = colorMode === "custom" ? accentColor : color;
      const extra = download ? ' download="event.ics"' : ' target="_blank" rel="noopener"';
      const btnStyle = outline
        ? `display:inline-flex;align-items:center;gap:${gap};background:transparent;` +
          `border:2px solid ${bg};color:${bg};border-radius:${radius};padding:${lp};` +
          `height:${h};box-sizing:border-box;text-decoration:none;font-size:${fs};` +
          `font-weight:600;font-family:sans-serif;white-space:nowrap;`
        : `display:inline-flex;align-items:center;gap:${gap};background:${bg};` +
          `border:none;color:#ffffff;border-radius:${radius};padding:${lp};` +
          `height:${h};box-sizing:border-box;text-decoration:none;font-size:${fs};` +
          `font-weight:600;font-family:sans-serif;white-space:nowrap;`;
      return (
        `<a href="${escHtmlAttr(href)}"${extra} style="${btnStyle}">` +
        `<img src="${iconUri}" width="${icon}" height="${icon}" alt="" style="display:block;flex-shrink:0;"/>` +
        label +
        `</a>`
      );
    }).join("\n  ");

    return `<div style="${wrapStyle}">\n  ${buttons}\n</div>`;
  }

  /**
   * Main router: delegates to the appropriate snippet builder based on opts.content.
   * opts: { content, outline, shape, size, align, colorMode, accentColor, addTo, fullWidth }
   */
  function buildHtmlSnippet(googleUrl, outlookLiveUrl, outlookUrl, yahooUrl, icsContent, opts = {}) {
    const content = opts.content ?? "icon-text";
    if (content === "icons-only") {
      return buildIconSnippet(googleUrl, outlookLiveUrl, outlookUrl, yahooUrl, icsContent, opts);
    }
    if (content === "icon-text") {
      return buildIconTextSnippet(googleUrl, outlookLiveUrl, outlookUrl, yahooUrl, icsContent, opts);
    }

    // ── text-only ──
    const {
      addTo       = true,
      fullWidth   = false,
      outline     = false,
      shape       = "rounded",
      size        = "36",
      align       = "center",
      colorMode   = "native",
      accentColor = "#0f766e",
    } = opts;

    const icsHref        = "data:text/calendar;charset=utf-8," + encodeURIComponent(icsContent);
    const radius         = _RADIUS_MAP[shape] ?? "8px";
    const { h, fs, lp } = _SIZE_MAP[size] ?? _SIZE_MAP["36"];
    const prefix         = addTo ? "Add to " : "";
    const justifyContent = align === "left" ? "flex-start" : "center";

    const entries = [
      { href: googleUrl,      color: "#4285f4", label: prefix + "Google Calendar",        download: false },
      { href: outlookLiveUrl,  color: "#0072c6", label: prefix + "Outlook.com",            download: false },
      { href: outlookUrl,      color: "#C84B3C", label: prefix + "Office\u00a0365",        download: false },
      { href: yahooUrl,        color: "#6001d2", label: prefix + "Yahoo Calendar",         download: false },
      { href: icsHref,         color: "#1c1c1e", label: prefix + "Apple\u00a0/\u00a0ICS", download: true  },
    ];

    function btnStyle(color) {
      const bg   = colorMode === "custom" ? accentColor : color;
      const base =
        `display:inline-flex;align-items:center;justify-content:center;` +
        `height:${h};padding:${lp};box-sizing:border-box;border-radius:${radius};` +
        `text-decoration:none;font-size:${fs};font-weight:600;font-family:sans-serif;` +
        (fullWidth ? "width:100%;" : "");
      return outline
        ? `${base}color:${bg};background:#ffffff;border:2px solid ${bg};`
        : `${base}color:#ffffff;background:${bg};border:none;`;
    }

    const wrapStyle = fullWidth
      ? `display:flex;flex-direction:column;gap:8px;`
      : `display:flex;flex-wrap:wrap;gap:8px;justify-content:${justifyContent};`;

    const buttons = entries.map(({ href, color, label, download }) => {
      const extra = download ? ' download="event.ics"' : ' target="_blank" rel="noopener"';
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
