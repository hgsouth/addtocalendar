(() => {
  "use strict";

  // ── DOM References ──
  const form = document.getElementById("eventForm");
  const outputSection = document.getElementById("outputSection");
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

    if (ev.description) params.set("details", ev.description);
    if (ev.location) params.set("location", ev.location);

    return `https://www.google.com/calendar/render?${params.toString()}`;
  }

  // ── Office 365 / Outlook ──
  function buildOutlookUrl(ev) {
    const base = "https://outlook.live.com/calendar/0/action/compose";
    const params = new URLSearchParams();
    params.set("rru", "addevent");
    params.set("subject", ev.title);

    if (ev.allDay) {
      params.set("startdt", ev.startDate);
      params.set("enddt", addDays(ev.endDate, 1));
      params.set("allday", "true");
    } else {
      params.set("startdt", toIsoLocal(ev.startDate, ev.startTime));
      params.set("enddt", toIsoLocal(ev.endDate, ev.endTime));
    }

    if (ev.description) params.set("body", ev.description);
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

    if (ev.description) params.set("desc", ev.description);
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

    if (ev.description) lines.push(`DESCRIPTION:${escapeIcs(ev.description)}`);
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
