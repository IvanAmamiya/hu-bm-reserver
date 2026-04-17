const venueSelect = document.getElementById("venueSelect");
const organizationInput = document.getElementById("organizationInput");
const applicantInput = document.getElementById("applicantInput");
const dateInput = document.getElementById("dateInput");
const daysInput = document.getElementById("daysInput");
const loadBtn = document.getElementById("loadBtn");
const exportBtn = document.getElementById("exportBtn");
const statusEl = document.getElementById("status");
const slotsEl = document.getElementById("slots");

const API_BASE_URL = String(window.APP_CONFIG?.API_BASE_URL || "").trim().replace(/\/$/, "");

function apiUrl(path) {
  return API_BASE_URL ? `${API_BASE_URL}${path}` : path;
}

let currentVenue = null;

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

dateInput.value = todayISO();

function setStatus(text) {
  statusEl.textContent = text;
}

function selectedVenueName() {
  const option = venueSelect.options[venueSelect.selectedIndex];
  return option ? option.text : "";
}

async function loadVenues() {
  setStatus("正在加载体育馆列表...");
  loadBtn.disabled = true;
  exportBtn.disabled = true;
  venueSelect.innerHTML = "";

  try {
    const response = await fetch(apiUrl("/api/venues"));
    if (!response.ok) {
      throw new Error("加载体育馆失败");
    }

    const data = await response.json();
    const venues = Array.isArray(data.venues) ? data.venues : [];
    if (venues.length === 0) {
      throw new Error("未配置体育馆，请检查 config/venues.json");
    }

    venues.forEach((venue, index) => {
      const option = document.createElement("option");
      option.value = venue.id;
      option.textContent = venue.name;
      if (index === 0) {
        option.selected = true;
      }
      venueSelect.appendChild(option);
    });

    currentVenue = {
      id: venueSelect.value,
      name: selectedVenueName(),
    };
    loadBtn.disabled = false;
    setStatus("请选择日期并查询空余时间");
  } catch (error) {
    setStatus(`体育馆加载失败：${error.message}`);
  }
}

function renderSlots(days) {
  slotsEl.innerHTML = "";
  let selectableCount = 0;

  days.forEach((day, dayIndex) => {
    const dayCard = document.createElement("article");
    dayCard.className = "day-card";

    const title = document.createElement("div");
    title.className = "day-title";
    title.textContent = `${day.date} (${day.window.start}-${day.window.end})`;
    dayCard.appendChild(title);

    if (!day.commonFree || day.commonFree.length === 0) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.textContent = "无共同空余时间";
      dayCard.appendChild(empty);
    } else {
      const list = document.createElement("div");
      list.className = "slot-list";

      day.commonFree.forEach((slot, slotIndex) => {
        selectableCount += 1;
        const item = document.createElement("label");
        item.className = "slot-item";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.className = "slot-checkbox";
        checkbox.value = JSON.stringify({
          date: day.date,
          start: slot.start,
          end: slot.end,
        });
        checkbox.id = `slot-${dayIndex}-${slotIndex}`;

        const text = document.createElement("span");
        text.textContent = `${slot.start}-${slot.end}`;

        item.appendChild(checkbox);
        item.appendChild(text);
        list.appendChild(item);
      });

      dayCard.appendChild(list);
    }

    slotsEl.appendChild(dayCard);
  });

  exportBtn.disabled = selectableCount === 0;
}

async function loadAvailability() {
  const venueId = venueSelect.value;
  const venueName = selectedVenueName();
  const date = dateInput.value;
  const days = Number(daysInput.value || 7);

  if (!venueId) {
    setStatus("请先选择体育馆");
    return;
  }

  if (!date) {
    setStatus("请先选择日期");
    return;
  }

  setStatus(`正在查询 ${venueName} 空余时间...`);
  loadBtn.disabled = true;
  exportBtn.disabled = true;
  currentVenue = {
    id: venueId,
    name: venueName,
  };

  try {
    const response = await fetch(
      apiUrl(
        `/api/availability?venueId=${encodeURIComponent(venueId)}&date=${encodeURIComponent(date)}&days=${days}`
      )
    );
    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || "查询失败");
    }

    const data = await response.json();
    renderSlots(data.data);
    setStatus(`查询完成：${data.venue.name}，共 ${data.data.length} 天，勾选后可导出`);
  } catch (error) {
    slotsEl.innerHTML = "";
    exportBtn.disabled = true;
    setStatus(`查询失败：${error.message}`);
  } finally {
    loadBtn.disabled = false;
  }
}

function selectedSlots() {
  return Array.from(document.querySelectorAll(".slot-checkbox:checked")).map((node) =>
    JSON.parse(node.value)
  );
}

async function exportXlsx() {
  const selected = selectedSlots();
  const organizationName = (organizationInput.value || "").trim();
  const applicantName = (applicantInput.value || "").trim();

  if (!organizationName) {
    setStatus("请先填写团体名（或个人名）");
    return;
  }

  if (!applicantName) {
    setStatus("请先填写申请者姓名");
    return;
  }

  if (selected.length === 0) {
    setStatus("请至少勾选一个时段再导出");
    return;
  }

  exportBtn.disabled = true;
  setStatus("正在导出 xlsx...");

  try {
    const response = await fetch(apiUrl("/api/export-xlsx"), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        venueId: currentVenue?.id,
        venueName: currentVenue?.name || selectedVenueName(),
        organizationName,
        applicantName,
        selected,
      }),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || "导出失败");
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;

    const disposition = response.headers.get("Content-Disposition") || "";
    const utf8NameMatch = disposition.match(/filename\*=UTF-8''([^;]+)/i);
    const plainNameMatch = disposition.match(/filename="?([^";]+)"?/i);
    const fileName = utf8NameMatch
      ? decodeURIComponent(utf8NameMatch[1])
      : plainNameMatch
        ? plainNameMatch[1]
        : "団体名_施設名_予定表.xlsx";

    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.URL.revokeObjectURL(url);

    const encodedName = response.headers.get("X-Export-File-Encoded") || "";
    const savedPath = encodedName ? `exports/${decodeURIComponent(encodedName)}` : "exports";
    setStatus(`导出成功，已下载并保存：${savedPath}`);
  } catch (error) {
    setStatus(`导出失败：${error.message}`);
  } finally {
    exportBtn.disabled = false;
  }
}

loadBtn.addEventListener("click", loadAvailability);
exportBtn.addEventListener("click", exportXlsx);
venueSelect.addEventListener("change", () => {
  currentVenue = {
    id: venueSelect.value,
    name: selectedVenueName(),
  };
});

loadVenues();
