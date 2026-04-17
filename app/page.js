"use client";

import { useEffect, useMemo, useState } from "react";

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

export default function HomePage() {
  const [organizationName, setOrganizationName] = useState("");
  const [applicantName, setApplicantName] = useState("");
  const [venues, setVenues] = useState([]);
  const [venueId, setVenueId] = useState("");
  const [date, setDate] = useState(todayISO());
  const [days, setDays] = useState(7);
  const [slotsByDay, setSlotsByDay] = useState([]);
  const [checkedMap, setCheckedMap] = useState({});
  const [status, setStatus] = useState("正在加载体育馆列表...");
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);

  const selectedVenue = useMemo(() => venues.find((v) => v.id === venueId) || null, [venues, venueId]);

  useEffect(() => {
    async function loadVenues() {
      try {
        const response = await fetch("/api/venues");
        if (!response.ok) {
          throw new Error("加载体育馆失败");
        }

        const data = await response.json();
        const list = Array.isArray(data.venues) ? data.venues : [];
        setVenues(list);
        if (list.length > 0) {
          setVenueId(list[0].id);
        }
        setStatus("请选择日期并查询空余时间");
      } catch (error) {
        setStatus(`体育馆加载失败：${error.message}`);
      }
    }

    loadVenues();
  }, []);

  function toggleSlot(value) {
    setCheckedMap((prev) => ({
      ...prev,
      [value]: !prev[value],
    }));
  }

  function selectedSlots() {
    return Object.keys(checkedMap)
      .filter((key) => checkedMap[key])
      .map((key) => JSON.parse(key));
  }

  async function handleLoadAvailability() {
    if (!venueId) {
      setStatus("请先选择体育馆");
      return;
    }

    if (!date) {
      setStatus("请先选择日期");
      return;
    }

    setLoading(true);
    setStatus(`正在查询 ${selectedVenue?.name || "体育馆"} 空余时间...`);
    setCheckedMap({});

    try {
      const response = await fetch(
        `/api/availability?venueId=${encodeURIComponent(venueId)}&date=${encodeURIComponent(date)}&days=${Math.max(1, Math.min(14, Number(days || 1)))}`
      );
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "查询失败");
      }

      const data = await response.json();
      const entries = Array.isArray(data.data) ? data.data : [];
      setSlotsByDay(entries);
      setStatus(`查询完成：${data.venue?.name || "体育馆"}，共 ${entries.length} 天，勾选后可导出`);
    } catch (error) {
      setSlotsByDay([]);
      setStatus(`查询失败：${error.message}`);
    } finally {
      setLoading(false);
    }
  }

  async function handleExport() {
    if (!organizationName.trim()) {
      setStatus("请先填写团体名（或个人名）");
      return;
    }

    if (!applicantName.trim()) {
      setStatus("请先填写申请者姓名");
      return;
    }

    const selected = selectedSlots();
    if (selected.length === 0) {
      setStatus("请至少勾选一个时段再导出");
      return;
    }

    setExporting(true);
    setStatus("正在导出 xlsx...");

    try {
      const response = await fetch("/api/export-xlsx", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          selected,
          venueId,
          venueName: selectedVenue?.name || "Unknown Venue",
          organizationName: organizationName.trim(),
          applicantName: applicantName.trim(),
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

      setStatus(`导出成功：${fileName}`);
    } catch (error) {
      setStatus(`导出失败：${error.message}`);
    } finally {
      setExporting(false);
    }
  }

  const selectableCount = slotsByDay.reduce((sum, day) => sum + (day.commonFree?.length || 0), 0);

  return (
    <main className="container">
      <header className="hero">
        <h1>空余时间整理</h1>
        <p>先选体育馆，再查看可选时段，最后导出到 XLSX。</p>
        <p className="hero-slogan">大家都能自由地约球，不必囿于人际关系</p>
      </header>

      <section className="panel">
        <div className="controls">
          <label>
            团体名（或个人名）
            <input value={organizationName} onChange={(e) => setOrganizationName(e.target.value)} placeholder="例如：篮球社 / 张三" />
          </label>
          <label>
            申请者姓名
            <input value={applicantName} onChange={(e) => setApplicantName(e.target.value)} placeholder="例如：李四" />
          </label>
          <label>
            体育馆
            <select value={venueId} onChange={(e) => setVenueId(e.target.value)}>
              {venues.map((v) => (
                <option key={v.id} value={v.id}>
                  {v.name}
                </option>
              ))}
            </select>
          </label>
          <label>
            起始日期
            <input type="date" value={date} onChange={(e) => setDate(e.target.value)} />
          </label>
          <label>
            天数
            <input
              type="number"
              min="1"
              max="14"
              value={days}
              onChange={(e) => setDays(Number(e.target.value || 1))}
            />
          </label>
          <button type="button" onClick={handleLoadAvailability} disabled={loading || exporting}>
            查询空余时间
          </button>
        </div>
      </section>

      <section className="panel">
        <div className="panel-head">
          <h2>可选时段</h2>
          <button type="button" disabled={exporting || loading || selectableCount === 0} onClick={handleExport}>
            导出到 XLSX
          </button>
        </div>
        <div className="status">{status}</div>
        <div className="slots">
          {slotsByDay.map((day, dayIndex) => (
            <article key={`${day.date}-${dayIndex}`} className="day-card">
              <div className="day-title">
                {day.date} ({day.window?.start}-{day.window?.end})
              </div>
              {!day.commonFree || day.commonFree.length === 0 ? (
                <div className="empty">无共同空余时间</div>
              ) : (
                <div className="slot-list">
                  {day.commonFree.map((slot, slotIndex) => {
                    const value = JSON.stringify({ date: day.date, start: slot.start, end: slot.end });
                    return (
                      <label key={`${day.date}-${slotIndex}`} className="slot-item">
                        <input
                          type="checkbox"
                          checked={Boolean(checkedMap[value])}
                          onChange={() => toggleSlot(value)}
                        />
                        <span>
                          {slot.start}-{slot.end}
                        </span>
                      </label>
                    );
                  })}
                </div>
              )}
            </article>
          ))}
        </div>
      </section>
    </main>
  );
}
