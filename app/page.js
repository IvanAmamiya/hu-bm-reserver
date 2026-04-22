"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const ADVANCE_BOOKING_DAYS = Math.max(0, Number(process.env.NEXT_PUBLIC_ADVANCE_BOOKING_DAYS || 7));

function toISODate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function earliestBookableISO() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  now.setDate(now.getDate() + ADVANCE_BOOKING_DAYS);
  return toISODate(now);
}

function latestBookableISO() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const nextMonthEnd = new Date(now.getFullYear(), now.getMonth() + 2, 0);
  return toISODate(nextMonthEnd);
}

function normalizeSlotKey(item) {
  return `${String(item.date || "").trim()}|${String(item.start || "").trim()}|${String(item.end || "").trim()}`;
}

function parseDateCellToISO(cell) {
  if (!cell) {
    return null;
  }

  if (typeof cell.v === "number") {
    const parsed = XLSX.SSF.parse_date_code(cell.v);
    if (!parsed || !parsed.y || !parsed.m || !parsed.d) {
      return null;
    }
    const y = String(parsed.y).padStart(4, "0");
    const m = String(parsed.m).padStart(2, "0");
    const d = String(parsed.d).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  if (cell.v instanceof Date) {
    return toISODate(cell.v);
  }

  if (typeof cell.v === "string") {
    const text = cell.v.trim();
    const m = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
    if (m) {
      return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
    }
  }

  return null;
}

function extractEntriesFromWorkbook(workbook) {
  const rangeRegex = /(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/g;
  const map = new Map();

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet || !sheet["!ref"]) {
      continue;
    }

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let r = range.s.r; r <= range.e.r; r += 1) {
      for (let c = range.s.c; c <= range.e.c; c += 1) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[addr];
        const text = String(cell?.v || "");
        if (!text.includes(":")) {
          continue;
        }

        const dateCell = sheet[XLSX.utils.encode_cell({ r: Math.max(r - 1, 0), c })];
        const date = parseDateCellToISO(dateCell);
        if (!date) {
          continue;
        }

        const matches = [...text.matchAll(rangeRegex)];
        for (const hit of matches) {
          const start = hit[1].padStart(5, "0");
          const end = hit[2].padStart(5, "0");
          const key = `${date}|${start}|${end}`;
          if (!map.has(key)) {
            map.set(key, {
              date,
              start,
              end,
              sheet: sheetName,
              cell: addr,
            });
          }
        }
      }
    }
  }

  return [...map.values()].sort((a, b) => normalizeSlotKey(a).localeCompare(normalizeSlotKey(b)));
}

export default function HomePage() {
  const [organizationName, setOrganizationName] = useState("");
  const [applicantName, setApplicantName] = useState("");
  const [venues, setVenues] = useState([]);
  const [venueId, setVenueId] = useState("");
  const [date, setDate] = useState(earliestBookableISO());
  const [days, setDays] = useState(7);
  const [slotsByDay, setSlotsByDay] = useState([]);
  const [checkedMap, setCheckedMap] = useState({});
  const [status, setStatus] = useState("正在加载体育馆列表...");
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [externalFileName, setExternalFileName] = useState("");
  const [externalEntries, setExternalEntries] = useState([]);
  const [mergeStatus, setMergeStatus] = useState("请上传外部表格（xlsx）以检测内容并对照当前勾选时段。");

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
          const firstReady = list.find((item) => item.ready);
          setVenueId((firstReady || list[0]).id);
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

  async function handleExternalTableChange(event) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setExternalFileName(file.name);
    try {
      const bytes = await file.arrayBuffer();
      const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
      const entries = extractEntriesFromWorkbook(workbook);
      setExternalEntries(entries);
      setMergeStatus(`已读取 ${file.name}：识别到 ${entries.length} 条可预约时段。`);
    } catch (error) {
      setExternalEntries([]);
      setMergeStatus(`读取失败：${error.message}`);
    }
  }

  async function handleLoadAvailability() {
    if (!venueId) {
      setStatus("请先选择体育馆");
      return;
    }

    if (selectedVenue && !selectedVenue.ready) {
      setStatus(`设施“${selectedVenue.name}”尚未配置日历源，暂时无法查询。请在 config/venues.json 中补充 embedUrls。`);
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
        `/api/availability?venueId=${encodeURIComponent(venueId)}&date=${encodeURIComponent(date)}&days=${Math.max(1, Number(days || 1))}`
      );
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "查询失败");
      }

      const data = await response.json();
      const entries = Array.isArray(data.data) ? data.data : [];
      setSlotsByDay(entries);
      const requestedDays = Number(data?.range?.requestedDays || days);
      const appliedDays = Number(data?.range?.appliedDays || entries.length);
      if (appliedDays < requestedDays) {
        setStatus(`查询完成：${data.venue?.name || "体育馆"}，申请了 ${requestedDays} 天，按预约规则可查询 ${appliedDays} 天（到 ${data?.range?.endDate || "规则上限"} 为止）`);
      } else {
        setStatus(`查询完成：${data.venue?.name || "体育馆"}，共 ${entries.length} 天，勾选后可导出`);
      }
    } catch (error) {
      setSlotsByDay([]);
      setStatus(`查询失败：${error.message}`);
    } finally {
      setLoading(false);
    }
  }

  async function handleExport() {
    if (!organizationName.trim() && !applicantName.trim()) {
      alert("请至少填写团体名或申请者名");
      return;
    }

    const selected = selectedSlots();
    console.log("[导出] 获取选中时段:", selected, "总数:", selected.length);
    if (selected.length === 0) {
      setStatus("请至少勾选一个时段再导出");
      return;
    }

    setExporting(true);
    setStatus("正在导出 xlsx...");

    try {
      console.log("[导出] 开始创建表单，场馆ID:", venueId, "名称:", selectedVenue?.name);
      const form = document.createElement("form");
      form.method = "POST";
      form.action = "/api/export-xlsx";
      form.style.display = "none";

      const fields = {
        selected: JSON.stringify(selected),
        venueId,
        venueName: selectedVenue?.name || "Unknown Venue",
        organizationName: organizationName.trim(),
        applicantName: applicantName.trim(),
      };

      console.log("[导出] 表单字段:", {
        selectedCount: selected.length,
        venueId: fields.venueId,
        venueName: fields.venueName,
        organizationName: fields.organizationName,
        applicantName: fields.applicantName,
      });

      Object.entries(fields).forEach(([key, value]) => {
        const input = document.createElement("input");
        input.type = "hidden";
        input.name = key;
        input.value = String(value);
        form.appendChild(input);
      });

      console.log("[导出] 表单已创建，现在加入DOM并提交");
      document.body.appendChild(form);
      console.log("[导出] form.submit() 即将调用");
      form.submit();
      console.log("[导出] form.submit() 已调用");
      // Delay form removal to allow submission to complete
      window.setTimeout(() => {
        try {
          console.log("[导出] 延迟500ms后移除表单DOM");
          form.remove();
        } catch (e) {
          console.log("[导出] 移除表单失败:", e.message);
        }
      }, 500);

      setStatus("已提交导出请求，请查看浏览器下载列表。若3秒内未见下载，请检查弹出窗口或下载文件夹。");
    } catch (error) {
      setStatus(`导出失败：${error.message}`);
    } finally {
      setExporting(false);
    }
  }

  // 检查时间段是否与教职员使用时间（12:00-13:00）重叠
  function hasConflictWithStaffTime(slotStart, slotEnd) {
    const timeToMinutes = (timeStr) => {
      const [h, m] = timeStr.split(':').map(Number);
      return h * 60 + m;
    };
    
    const staffStart = 12 * 60;  // 12:00
    const staffEnd = 13 * 60;    // 13:00
    const slotStartMin = timeToMinutes(slotStart);
    const slotEndMin = timeToMinutes(slotEnd);
    
    return slotStartMin < staffEnd && slotEndMin > staffStart;
  }

  const selectableCount = slotsByDay.reduce((sum, day) => sum + (day.commonFree?.filter(slot => !hasConflictWithStaffTime(slot.start, slot.end)).length || 0), 0);
  const selectedCount = Object.values(checkedMap).filter(Boolean).length;
  const selectedEntries = useMemo(() => selectedSlots().sort((a, b) => normalizeSlotKey(a).localeCompare(normalizeSlotKey(b))), [checkedMap]);
  const externalKeySet = useMemo(() => new Set(externalEntries.map((item) => normalizeSlotKey(item))), [externalEntries]);
  const mergeRows = useMemo(
    () => selectedEntries.map((item) => ({ ...item, matched: externalKeySet.has(normalizeSlotKey(item)) })),
    [selectedEntries, externalKeySet]
  );
  const matchedCount = mergeRows.filter((row) => row.matched).length;

  return (
    <main className="container">
      <header className="hero">
        <div className="hero-copy">
          <h1>空余时间整理</h1>
          <p>先选体育馆，再查看可选时段，最后导出到 XLSX。</p>
        </div>
        <p className="hero-slogan">自由约球，权力不系一人之手！</p>
      </header>

      <div className="notice">
        <strong>※预约规则：</strong>
        <div>利用の予約は申請日の1週間先から申請日の翌月末までの間の予約が可能です。</div>
        <div style={{ fontSize: '0.85rem', marginTop: '6px', color: '#7a6d66' }}>
          例）2024/4/1に申請する場合、4/8～5/31までの申込が可能。申請日から直近1週間は施設が空いていても予約できません。
        </div>
        <div style={{ fontSize: '0.85rem', marginTop: '6px', color: '#7a6d66' }}>
          体育会所属団体は優先利用。別方法で予約調整を行います。
        </div>
      </div>

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
                  {v.name}{v.ready ? "" : "（未配置）"}
                </option>
              ))}
            </select>
          </label>
          <label>
            起始日期
            <input
              type="date"
              min={earliestBookableISO()}
              max={latestBookableISO()}
              value={date}
              onChange={(e) => setDate(e.target.value)}
            />
          </label>
          <label>
            天数
            <input
              type="number"
              min="1"
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
          <button type="button" disabled={exporting || loading || selectedCount === 0} onClick={handleExport}>
            导出到 XLSX{selectedCount > 0 ? `（已选 ${selectedCount}）` : ""}
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
                  {day.commonFree
                    .filter((slot) => !hasConflictWithStaffTime(slot.start, slot.end))
                    .map((slot, slotIndex) => {
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

      <section className="panel">
        <div className="panel-head">
          <h2>表格识别与融合预览</h2>
        </div>
        <div className="merge-tools">
          <label className="file-picker">
            外部表格（SharePoint导出的 xlsx）
            <input type="file" accept=".xlsx,.xls" onChange={handleExternalTableChange} />
          </label>
        </div>
        <div className="status">{mergeStatus}</div>
        <div className="merge-summary">
          <span>外部表格：{externalFileName || "未上传"}</span>
          <span>外部识别时段：{externalEntries.length}</span>
          <span>当前已勾选：{selectedEntries.length}</span>
          <span>匹配成功：{matchedCount}</span>
        </div>
        {selectedEntries.length > 0 ? (
          <div className="merge-table-wrap">
            <table className="merge-table">
              <thead>
                <tr>
                  <th>日期</th>
                  <th>时间段</th>
                  <th>外部表格中是否存在</th>
                </tr>
              </thead>
              <tbody>
                {mergeRows.map((row, idx) => (
                  <tr key={`${normalizeSlotKey(row)}-${idx}`}>
                    <td>{row.date}</td>
                    <td>{row.start}-{row.end}</td>
                    <td>
                      <span className={row.matched ? "tag-ok" : "tag-miss"}>{row.matched ? "已匹配" : "未匹配"}</span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="empty">请先在上方勾选时段，再进行融合对照。</div>
        )}
      </section>
    </main>
  );
}
