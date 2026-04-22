"use client";

import { useEffect, useMemo, useState } from "react";

const ADVANCE_BOOKING_DAYS = Math.max(0, Number(process.env.NEXT_PUBLIC_ADVANCE_BOOKING_DAYS || 7));
const FORM_URL = "https://forms.office.com/r/0fQ6dvkcDm";

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

const HERO_SLOGANS = [
  "安安静静打球，不搞人际关系",
  "自由约球，权力不系一人之手！",
  "闪闪发亮，心潮澎湃",
  "释放多巴胺！",
];

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
  const [showPostExportGuide, setShowPostExportGuide] = useState(false);
  const [sloganIndex, setSloganIndex] = useState(0);
  const [activeMainTab, setActiveMainTab] = useState("free");
  const [bookingsViewTab, setBookingsViewTab] = useState("day");
  const [bookingsByDay, setBookingsByDay] = useState([]);
  const [bookingsLoading, setBookingsLoading] = useState(false);
  const [bookingsStatus, setBookingsStatus] = useState("请先点击“查询空余时间”加载预约看板。");

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

  useEffect(() => {
    const timer = window.setInterval(() => {
      setSloganIndex((prev) => (prev + 1) % HERO_SLOGANS.length);
    }, 3200);

    return () => window.clearInterval(timer);
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

  async function loadBookingsBoard(queryDate, queryDays) {
    setBookingsLoading(true);
    setBookingsStatus("正在加载预约看板...");

    try {
      const response = await fetch(
        `/api/bookings?date=${encodeURIComponent(queryDate)}&days=${Math.max(1, Number(queryDays || 1))}`
      );
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "加载预约看板失败");
      }

      const data = await response.json();
      const entries = Array.isArray(data.data) ? data.data : [];
      setBookingsByDay(entries);
      const appliedDays = Number(data?.range?.appliedDays || entries.length);
      setBookingsStatus(`预约看板已加载，共 ${appliedDays} 天。可切换按天/按月查看各体育馆占用情况。`);
    } catch (error) {
      setBookingsByDay([]);
      setBookingsStatus(`预约看板加载失败：${error.message}`);
    } finally {
      setBookingsLoading(false);
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
    setShowPostExportGuide(false);

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
      loadBookingsBoard(date, days);
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
    setShowPostExportGuide(false);

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

      setStatus("已提交导出请求，请查看浏览器下载列表。下载完成后请按下方“下一步”指引提交预约表单。");
      setShowPostExportGuide(true);
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
  const bookingsByMonth = useMemo(() => {
    const monthMap = new Map();
    for (const day of bookingsByDay) {
      const monthKey = String(day.date || "").slice(0, 7);
      if (!monthMap.has(monthKey)) {
        monthMap.set(monthKey, new Map());
      }
      const venueMap = monthMap.get(monthKey);
      for (const venue of day.venues || []) {
        if (!venueMap.has(venue.id)) {
          venueMap.set(venue.id, {
            id: venue.id,
            name: venue.name,
            ready: venue.ready !== false,
            entries: [],
          });
        }
        const target = venueMap.get(venue.id);
        const bookings = Array.isArray(venue.bookings) ? venue.bookings : [];
        if (bookings.length > 0) {
          target.entries.push({ date: day.date, bookings });
        }
      }
    }

    return [...monthMap.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([month, venueMap]) => ({
        month,
        venues: [...venueMap.values()].sort((a, b) => a.name.localeCompare(b.name)),
      }));
  }, [bookingsByDay]);

  return (
    <main className="container">
      <header className="hero">
        <div className="hero-copy">
          <h1>空余时间整理</h1>
          <p>先选体育馆，再查看可选时段，最后导出到 XLSX，然后填写预约表单</p>
        </div>
        <p className="hero-slogan">{HERO_SLOGANS[sloganIndex]}</p>
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
          <label className="control-venue">
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
          <h2>{activeMainTab === "free" ? "可选时段" : "预约看板"}</h2>
          {activeMainTab === "free" ? (
            <button type="button" disabled={exporting || loading || selectedCount === 0} onClick={handleExport}>
              导出到 XLSX{selectedCount > 0 ? `（已选 ${selectedCount}）` : ""}
            </button>
          ) : null}
        </div>
        <div className="tab-row">
          <button
            type="button"
            className={`tab-btn ${activeMainTab === "free" ? "active" : ""}`}
            onClick={() => setActiveMainTab("free")}
          >
            可选时段
          </button>
          <button
            type="button"
            className={`tab-btn ${activeMainTab === "board" ? "active" : ""}`}
            onClick={() => setActiveMainTab("board")}
          >
            预约看板
          </button>
        </div>

        {activeMainTab === "free" ? (
          <>
            <div className="status">{status}</div>
            {showPostExportGuide ? (
          <div className="next-step-guide">
            <strong>下一步：</strong>
            <div>
              请打开
              <a href={FORM_URL} target="_blank" rel="noopener noreferrer">
                https://forms.office.com/r/0fQ6dvkcDm
              </a>
              （体育场馆预约表单），并上传你刚下载的 Excel 文件后再提交。
            </div>
          </div>
            ) : null}
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
          </>
        ) : (
          <>
            <div className="status">{bookingsStatus}</div>
            <div className="tab-row sub">
              <button
                type="button"
                className={`tab-btn ${bookingsViewTab === "day" ? "active" : ""}`}
                onClick={() => setBookingsViewTab("day")}
              >
                按天
              </button>
              <button
                type="button"
                className={`tab-btn ${bookingsViewTab === "month" ? "active" : ""}`}
                onClick={() => setBookingsViewTab("month")}
              >
                按月
              </button>
            </div>

            {bookingsLoading ? <div className="empty">预约看板加载中...</div> : null}

            {!bookingsLoading && bookingsViewTab === "day" ? (
              <div className="slots">
                {bookingsByDay.map((day) => (
                  <article key={day.date} className="day-card">
                    <div className="day-title">{day.date}</div>
                    <div className="board-venue-list">
                      {(day.venues || []).map((venue) => (
                        <div key={`${day.date}-${venue.id}`} className="board-venue-item">
                          <div className="board-venue-name">
                            {venue.name}
                            {venue.ready === false ? "（未配置）" : ""}
                          </div>
                          {venue.ready === false ? (
                            <div className="empty">未配置日历</div>
                          ) : !venue.bookings || venue.bookings.length === 0 ? (
                            <div className="empty">当日无预约</div>
                          ) : (
                            <div className="board-bookings">
                              {venue.bookings.map((booking, idx) => (
                                <div key={`${day.date}-${venue.id}-${idx}`} className="board-booking-item">
                                  <strong>{booking.start}-{booking.end}</strong>
                                  <span>{booking.summary}</span>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  </article>
                ))}
              </div>
            ) : null}

            {!bookingsLoading && bookingsViewTab === "month" ? (
              <div className="slots">
                {bookingsByMonth.map((monthItem) => (
                  <article key={monthItem.month} className="day-card">
                    <div className="day-title">{monthItem.month}</div>
                    <div className="board-venue-list">
                      {monthItem.venues.map((venue) => (
                        <div key={`${monthItem.month}-${venue.id}`} className="board-venue-item">
                          <div className="board-venue-name">
                            {venue.name}
                            {venue.ready === false ? "（未配置）" : ""}
                          </div>
                          {venue.ready === false ? (
                            <div className="empty">未配置日历</div>
                          ) : venue.entries.length === 0 ? (
                            <div className="empty">本月暂无预约</div>
                          ) : (
                            <div className="board-bookings">
                              {venue.entries.map((entry) => (
                                <div key={`${monthItem.month}-${venue.id}-${entry.date}`} className="board-booking-item">
                                  <strong>{entry.date}</strong>
                                  <span>{entry.bookings.map((booking) => `${booking.start}-${booking.end} ${booking.summary}`).join(" / ")}</span>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  </article>
                ))}
              </div>
            ) : null}
          </>
        )}
      </section>
    </main>
  );
}
