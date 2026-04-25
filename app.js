// 節慶與日程設定
const SOLAR_HOLIDAYS = [
    { name: "元旦", month: 1, day: 1 }, { name: "和平紀念日", month: 2, day: 28 },
    { name: "兒童節/清明節", month: 4, day: 4 }, { name: "勞動節", month: 5, day: 1 },
    { name: "父親節", month: 8, day: 8 }, { name: "教師節", month: 9, day: 28 },
    { name: "國慶日", month: 10, day: 10 }, { name: "台灣光復節", month: 10, day: 25 },
    { name: "跨年夜", month: 12, day: 31 }
];

const LUNAR_HOLIDAYS = [
    { name: "春節", lMonth: 1, lDay: 1 }, { name: "初二", lMonth: 1, lDay: 2 },
    { name: "初三", lMonth: 1, lDay: 3 }, { name: "初四", lMonth: 1, lDay: 4 },
    { name: "元宵節", lMonth: 1, lDay: 15 }, { name: "端午節", lMonth: 5, lDay: 5 },
    { name: "中秋節", lMonth: 8, lDay: 15 }
];

class DailyDashboard {
    constructor() {
        this.initConfiguration();
        this.birthdays = [];
        this.todos = [];
        this.pickup = [];
        this.calYear = new Date().getFullYear();
        this.calMonth = new Date().getMonth();
        
        if (document.readyState === 'complete' || document.readyState === 'interactive') { this.init(); } 
        else { document.addEventListener('DOMContentLoaded', () => this.init()); }
    }

    initConfiguration() {
        // 先嘗試從本地儲存 (localStorage) 讀取，方便在 GitHub Pages 上使用
        const savedApi = localStorage.getItem('DASHBOARD_API_URL');
        const savedToken = localStorage.getItem('DASHBOARD_SECRET_TOKEN');
        const savedPwd = localStorage.getItem('DASHBOARD_EDIT_PASSWORD');

        if (savedApi) CONFIG.API_URL = savedApi;
        if (savedToken) CONFIG.SECRET_TOKEN = savedToken;
        if (savedPwd) CONFIG.EDIT_PASSWORD = savedPwd;

        // 如果連基本的 API 網址都沒有 (通常是 GitHub Pages 環境)
        if (!CONFIG.API_URL || CONFIG.API_URL.includes("你的網址")) {
            const url = prompt("檢測到尚未配置連接資訊。\n請輸入您的 Google Apps Script 部署網址 (API_URL):", "");
            if (url) {
                localStorage.setItem('DASHBOARD_API_URL', url);
                CONFIG.API_URL = url;
            }
            const token = prompt("請輸入您的 SECRET_TOKEN (GAS 安全金鑰):", "");
            if (token) {
                localStorage.setItem('DASHBOARD_SECRET_TOKEN', token);
                CONFIG.SECRET_TOKEN = token;
            }
            const pwd = prompt("請輸入您的 EDIT_PASSWORD (編輯密碼):", "");
            if (pwd) {
                localStorage.setItem('DASHBOARD_EDIT_PASSWORD', pwd);
                CONFIG.EDIT_PASSWORD = pwd;
            }
        }
    }

    async init() {
        this.setupEventListeners();
        this.safeUpdateTime();
        setInterval(() => this.safeUpdateTime(), 1000);
        
        // Load initial data
        await this.fetchCloudData();
        this.refreshAll();
        
        // Periodic refresh (every 5 mins)
        setInterval(() => this.fetchCloudData().then(() => this.refreshAll()), 300000);
    }

    setupEventListeners() {
        document.getElementById('prev-month')?.addEventListener('click', () => this.changeMonth(-1));
        document.getElementById('next-month')?.addEventListener('click', () => this.changeMonth(1));
        document.getElementById('add-birthday-btn')?.addEventListener('click', () => this.openBirthdayModal());
        document.getElementById('close-modal')?.addEventListener('click', () => this.closeModal());
        document.getElementById('save-btn')?.addEventListener('click', () => this.handleSave());
        document.getElementById('delete-btn')?.addEventListener('click', () => this.handleDelete());
        
        // Click outside to close modal
        window.onclick = (event) => {
            if (event.target == document.getElementById('edit-modal')) this.closeModal();
        }
    }

    async fetchCloudData() {
        if (typeof CONFIG === 'undefined' || !CONFIG.API_URL.includes("https")) {
            console.warn("Config not set or invalid. Using local data.");
            this.birthdays = (typeof USER_BIRTHDAYS !== 'undefined') ? USER_BIRTHDAYS : [];
            this.todos = (typeof USER_TODOS !== 'undefined') ? USER_TODOS : [];
            this.pickup = []; // Default pickup from data.js is by week, we will prioritize cloud dates
            return;
        }
        
        try {
            const response = await fetch(`${CONFIG.API_URL}?token=${CONFIG.SECRET_TOKEN}`);
            const data = await response.json();
            if (data.error) throw new Error(data.error);
            
            this.birthdays = data.birthdays || [];
            this.todos = data.todos || [];
            this.pickup = data.pickup || [];
            console.log("Cloud data synced.");
        } catch (e) {
            console.error("Sync failed:", e);
        }
    }

    async saveToCloud(sheet, data, password, action = "update", shouldClose = true) {
        const btn = document.getElementById('save-btn');
        const originalText = btn.textContent;
        if (shouldClose) { // 只有在大儲存時才顯示按鈕讀取狀態
            btn.textContent = "傳送中...";
            btn.disabled = true;
        }

        try {
            const body = {
                token: CONFIG.SECRET_TOKEN,
                password: password,
                action: action,
                sheet: sheet,
                data: data
            };

            await fetch(CONFIG.API_URL, {
                method: 'POST',
                mode: 'no-cors',
                body: JSON.stringify(body)
            });
            
            console.log(`Cloud ${action} requested for ${sheet}`);
            
            setTimeout(async () => {
                await this.fetchCloudData();
                this.refreshAll();
                if (shouldClose) {
                    this.closeModal();
                    btn.textContent = originalText;
                    btn.disabled = false;
                }
            }, 1000);
        } catch (e) {
            alert("操作失敗: " + e.message);
            if (shouldClose) {
                btn.textContent = originalText;
                btn.disabled = false;
            }
        }
    }

    refreshAll() {
        this.renderCalendar();
        this.renderBirthdays();
        this.renderHolidays();
        this.renderPickup();
        this.renderTodos();
    }

    switchModalTab(tab) {
        // Update Buttons
        const todoTab = document.getElementById('modal-tab-todo');
        const pickupTab = document.getElementById('modal-tab-pickup');
        if (todoTab) todoTab.classList.remove('active');
        if (pickupTab) pickupTab.classList.remove('active');
        
        const activeTab = document.getElementById(`modal-tab-${tab}`);
        if (activeTab) activeTab.classList.add('active');

        // Update Views
        const todoView = document.getElementById('modal-todo-view');
        const pickupView = document.getElementById('modal-pickup-view');
        if (todoView) todoView.style.display = 'none';
        if (pickupView) pickupView.style.display = 'none';
        
        const activeView = document.getElementById(`modal-${tab}-view`);
        if (activeView) activeView.style.display = 'block';
    }

    // --- RENDERERS ---
    renderCalendar() {
        const grid = document.getElementById('calendar-grid');
        const title = document.getElementById('calendar-month-name');
        if (!grid || !title) return;

        title.textContent = `${this.calYear}年 ${this.calMonth + 1}月`;
        grid.innerHTML = '';
        ['日','一','二','三','四','五','六'].forEach(d => {
            const el = document.createElement('div'); el.className = 'day-name'; el.textContent = d; grid.appendChild(el);
        });

        const first = new Date(this.calYear, this.calMonth, 1).getDay();
        const total = new Date(this.calYear, this.calMonth + 1, 0).getDate();
        for (let i = 0; i < first; i++) grid.appendChild(document.createElement('div'));

        const now = new Date();
        const todayStr = this.formatDate(now);

        for (let i = 1; i <= total; i++) {
            const dateStr = `${this.calYear}-${String(this.calMonth + 1).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
            const el = document.createElement('div');
            el.className = 'day clickable';
            if (dateStr === todayStr) el.classList.add('today');
            
            el.innerHTML = `<span>${i}</span>`;
            
            // Add Lunar and Labels
            if (typeof Solar !== 'undefined') {
                const s = Solar.fromYmd(this.calYear, this.calMonth + 1, i);
                const l = s.getLunar();
                const lunarDayStr = l.getDay() === 1 ? l.getMonthInChinese() + '月' : l.getDayInChinese();
                const lSpan = document.createElement('span'); lSpan.className = 'lunar-day'; lSpan.textContent = lunarDayStr;
                el.appendChild(lSpan);

                // Holidays
                let h = this.getHolidayAt(this.calYear, this.calMonth + 1, i);
                
                // Buddhist Fasting Days (六齋日: Lunar 8, 14, 15, 23, 29, 30)
                // If small month (29 days), use 28, 29
                const lunarDay = l.getDay();
                const isLastDay = l.next(1).getDay() === 1;
                const isSecondLastDay = l.next(2).getDay() === 1;
                
                let isFastingDay = [8, 14, 15, 23].includes(lunarDay);
                if (isLastDay || isSecondLastDay) isFastingDay = true;

                if (isFastingDay) {
                    const fl = document.createElement('span');
                    fl.className = 'fasting-label';
                    fl.textContent = "齋";
                    el.appendChild(fl);
                }
                
                // Make-up holiday logic
                if (!h) {
                    const dayOfWeek = s.getWeek(); // 0 is Sunday, 6 is Saturday
                    if (dayOfWeek === 1) { // Monday
                        const yesterdaySolar = s.next(-1);
                        const yesterdayHoliday = this.getHolidayAt(yesterdaySolar.getYear(), yesterdaySolar.getMonth(), yesterdaySolar.getDay());
                        if (yesterdayHoliday && yesterdaySolar.getWeek() === 0) h = "補假";
                    } else if (dayOfWeek === 5) { // Friday
                        const tomorrowSolar = s.next(1);
                        const tomorrowHoliday = this.getHolidayAt(tomorrowSolar.getYear(), tomorrowSolar.getMonth(), tomorrowSolar.getDay());
                        if (tomorrowHoliday && tomorrowSolar.getWeek() === 6) h = "補假";
                    }
                }

                if (h) { const hl = document.createElement('span'); hl.className = 'holiday-label'; hl.textContent = h; el.appendChild(hl); }
            }

            // Create a container for horizontal labels (Task/Pickup)
            const labelContainer = document.createElement('div');
            labelContainer.className = 'label-container';

            // Birthday indicator (showing names)
            const bdaysToday = this.birthdays.filter(b => {
                if (b.type === 'solar' || !b.type) return (parseInt(b.month) === this.calMonth + 1 && parseInt(b.day) === i);
                if (typeof Solar !== 'undefined') {
                    const l = Solar.fromYmd(this.calYear, this.calMonth + 1, i).getLunar();
                    return (parseInt(b.month) === l.getMonth() && parseInt(b.day) === l.getDay());
                }
                return false;
            });
            
            if (bdaysToday.length > 0) { 
                const names = bdaysToday.map(b => b.name).join(' ');
                const bl = document.createElement('span'); 
                bl.className = 'birthday-label'; 
                bl.innerHTML = `🎂 ${names}`; 
                el.appendChild(bl); 
            }
            
            // Task indicator
            const hasTask = this.todos.some(t => {
                if (!t.date || !t.task) return false;
                return this.isDateMatch(t.date, dateStr, t.repeat, t.endType, t.endValue, t.rangeUntil);
            });
            if (hasTask) {
                const tl = document.createElement('span');
                tl.className = 'task-label';
                tl.innerHTML = '<span>📝</span>';
                labelContainer.appendChild(tl);
            }

            // Pickup indicator
            const hasPickup = this.pickup.some(p => {
                if (!p.date || !p.time) return false;
                return this.isDateMatch(p.date, dateStr, p.repeat, p.endType, p.endValue, p.rangeUntil);
            });
            if (hasPickup) {
                const pl = document.createElement('span');
                pl.className = 'pickup-label';
                pl.innerHTML = '<span>🚗</span>';
                labelContainer.appendChild(pl);
            }

            if (labelContainer.hasChildNodes()) {
                el.appendChild(labelContainer);
            }

            el.onclick = () => this.openDateModal(dateStr);
            grid.appendChild(el);
        }
    }

    renderBirthdays() {
        const list = document.getElementById('birthday-list');
        if (!list) return;
        list.innerHTML = '';
        const now = new Date();
        
        const sorted = this.birthdays.map(b => {
            let target;
            if (b.type === 'solar' || !b.type) {
                target = new Date(now.getFullYear(), b.month - 1, b.day);
                if (target < now && !this.isSameDay(target, now)) target.setFullYear(now.getFullYear() + 1);
            } else if (typeof Lunar !== 'undefined') {
                const s = Lunar.fromYmd(now.getFullYear(), b.month, b.day).getSolar();
                target = new Date(s.getYear(), s.getMonth() - 1, s.getDay());
                if (target < now && !this.isSameDay(target, now)) {
                    const sNext = Lunar.fromYmd(now.getFullYear() + 1, b.month, b.day).getSolar();
                    target = new Date(sNext.getYear(), sNext.getMonth() - 1, sNext.getDay());
                }
            }
            const diff = Math.ceil((target - now) / (86400000));
            return { ...b, target, diff };
        }).sort((a,b) => a.diff - b.diff).slice(0, 3);

        sorted.forEach(b => {
            const li = document.createElement('li');
            li.className = 'info-item clickable';
            li.innerHTML = `
                <div class="item-date">${b.month}/${b.day}${b.type === 'lunar' ? '(農)' : ''}</div>
                <div class="item-content">
                    <h4>${b.name} (${b.target.getFullYear() - b.birthYear}歲)</h4>
                    <p>${b.relation} • 倒數 ${b.diff} 天</p>
                </div>
            `;
            li.onclick = (e) => { e.stopPropagation(); this.openBirthdayModal(b); };
            list.appendChild(li);
        });
    }

    renderPickup() {
        const container = document.getElementById('pickup-info');
        if (!container) return;
        const todayStr = this.formatDate(new Date());
        
        // Find specific pickup entries for today
        const items = this.pickup.filter(p => this.isDateMatch(p.date, todayStr, p.repeat, p.endType, p.endValue, p.rangeUntil));
        container.innerHTML = '';
        
        if (items.length === 0) {
            container.innerHTML = '<div class="pickup-note" style="text-align:center; padding: 10px; color: grey; font-size: 0.9rem;">今日無特定接送安排</div>';
            return;
        }

        const list = document.createElement('ul');
        list.className = 'info-list'; // 使用與壽星/假日一致的列表樣式
        list.style.padding = '0';
        list.style.listStyle = 'none';

        items.forEach(p => {
            const li = document.createElement('li');
            li.style = "display: flex; align-items: center; justify-content: space-between; padding: 12px 5px; border-bottom: 1px solid rgba(0,0,0,0.05); font-size: 0.95rem;";
            
            li.innerHTML = `
                <div style="display: flex; gap: 15px; align-items: center;">
                    <span style="font-weight: 700; color: var(--text-primary); min-width: 60px;">${p.name || ""}</span>
                    <span style="color: var(--text-secondary);">${p.purpose || ""}</span>
                </div>
                <div style="font-weight: 800; color: var(--accent-color);">${this.formatDisplayTime(p.time)}</div>
            `;
            list.appendChild(li);
        });
        container.appendChild(list);
    }

    renderTodos() {
        const list = document.getElementById('todo-list');
        if (!list) return;
        const todayStr = this.formatDate(new Date());
        const items = this.todos.filter(t => {
            if (!t.date) return false;
            return this.isDateMatch(t.date, todayStr, t.repeat, t.endType, t.endValue, t.rangeUntil);
        });
        list.innerHTML = '';
        
        if (items.length === 0) {
            list.innerHTML = '<li class="todo-note" style="text-align:center; padding: 10px; color: grey;">今日尚無待辦事項</li>';
            return;
        }

        items.forEach(t => {
            const li = document.createElement('li');
            li.className = 'todo-item';
            li.style.padding = '10px 15px';
            li.innerHTML = `
                <div class="todo-text">
                    ${(t.person || t.urgent) ? `<span class="person-tag">${t.person || t.urgent}</span>` : ''}
                    ${t.task || ''}
                </div>
            `;
            list.appendChild(li);
        });
    }

    renderHolidays() {
        const list = document.getElementById('holiday-list');
        if (!list || typeof Solar === 'undefined') return;
        list.innerHTML = '';
        const now = new Date();
        const futureHolidays = [];

        // Check next 60 days
        for (let i = 0; i < 60; i++) {
            const d = new Date(now.getTime() + i * 86400000);
            const s = Solar.fromYmd(d.getFullYear(), d.getMonth() + 1, d.getDate());
            let h = this.getHolidayAt(s.getYear(), s.getMonth(), s.getDay());
            
            // Make-up check
            if (!h) {
                if (s.getWeek() === 1) {
                    const y = s.next(-1);
                    const yh = this.getHolidayAt(y.getYear(), y.getMonth(), y.getDay());
                    if (yh && y.getWeek() === 0) h = yh + " 補假";
                } else if (s.getWeek() === 5) {
                    const t = s.next(1);
                    const th = this.getHolidayAt(t.getYear(), t.getMonth(), t.getDay());
                    if (th && t.getWeek() === 6) h = th + " 補假";
                }
            }

            if (h) {
                const diff = Math.ceil((d.getTime() - now.getTime()) / 86400000);
                futureHolidays.push({ name: h, m: s.getMonth(), d: s.getDay(), diff });
            }
        }

        // De-duplicate same holiday (e.g. holiday and its makeup showing up)
        // Actually showing both is fine if they are on different days.
        futureHolidays.slice(0, 3).forEach(x => {
            const li = document.createElement('li');
            li.className = 'info-item';
            li.innerHTML = `
                <div class="item-date">${x.m}/${x.d}</div>
                <div class="item-content">
                    <h4>${x.name}</h4>
                    <p>還有 ${x.diff === 0 ? '今' : x.diff} 天</p>
                </div>
            `;
            list.appendChild(li);
        });
    }

    getHolidayAt(year, month, day) {
        const s = Solar.fromYmd(year, month, day);
        const l = s.getLunar();
        let h = "";
        SOLAR_HOLIDAYS.forEach(sh => { if (month === sh.month && day === sh.day) h = sh.name; });
        LUNAR_HOLIDAYS.forEach(lh => { if (l.getMonth() === lh.lMonth && l.getDay() === lh.lDay) h = lh.name; });
        return h;
    }

    // --- MODAL LOGIC ---
    openBirthdayModal(data = null) {
        const modal = document.getElementById('edit-modal');
        const body = document.getElementById('modal-body');
        const title = document.getElementById('modal-title');
        title.textContent = data ? "編輯壽星" : "新增壽星";
        
        body.innerHTML = `
            <div class="form-group"><label>姓名</label><input id="b-name" value="${data?.name || ''}"></div>
            <div class="form-group"><label>關係</label><input id="b-relation" value="${data?.relation || ''}"></div>
            <div class="form-group"><label>類型</label>
                <select id="b-type"><option value="solar" ${data?.type === 'solar' ? 'selected' : ''}>國曆</option>
                <option value="lunar" ${data?.type === 'lunar' ? 'selected' : ''}>農曆</option></select>
            </div>
            <div style="display:flex; gap:10px;">
                <div class="form-group"><label>西元出生年</label><input type="number" id="b-year" value="${data?.birthYear || 1990}"></div>
                <div class="form-group"><label>月</label><input type="number" id="b-month" value="${data?.month || 1}"></div>
                <div class="form-group"><label>日</label><input type="number" id="b-day" value="${data?.day || 1}"></div>
            </div>
            <input type="hidden" id="b-id" value="${data?.id || ''}">
            <input type="hidden" id="entry-type" value="birthday">
        `;
        
        document.getElementById('delete-btn').style.display = data ? 'block' : 'none';
        modal.style.display = 'flex';
    }

    openDateModal(dateStr) {
        const modal = document.getElementById('edit-modal');
        const body = document.getElementById('modal-body');
        const title = document.getElementById('modal-title');
        title.textContent = `${dateStr} 編輯`;
        
        const sBtn = document.getElementById('save-btn');
        if (sBtn) { sBtn.textContent = "儲存"; sBtn.disabled = false; }

        const dayTodos = this.todos.filter(t => this.isDateMatch(t.date, dateStr, t.repeat, t.endType, t.endValue, t.rangeUntil));
        const dayPickups = this.pickup.filter(p => this.isDateMatch(p.date, dateStr, p.repeat, p.endType, p.endValue, p.rangeUntil));

        let tasksHtml = dayTodos.map(t => this.generateTaskRowHtml(t)).join('');
        if (dayTodos.length === 0) tasksHtml = this.generateTaskRowHtml();

        let pickupsHtml = dayPickups.map(p => this.generatePickupRowHtml(p)).join('');
        if (dayPickups.length === 0) pickupsHtml = this.generatePickupRowHtml();

        body.innerHTML = `
            <div class="tabs-header">
                <button class="tab-btn active" id="modal-tab-todo" onclick="app.switchModalTab('todo')">
                    📝 待辦事項
                </button>
                <button class="tab-btn" id="modal-tab-pickup" onclick="app.switchModalTab('pickup')">
                    🚗 接送行程
                </button>
            </div>

            <div id="modal-todo-view">
                <h4 style="margin: 20px 0 15px 0; color:var(--accent-color); display: flex; align-items: center; gap: 8px;">
                    <span>📝</span> 今日待辦事項
                </h4>
                <div id="modal-task-list" class="task-list-container">
                    ${tasksHtml}
                </div>
                <button type="button" class="add-task-line" id="add-task-row-btn">+ 新增待辦事項</button>
            </div>

            <div id="modal-pickup-view" style="display:none;">
                <h4 style="margin: 20px 0 15px 0; color:var(--accent-color); display: flex; align-items: center; gap: 8px;">
                    <span>🚗</span> 接送行程
                </h4>
                <div id="modal-pickup-list" class="task-list-container">
                    ${pickupsHtml}
                </div>
                <button type="button" class="add-task-line" id="add-pickup-row-btn">+ 新增接送行程</button>
            </div>
            
            <input type="hidden" id="entry-date" value="${dateStr}">
            <input type="hidden" id="entry-type" value="date-override">
        `;
        
        // 綁定待辦事件
        document.getElementById('add-task-row-btn').onclick = () => {
            const container = document.getElementById('modal-task-list');
            const newRow = document.createElement('div');
            newRow.innerHTML = this.generateTaskRowHtml();
            container.appendChild(newRow.firstElementChild);
        };

        // 綁定接送事件
        document.getElementById('add-pickup-row-btn').onclick = () => {
            const container = document.getElementById('modal-pickup-list');
            const newRow = document.createElement('div');
            newRow.innerHTML = this.generatePickupRowHtml();
            container.appendChild(newRow.firstElementChild);
        };
        // 隱藏原本的大刪除按鈕，因為現在待辦有自己獨立的刪除鍵
        document.getElementById('delete-btn').style.display = 'none';
        modal.style.display = 'flex';
    }

    generateTaskRowHtml(t = {id:'', person:'', task:'', repeat: 'none', endType: 'never', endValue: '', rangeUntil: ''}) {
        const rowId = t.id || 'new-' + Math.random().toString(36).substr(2, 9);
        return `
            <div class="modal-task-row task-input-row" id="row-${rowId}">
                <div class="form-group field-1"><label>人員</label><input class="row-person" value="${t.person || ''}" placeholder="誰?"></div>
                <div class="form-group field-2"><label>任務</label><input class="row-task" value="${t.task || ''}" placeholder="要做什麼?"></div>
                <div class="form-group field-3"><label>重複</label>
                    <select class="row-repeat" onchange="app.toggleRecurrenceDetails('${rowId}')">
                        <option value="none" ${t.repeat === 'none' ? 'selected' : ''}>不重複</option>
                        <option value="daily" ${t.repeat === 'daily' || t.repeat === '每天' ? 'selected' : ''}>每天</option>
                        <option value="weekly" ${t.repeat === 'weekly' || t.repeat === '每週' ? 'selected' : ''}>每週</option>
                        <option value="monthly" ${t.repeat === 'monthly' || t.repeat === '每月' ? 'selected' : ''}>每月</option>
                    </select>
                </div>
                
                <div class="form-group field-4" style="display:flex; gap:15px; flex-direction:column;">
                    <div><label>截止日期 (期間)</label><input type="date" class="row-range-until" value="${t.rangeUntil || ''}" style="width: 200px;"></div>
                    
                    <div class="recurrence-details" id="rec-${rowId}" style="display: ${t.repeat && t.repeat !== 'none' ? 'block' : 'none'}; border-top: 1px dashed #eee; padding-top: 10px;">
                        <div style="display: flex; gap: 15px; align-items: center; flex-wrap: wrap; font-size: 0.85rem;">
                            <label style="display: flex; align-items: center; gap: 5px; cursor: pointer; font-weight:normal;">
                                <input type="radio" name="endType-${rowId}" class="row-end-type" value="never" ${t.endType === 'never' || !t.endType ? 'checked' : ''}> 持續不停
                            </label>
                            <label style="display: flex; align-items: center; gap: 5px; cursor: pointer; font-weight:normal;">
                                <input type="radio" name="endType-${rowId}" class="row-end-type" value="date" ${t.endType === 'date' ? 'checked' : ''}> 於
                                <input type="date" class="row-end-date" value="${t.endType === 'date' ? t.endValue : ''}" style="width: 130px; padding: 2px 5px;">
                            </label>
                            <label style="display: flex; align-items: center; gap: 5px; cursor: pointer; font-weight:normal;">
                                <input type="radio" name="endType-${rowId}" class="row-end-type" value="count" ${t.endType === 'count' ? 'checked' : ''}> 重複
                                <input type="number" class="row-end-count" value="${t.endType === 'count' ? t.endValue : '10'}" style="width: 45px; padding: 2px 5px;"> 次
                            </label>
                        </div>
                    </div>
                </div>

                <input type="hidden" class="row-id" value="${t.id || ''}">
                ${t.id ? `<button type="button" class="item-delete-btn field-5" title="刪除" onclick="app.deleteSingleTask('${t.id}', this)">×</button>` : ''}
            </div>
        `;
    }

    toggleRecurrenceDetails(rowId) {
        const row = document.getElementById('row-' + rowId);
        if (!row) return;
        const select = row.querySelector('.row-repeat') || row.querySelector('.row-p-repeat');
        const details = document.getElementById('rec-' + rowId);
        if (select && details) {
            details.style.display = (select.value === 'none') ? 'none' : 'block';
        }
    }

    async deleteSingleTask(id, btnElement) {
        if (!confirm("確定要刪除這項任務嗎？")) return;
        this.optimisticRemove(btnElement);
        const pwd = CONFIG.EDIT_PASSWORD || CONFIG.SECRET_TOKEN;
        await this.saveToCloud("Todos", { id: id }, pwd, "delete", false);
    }

    generatePickupRowHtml(p = {id:'', name:'', time:'', purpose:'', repeat: 'none', endType: 'never', endValue: '', rangeUntil: ''}) {
        const rowId = p.id || 'new-' + Math.random().toString(36).substr(2, 9);
        return `
            <div class="modal-task-row pickup-input-row" id="row-${rowId}">
                <div class="form-group field-1"><label>對象</label><input class="row-p-name" value="${p.name || ''}" placeholder="姓名"></div>
                <div class="form-group field-2"><label>時間</label><input type="time" class="row-p-time" value="${this.formatDisplayTime(p.time) || '07:30'}"></div>
                <div class="form-group field-3"><label>目的/地點</label><input class="row-p-purpose" value="${p.purpose || ''}" placeholder="如：補習班"></div>
                
                <div class="form-group field-4" style="display:flex; gap:15px; flex-direction:column;">
                    <div style="display:flex; gap:20px; align-items: flex-start;">
                        <div style="flex:1;"><label>重複</label>
                            <select class="row-p-repeat" onchange="app.toggleRecurrenceDetails('${rowId}')">
                                <option value="none" ${p.repeat === 'none' ? 'selected' : ''}>不重複</option>
                                <option value="daily" ${p.repeat === 'daily' || p.repeat === '每天' ? 'selected' : ''}>每天</option>
                                <option value="weekly" ${p.repeat === 'weekly' || p.repeat === '每週' ? 'selected' : ''}>每週</option>
                                <option value="monthly" ${p.repeat === 'monthly' || p.repeat === '每月' ? 'selected' : ''}>每月</option>
                            </select>
                        </div>
                        <div style="flex:1.2;"><label>截止日期 (期間)</label><input type="date" class="row-p-range-until" value="${p.rangeUntil || ''}"></div>
                    </div>
                    
                    <div class="recurrence-details" id="rec-${rowId}" style="display: ${p.repeat && p.repeat !== 'none' ? 'block' : 'none'}; border-top: 1px dashed #eee; padding-top: 10px;">
                        <div style="display: flex; gap: 15px; align-items: center; flex-wrap: wrap; font-size: 0.85rem;">
                            <label style="display: flex; align-items: center; gap: 5px; cursor: pointer; font-weight:normal;">
                                <input type="radio" name="endType-${rowId}" class="row-end-type" value="never" ${p.endType === 'never' || !p.endType ? 'checked' : ''}> 持續不停
                            </label>
                            <label style="display: flex; align-items: center; gap: 5px; cursor: pointer; font-weight:normal;">
                                <input type="radio" name="endType-${rowId}" class="row-end-type" value="date" ${p.endType === 'date' ? 'checked' : ''}> 於
                                <input type="date" class="row-end-date" value="${p.endType === 'date' ? p.endValue : ''}" style="width: 130px; padding: 2px 5px;">
                            </label>
                            <label style="display: flex; align-items: center; gap: 5px; cursor: pointer; font-weight:normal;">
                                <input type="radio" name="endType-${rowId}" class="row-end-type" value="count" ${p.endType === 'count' ? 'checked' : ''}> 重複
                                <input type="number" class="row-end-count" value="${p.endType === 'count' ? p.endValue : '10'}" style="width: 45px; padding: 2px 5px;"> 次
                            </label>
                        </div>
                    </div>
                </div>

                <input type="hidden" class="row-p-id" value="${p.id || ''}">
                ${p.id ? `<button type="button" class="item-delete-btn field-5" title="刪除" onclick="app.deleteSinglePickup('${p.id}', this)">×</button>` : ''}
            </div>
        `;
    }


    async deleteSinglePickup(id, btnElement) {
        if (!confirm("確定要刪除這條接送行程嗎？")) return;
        this.optimisticRemove(btnElement);
        const pwd = CONFIG.EDIT_PASSWORD || CONFIG.SECRET_TOKEN;
        await this.saveToCloud("Pickup", { id: id }, pwd, "delete", false);
    }

    optimisticRemove(btnElement) {
        const row = btnElement.closest('.modal-task-row');
        if (row) {
            row.style.transition = 'all 0.3s ease';
            row.style.opacity = '0';
            row.style.transform = 'scale(0.9)';
            setTimeout(() => row.remove(), 300);
        }
    }

    async handleSave() {
        const pwd = CONFIG.EDIT_PASSWORD || CONFIG.SECRET_TOKEN;
        const type = document.getElementById('entry-type').value;

        if (type === 'birthday') {
            const data = {
                id: document.getElementById('b-id').value,
                name: document.getElementById('b-name').value,
                relation: document.getElementById('b-relation').value,
                type: document.getElementById('b-type').value,
                birthYear: parseInt(document.getElementById('b-year').value),
                month: parseInt(document.getElementById('b-month').value),
                day: parseInt(document.getElementById('b-day').value)
            };
            await this.saveToCloud("Birthdays", data, pwd);
        } else if (type === 'date-override') {
            const date = document.getElementById('entry-date').value;
            const promises = [];
            
            // 處理所有待辦事項
            const taskRows = document.querySelectorAll('.task-input-row');
            for (let row of taskRows) {
                const tTask = row.querySelector('.row-task').value;
                const tPerson = row.querySelector('.row-person').value;
                const tRepeat = row.querySelector('.row-repeat').value;
                const tRangeUntil = row.querySelector('.row-range-until').value;
                const tId = row.querySelector('.row-id').value;
                
                let tEndType = 'never';
                let tEndValue = '';
                if (tRepeat !== 'none') {
                    const checkedRadio = row.querySelector('.row-end-type:checked');
                    tEndType = checkedRadio ? checkedRadio.value : 'never';
                    if (tEndType === 'date') tEndValue = row.querySelector('.row-end-date').value;
                    if (tEndType === 'count') tEndValue = row.querySelector('.row-end-count').value;
                }

                if (tTask) {
                    promises.push(this.saveToCloud("Todos", { 
                        id: tId, date: date, task: tTask, person: tPerson, 
                        repeat: tRepeat, endType: tEndType, endValue: tEndValue,
                        rangeUntil: tRangeUntil,
                        completed: 'FALSE' 
                    }, pwd, "update", false));
                    await new Promise(r => setTimeout(r, 150));
                }
            }
            
            // 處理所有接送行程
            const pickupRows = document.querySelectorAll('.pickup-input-row');
            for (let row of pickupRows) {
                const pName = row.querySelector('.row-p-name').value;
                const pTime = row.querySelector('.row-p-time').value;
                const pPurpose = row.querySelector('.row-p-purpose').value;
                const pRepeat = row.querySelector('.row-p-repeat').value;
                const pRangeUntil = row.querySelector('.row-p-range-until').value;
                const pId = row.querySelector('.row-p-id').value;

                let pEndType = 'never';
                let pEndValue = '';
                if (pRepeat !== 'none') {
                    const checkedRadio = row.querySelector('.row-end-type:checked');
                    pEndType = checkedRadio ? checkedRadio.value : 'never';
                    if (pEndType === 'date') pEndValue = row.querySelector('.row-end-date').value;
                    if (pEndType === 'count') pEndValue = row.querySelector('.row-end-count').value;
                }

                if (pName && pTime) {
                    promises.push(this.saveToCloud("Pickup", { 
                        id: pId, date: date, name: pName, time: pTime, 
                        purpose: pPurpose, repeat: pRepeat, 
                        endType: pEndType, endValue: pEndValue,
                        rangeUntil: pRangeUntil
                    }, pwd, "update", false));
                    await new Promise(r => setTimeout(r, 150));
                }
            }
            
            if (promises.length === 0) {
                alert("請填寫內容後再儲存");
                return;
            }

            // 最後一筆帶動關閉視窗
            setTimeout(() => {
                this.closeModal();
            }, 800);
        }
    }

    async handleDelete() {
        const type = document.getElementById('entry-type').value;
        const pwd = CONFIG.EDIT_PASSWORD || CONFIG.SECRET_TOKEN;

        if (type === 'birthday') {
            const id = document.getElementById('b-id').value;
            if (id && confirm("確定要刪除此生日記錄嗎？")) {
                await this.saveToCloud("Birthdays", { id: id }, pwd, "delete");
            }
        } else if (type === 'date-override') {
            const tId = document.getElementById('t-id').value;
            const pId = document.getElementById('p-id').value;
            
            if (!tId && !pId) return;

            if (confirm("確定要刪除此日期的資料嗎？")) {
                // 依序執行刪除，確保 UI 狀態同步
                if (tId) {
                    await this.saveToCloud("Todos", { id: tId }, pwd, "delete");
                }
                if (pId) {
                    if (tId) await new Promise(r => setTimeout(r, 200)); // 只有當兩項都要刪時才延遲
                    await this.saveToCloud("Pickup", { id: pId }, pwd, "delete");
                }
            }
        }
    }

    closeModal() {
        document.getElementById('edit-modal').style.display = 'none';
    }

    // --- HELPERS ---
    formatDate(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
    isSameDay(d1, d2) { return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate(); }
    changeMonth(offset) { let d = new Date(this.calYear, this.calMonth + offset, 1); this.calYear = d.getFullYear(); this.calMonth = d.getMonth(); this.renderCalendar(); }
    safeUpdateTime() {
        const now = new Date();
        const cEl = document.getElementById('digital-clock');
        const dEl = document.getElementById('current-date');
        if (!cEl || !dEl) return;
        cEl.textContent = now.toLocaleTimeString('zh-TW', { hour12: false });
        const greg = now.toLocaleDateString('zh-TW', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' });
        if (typeof Solar !== 'undefined') {
            try {
                const l = Solar.fromYmd(now.getFullYear(), now.getMonth() + 1, now.getDate()).getLunar();
                dEl.textContent = `${greg} (農曆 ${l.getMonthInChinese()}月${l.getDayInChinese()})`;
            } catch(e) { dEl.textContent = greg; }
        } else { dEl.textContent = greg; }
    }
    
    // 格式化顯示時間，處理 ISO 格式字串
    formatDisplayTime(timeStr) {
        if (!timeStr) return "--:--";
        // 如果是完整的 ISO 日期時間字串 (例如 1899-12-30T00:31:00.000Z)
        if (timeStr.includes('T')) {
            const timePart = timeStr.split('T')[1];
            return timePart.substring(0, 5); // 取得 HH:mm
        }
        return timeStr;
    }

    // 改良版比對函式：支援每天、每週、每月重複，且支援結束條件與期間(Range)
    isDateMatch(itemDate, targetDateStr, repeatType, endType, endValue, rangeUntil) {
        if (!itemDate || !targetDateStr) return false;
        
        const d1 = this.parseLocalDate(itemDate);
        const d2 = this.parseLocalDate(targetDateStr);
        if (!d1 || !d2) return false;
        
        // 如果目標日期在起始日期之前，不匹配
        if (d2 < d1) return false;

        // --- 期間檢查 (Range Check) ---
        if (rangeUntil) {
            const dRangeEnd = this.parseLocalDate(rangeUntil);
            if (d2 >= d1 && d2 <= dRangeEnd) return true;
        }

        // --- 結束條件檢查 ---
        if (repeatType && repeatType !== 'none') {
            if (endType === 'date' && endValue) {
                const dEnd = this.parseLocalDate(endValue);
                if (d2 > dEnd) return false;
            } else if (endType === 'count' && endValue) {
                const count = parseInt(endValue);
                if (repeatType === 'daily' || repeatType === '每天') {
                    const diffDays = Math.floor((d2 - d1) / (24 * 60 * 60 * 1000));
                    if (diffDays >= count) return false;
                } else if (repeatType === 'weekly' || repeatType === '每週') {
                    const diffDays = Math.floor((d2 - d1) / (24 * 60 * 60 * 1000));
                    const diffWeeks = Math.floor(diffDays / 7);
                    if (diffWeeks >= count) return false;
                } else if (repeatType === 'monthly' || repeatType === '每月') {
                    const diffMonths = (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
                    if (diffMonths >= count) return false;
                }
            }
        } else {
            // 不重複的情況
            return this.getLocalDateStr(d1) === this.getLocalDateStr(d2);
        }
        
        // --- 重複規律檢查 ---
        if (this.getLocalDateStr(d1) === this.getLocalDateStr(d2)) return true;
        
        if (repeatType === 'daily' || repeatType === '每天') {
            return true;
        }
        
        if (repeatType === 'weekly' || repeatType === '每週') {
            return d1.getDay() === d2.getDay();
        }
        
        if (repeatType === 'monthly' || repeatType === '每月') {
            return d1.getDate() === d2.getDate();
        }
        
        return false;
    }

    getLocalDateStr(input) {
        const d = this.parseLocalDate(input);
        if (!d) return null;
        return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }

    parseLocalDate(input) {
        if (input instanceof Date) return input;
        const s = String(input).trim();
        if (s.includes('T') || s.includes('Z')) return new Date(s);
        const parts = s.replace(/\//g, '-').split('-');
        if (parts.length === 3) return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        return new Date(s);
    }
}
const app = new DailyDashboard();
