// ВСТАВЬТЕ СЮДА ВАШУ ССЫЛКУ НА APPS SCRIPT
const GAS_URL = "https://script.google.com/macros/s/AKfycbw165Z4yzi8xk8MEF5XnqE0LB2tANhIOBUcN-5nqG7bqoBBOTShYn2fAQQxh4S-Tj7mHw/exec"; 

let map;
let layers = { depo: L.layerGroup(), polygons: L.layerGroup(), points: L.layerGroup() };
let depoData = [];
let rawPointsData = [];
let filteredPoints = [];
let workingDepo = [];

// Оновлена, більш контрастна палітра (без схожих відтінків підряд)
const distinctColors = [
    '#FF0000', // Червоний
    '#0000FF', // Синій
    '#00AA00', // Насичений зелений
    '#FF8800', // Помаранчевий
    '#8800FF', // Фіолетовий
    '#00FFFF', // Блакитний (Ціан)
    '#FF00FF', // Маджента
    '#DDDD00', // Жовтий
    '#008080', // Тіл (Морська хвиля)
    '#800000', // Бордовий
    '#000080', // Темно-синій
    '#FF8080', // Рожевий
    '#80FF80', // Світло-зелений
    '#808000'  // Оливковий
];

// Ініціалізація карти
function initMap() {
    map = L.map('map', {zoomControl: false}).setView([50.4501, 30.5234], 11);
    L.control.zoom({ position: 'topright' }).addTo(map);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { attribution: '© OS' }).addTo(map);
    
    // Порядок шарів: полігони знизу, потім депо, зверху точки
    layers.polygons.addTo(map);
    layers.depo.addTo(map);
    layers.points.addTo(map);
    
    const cached = localStorage.getItem('deposhker_depo_cache');
    if (cached) {
        depoData = JSON.parse(cached);
        document.getElementById('depoStatus').innerHTML = `Завантажено з кешу (${depoData.length} шт)`;
        document.getElementById('depoStatus').className = 'status ok';
        updateDepoSelect();
        drawMap();
    }
}
initMap();

// 1. Завантаження Депо
async function fetchDepoData() {
    const statusEl = document.getElementById('depoStatus');
    
    if (!GAS_URL || GAS_URL.includes("ВАШ_ID")) {
        statusEl.innerText = 'Вкажіть URL у файлі script.js!';
        statusEl.className = 'status error';
        return;
    }
    
    statusEl.innerText = 'Завантаження...';
    statusEl.className = 'status';
    
    try {
        const response = await fetch(GAS_URL);
        const json = await response.json();
        
        if (json.status === 'success') {
            depoData = json.data;
            localStorage.setItem('deposhker_depo_cache', JSON.stringify(depoData));
            statusEl.innerHTML = `Успішно (${depoData.length} шт)`;
            statusEl.className = 'status ok';
            updateDepoSelect();
            drawMap();
        } else {
            throw new Error(json.message);
        }
    } catch (e) {
        statusEl.innerText = 'Помилка: ' + e.message;
        statusEl.className = 'status error';
    }
}

function updateDepoSelect() {
    const sel = document.getElementById('priorityDepo');
    sel.innerHTML = '<option value="auto">Авто (Північ -> Південь)</option>';
    depoData.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d['Вузол'];
        opt.innerText = d['Вузол'];
        sel.appendChild(opt);
    });
}

// 2. Читання Excel та відображення
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById('fileName').innerText = file.name;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type: 'array'});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws);
        
        rawPointsData = json.map(r => ({
            id: r['Вузол'] || r['ID'],
            lat: parseFloat(r['Широта']),
            lng: parseFloat(r['Довгота']),
            hour: parseInt(r['Час'] || r['Година'] || 0),
            dv: parseFloat(r['ДВ']) || 0,
            p: parseFloat(r['П']) || 0,
            v: parseFloat(r['В']) || 0
        })).filter(r => !isNaN(r.lat) && !isNaN(r.lng));
        
        const uniquePoints = {};
        rawPointsData.forEach(p => {
            if (!uniquePoints[p.id]) {
                uniquePoints[p.id] = { id: p.id, lat: p.lat, lng: p.lng, dv: 0, p: 0, v: 0, assigned: null, color: '#64748b' };
            }
        });
        filteredPoints = Object.values(uniquePoints);
        
        document.getElementById('pointsStatus').innerHTML = `Завантажено рядків: ${rawPointsData.length}`;
        document.getElementById('pointsStatus').className = 'status ok';
        
        drawMap();
    };
    reader.readAsArrayBuffer(file);
});

// 3. Алгоритм розподілу
// 3. Алгоритм розподілу
function runDistribution() {
    if (depoData.length === 0 || rawPointsData.length === 0) {
        return alert("Завантажте дані депо та відділень!");
    }

    const hStart = parseInt(document.getElementById('hourStart').value);
    const hEnd = parseInt(document.getElementById('hourEnd').value);
    const useDV = document.getElementById('chkDV').checked;
    const useP = document.getElementById('chkP').checked;
    const useV = document.getElementById('chkV').checked;
    const priorityId = document.getElementById('priorityDepo').value;
    const algoType = document.getElementById('algoType').value;

    if (!useDV && !useP && !useV) return alert("Оберіть хоча б один тип вантажу!");
    if (hStart > hEnd) return alert("Невірний діапазон часу!");

    const hoursMultiplier = (hEnd - hStart) + 1;

    // Агрегація точок
    const aggregated = {};
    rawPointsData.forEach(r => {
        if (r.hour >= hStart && r.hour <= hEnd) {
            if (!aggregated[r.id]) {
                aggregated[r.id] = { id: r.id, lat: r.lat, lng: r.lng, dv: 0, p: 0, v: 0, assigned: null };
            }
            aggregated[r.id].dv += r.dv;
            aggregated[r.id].p += r.p;
            aggregated[r.id].v += r.v;
        }
    });
    filteredPoints = Object.values(aggregated);

    // Підготовка Депо
    workingDepo = depoData
        .filter(d => d.isActive !== false)
        .map((d, i) => ({
            ...d,
            color: distinctColors[i % distinctColors.length],
            capDV: (parseFloat(d['ДВ']) || 0) * hoursMultiplier,
            capP: (parseFloat(d['П']) || 0) * hoursMultiplier,
            capV: (parseFloat(d['В']) || 0) * hoursMultiplier,
            curDV: 0, curP: 0, curV: 0
        }));

    // Сортування депо (має значення в основному для послідовного алгоритму)
    workingDepo.sort((a, b) => {
        if (priorityId !== 'auto') {
            if (a['Вузол'] === priorityId) return -1;
            if (b['Вузол'] === priorityId) return 1;
        }
        return b['Широта'] - a['Широта']; 
    });

    // --- ФАЗА 1: РОЗПОДІЛ ЗА ЛІМІТАМИ ---
    
    if (algoType === 'sequential') {
        // 1. Послідовний алгоритм
        workingDepo.forEach(depo => {
            let pool = filteredPoints.filter(p => !p.assigned);
            const depoPt = turf.point([parseFloat(depo['Довгота']), parseFloat(depo['Широта'])]);
            
            pool.sort((a, b) => {
                return turf.distance(depoPt, turf.point([a.lng, a.lat])) - turf.distance(depoPt, turf.point([b.lng, b.lat]));
            });

            for (let pt of pool) {
                let hitLimit = false;
                if (useDV && (depo.curDV + pt.dv > depo.capDV)) hitLimit = true;
                if (useP && (depo.curP + pt.p > depo.capP)) hitLimit = true;
                if (useV && (depo.curV + pt.v > depo.capV)) hitLimit = true;

                if (hitLimit) break; 

                pt.assigned = depo['Вузол'];
                pt.color = depo.color;
                depo.curDV += pt.dv;
                depo.curP += pt.p;
                depo.curV += pt.v;
            }
        });
    } 
    else if (algoType === 'bubbles') {
        // 2. Зростаючі бульбашки (крок 1 км)
        const MAX_RADIUS = 200; // обмежувач, щоб не зациклитися
        for (let radius = 1; radius <= MAX_RADIUS; radius++) {
            let pool = filteredPoints.filter(p => !p.assigned);
            if (pool.length === 0) break; // всі розподілені

            let candidates = [];
            
            pool.forEach(pt => {
                const ptGeo = turf.point([pt.lng, pt.lat]);
                let bestDepo = null;
                let minDist = Infinity;

                workingDepo.forEach(depo => {
                    // Перевіряємо, чи є ще місце в депо
                    let hasSpace = true;
                    if (useDV && depo.curDV >= depo.capDV) hasSpace = false;
                    if (useP && depo.curP >= depo.capP) hasSpace = false;
                    if (useV && depo.curV >= depo.capV) hasSpace = false;
                    
                    if (!hasSpace) return;

                    const dist = turf.distance(ptGeo, turf.point([parseFloat(depo['Довгота']), parseFloat(depo['Широта'])]));
                    // Якщо точка потрапляє в поточний радіус і вона ближче, ніж інші варіанти
                    if (dist <= radius && dist < minDist) {
                        minDist = dist;
                        bestDepo = depo;
                    }
                });

                if (bestDepo) {
                    candidates.push({ pt, depo: bestDepo, dist: minDist });
                }
            });

            // Сортуємо кандидатів за реальною відстанню, щоб вирішити конфлікти всередині одного кроку
            candidates.sort((a, b) => a.dist - b.dist);

            candidates.forEach(cand => {
                let { pt, depo } = cand;
                if (pt.assigned) return; // вже забрали на цьому кроці
                
                let hitLimit = false;
                if (useDV && (depo.curDV + pt.dv > depo.capDV)) hitLimit = true;
                if (useP && (depo.curP + pt.p > depo.capP)) hitLimit = true;
                if (useV && (depo.curV + pt.v > depo.capV)) hitLimit = true;

                if (!hitLimit) {
                    pt.assigned = depo['Вузол'];
                    pt.color = depo.color;
                    depo.curDV += pt.dv;
                    depo.curP += pt.p;
                    depo.curV += pt.v;
                }
            });
        }
    }
    else if (algoType === 'global') {
        // 3. Глобальна відстань (математичний аналог бульбашок, але без кроків)
        let pairs = [];
        filteredPoints.forEach(pt => {
            const ptGeo = turf.point([pt.lng, pt.lat]);
            workingDepo.forEach(depo => {
                const dist = turf.distance(ptGeo, turf.point([parseFloat(depo['Довгота']), parseFloat(depo['Широта'])]));
                pairs.push({ pt, depo, dist });
            });
        });

        // Сортуємо всі можливі пари "Відділення-Депо" від найменшої відстані до найбільшої
        pairs.sort((a, b) => a.dist - b.dist);

        for (let pair of pairs) {
            let { pt, depo } = pair;
            if (pt.assigned) continue; // точка вже має депо

            let hitLimit = false;
            if (useDV && (depo.curDV + pt.dv > depo.capDV)) hitLimit = true;
            if (useP && (depo.curP + pt.p > depo.capP)) hitLimit = true;
            if (useV && (depo.curV + pt.v > depo.capV)) hitLimit = true;

            if (!hitLimit) {
                pt.assigned = depo['Вузол'];
                pt.color = depo.color;
                depo.curDV += pt.dv;
                depo.curP += pt.p;
                depo.curV += pt.v;
            }
        }
    }

    // --- ФАЗА 2: РОЗПОДІЛ ЗАЛИШКІВ (діє для всіх алгоритмів) ---
    let remainders = filteredPoints.filter(p => !p.assigned);
    remainders.forEach(pt => {
        const ptGeo = turf.point([pt.lng, pt.lat]);
        let bestDepo = null;
        let bestScore = Infinity;

        workingDepo.forEach(depo => {
            const dist = turf.distance(ptGeo, turf.point([parseFloat(depo['Довгота']), parseFloat(depo['Широта'])]));
            
            let maxOverload = 0;
            if (useDV && depo.capDV > 0) maxOverload = Math.max(maxOverload, depo.curDV / depo.capDV);
            if (useP && depo.capP > 0) maxOverload = Math.max(maxOverload, depo.curP / depo.capP);
            if (useV && depo.capV > 0) maxOverload = Math.max(maxOverload, depo.curV / depo.capV);

            // Чим більший перевантаж, тим жорсткіший штраф на відстань
            const penalty = Math.pow(Math.max(0, maxOverload - 1), 2) * 50; 
            const score = dist * (1 + penalty);

            if (score < bestScore) {
                bestScore = score;
                bestDepo = depo;
            }
        });

        if (bestDepo) {
            pt.assigned = bestDepo['Вузол'];
            pt.color = bestDepo.color;
            bestDepo.curDV += pt.dv;
            bestDepo.curP += pt.p;
            bestDepo.curV += pt.v;
        }
    });

    drawMap(workingDepo);
}

// --- ВИВАНТАЖЕННЯ В EXCEL ---
function exportToExcel() {
    if (!filteredPoints || filteredPoints.length === 0) {
        return alert("Немає даних для вивантаження! Спочатку завантажте файл.");
    }

    // Формуємо масив об'єктів для Excel з потрібними назвами колонок
    const exportData = filteredPoints.map(p => ({
        "Відділення (Вузол)": p.id,
        "Широта": p.lat,
        "Довгота": p.lng,
        "ДВ (сума)": p.dv,
        "П (сума)": p.p,
        "В (сума)": p.v,
        "Призначене Депо": p.assigned || "Не розподілено"
    }));

    // Створюємо книгу та аркуш
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Розподіл_Депо");

    // Зберігаємо файл
    XLSX.writeFile(wb, "Депошкер_Результат.xlsx");
}

// --- РУЧНЕ ПЕРЕМІЩЕННЯ ---
window.movePoint = function(pointId) {
    const sel = document.getElementById('move-sel-' + pointId);
    if (!sel) return;
    
    const newDepoId = sel.value;
    const pt = filteredPoints.find(p => p.id === pointId);
    const newDepo = workingDepo.find(d => d['Вузол'] === newDepoId);
    const oldDepo = workingDepo.find(d => d['Вузол'] === pt.assigned);

    if (pt && newDepo && pt.assigned !== newDepoId) {
        // Коригуємо завантаження старого депо
        if (oldDepo) {
            oldDepo.curDV -= pt.dv;
            oldDepo.curP -= pt.p;
            oldDepo.curV -= pt.v;
        }
        // Додаємо вантаж новому депо
        newDepo.curDV += pt.dv;
        newDepo.curP += pt.p;
        newDepo.curV += pt.v;

        // Змінюємо прив'язку та колір точки
        pt.assigned = newDepoId;
        pt.color = newDepo.color;
        
        map.closePopup();
        drawMap(workingDepo); // Перемальовуємо карту з новими полігонами
    }
};

// Малювання на карті (ОНОВЛЕНО З POPUP)
function drawMap(processedDepo = null) {
    layers.depo.clearLayers();
    layers.points.clearLayers();
    layers.polygons.clearLayers();

    const dList = processedDepo || depoData;

    // 1. Полігони
    // 1. Полігони
    if (processedDepo) {
        processedDepo.forEach(depo => {
            const depoPoints = filteredPoints.filter(p => p.assigned === depo['Вузол']);
            
            if (depoPoints.length >= 3) {
                const pts = depoPoints.map(p => turf.point([p.lng, p.lat]));
                pts.push(turf.point([parseFloat(depo['Довгота']), parseFloat(depo['Широта'])]));
                
                const fc = turf.featureCollection(pts);
                let hull = turf.convex(fc);
                
                if (hull) {
                    // 1. Відступаємо від крайніх точок на 1 кілометр назовні
                    hull = turf.buffer(hull, 0.3, { units: 'kilometers' });
                    
                    // 2. Згладжуємо лінії, роблячи полігон плавним (2 ітерації)
                    hull = turf.polygonSmooth(hull, { iterations: 2 });

                    L.geoJSON(hull, {
                        style: { 
                            color: depo.color, 
                            weight: 4,          
                            opacity: 1,         
                            fillOpacity: 0.3    
                        }
                    }).addTo(layers.polygons);
                }
            }
        });
    }

    // 2. Депо
    depoData.forEach(d => {
        const isOff = d.isActive === false;
        
        // Стилі залежно від статусу
        const bgColor = isOff ? '#334155' : '#0f172a';
        const borderColor = isOff ? '#64748b' : '#38bdf8';
        const opacity = isOff ? 0.6 : 1;

        // Внутрішній div розтягується на 100% батьківського контейнера, 
        // текст жорстко центрується без відступів
        const iconHtml = `<div style="
            background: ${bgColor}; 
            color: #fff; 
            border: 2px solid ${borderColor}; 
            opacity: ${opacity}; 
            border-radius: 50%; 
            display: flex; 
            align-items: center; 
            justify-content: center; 
            font-weight: bold; 
            font-size: 10px; 
            width: 100%; 
            height: 100%; 
            box-sizing: border-box; 
            margin: 0; 
            padding: 0; 
            line-height: 1;
            white-space: nowrap;
        ">${d['Вузол']}</div>`;

        // Задаємо розмір 36х36 (щоб влізли довгі назви) та чіткі координати центрів
        const icon = L.divIcon({ 
            className: '', 
            html: iconHtml, 
            iconSize: [36, 36], 
            iconAnchor: [18, 18],   // Центруємо саму іконку точно по координатам
            popupAnchor: [0, -18]   // Попап з меню буде відкриватися рівно над іконкою
        });
        
        const marker = L.marker([d['Широта'], d['Довгота']], {icon});
        
        // HTML для Popup меню
        const popupHtml = `
            <div style="text-align:center; min-width: 140px;">
                <b style="color: ${isOff ? '#94a3b8' : '#f1f5f9'}">${d['Вузол']}</b><br>
                <hr style="margin:8px 0; border-color:#334155;">
                
                <button onclick="toggleDepo('${d['Вузол']}')" style="width:100%; padding:6px; background:${isOff ? '#4ade80' : '#ef4444'}; color:${isOff ? '#000' : '#fff'}; border:none; border-radius:4px; cursor:pointer; font-weight:bold; margin-bottom:8px;">
                    ${isOff ? 'ВКЛЮЧИТИ' : 'ВИКЛЮЧИТИ'}
                </button>
                
                <button onclick="openDepotChart('${d['Вузол']}')" style="width:100%; padding:6px; background:#1e293b; color:#fff; border:1px solid #38bdf8; border-radius:4px; cursor:pointer;" ${isOff ? 'disabled' : ''}>
                    📊 Графік
                </button>
            </div>
        `;
        
        // Leaflet за замовчуванням відкриває bindPopup саме ПО КЛІКУ
        marker.bindPopup(popupHtml);
        
        marker.addTo(layers.depo);
    });

    // 3. Точки відділень (з можливістю кліку)
    filteredPoints.forEach(p => {
        const col = p.color || '#64748b';
        const icon = L.divIcon({
            className: '',
            html: `<div class="point-marker" style="
                background:${col}; 
                width:14px; 
                height:14px; 
                border: 2px solid #ffffff; 
                border-radius: 50%;
                box-shadow: 0 0 4px rgba(0,0,0,0.6);"></div>`,
            iconSize: [18, 18] 
        });
        
        // Генеруємо список опцій для випадаючого списку
        let optionsHtml = '';
        if (workingDepo && workingDepo.length > 0) {
            optionsHtml = workingDepo.map(d => {
                const selected = (d['Вузол'] === p.assigned) ? 'selected' : '';
                return `<option value="${d['Вузол']}" ${selected}>${d['Вузол']}</option>`;
            }).join('');
        }

        // Вміст Popup (з'являється по кліку)
        const popupContent = `
            <div style="font-size:0.85rem; color:#f1f5f9;">
                <b>${p.id}</b><br>
                Депо: <b>${p.assigned || 'Немає'}</b><br>
                ДВ: ${p.dv} | П: ${p.p} | В: ${p.v}
                <hr style="margin:8px 0; border-color:#334155;">
                
                <div style="font-size:0.75rem; color:#94a3b8; margin-bottom:4px;">Перемістити до іншого депо:</div>
                <select id="move-sel-${p.id}" style="width:100%; padding:6px; background:#0f172a; color:#fff; border:1px solid #334155; border-radius:4px; margin-bottom:8px;">
                    ${optionsHtml}
                </select>
                <button onclick="movePoint('${p.id}')" style="width:100%; padding:6px; background:#1e293b; color:#fff; border:1px solid #4ade80; border-radius:4px; cursor:pointer;">
                    ЗБЕРЕГТИ
                </button>
            </div>
        `;

        const m = L.marker([p.lat, p.lng], {icon});
        
        // Tooltip для швидкого перегляду при наведенні
        m.bindTooltip(`<b>${p.id}</b><br>Депо: ${p.assigned || 'Не розподілено'}`);
        
        // Popup для переміщення при кліку
        if (workingDepo && workingDepo.length > 0) {
            m.bindPopup(popupContent);
        }

        m.addTo(layers.points);
    });
}

// --- ГРАФІКИ ДЛЯ ДЕПО ---
let depotChartInstance = null;

window.openDepotChart = function(depotId) {
    if (!filteredPoints || filteredPoints.length === 0) return alert("Спочатку запустіть розподіл!");

    document.getElementById('chartDepotTitle').innerText = `Навантаження: ${depotId} (24 години)`;

    const assignedBranchIds = filteredPoints.filter(p => p.assigned === depotId).map(p => p.id);
    const depoInfo = depoData.find(d => d['Вузол'] === depotId);

    // 1. Збираємо базові ліміти для кожного типу
    const caps = {
        dv: parseFloat(depoInfo['ДВ']) || 0,
        p: parseFloat(depoInfo['П']) || 0,
        v: parseFloat(depoInfo['В']) || 0
    };

    // 2. Збираємо 24-годинну статистику для кожного типу окремо
    const loads = {
        dv: new Array(24).fill(0),
        p: new Array(24).fill(0),
        v: new Array(24).fill(0)
    };
    
    rawPointsData.forEach(r => {
        if (assignedBranchIds.includes(r.id)) {
            const h = r.hour;
            if (h >= 0 && h <= 23) {
                loads.dv[h] += r.dv || 0;
                loads.p[h] += r.p || 0;
                loads.v[h] += r.v || 0;
            }
        }
    });

    renderChart(loads, caps);
    document.getElementById('chartModal').style.display = 'flex';
};

window.closeChartModal = function() {
    document.getElementById('chartModal').style.display = 'none';
};

function renderChart(loads, caps) {
    const ctx = document.getElementById('depotChart').getContext('2d');
    if (depotChartInstance) depotChartInstance.destroy();

    const labels = Array.from({length: 24}, (_, i) => `${i}:00`);

    depotChartInstance = new Chart(ctx, {
        type: 'line', 
        data: {
            labels: labels,
            datasets: [
                // --- ДВ (Синій) ---
                {
                    label: 'Фактично ДВ', data: loads.dv,
                    borderColor: '#38bdf8', borderWidth: 3, tension: 0.3, fill: false
                },
                {
                    label: 'Ліміт ДВ', data: new Array(24).fill(caps.dv),
                    borderColor: '#38bdf8', borderWidth: 2, borderDash: [5, 5], pointRadius: 0, fill: false
                },
                // --- П (Зелений) ---
                {
                    label: 'Фактично П', data: loads.p,
                    borderColor: '#4ade80', borderWidth: 3, tension: 0.3, fill: false
                },
                {
                    label: 'Ліміт П', data: new Array(24).fill(caps.p),
                    borderColor: '#4ade80', borderWidth: 2, borderDash: [5, 5], pointRadius: 0, fill: false
                },
                // --- В (Жовтий) ---
                {
                    label: 'Фактично В', data: loads.v,
                    borderColor: '#facc15', borderWidth: 3, tension: 0.3, fill: false
                },
                {
                    label: 'Ліміт В', data: new Array(24).fill(caps.v),
                    borderColor: '#facc15', borderWidth: 2, borderDash: [5, 5], pointRadius: 0, fill: false
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { 
                mode: 'index', 
                intersect: false // Показуватиме тултип для всіх ліній одразу
            },
            plugins: {
                legend: { 
                    labels: { color: '#f1f5f9', usePointStyle: true, boxWidth: 8 } 
                }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' }, beginAtZero: true }
            }
        }
    });
}

// --- УПРАВЛІННЯ СТАТУСОМ ДЕПО ---
window.toggleDepo = function(depoId) {
    const d = depoData.find(x => x['Вузол'] === depoId);
    if (d) {
        // Змінюємо статус на протилежний (за замовчуванням true)
        d.isActive = (d.isActive === false) ? true : false;
        
        // Зберігаємо в кеш, щоб після оновлення сторінки депо залишалося вимкненим
        localStorage.setItem('deposhker_depo_cache', JSON.stringify(depoData));
        
        // Якщо відділення вже були завантажені і розподілені — автоматично перераховуємо
        if (filteredPoints && filteredPoints.length > 0) {
            // Скидаємо старі прив'язки
            filteredPoints.forEach(p => { p.assigned = null; p.color = '#64748b'; });
            runDistribution(); 
        } else {
            drawMap(); // Якщо відділень ще немає, просто перемальовуємо колір маркера депо
        }
    }
};