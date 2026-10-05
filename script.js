// ВСТАВЬТЕ СЮДА ВАШУ ССЫЛКУ НА APPS SCRIPT
const GAS_URL = "https://script.google.com/macros/s/AKfycbw165Z4yzi8xk8MEF5XnqE0LB2tANhIOBUcN-5nqG7bqoBBOTShYn2fAQQxh4S-Tj7mHw/exec"; 

let map;
let layers = { depo: L.layerGroup(), polygons: L.layerGroup(), lines: L.layerGroup(), points: L.layerGroup() };
let depoData = [];
let rawPointsData = [];
let filteredPoints = [];
let workingDepo = [];
let areaData = [];

window.activeDepotFilters = new Set(); // Порожній = показувати всі

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
let oms; // Глобальна змінна для Spiderfier
let drawnItems; // Шар для малювання полігонів

function initMap() {
    map = L.map('map', {zoomControl: false}).setView([50.4501, 30.5234], 11);
    L.control.zoom({ position: 'topright' }).addTo(map);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { attribution: '© OS' }).addTo(map);
    
    layers.polygons.addTo(map);
    layers.lines.addTo(map);
    layers.depo.addTo(map);
    layers.points.addTo(map);
    
    // Ініціалізація Spiderfier (Веер для однакових координат)
    // Ініціалізація Spiderfier (Веер для однакових координат)
    oms = new OverlappingMarkerSpiderfier(map, {
        keepSpiderfied: true, 
        nearbyDistance: 15, // Зменшили радіус, щоб ловив тільки реальні накладання
        legWeight: 2, 
        circleSpiralSwitchover: 9
    });

    // Створюємо єдиний попап для всіх відділень
    const globalPopup = new L.Popup({ minWidth: 260 });
    
    // Перехоплюємо клік через плагін
    oms.addListener('click', function(marker) {
        // Якщо клік був імітований просто наведенням мишки - ігноруємо відкриття попапу
        if (window.isHoverClick) return; 
        
        if (marker.customPopupHtml) {
            globalPopup.setContent(marker.customPopupHtml);
            globalPopup.setLatLng(marker.getLatLng());
            map.openPopup(globalPopup);
        }
    });

    // Ініціалізація Leaflet.Draw (Масове виділення)
    drawnItems = new L.FeatureGroup();
    map.addLayer(drawnItems);
    const drawControl = new L.Control.Draw({
        draw: { polyline: false, marker: false, circlemarker: false, circle: false,
            polygon: { allowIntersection: false, showArea: true, shapeOptions: { color: '#38bdf8' } },
            rectangle: { shapeOptions: { color: '#38bdf8' } }
        },
        edit: { featureGroup: drawnItems, remove: true }
    });
    map.addControl(drawControl);

    // Обробник виділення полігоном
    map.on(L.Draw.Event.CREATED, function (e) {
        drawnItems.clearLayers(); // Залишаємо лише один полігон
        drawnItems.addLayer(e.layer);
        
        const polyGeoJson = e.layer.toGeoJSON();
        window.selectedPointsForMassMove = [];
        
        filteredPoints.forEach(p => {
            if (turf.booleanPointInPolygon(turf.point([p.lng, p.lat]), polyGeoJson)) {
                window.selectedPointsForMassMove.push(p);
            }
        });

        if(window.selectedPointsForMassMove.length > 0) {
            document.getElementById('massMoveCount').innerText = window.selectedPointsForMassMove.length;
            populateMassMoveDatalist();
            document.getElementById('massMoveModal').style.display = 'flex';
        } else {
            alert("В обраній зоні немає відділень.");
            drawnItems.clearLayers();
        }
    });
    
    const cachedDepo = localStorage.getItem('deposhker_depo_cache');
    const cachedArea = localStorage.getItem('deposhker_area_cache');
    if (cachedArea) areaData = JSON.parse(cachedArea);
    
    if (cachedDepo) {
        depoData = JSON.parse(cachedDepo);
        document.getElementById('depoStatus').innerHTML = `З кешу: ${depoData.length} депо, ${areaData.length} площ`;
        document.getElementById('depoStatus').className = 'status ok';
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
    
    statusEl.innerText = 'Завантаження Даних...';
    statusEl.className = 'status';
    
    try {
        const response = await fetch(GAS_URL);
        const json = await response.json();
        
        if (json.status === 'success') {
            // Оновлена логіка під нову структуру JSON
            depoData = json.data.depots;
            areaData = json.data.areas || [];
            
            // Зберігаємо обидва масиви в кеш
            localStorage.setItem('deposhker_depo_cache', JSON.stringify(depoData));
            localStorage.setItem('deposhker_area_cache', JSON.stringify(areaData));
            
            statusEl.innerHTML = `Успішно: ${depoData.length} депо, ${areaData.length} площ`;
            statusEl.className = 'status ok';
            drawMap();
        } else {
            throw new Error(json.message);
        }
    } catch (e) {
        statusEl.innerText = 'Помилка: ' + e.message;
        statusEl.className = 'status error';
    }
}


// 2. Читання Excel та відображення
// 2. Читання Excel та відображення
// 2. Читання Excel та відображення
let currentWorkbook = null;

document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById('fileName').innerText = file.name;
    
    if (depoData.length === 0) {
        alert("Увага: Спочатку завантажте дані Депо (Крок 1), інакше збережена сесія може завантажитись некоректно.");
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        currentWorkbook = XLSX.read(data, {type: 'array'});
        
        const select = document.getElementById('sheetSelect');
        select.innerHTML = '';
        let validSheetsCount = 0;

        // Перевіряємо всі аркуші у файлі
        currentWorkbook.SheetNames.forEach(sheetName => {
            const ws = currentWorkbook.Sheets[sheetName];
            // Читаємо тільки перший рядок (заголовки)
            const headers = XLSX.utils.sheet_to_json(ws, {header: 1})[0] || [];
            
            // Аркуш підходить, якщо в ньому є координати і хоча б один ідентифікатор
            const hasLat = headers.includes('Широта');
            const hasLng = headers.includes('Довгота');
            const hasId = headers.includes('Вузол') || headers.includes('ID') || headers.includes('Відділення (Вузол)');

            if (hasLat && hasLng && hasId) {
                const opt = document.createElement('option');
                opt.value = sheetName;
                opt.innerText = sheetName;
                select.appendChild(opt);
                validSheetsCount++;
            }
        });

        const selectionBlock = document.getElementById('sheetSelectionBlock');

        if (validSheetsCount === 0) {
            alert("У файлі не знайдено аркушів із потрібним форматом (мають бути колонки Широта, Довгота та Вузол/ID).");
            selectionBlock.style.display = 'none';
        } else if (validSheetsCount === 1) {
            // Якщо підходить тільки один аркуш — вантажимо його автоматично
            selectionBlock.style.display = 'none';
            processSelectedSheet(select.options[0].value);
        } else {
            // Якщо підходять декілька — показуємо селект
            selectionBlock.style.display = 'block';
            document.getElementById('pointsStatus').innerHTML = `Знайдено аркушів: ${validSheetsCount}. Оберіть потрібний.`;
            document.getElementById('pointsStatus').className = 'status';
        }
    };
    reader.readAsArrayBuffer(file);
});

// Функція обробки конкретного обраного аркуша
window.processSelectedSheet = function(sheetNameFromArg = null) {
    const sheetName = sheetNameFromArg || document.getElementById('sheetSelect').value;
    if (!sheetName || !currentWorkbook) return;

    // Ховаємо блок вибору після натискання ОК
    document.getElementById('sheetSelectionBlock').style.display = 'none';

    const ws = currentWorkbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws);
    
    const isSavedFile = json.length > 0 && json[0]['Призначене Депо'] !== undefined;

    if (isSavedFile) {
        // --- ВІДНОВЛЕННЯ ЗБЕРЕЖЕНОГО СТАНУ ---
        rawPointsData = json.map(r => ({
            id: String(r['Відділення (Вузол)']),
            lat: parseFloat(r['Широта']),
            lng: parseFloat(r['Довгота']),
            hour: 10, 
            dv: parseFloat(r['ДВ (сума)']) || 0,
            p: parseFloat(r['П (сума)']) || 0,
            v: parseFloat(r['В (сума)']) || 0
        }));
        
        filteredPoints = json.map(r => ({
            id: String(r['Відділення (Вузол)']),
            lat: parseFloat(r['Широта']),
            lng: parseFloat(r['Довгота']),
            dv: parseFloat(r['ДВ (сума)']) || 0,
            p: parseFloat(r['П (сума)']) || 0,
            v: parseFloat(r['В (сума)']) || 0,
            assigned: r['Призначене Депо'] === "Не розподілено" ? null : r['Призначене Депо'],
            color: r['Колір'] || '#64748b',
            isStar: (r['Є Звіздою'] === 'Так' || r['Є Зірочкою'] === 'Так'),
            assignedStar: r["Прив'язана Звіздочка"] || r["Прив'язана Зірочка"] || null,
            starChildren: 0 
        }));
        
        filteredPoints.forEach(p => {
            if (p.assignedStar) {
                const star = filteredPoints.find(s => s.id === p.assignedStar);
                if (star) star.starChildren++;
            }
        });

        const hStart = parseInt(document.getElementById('hourStart').value) || 10;
        const hEnd = parseInt(document.getElementById('hourEnd').value) || 12;
        const hoursMultiplier = Math.max(1, (hEnd - hStart) + 1);

        workingDepo = depoData.map(d => ({
            ...d,
            color: '#38bdf8', 
            capDV: (parseFloat(d['ДВ']) || 0) * hoursMultiplier,
            capP: (parseFloat(d['П']) || 0) * hoursMultiplier,
            capV: (parseFloat(d['В']) || 0) * hoursMultiplier,
            curDV: 0, curP: 0, curV: 0
        }));
        
        filteredPoints.forEach(p => {
            if (p.assigned) {
                const depo = workingDepo.find(d => d['Вузол'] === p.assigned);
                if (depo) {
                    depo.curDV += p.dv;
                    depo.curP += p.p;
                    depo.curV += p.v;
                    depo.color = p.color; 
                }
            }
        });

        document.getElementById('pointsStatus').innerHTML = `Відновлено сесію: ${filteredPoints.length} точок`;
        document.getElementById('pointsStatus').className = 'status ok';
        
        drawMap(workingDepo);

    } else {
        // --- ЗВИЧАЙНЕ ЗАВАНТАЖЕННЯ СИРИХ ДАНИХ ---
        rawPointsData = json.map(r => ({
            id: String(r['Вузол'] || r['ID']),
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
    }
};
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
    
    // Жорстко задані налаштування
    const priorityId = 'auto';
    const algoType = 'global';

    if (!useDV && !useP && !useV) return alert("Оберіть хоча б один тип вантажу!");
    if (hStart > hEnd) return alert("Невірний діапазон часу!");

    const hoursMultiplier = (hEnd - hStart) + 1;
    
    // Зберігаємо статуси виключення перед агрегацією
    const excludedPointsMap = {};
    if (filteredPoints.length > 0) {
        filteredPoints.forEach(p => {
            if (p.isExcluded) excludedPointsMap[p.id] = true;
        });
    }

    // Агрегація точок (ЗБЕРІГАЄМО НУЛЬОВІ ВІДДІЛЕННЯ)
    const aggregated = {};
    rawPointsData.forEach(r => {
        if (!aggregated[r.id]) {
            // Ініціалізуємо всі точки, навіть якщо обіг 0
            aggregated[r.id] = { id: r.id, lat: r.lat, lng: r.lng, dv: 0, p: 0, v: 0, assigned: null };
        }
    });

    rawPointsData.forEach(r => {
        if (r.hour >= hStart && r.hour <= hEnd) {
            aggregated[r.id].dv += r.dv || 0;
            aggregated[r.id].p += r.p || 0;
            aggregated[r.id].v += r.v || 0;
        }
    });
    filteredPoints = Object.values(aggregated);

    // Відновлюємо статуси виключення
    filteredPoints.forEach(p => {
        if (excludedPointsMap[p.id]) p.isExcluded = true;
    });
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

    // Сортування депо (Північ -> Південь)
    workingDepo.sort((a, b) => b['Широта'] - a['Широта']);

    // --- ФАЗА 1: РОЗПОДІЛ (Глобальна відстань) ---
    let pairs = [];
    filteredPoints.forEach(pt => {
        const ptGeo = turf.point([pt.lng, pt.lat]);
        workingDepo.forEach(depo => {
            const dist = turf.distance(ptGeo, turf.point([parseFloat(depo['Довгота']), parseFloat(depo['Широта'])]));
            pairs.push({ pt, depo, dist });
        });
    });

    pairs.sort((a, b) => a.dist - b.dist);

    for (let pair of pairs) {
        let { pt, depo } = pair;
        if (pt.assigned || pt.isExcluded) continue; 

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

    // --- ФАЗА 2: РОЗПОДІЛ ЗАЛИШКІВ ---
    let remainders = filteredPoints.filter(p => !p.assigned && !p.isExcluded);
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
// --- ВИВАНТАЖЕННЯ В EXCEL ---
window.exportToExcel = function() {
    if (!filteredPoints || filteredPoints.length === 0) {
        return alert("Немає даних для вивантаження! Спочатку завантажте файл.");
    }

    // Формуємо масив з розширеними метаданими
    const exportData = filteredPoints.map(p => ({
        "Відділення (Вузол)": p.id,
        "Широта": p.lat,
        "Довгота": p.lng,
        "ДВ (сума)": p.dv,
        "П (сума)": p.p,
        "В (сума)": p.v,
        "Призначене Депо": p.assigned || "Не розподілено",
        "Колір": p.color || "",
        "Є Звіздою": p.isStar ? "Так" : "Ні",
        "Прив'язана Звіздочка": p.assignedStar || ""
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Розподіл_Депо");

    XLSX.writeFile(wb, "Депошкер_Результат.xlsx");
};

// --- РУЧНЕ ПЕРЕМІЩЕННЯ ---
// --- РУЧНЕ ПЕРЕМІЩЕННЯ МІЖ ДЕПО ---
window.movePoint = function(pointId) {
    const input = document.getElementById('move-input-' + pointId);
    if (!input) return;
    const newDepoId = input.value;
    const pt = filteredPoints.find(p => p.id === pointId);
    const newDepo = workingDepo.find(d => d['Вузол'] === newDepoId);
    const oldDepo = workingDepo.find(d => d['Вузол'] === pt.assigned);

    if (pt && newDepo && pt.assigned !== newDepoId) {
        // 1. АВТОВІДВ'ЯЗУВАННЯ ВІД ЗІРОЧКИ
        if (pt.assignedStar) {
            const oldStar = filteredPoints.find(s => s.id === pt.assignedStar);
            if (oldStar) oldStar.starChildren--;
            pt.assignedStar = null;
        }
        
        // Якщо сама точка була зірочкою, знімаємо статус і відв'язуємо її "дітей"
        if (pt.isStar) {
            pt.isStar = false;
            pt.starChildren = 0;
            filteredPoints.forEach(child => {
                if (child.assignedStar === pt.id) child.assignedStar = null;
            });
        }

        // 2. Коригуємо завантаження старого депо
        if (oldDepo) {
            oldDepo.curDV -= pt.dv;
            oldDepo.curP -= pt.p;
            oldDepo.curV -= pt.v;
        }
        // 3. Додаємо вантаж новому депо
        newDepo.curDV += pt.dv;
        newDepo.curP += pt.p;
        newDepo.curV += pt.v;

        // 4. Змінюємо прив'язку та колір точки
        pt.assigned = newDepoId;
        pt.color = newDepo.color;
        
        map.closePopup();
        drawMap(workingDepo); // Перемальовуємо карту
    }
};

// --- ВИКЛЮЧЕННЯ ТОЧКИ З РОЗРАХУНКУ ---
window.toggleExcludePoint = function(pointId) {
    const pt = filteredPoints.find(p => p.id === pointId);
    if (!pt) return;
    
    pt.isExcluded = !pt.isExcluded;

    if (pt.isExcluded) {
        // Знімаємо вантаж з депо та відв'язуємо
        if (pt.assigned) {
            const depo = workingDepo.find(d => d['Вузол'] === pt.assigned);
            if (depo) {
                depo.curDV -= pt.dv; depo.curP -= pt.p; depo.curV -= pt.v;
            }
            pt.assigned = null;
        }
        // Відв'язуємо від зірочок
        if (pt.isStar) {
            pt.isStar = false; pt.starChildren = 0;
            filteredPoints.forEach(child => { if (child.assignedStar === pt.id) child.assignedStar = null; });
        }
        if (pt.assignedStar) {
            const oldStar = filteredPoints.find(s => s.id === pt.assignedStar);
            if (oldStar) oldStar.starChildren--;
            pt.assignedStar = null;
        }
    }
    
    map.closePopup();
    drawMap(workingDepo);
};

// --- РУЧНЕ КЕРУВАННЯ ЗІРОЧКАМИ ---
window.movePointStar = function(pointId) {
    const sel = document.getElementById('move-star-sel-' + pointId);
    if (!sel) return;

    const newStarId = sel.value;
    const pt = filteredPoints.find(p => p.id === pointId);

    if (!pt || pt.isStar) return; // Зірочку не можна прив'язати до іншої зірочки

    const oldStarId = pt.assignedStar;

    if (oldStarId !== newStarId) {
        // Віднімаємо лічильник від старої зірочки
        if (oldStarId && oldStarId !== 'none') {
            const oldStar = filteredPoints.find(s => s.id === oldStarId);
            if (oldStar) oldStar.starChildren--;
        }

        // Додаємо до нової зірочки або просто скидаємо
        if (newStarId !== 'none') {
            const newStar = filteredPoints.find(s => s.id === newStarId);
            if (newStar) newStar.starChildren++;
            pt.assignedStar = newStarId;
        } else {
            pt.assignedStar = null; // Скинули зірочку
        }

        map.closePopup();
        drawMap(workingDepo); // Перемальовуємо лінії
    }
};

// --- РУЧНЕ СТВОРЕННЯ ЗІРОЧКИ З АВТОДОБОРОМ ---
window.makeManualStar = function(pointId) {
    const pt = filteredPoints.find(p => p.id === pointId);
    if (!pt || pt.isStar) return;

    // Знімаємо з поточної зірочки, якщо була
    if (pt.assignedStar) {
        const oldStar = filteredPoints.find(s => s.id === pt.assignedStar);
        if (oldStar) oldStar.starChildren--;
        pt.assignedStar = null;
    }

    pt.isStar = true;
    pt.starChildren = 0;

    const depotId = pt.assigned;
    if (!depotId) return drawMap(workingDepo); // Якщо немає депо, просто малюємо

    const depotInfo = depoData.find(d => d['Вузол'] === depotId);
    if (!depotInfo) return drawMap(workingDepo);

    const depotPt = turf.point([parseFloat(depotInfo['Довгота']), parseFloat(depotInfo['Широта'])]);
    const starGeo = turf.point([pt.lng, pt.lat]);
    const distDepotToStar = turf.distance(depotPt, starGeo);
    const bearingDepotToStar = turf.bearing(depotPt, starGeo);
    
    const maxPointsPerStar = parseInt(document.getElementById('starMaxPoints').value) || 15;

    // Шукаємо вільні відділення у цьому ж депо
    let regularPoints = filteredPoints.filter(p => p.assigned === depotId && !p.isStar && !p.assignedStar);

    regularPoints.forEach(child => {
        if (pt.starChildren >= maxPointsPerStar) return; // Ліміт

        const childGeo = turf.point([child.lng, child.lat]);
        const distDepotToChild = turf.distance(depotPt, childGeo);
        
        // Правило: точка далі від депо, ніж зірочка, і в тому ж напрямку
        if (distDepotToChild > distDepotToStar) {
            const bearingDepotToChild = turf.bearing(depotPt, childGeo);
            let diff = Math.abs(bearingDepotToChild - bearingDepotToStar);
            if (diff > 180) diff = 360 - diff;
            
            if (diff <= 35) {
                child.assignedStar = pt.id;
                pt.starChildren++;
            }
        }
    });

    map.closePopup();
    drawMap(workingDepo);
};

// --- РУЧНЕ РОЗЗВІЗДОВУВАННЯ ---
window.unmakeManualStar = function(pointId) {
    const pt = filteredPoints.find(p => p.id === pointId);
    if (!pt || !pt.isStar) return;

    // Знімаємо статус зірочки
    pt.isStar = false;
    pt.starChildren = 0;

    // Відв'язуємо всі точки (дітей), що були прив'язані до цієї зірочки
    filteredPoints.forEach(child => {
        if (child.assignedStar === pt.id) {
            child.assignedStar = null;
        }
    });

    map.closePopup();
    drawMap(workingDepo); // Перемальовуємо карту, щоб прибрати лінії
};

// --- АЛГОРИТМ РОЗПОДІЛУ ЗІРОЧОК ---
window.distributeStars = function(depotId) {
    // 1. Беремо налаштування з UI
    const minArea = parseFloat(document.getElementById('starMinArea').value) || 150;
    const maxStars = parseInt(document.getElementById('starMaxCount').value) || 15;
    const maxPointsPerStar = parseInt(document.getElementById('starMaxPoints').value) || 15;

    // 2. Знаходимо Депо
    const depotInfo = depoData.find(d => d['Вузол'] === depotId);
    if (!depotInfo) return;
    const depotPt = turf.point([parseFloat(depotInfo['Довгота']), parseFloat(depotInfo['Широта'])]);

    // 3. Беремо всі відділення, які вже прив'язані до цього Депо
    let depotPoints = filteredPoints.filter(p => p.assigned === depotId);
    
    // Скидаємо попередні налаштування зірочок для цієї зони (якщо користувач натиснув вдруге)
    depotPoints.forEach(p => { p.isStar = false; p.assignedStar = null; });

    // 4. Шукаємо кандидатів у Зірочки (за площею)
    let potentialStars = [];
    depotPoints.forEach(p => {
        const areaInfo = areaData.find(a => String(a['Вузол']) === String(p.id));
        // В таблиці колонка називається "Площа"
        const area = areaInfo ? parseFloat(areaInfo['Площа']) : 0; 
        
        if (area >= minArea) {
            potentialStars.push({ 
                pt: p, 
                area: area, 
                dist: turf.distance(depotPt, turf.point([p.lng, p.lat])) 
            });
        }
    });

    // Сортуємо: чим більша площа, тим вищий пріоритет
    potentialStars.sort((a, b) => b.area - a.area);
    
    // Відрізаємо зайві по ліміту (maxStars)
    let actualStars = potentialStars.slice(0, maxStars).map(s => s.pt);
    actualStars.forEach(s => { s.isStar = true; s.starChildren = 0; });

    // 5. Розподіляємо ЗАЛИШКИ (відділення, що не стали зірочками)
    let regularPoints = depotPoints.filter(p => !p.isStar);

    regularPoints.forEach(pt => {
        const ptGeo = turf.point([pt.lng, pt.lat]);
        const distDepotToPt = turf.distance(depotPt, ptGeo);
        const bearingDepotToPt = turf.bearing(depotPt, ptGeo); // Вектор (кут) від Депо до Точки

        let bestStar = null;
        let minStarDist = Infinity;

        actualStars.forEach(star => {
            if (star.starChildren >= maxPointsPerStar) return; // Ліміт переповнено

            const starGeo = turf.point([star.lng, star.lat]);
            const distDepotToStar = turf.distance(depotPt, starGeo);
            
            // ПРАВИЛО 1: Точка має бути ДАЛІ ВІД ДЕПО, ніж Зірочка (щоб не повертатися назад)
            if (distDepotToPt > distDepotToStar) {
                const bearingDepotToStar = turf.bearing(depotPt, starGeo);
                
                // ПРАВИЛО 2: Вектор. Різниця кутів (азимутів). Беремо дельту в межах 35 градусів.
                let diff = Math.abs(bearingDepotToPt - bearingDepotToStar);
                if (diff > 180) diff = 360 - diff; // Коригування переходу через полюс (-180/180)

                if (diff <= 35) { // Входить в "конус" напрямку
                    const distStarToPt = turf.distance(starGeo, ptGeo);
                    // Вибираємо найближчу зірочку на цьому векторі
                    if (distStarToPt < minStarDist) {
                        minStarDist = distStarToPt;
                        bestStar = star;
                    }
                }
            }
        });

        // Якщо знайшли підходящу зірочку - прив'язуємо
        if (bestStar) {
            pt.assignedStar = bestStar.id;
            bestStar.starChildren++;
        }
    });

    // 6. Перемальовуємо карту з новими зв'язками
    map.closePopup();
    drawMap(workingDepo);
};

// --- ЗМІНА КОЛЬОРУ ЗОНИ ---
window.changeDepoColor = function(depoId, newColor) {
    if (!workingDepo || workingDepo.length === 0) return;

    // Шукаємо депо в робочому масиві
    const depo = workingDepo.find(d => d['Вузол'] === depoId);
    if (depo) {
        depo.color = newColor; // Оновлюємо колір депо
        
        // Оновлюємо колір усіх відділень, що належать до цього депо
        filteredPoints.forEach(p => {
            if (p.assigned === depoId) {
                p.color = newColor;
            }
        });
        
        // Перемальовуємо карту з новими кольорами
        drawMap(workingDepo);
    }
};

// Малювання на карті
// Малювання на карті
function drawMap(processedDepo = null) {
    layers.depo.clearLayers();
    layers.points.clearLayers();
    layers.polygons.clearLayers();
    layers.lines.clearLayers(); 
    oms.clearMarkers(); // Очищаємо веєр

    const dList = processedDepo || depoData;

    // 1. Полігони
    if (processedDepo) {
        processedDepo.forEach(depo => {
            if (window.activeDepotFilters.size > 0 && !window.activeDepotFilters.has(depo['Вузол'])) return;
            const depoPoints = filteredPoints.filter(p => p.assigned === depo['Вузол']);
            if (depoPoints.length >= 3) {
                const pts = depoPoints.map(p => turf.point([p.lng, p.lat]));
                pts.push(turf.point([parseFloat(depo['Довгота']), parseFloat(depo['Широта'])]));
                let hull = turf.convex(turf.featureCollection(pts));
                if (hull) {
                    hull = turf.buffer(hull, 0.3, { units: 'kilometers' });
                    hull = turf.polygonSmooth(hull, { iterations: 2 });
                    L.geoJSON(hull, { style: { color: depo.color, weight: 4, opacity: 1, fillOpacity: 0.3 } }).addTo(layers.polygons);
                }
            }
        });
    }

    // 2. Депо (Зі статистикою та підтримкою веєра)
    // 2. Депо (Зі статистикою та підтримкою веєра)
    depoData.forEach(d => {
        if (window.activeDepotFilters.size > 0 && !window.activeDepotFilters.has(d['Вузол'])) return;
        const isOff = d.isActive === false;
        const bgColor = isOff ? '#334155' : '#0f172a';
        const borderColor = isOff ? '#64748b' : '#38bdf8';
        
        const iconHtml = `<div style="background: ${bgColor}; color: #fff; border: 2px solid ${borderColor}; opacity: ${isOff ? 0.6 : 1}; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 10px; width: 100%; height: 100%;">${d['Вузол']}</div>`;
        const icon = L.divIcon({ className: '', html: iconHtml, iconSize: [36, 36], iconAnchor: [18, 18], popupAnchor: [0, -18] });
        const marker = L.marker([d['Широта'], d['Довгота']], {icon});
        
        const workDepoInfo = workingDepo ? workingDepo.find(wd => wd['Вузол'] === d['Вузол']) : null;
        const isDist = !!workDepoInfo;

        // Збір статистики Депо
        const dPts = filteredPoints.filter(p => p.assigned === d['Вузол']);
        const starsCount = dPts.filter(p => p.isStar).length;
        const orphansCount = dPts.filter(p => !p.isStar && !p.assignedStar).length;
        const depotRamps = d['Рампи'] || '0'; // <--- Читаємо рампи з Депо

        const popupHtml = `
            <div style="min-width: 200px;">
                <b style="color: ${isOff ? '#94a3b8' : '#f1f5f9'}">${d['Вузол']}</b><br>
                <div class="stats-grid">
                    <div class="stats-item"><span>Всього точок</span><b>${dPts.length}</b></div>
                    <div class="stats-item"><span>Звіздочок</span><b style="color:#eab308;">${starsCount}</b></div>
                    <div class="stats-item"><span>Без звізди</span><b>${orphansCount}</b></div>
                    <div class="stats-item"><span>Рампи</span><b style="color:#38bdf8;">${depotRamps}</b></div>
                </div>
                
                ${isDist ? `<div class="popup-row" style="font-size:0.75rem;">Колір зони: <input type="color" value="${workDepoInfo.color}" onchange="changeDepoColor('${d['Вузол']}', this.value)" style="border:none; background:transparent; cursor:pointer;"></div>` : ''}

                <button onclick="distributeStars('${d['Вузол']}')" class="popup-btn btn-blue" style="color:#000; background:#eab308; border-color:#eab308;" ${isOff || !isDist ? 'disabled' : ''}>🌟 Шукати Звіздочки</button>
                <div class="popup-row">
                    <button onclick="toggleDepo('${d['Вузол']}')" class="popup-btn btn-blue" style="flex:1; background:${isOff ? '#4ade80' : '#ef4444'}; color:${isOff ? '#000' : '#fff'};">${isOff ? 'ВКЛ' : 'ВИКЛ'}</button>
                    <button onclick="openDepotChart('${d['Вузол']}')" class="popup-btn btn-blue" style="flex:1;" ${isOff ? 'disabled' : ''}>📊 Графік</button>
                </div>
            </div>
        `;
        
        // Передаємо HTML у маркер, але не відкриваємо стандартним способом
        marker.customPopupHtml = popupHtml;

        // Імітуємо клік при наведенні для плагіна OMS (Веєр)
        marker.on('mouseover', function() {
            window.isHoverClick = true; 
            marker.fire('click'); 
            window.isHoverClick = false; 
        });

        layers.depo.addLayer(marker); 
        oms.addMarker(marker); // Додаємо Депо у веєр!
    });

    // 3. Точки відділень та Зірочки
    filteredPoints.forEach(p => {
        if (window.activeDepotFilters.size > 0 && p.assigned && !window.activeDepotFilters.has(p.assigned)) return;
        const isZero = (p.dv + p.p + p.v) === 0;
        const col = isZero ? '#334155' : (p.color || '#64748b'); // Сірий, якщо нульовий
        
        let iconHtml, iconSize;
        if (p.isExcluded) {
            iconHtml = `<div style="background:#000; width:16px; height:16px; border: 2px solid #ef4444; border-radius: 50%; display:flex; align-items:center; justify-content:center; color:#ef4444; font-size:12px; font-weight:bold; box-shadow: 0 0 8px #ef4444;">X</div>`;
            iconSize = [20, 20];
        } else if (p.isStar) {
            iconHtml = `<div style="background:${p.color}; width:20px; height:20px; border: 2px solid #fff; border-radius: 50%; box-shadow: 0 0 10px ${p.color}; display:flex; align-items:center; justify-content:center; font-size:12px; text-shadow: 1px 1px 2px #000;">⭐</div>`;
            iconSize = [24, 24];
        } else {
            const starColor = p.assignedStar ? p.color : (isZero ? '#1e293b' : '#ffffff'); 
            iconHtml = `<div style="background:${col}; width:14px; height:14px; border: 2px solid ${starColor}; border-radius: 50%;"></div>`;
            iconSize = [18, 18];
        }

        const icon = L.divIcon({ className: '', html: iconHtml, iconSize: iconSize });
        
        if (p.assignedStar) {
            const star = filteredPoints.find(s => s.id === p.assignedStar);
            if (star) L.polyline([[star.lat, star.lng], [p.lat, p.lng]], { color: p.color, weight: 2, dashArray: '5, 5', opacity: 0.8 }).addTo(layers.lines);
        }
        if (p.isStar) {
            const dPt = depoData.find(d => d['Вузол'] === p.assigned);
            if (dPt) L.polyline([[dPt['Широта'], dPt['Довгота']], [p.lat, p.lng]], { color: p.color, weight: 3, opacity: 0.9 }).addTo(layers.lines);
        }

        // Обчислення відстаней до депо (Радіус 200 км)
        const ptGeo = turf.point([p.lng, p.lat]);
        let depoSelectHtml = '';
        if (workingDepo) {
            workingDepo.forEach(d => {
                const dist = turf.distance(ptGeo, turf.point([d['Довгота'], d['Широта']]));
                if (dist <= 200) {
                    const isSelected = p.assigned === d['Вузол'] ? 'selected' : '';
                    depoSelectHtml += `<option value="${d['Вузол']}" ${isSelected}>${d['Вузол']} (${dist.toFixed(1)} км)</option>`;
                }
            });
        }

        let starOptionsHtml = '<option value="none">Немає</option>';
        if (!p.isStar) {
            filteredPoints.filter(s => s.isStar && s.assigned === p.assigned).forEach(s => {
                starOptionsHtml += `<option value="${s.id}" ${s.id === p.assignedStar ? 'selected' : ''}>${s.id}</option>`;
            });
        }

        // Дані про площу тепер тягнемо для ВСІХ відділень
        const areaInfo = areaData.find(a => String(a['Вузол']) === String(p.id)) || {};
        const ptArea = areaInfo['Площа'] || 0;
        const ptRamps = areaInfo['Рампи'] || '0';

        let starStatsHtml = '';
        if (p.isStar) {
            starStatsHtml = `
                <div class="stats-grid">
                    <div class="stats-item"><span>Підлеглих</span><b>${p.starChildren}</b></div>
                    <div class="stats-item"><span>Площа Звіздочки</span><b>${ptArea}</b></div>
                </div>
                <button onclick="openStarChart('${p.id}')" class="popup-btn btn-blue">📊 Графік Звіздочки</button>
            `;
        }

        const popupContent = `
            <div style="min-width: 250px;">
                <b style="font-size: 1.1rem; color: #fff;">${p.id}</b>
                ${isZero ? '<span style="color:#ef4444; font-size:0.75rem; float:right;">Нульовий обіг</span>' : ''}
                <hr style="margin:8px 0; border-color:var(--border);">
                
                ${starStatsHtml}
                
                <div style="font-size:0.75rem; color:var(--text-muted);">
                    ДВ: ${p.dv} | П: ${p.p} | В: ${p.v}<br>
                    Площа: <b>${ptArea}</b> | Рампи: <b>${ptRamps}</b>
                </div>
                <hr style="margin:8px 0; border-color:var(--border);">
                
                ${!p.isStar ? `
                <button onclick="makeManualStar('${p.id}')" class="popup-btn btn-yellow" style="width:100%; margin-bottom:10px;">⭐ Зазвіздувати</button>
                <div style="font-size:0.75rem; color:var(--text-muted); margin-bottom:4px;">Прив'язка до звіздочки:</div>
                <div class="popup-row">
                    <select id="move-star-sel-${p.id}" class="smart-input">${starOptionsHtml}</select>
                    <button onclick="movePointStar('${p.id}')" class="popup-btn btn-yellow">ОК</button>
                </div>` : `
                <button onclick="unmakeManualStar('${p.id}')" class="popup-btn btn-red" style="width:100%; margin-bottom:10px;">❌ Роззвіздувати</button>
                `}

                <div style="font-size:0.75rem; color:var(--text-muted); margin-bottom:4px;">Змінити Депо (< 200км):</div>
                <div class="popup-row">
                    <select id="move-input-${p.id}" class="smart-input">
                        ${depoSelectHtml}
                    </select>
                    <button onclick="movePoint('${p.id}')" class="popup-btn btn-green">ОК</button>
                </div>
                
                <button onclick="toggleExcludePoint('${p.id}')" class="popup-btn ${p.isExcluded ? 'btn-green' : 'btn-red'}" style="width:100%; margin-top:5px;">
                    ${p.isExcluded ? '✅ Повернути до розрахунку' : '⛔ Виключити з розрахунку'}
                </button>
            </div>
        `;

        const m = L.marker([p.lat, p.lng], {icon, zIndexOffset: p.isStar ? 1000 : 0});
        
        m.customPopupHtml = popupContent; 
        
        m.on('mouseover', function() {
            window.isHoverClick = true; 
            m.fire('click'); 
            window.isHoverClick = false; 
        });

        layers.points.addLayer(m);

        layers.points.addLayer(m); // Додаємо на шар карти
        oms.addMarker(m); // Додаємо до Spiderfier
    });
}


// --- ГРАФІКИ ДЛЯ ДЕПО ТА ЗІРОЧОК ---
let depotChartInstance = null;
let currentChartData = null; // Глобально зберігаємо дані поточного відкритого графіка
let currentChartTab = 'all';

window.openDepotChart = function(depotId) {
    if (!filteredPoints || filteredPoints.length === 0) return alert("Спочатку запустіть розподіл!");

    document.getElementById('chartDepotTitle').innerText = `Навантаження: ${depotId}`;

    const assignedBranchIds = filteredPoints.filter(p => p.assigned === depotId).map(p => p.id);
    const depoInfo = depoData.find(d => d['Вузол'] === depotId);

    const caps = { dv: parseFloat(depoInfo['ДВ']) || 0, p: parseFloat(depoInfo['П']) || 0, v: parseFloat(depoInfo['В']) || 0 };
    const loads = { dv: new Array(24).fill(0), p: new Array(24).fill(0), v: new Array(24).fill(0) };
    
    rawPointsData.forEach(r => {
        if (assignedBranchIds.includes(r.id) && r.hour >= 0 && r.hour <= 23) {
            loads.dv[r.hour] += r.dv || 0;
            loads.p[r.hour] += r.p || 0;
            loads.v[r.hour] += r.v || 0;
        }
    });

    currentChartData = { loads, caps };
    switchChartTab('all'); // Завжди відкриваємо на загальній вкладці
    document.getElementById('chartModal').style.display = 'flex';
};

window.openStarChart = function(starId) {
    document.getElementById('chartDepotTitle').innerText = `Навантаження Звіздочки: ${starId}`;
    
    const childrenIds = filteredPoints.filter(p => p.assignedStar === starId || p.id === starId).map(p => p.id);
    const loads = { dv: new Array(24).fill(0), p: new Array(24).fill(0), v: new Array(24).fill(0) };
    
    rawPointsData.forEach(r => {
        if (childrenIds.includes(r.id) && r.hour >= 0 && r.hour <= 23) {
            loads.dv[r.hour] += r.dv || 0;
            loads.p[r.hour] += r.p || 0;
            loads.v[r.hour] += r.v || 0;
        }
    });

    currentChartData = { loads, caps: {dv: 0, p: 0, v: 0} }; // Зірочки не мають лімітів, передаємо нулі
    switchChartTab('all');
    document.getElementById('chartModal').style.display = 'flex';
};

window.switchChartTab = function(tabId) {
    currentChartTab = tabId;
    
    // Оновлюємо UI кнопок
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.getAttribute('data-tab') === tabId) btn.classList.add('active');
    });

    renderChart();
};

window.closeChartModal = function() {
    document.getElementById('chartModal').style.display = 'none';
};

function renderChart() {
    const ctx = document.getElementById('depotChart').getContext('2d');
    if (depotChartInstance) depotChartInstance.destroy();

    const labels = Array.from({length: 24}, (_, i) => `${i}:00`);
    const { loads, caps } = currentChartData;
    let datasets = [];

    // Функції для додавання конкретних кривих
    const pushDV = () => {
        datasets.push({ label: 'Фактично ДВ', data: loads.dv, borderColor: '#38bdf8', borderWidth: 3, tension: 0.3, fill: false });
        if (caps.dv > 0) {
            datasets.push({ label: 'Ліміт ДВ', data: new Array(24).fill(caps.dv), borderColor: '#38bdf8', borderWidth: 2, borderDash: [5, 5], pointRadius: 0, fill: false });
            datasets.push({ label: '50% ДВ', data: new Array(24).fill(caps.dv / 2), borderColor: '#38bdf8', borderWidth: 1, borderDash: [2, 4], pointRadius: 0, fill: false });
        }
    };
    
    const pushP = () => {
        datasets.push({ label: 'Фактично П', data: loads.p, borderColor: '#4ade80', borderWidth: 3, tension: 0.3, fill: false });
        if (caps.p > 0) {
            datasets.push({ label: 'Ліміт П', data: new Array(24).fill(caps.p), borderColor: '#4ade80', borderWidth: 2, borderDash: [5, 5], pointRadius: 0, fill: false });
            datasets.push({ label: '50% П', data: new Array(24).fill(caps.p / 2), borderColor: '#4ade80', borderWidth: 1, borderDash: [2, 4], pointRadius: 0, fill: false });
        }
    };
    
    const pushV = () => {
        datasets.push({ label: 'Фактично В', data: loads.v, borderColor: '#facc15', borderWidth: 3, tension: 0.3, fill: false });
        if (caps.v > 0) {
            datasets.push({ label: 'Ліміт В', data: new Array(24).fill(caps.v), borderColor: '#facc15', borderWidth: 2, borderDash: [5, 5], pointRadius: 0, fill: false });
            datasets.push({ label: '50% В', data: new Array(24).fill(caps.v / 2), borderColor: '#facc15', borderWidth: 1, borderDash: [2, 4], pointRadius: 0, fill: false });
        }
    };

    // Фільтруємо залежно від обраної вкладки
    if (currentChartTab === 'all' || currentChartTab === 'dv') pushDV();
    if (currentChartTab === 'all' || currentChartTab === 'p') pushP();
    if (currentChartTab === 'all' || currentChartTab === 'v') pushV();

    depotChartInstance = new Chart(ctx, {
        type: 'line', 
        data: { labels: labels, datasets: datasets },
        options: {
            responsive: true, maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: { legend: { labels: { color: '#f1f5f9', usePointStyle: true, boxWidth: 8 } } },
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

// --- МАСОВЕ ПЕРЕНЕСЕННЯ ---
window.populateMassMoveDatalist = function() {
    const dl = document.getElementById('mass-depo-list');
    let html = '';
    workingDepo.forEach(d => { html += `<option value="${d['Вузол']}"></option>`; });
    dl.innerHTML = html;
};

window.closeMassMoveModal = function() {
    document.getElementById('massMoveModal').style.display = 'none';
    if(drawnItems) drawnItems.clearLayers();
};

window.applyMassMove = function() {
    const newDepoId = document.getElementById('massMoveDepoInput').value;
    const newDepo = workingDepo.find(d => d['Вузол'] === newDepoId);
    
    if (!newDepo) return alert("Оберіть існуюче депо зі списку!");

    window.selectedPointsForMassMove.forEach(pt => {
        if (pt.assigned === newDepoId) return; // Вже там

        const oldDepo = workingDepo.find(d => d['Вузол'] === pt.assigned);

        // Знімаємо з зірочок
        if (pt.assignedStar) {
            const oldStar = filteredPoints.find(s => s.id === pt.assignedStar);
            if (oldStar) oldStar.starChildren--;
            pt.assignedStar = null;
        }
        if (pt.isStar) {
            pt.isStar = false; pt.starChildren = 0;
            filteredPoints.forEach(child => { if (child.assignedStar === pt.id) child.assignedStar = null; });
        }

        // Перекидаємо вантаж
        if (oldDepo) {
            oldDepo.curDV -= pt.dv; oldDepo.curP -= pt.p; oldDepo.curV -= pt.v;
        }
        newDepo.curDV += pt.dv; newDepo.curP += pt.p; newDepo.curV += pt.v;

        pt.assigned = newDepoId;
        pt.color = newDepo.color;
    });

    closeMassMoveModal();
    drawMap(workingDepo);
};

// --- ФІЛЬТР ДЕПО НА КАРТІ ---


window.toggleFilterMenu = function() {
    const menu = document.getElementById('mapFilterMenu');
    if (menu.style.display === 'none' || menu.style.display === '') {
        buildFilterMenu();
        menu.style.display = 'flex';
    } else {
        menu.style.display = 'none';
    }
};

window.buildFilterMenu = function() {
    const menu = document.getElementById('mapFilterMenu');
    // Знаходимо тільки ті депо, у яких є відділення
    const activeDepots = [...new Set(filteredPoints.filter(p => p.assigned).map(p => p.assigned))].sort();
    
    if (activeDepots.length === 0) {
        menu.innerHTML = '<div style="color:var(--text-muted); font-size:0.8rem;">Немає призначених депо</div>';
        return;
    }

    // Перевіряємо, чи всі депо наразі вибрані
    const isAllChecked = window.activeDepotFilters.size === 0 || window.activeDepotFilters.size === activeDepots.length;

    // Додаємо чекбокс "Обрати всі" з невеликим візуальним відділенням
    let html = `<label style="font-weight: bold; border-bottom: 1px solid var(--border); padding-bottom: 8px; margin-bottom: 8px;">
                    <input type="checkbox" id="selectAllDepots" onchange="toggleAllDepots(this)" ${isAllChecked ? 'checked' : ''}> 
                    Обрати всі
                </label>`;
    
    activeDepots.forEach(depoId => {
        // Якщо фільтр порожній (size === 0), це означає, що показуються всі, тому галочки стоять
        const isChecked = window.activeDepotFilters.size === 0 || window.activeDepotFilters.has(depoId);
        
        // Додаємо клас "depot-filter-cb", щоб легко масово ними керувати
        html += `<label><input type="checkbox" class="depot-filter-cb" value="${depoId}" onchange="updateDepotFilter()" ${isChecked ? 'checked' : ''}> ${depoId}</label>`;
    });
    
    menu.innerHTML = html;
};

// Нова функція для керування масовим виділенням
window.toggleAllDepots = function(selectAllCb) {
    const checkboxes = document.querySelectorAll('.depot-filter-cb');
    // Ставимо або знімаємо всі галочки залежно від стану "Обрати всі"
    checkboxes.forEach(cb => {
        cb.checked = selectAllCb.checked;
    });
    // Запускаємо оновлення карти
    updateDepotFilter();
};

window.updateDepotFilter = function() {
    const checkboxes = document.querySelectorAll('.depot-filter-cb');
    const selectAllCb = document.getElementById('selectAllDepots');
    
    window.activeDepotFilters.clear();
    let selectedNames = [];
    let allChecked = true;

    checkboxes.forEach(cb => {
        if (cb.checked) {
            window.activeDepotFilters.add(cb.value);
            selectedNames.push(cb.value);
        } else {
            allChecked = false;
        }
    });

    // Синхронізуємо стан головного чекбокса (якщо користувач вручну проклікав усі)
    if (selectAllCb) {
        selectAllCb.checked = allChecked;
    }

    const input = document.getElementById('mapFilterInput');
    
    if (allChecked) {
        // Якщо вибрані всі, очищаємо фільтр (карта покаже все без обмежень)
        window.activeDepotFilters.clear(); 
        input.value = ''; 
    } else if (selectedNames.length === 0) {
        // Якщо користувач зняв усі галочки - додаємо фейковий ID, щоб на карті не відображалося нічого
        window.activeDepotFilters.add('__NONE__');
        input.value = 'Нічого не обрано';
    } else {
        // Якщо вибрана тільки частина депо
        input.value = selectedNames.join(', ');
    }

    drawMap(workingDepo.length > 0 ? workingDepo : null);
};

// Закриваємо меню кліком поза ним
document.addEventListener('click', function(e) {
    const container = document.getElementById('mapFilterContainer');
    if (container && !container.contains(e.target)) {
        document.getElementById('mapFilterMenu').style.display = 'none';
    }
});