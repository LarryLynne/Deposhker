// ВСТАВЬТЕ СЮДА ВАШУ ССЫЛКУ НА APPS SCRIPT
const GAS_URL = "https://script.google.com/macros/s/AKfycbw165Z4yzi8xk8MEF5XnqE0LB2tANhIOBUcN-5nqG7bqoBBOTShYn2fAQQxh4S-Tj7mHw/exec"; 

let map;
let layers = { depo: L.layerGroup(), polygons: L.layerGroup(), lines: L.layerGroup(), points: L.layerGroup() };
let depoData = [];
let rawPointsData = [];
let filteredPoints = [];
let workingDepo = [];
let areaData = [];

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
    
    layers.polygons.addTo(map);
    layers.lines.addTo(map);
    layers.depo.addTo(map);
    layers.points.addTo(map);
    
    const cachedDepo = localStorage.getItem('deposhker_depo_cache');
    const cachedArea = localStorage.getItem('deposhker_area_cache'); // Читаем кэш площадей
    
    if (cachedArea) {
        areaData = JSON.parse(cachedArea);
    }

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
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById('fileName').innerText = file.name;
    
    // Перевірка, чи завантажені депо, оскільки вони потрібні для відновлення робочого стану
    if (depoData.length === 0) {
        alert("Увага: Спочатку завантажте дані Депо (Крок 1), інакше збережена сесія може завантажитись некоректно.");
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type: 'array'});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws);
        
        // Визначаємо, чи це збережений файл результатів, чи сирі дані
        const isSavedFile = json.length > 0 && json[0]['Призначене Депо'] !== undefined;

        if (isSavedFile) {
            // --- ВІДНОВЛЕННЯ ЗБЕРЕЖЕНОГО СТАНУ ---
            
            // 1. Відновлюємо базові точки (фіктивний час потрібен, щоб не впав перерозподіл, якщо юзер захоче його запустити)
            rawPointsData = json.map(r => ({
                id: r['Відділення (Вузол)'],
                lat: parseFloat(r['Широта']),
                lng: parseFloat(r['Довгота']),
                hour: 10, 
                dv: parseFloat(r['ДВ (сума)']) || 0,
                p: parseFloat(r['П (сума)']) || 0,
                v: parseFloat(r['В (сума)']) || 0
            }));
            
            // 2. Відновлюємо розподілені точки зі всіма метаданими
            filteredPoints = json.map(r => ({
                id: r['Відділення (Вузол)'],
                lat: parseFloat(r['Широта']),
                lng: parseFloat(r['Довгота']),
                dv: parseFloat(r['ДВ (сума)']) || 0,
                p: parseFloat(r['П (сума)']) || 0,
                v: parseFloat(r['В (сума)']) || 0,
                assigned: r['Призначене Депо'] === "Не розподілено" ? null : r['Призначене Депо'],
                color: r['Колір'] || '#64748b',
                isStar: r['Є Зірочкою'] === 'Так',
                assignedStar: r["Прив'язана Зірочка"] || null,
                starChildren: 0 
            }));
            
            // Відновлюємо лічильники кількості підлеглих у зірочок
            filteredPoints.forEach(p => {
                if (p.assignedStar) {
                    const star = filteredPoints.find(s => s.id === p.assignedStar);
                    if (star) star.starChildren++;
                }
            });

            // 3. Відновлюємо робочий масив Депо (workingDepo)
            const hStart = parseInt(document.getElementById('hourStart').value) || 10;
            const hEnd = parseInt(document.getElementById('hourEnd').value) || 12;
            const hoursMultiplier = Math.max(1, (hEnd - hStart) + 1);

            workingDepo = depoData.map(d => ({
                ...d,
                color: '#38bdf8', // Колір за замовчуванням
                capDV: (parseFloat(d['ДВ']) || 0) * hoursMultiplier,
                capP: (parseFloat(d['П']) || 0) * hoursMultiplier,
                capV: (parseFloat(d['В']) || 0) * hoursMultiplier,
                curDV: 0, curP: 0, curV: 0
            }));
            
            // Наповнюємо депо вантажами та кольорами з відновлених відділень
            filteredPoints.forEach(p => {
                if (p.assigned) {
                    const depo = workingDepo.find(d => d['Вузол'] === p.assigned);
                    if (depo) {
                        depo.curDV += p.dv;
                        depo.curP += p.p;
                        depo.curV += p.v;
                        depo.color = p.color; // Депо наслідує колір своїх точок
                    }
                }
            });

            document.getElementById('pointsStatus').innerHTML = `Відновлено сесію: ${filteredPoints.length} точок`;
            document.getElementById('pointsStatus').className = 'status ok';
            
            // Одразу малюємо карту зі збереженими зонами
            drawMap(workingDepo);

        } else {
            // --- ЗВИЧАЙНЕ ЗАВАНТАЖЕННЯ СИРИХ ДАНИХ (Стара логіка) ---
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
        }
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
    
    // Жорстко задані налаштування
    const priorityId = 'auto';
    const algoType = 'global';

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
        if (pt.assigned) continue; 

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
        "Є Зірочкою": p.isStar ? "Так" : "Ні",
        "Прив'язана Зірочка": p.assignedStar || ""
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Розподіл_Депо");

    XLSX.writeFile(wb, "Депошкер_Результат.xlsx");
};

// --- РУЧНЕ ПЕРЕМІЩЕННЯ ---
// --- РУЧНЕ ПЕРЕМІЩЕННЯ МІЖ ДЕПО ---
window.movePoint = function(pointId) {
    const sel = document.getElementById('move-sel-' + pointId);
    if (!sel) return;
    
    const newDepoId = sel.value;
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
function drawMap(processedDepo = null) {
    layers.depo.clearLayers();
    layers.points.clearLayers();
    layers.polygons.clearLayers();
    layers.lines.clearLayers(); 

    const dList = processedDepo || depoData;

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
                    hull = turf.buffer(hull, 0.3, { units: 'kilometers' });
                    hull = turf.polygonSmooth(hull, { iterations: 2 });

                    L.geoJSON(hull, {
                        style: { color: depo.color, weight: 4, opacity: 1, fillOpacity: 0.3 }
                    }).addTo(layers.polygons);
                }
            }
        });
    }

    // 2. Депо
    depoData.forEach(d => {
        const isOff = d.isActive === false;
        
        const bgColor = isOff ? '#334155' : '#0f172a';
        const borderColor = isOff ? '#64748b' : '#38bdf8';
        const opacity = isOff ? 0.6 : 1;

        const iconHtml = `<div style="
            background: ${bgColor}; color: #fff; border: 2px solid ${borderColor}; 
            opacity: ${opacity}; border-radius: 50%; display: flex; align-items: center; 
            justify-content: center; font-weight: bold; font-size: 10px; width: 100%; 
            height: 100%; box-sizing: border-box; margin: 0; padding: 0; 
            line-height: 1; white-space: nowrap;">${d['Вузол']}</div>`;

        const icon = L.divIcon({ 
            className: '', html: iconHtml, iconSize: [36, 36], 
            iconAnchor: [18, 18], popupAnchor: [0, -18]
        });
        
        const marker = L.marker([d['Широта'], d['Довгота']], {icon});
        
        // Визначаємо поточний колір депо (якщо розподіл вже був)
        const workDepoInfo = workingDepo ? workingDepo.find(wd => wd['Вузол'] === d['Вузол']) : null;
        const currentColor = workDepoInfo ? workDepoInfo.color : '#38bdf8'; 
        const isDistributed = !!workDepoInfo; // Перевірка, чи можна змінювати колір

        // HTML для Popup меню (ДОДАНО ПІКЕР КОЛЬОРУ)
        const popupHtml = `
            <div style="text-align:center; min-width: 150px;">
                <b style="color: ${isOff ? '#94a3b8' : '#f1f5f9'}">${d['Вузол']}</b><br>
                <hr style="margin:8px 0; border-color:#334155;">
                
                ${isDistributed ? `
                <div style="margin-bottom: 8px; text-align: left; font-size: 0.75rem; color: #94a3b8; display:flex; align-items:center; justify-content:space-between;">
                    Колір зони: 
                    <input type="color" value="${currentColor}" onchange="changeDepoColor('${d['Вузол']}', this.value)" style="width:40px; height: 25px; border:none; padding: 0; background: transparent; cursor: pointer;">
                </div>
                ` : ''}

                <button onclick="distributeStars('${d['Вузол']}')" style="width:100%; padding:6px; background:#eab308; color:#000; border:none; border-radius:4px; cursor:pointer; font-weight:bold; margin-bottom:8px;" ${isOff || !isDistributed ? 'disabled' : ''}>
                    🌟 Шукати Зірочки
                </button>

                <button onclick="toggleDepo('${d['Вузол']}')" style="width:100%; padding:6px; background:${isOff ? '#4ade80' : '#ef4444'}; color:${isOff ? '#000' : '#fff'}; border:none; border-radius:4px; cursor:pointer; font-weight:bold; margin-bottom:8px;">
                    ${isOff ? 'ВКЛЮЧИТИ' : 'ВИКЛЮЧИТИ'}
                </button>
                
                <button onclick="openDepotChart('${d['Вузол']}')" style="width:100%; padding:6px; background:#1e293b; color:#fff; border:1px solid #38bdf8; border-radius:4px; cursor:pointer;" ${isOff ? 'disabled' : ''}>
                    📊 Графік
                </button>
            </div>
        `;
        
        marker.bindPopup(popupHtml);
        marker.addTo(layers.depo);
    });

    // 3. Точки відділень та Зірочки
    filteredPoints.forEach(p => {
        const col = p.color || '#64748b';
        let iconHtml, iconSize;
        let zIndexOffset = 0;

        // --- ВІЗУАЛІЗАЦІЯ ЗІРОЧОК (Колір зони) ---
        if (p.isStar) {
            iconHtml = `<div style="
                background:${p.color}; width:20px; height:20px; border: 2px solid #fff; 
                border-radius: 50%; box-shadow: 0 0 10px ${p.color}; display:flex; 
                align-items:center; justify-content:center; font-size:12px; text-shadow: 1px 1px 2px #000;">⭐</div>`;
            iconSize = [24, 24];
            zIndexOffset = 1000; 
        } else {
            // Обводка кольором зони, якщо належить до зірочки
            const isAssignedToStar = p.assignedStar ? p.color : '#ffffff'; 
            iconHtml = `<div class="point-marker" style="
                background:${col}; width:14px; height:14px; 
                border: 2px solid ${isAssignedToStar}; border-radius: 50%;
                box-shadow: 0 0 4px rgba(0,0,0,0.6);"></div>`;
            iconSize = [18, 18];
        }

        const icon = L.divIcon({ className: '', html: iconHtml, iconSize: iconSize });
        
        // Малюємо лінію від Зірочки до Точки (Колір зони)
        if (p.assignedStar) {
            const starPoint = filteredPoints.find(s => s.id === p.assignedStar);
            if (starPoint) {
                L.polyline([
                    [starPoint.lat, starPoint.lng],
                    [p.lat, p.lng]
                ], { color: p.color, weight: 2, dashArray: '5, 5', opacity: 0.8 }).addTo(layers.lines);
            }
        }
        
        // Малюємо лінію від Депо до Зірочки (Колір зони)
        if (p.isStar) {
            const depoDataPoint = depoData.find(d => d['Вузол'] === p.assigned);
            if (depoDataPoint) {
                L.polyline([
                    [depoDataPoint['Широта'], depoDataPoint['Довгота']],
                    [p.lat, p.lng]
                ], { color: p.color, weight: 3, opacity: 0.9 }).addTo(layers.lines);
            }
        }

        // Генеруємо список Депо
        let optionsHtml = '';
        if (workingDepo && workingDepo.length > 0) {
            optionsHtml = workingDepo.map(d => {
                const selected = (d['Вузол'] === p.assigned) ? 'selected' : '';
                return `<option value="${d['Вузол']}" ${selected}>${d['Вузол']}</option>`;
            }).join('');
        }

        // Генеруємо список доступних Зірочок у цій зоні
        let starOptionsHtml = '<option value="none">Немає</option>';
        if (!p.isStar) {
            const depotStars = filteredPoints.filter(s => s.isStar && s.assigned === p.assigned);
            depotStars.forEach(s => {
                const selected = (s.id === p.assignedStar) ? 'selected' : '';
                starOptionsHtml += `<option value="${s.id}" ${selected}>${s.id}</option>`;
            });
        }

        const starInfo = p.isStar ? `<b style="color:${p.color};">Це Зірочка! (В підпорядкуванні: ${p.starChildren})</b><hr>` : '';
        const assignedStarInfo = p.assignedStar ? `<span style="color:${p.color};">Підпорядковано зірочці: <b>${p.assignedStar}</b></span><hr>` : '';

        // Вміст Popup відділення
        const popupContent = `
            <div style="font-size:0.85rem; color:#f1f5f9; min-width: 220px;">
                <b>${p.id}</b><br>
                Депо: <b>${p.assigned || 'Немає'}</b><br>
                ${starInfo}
                ${assignedStarInfo}
                ДВ: ${p.dv} | П: ${p.p} | В: ${p.v}
                <hr style="margin:8px 0; border-color:#334155;">
                
                ${!p.isStar ? `
                <div style="font-size:0.75rem; color:#94a3b8; margin-bottom:4px;">Прив'язка до зірочки:</div>
                <div style="display:flex; gap:5px; margin-bottom:12px;">
                    <select id="move-star-sel-${p.id}" style="flex-grow:1; padding:6px; background:#0f172a; color:#fff; border:1px solid #334155; border-radius:4px;">
                        ${starOptionsHtml}
                    </select>
                    <button onclick="movePointStar('${p.id}')" style="padding:6px 10px; background:#1e293b; color:#eab308; border:1px solid #eab308; border-radius:4px; cursor:pointer;">
                        ОК
                    </button>
                </div>
                ` : ''}

                <div style="font-size:0.75rem; color:#94a3b8; margin-bottom:4px;">Змінити Депо:</div>
                <div style="display:flex; gap:5px;">
                    <select id="move-sel-${p.id}" style="flex-grow:1; padding:6px; background:#0f172a; color:#fff; border:1px solid #334155; border-radius:4px;">
                        ${optionsHtml}
                    </select>
                    <button onclick="movePoint('${p.id}')" style="padding:6px 10px; background:#1e293b; color:#4ade80; border:1px solid #4ade80; border-radius:4px; cursor:pointer;">
                        ОК
                    </button>
                </div>
            </div>
        `;

        const m = L.marker([p.lat, p.lng], {icon, zIndexOffset});
        m.bindTooltip(`<b>${p.id}</b><br>Депо: ${p.assigned || 'Немає'}${p.assignedStar ? '<br>Зірочка: ' + p.assignedStar : ''}`);
        
        if (workingDepo && workingDepo.length > 0) {
            m.bindPopup(popupContent, { minWidth: 220 });
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