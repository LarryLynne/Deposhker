// ВСТАВЬТЕ СЮДА ВАШУ ССЫЛКУ НА APPS SCRIPT
const GAS_URL = "https://script.google.com/macros/s/AKfycbw165Z4yzi8xk8MEF5XnqE0LB2tANhIOBUcN-5nqG7bqoBBOTShYn2fAQQxh4S-Tj7mHw/exec"; 

let map;
let layers = { 
    depo: L.layerGroup(), polygons: L.layerGroup(), lines: L.layerGroup(), 
    points: L.layerGroup(), route: L.layerGroup(), routeHalos: L.layerGroup(),
    potentialStarHalos: L.layerGroup() // <--- Додано шар для підсвітки
};
let depoData = [];
let rawPointsData = [];
let filteredPoints = [];
let workingDepo = [];
let areaData = [];

window.activeDepotFilters = new Set(); // Порожній = показувати всі

const distinctColors = [
    '#FF0000', '#0000FF', '#00AA00', '#FF8800', '#8800FF', '#00FFFF', '#FF00FF', 
    '#DDDD00', '#008080', '#800000', '#000080', '#FF8080', '#80FF80', '#808000'
];

let oms; 
let drawnItems; 

function initMap() {
    map = L.map('map', {zoomControl: false}).setView([50.4501, 30.5234], 11);
    L.control.zoom({ position: 'topright' }).addTo(map);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { attribution: '© OS' }).addTo(map);
    
    layers.polygons.addTo(map);
    layers.lines.addTo(map);
    layers.depo.addTo(map);
    layers.points.addTo(map);
    layers.route.addTo(map);
    layers.routeHalos.addTo(map);
    layers.potentialStarHalos.addTo(map);
    
    oms = new OverlappingMarkerSpiderfier(map, {
        keepSpiderfied: true, 
        nearbyDistance: 15,
        legWeight: 2, 
        circleSpiralSwitchover: 9
    });

    const globalPopup = new L.Popup({ minWidth: 260 });
    
    oms.addListener('click', function(marker) {
        if (window.isHoverClick) return; 
        
        if (window.isRoutingMode) {
            if (marker.pointData) handleRouteClick(marker.pointData, 'point');
            if (marker.depotData) handleRouteClick(marker.depotData, 'depot');
            return;
        }

        if (marker.customPopupHtml) {
            globalPopup.setContent(marker.customPopupHtml);
            globalPopup.setLatLng(marker.getLatLng());
            map.openPopup(globalPopup);
        }
    });

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

    map.on(L.Draw.Event.CREATED, function (e) {
        drawnItems.clearLayers(); 
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

async function fetchDepoData() {
    const statusEl = document.getElementById('depoStatus');
    if (!GAS_URL || GAS_URL.includes("ВАШ_ID")) {
        statusEl.innerText = 'Вкажіть URL у файлі script.js!'; statusEl.className = 'status error'; return;
    }
    statusEl.innerText = 'Завантаження Даних...'; statusEl.className = 'status';
    try {
        const response = await fetch(GAS_URL);
        const json = await response.json();
        if (json.status === 'success') {
            depoData = json.data.depots;
            areaData = json.data.areas || [];
            localStorage.setItem('deposhker_depo_cache', JSON.stringify(depoData));
            localStorage.setItem('deposhker_area_cache', JSON.stringify(areaData));
            statusEl.innerHTML = `Успішно: ${depoData.length} депо, ${areaData.length} площ`;
            statusEl.className = 'status ok';
            drawMap();
        } else { throw new Error(json.message); }
    } catch (e) {
        statusEl.innerText = 'Помилка: ' + e.message; statusEl.className = 'status error';
    }
}

let currentWorkbook = null;

document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById('fileName').innerText = file.name;
    if (depoData.length === 0) alert("Увага: Спочатку завантажте дані Депо (Крок 1), інакше збережена сесія може завантажитись некоректно.");
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        currentWorkbook = XLSX.read(data, {type: 'array'});
        const select = document.getElementById('sheetSelect');
        select.innerHTML = '';
        let validSheetsCount = 0;

        currentWorkbook.SheetNames.forEach(sheetName => {
            if (sheetName === "Сирі_Дані") return;
            const ws = currentWorkbook.Sheets[sheetName];
            const headers = XLSX.utils.sheet_to_json(ws, {header: 1})[0] || [];
            const hasLat = headers.includes('Широта');
            const hasLng = headers.includes('Довгота');
            const hasId = headers.includes('Вузол') || headers.includes('ID') || headers.includes('Відділення (Вузол)');

            if (hasLat && hasLng && hasId) {
                const opt = document.createElement('option'); opt.value = sheetName; opt.innerText = sheetName;
                select.appendChild(opt); validSheetsCount++;
            }
        });

        const selectionBlock = document.getElementById('sheetSelectionBlock');
        if (validSheetsCount === 0) {
            alert("У файлі не знайдено аркушів із потрібним форматом (мають бути колонки Широта, Довгота та Вузол/ID).");
            selectionBlock.style.display = 'none';
        } else if (validSheetsCount === 1) {
            selectionBlock.style.display = 'none'; processSelectedSheet(select.options[0].value);
        } else {
            selectionBlock.style.display = 'block';
            document.getElementById('pointsStatus').innerHTML = `Знайдено аркушів: ${validSheetsCount}. Оберіть потрібний.`;
            document.getElementById('pointsStatus').className = 'status';
        }
    };
    reader.readAsArrayBuffer(file);
});

window.processSelectedSheet = function(sheetNameFromArg = null) {
    const sheetName = sheetNameFromArg || document.getElementById('sheetSelect').value;
    if (!sheetName || !currentWorkbook) return;
    document.getElementById('sheetSelectionBlock').style.display = 'none';

    const ws = currentWorkbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws);
    
    const getVal = (row, keys) => {
        for (let k in row) { if (keys.includes(k.trim())) return row[k]; }
        return undefined;
    };
    
    const isSavedFile = json.length > 0 && getVal(json[0], ['Призначене Депо']) !== undefined;

    if (isSavedFile) {
        const hStart = parseInt(document.getElementById('hourStart').value) || 10;
        const hEnd = parseInt(document.getElementById('hourEnd').value) || 12;
        const hoursMultiplier = Math.max(1, (hEnd - hStart) + 1);

        if (currentWorkbook.Sheets["Сирі_Дані"]) {
            const rawJson = XLSX.utils.sheet_to_json(currentWorkbook.Sheets["Сирі_Дані"]);
            rawPointsData = rawJson.map(r => ({
                id: String(getVal(r, ['Відділення (Вузол)', 'Вузол', 'ID'])).trim(),
                lat: parseFloat(getVal(r, ['Широта'])) || 0,
                lng: parseFloat(getVal(r, ['Довгота'])) || 0,
                hour: parseInt(getVal(r, ['Година', 'Час'])) || 0,
                dv: parseFloat(getVal(r, ['ДВ'])) || 0,
                p: parseFloat(getVal(r, ['П'])) || 0,
                v: parseFloat(getVal(r, ['В'])) || 0
            }));
        } else {
            rawPointsData = [];
            json.forEach(r => {
                const dvTotal = parseFloat(getVal(r, ['ДВ (сума)', 'ДВ'])) || 0;
                const pTotal = parseFloat(getVal(r, ['П (сума)', 'П'])) || 0;
                const vTotal = parseFloat(getVal(r, ['В (сума)', 'В'])) || 0;
                for (let h = hStart; h <= hEnd; h++) {
                    rawPointsData.push({
                        id: String(getVal(r, ['Відділення (Вузол)', 'Вузол', 'ID'])).trim(),
                        lat: parseFloat(getVal(r, ['Широта'])) || 0,
                        lng: parseFloat(getVal(r, ['Довгота'])) || 0,
                        hour: h, dv: dvTotal / hoursMultiplier, p: pTotal / hoursMultiplier, v: vTotal / hoursMultiplier
                    });
                }
            });
        }
        
        filteredPoints = json.map(r => {
            const isStar = (getVal(r, ['Є Звіздою']) === 'Так' || getVal(r, ['Є Зірочкою']) === 'Так');
            const explicitlyApproved = getVal(r, ['Затверджена Зірочка']);
            // Якщо є колонка Затверджена Зірочка - беремо її, інакше якщо це зірка з файлу, вона за замовчуванням затверджена
            const isApprovedStar = explicitlyApproved !== undefined ? (explicitlyApproved === 'Так') : isStar;

            return {
                id: String(getVal(r, ['Відділення (Вузол)', 'Вузол', 'ID'])).trim(),
                lat: parseFloat(getVal(r, ['Широта'])) || 0,
                lng: parseFloat(getVal(r, ['Довгота'])) || 0,
                dv: parseFloat(getVal(r, ['ДВ (сума)', 'ДВ'])) || 0,
                p: parseFloat(getVal(r, ['П (сума)', 'П'])) || 0,
                v: parseFloat(getVal(r, ['В (сума)', 'В'])) || 0,
                assigned: getVal(r, ['Призначене Депо']) === "Не розподілено" ? null : getVal(r, ['Призначене Депо']),
                color: getVal(r, ['Колір']) || '#64748b',
                isStar: isStar,
                isApprovedStar: isApprovedStar,
                assignedStar: getVal(r, ["Прив'язана Звіздочка"]) || getVal(r, ["Прив'язана Зірочка"]) || null,
                starColor: getVal(r, ['Колір Зірки']) || null, 
                starChildren: 0 
            };
        });
        
        filteredPoints.forEach(p => {
            if (p.assignedStar) {
                const star = filteredPoints.find(s => s.id === p.assignedStar);
                if (star) star.starChildren++;
            }
        });

        workingDepo = depoData.map(d => ({
            ...d, color: '#38bdf8', 
            capDV: (parseFloat(d['ДВ']) || 0) * hoursMultiplier,
            capP: (parseFloat(d['П']) || 0) * hoursMultiplier,
            capV: (parseFloat(d['В']) || 0) * hoursMultiplier,
            curDV: 0, curP: 0, curV: 0
        }));
        
        filteredPoints.forEach(p => {
            if (p.assigned) {
                const depo = workingDepo.find(d => String(d['Вузол']).trim() === String(p.assigned).trim());
                if (depo) { depo.curDV += p.dv; depo.curP += p.p; depo.curV += p.v; depo.color = p.color; }
            }
        });

        document.getElementById('pointsStatus').innerHTML = `Відновлено сесію: ${filteredPoints.length} точок`;
        document.getElementById('pointsStatus').className = 'status ok';
        drawMap(workingDepo);

    } else {
        rawPointsData = json.map(r => ({
            id: String(getVal(r, ['Вузол', 'ID', 'Відділення (Вузол)'])).trim(),
            lat: parseFloat(getVal(r, ['Широта'])) || 0,
            lng: parseFloat(getVal(r, ['Довгота'])) || 0,
            hour: parseInt(getVal(r, ['Час', 'Година'])) || 0,
            dv: parseFloat(getVal(r, ['ДВ'])) || 0, p: parseFloat(getVal(r, ['П'])) || 0, v: parseFloat(getVal(r, ['В'])) || 0
        })).filter(r => r.lat !== 0 && r.lng !== 0);
        
        const uniquePoints = {};
        rawPointsData.forEach(p => {
            if (!uniquePoints[p.id]) {
                uniquePoints[p.id] = { id: p.id, lat: p.lat, lng: p.lng, dv: 0, p: 0, v: 0, assigned: null, color: '#64748b', isApprovedStar: false };
            }
        });
        filteredPoints = Object.values(uniquePoints);
        document.getElementById('pointsStatus').innerHTML = `Завантажено рядків: ${rawPointsData.length}`;
        document.getElementById('pointsStatus').className = 'status ok';
        drawMap();
    }
};

function runDistribution() {
    if (depoData.length === 0 || rawPointsData.length === 0) return alert("Завантажте дані депо та відділень!");

    const hStart = parseInt(document.getElementById('hourStart').value);
    const hEnd = parseInt(document.getElementById('hourEnd').value);
    const useDV = document.getElementById('chkDV').checked;
    const useP = document.getElementById('chkP').checked;
    const useV = document.getElementById('chkV').checked;
    
    if (!useDV && !useP && !useV) return alert("Оберіть хоча б один тип вантажу!");
    if (hStart > hEnd) return alert("Невірний діапазон часу!");

    const hoursMultiplier = (hEnd - hStart) + 1;
    
    const excludedPointsMap = {};
    if (filteredPoints.length > 0) {
        filteredPoints.forEach(p => { if (p.isExcluded) excludedPointsMap[p.id] = true; });
    }

    const starsConfig = {};
    if (filteredPoints.length > 0) {
        filteredPoints.forEach(p => {
            if (p.isStar || p.assignedStar) {
                starsConfig[p.id] = { isStar: p.isStar, isApprovedStar: p.isApprovedStar, assignedStar: p.assignedStar, starColor: p.starColor, starChildren: p.starChildren };
            }
        });
    }

    const aggregated = {};
    rawPointsData.forEach(r => {
        if (!aggregated[r.id]) aggregated[r.id] = { id: r.id, lat: r.lat, lng: r.lng, dv: 0, p: 0, v: 0, assigned: null };
    });

    rawPointsData.forEach(r => {
        if (r.hour >= hStart && r.hour <= hEnd) {
            aggregated[r.id].dv += r.dv || 0; aggregated[r.id].p += r.p || 0; aggregated[r.id].v += r.v || 0;
        }
    });
    filteredPoints = Object.values(aggregated);

    filteredPoints.forEach(p => {
        if (excludedPointsMap[p.id]) p.isExcluded = true;
        if (starsConfig[p.id]) {
            p.isStar = starsConfig[p.id].isStar;
            p.isApprovedStar = starsConfig[p.id].isApprovedStar;
            p.assignedStar = starsConfig[p.id].assignedStar;
            p.starColor = starsConfig[p.id].starColor;
            p.starChildren = starsConfig[p.id].starChildren;
        }
    });

    workingDepo = depoData.filter(d => d.isActive !== false).map((d, i) => ({
        ...d, color: distinctColors[i % distinctColors.length],
        capDV: (parseFloat(d['ДВ']) || 0) * hoursMultiplier, capP: (parseFloat(d['П']) || 0) * hoursMultiplier, capV: (parseFloat(d['В']) || 0) * hoursMultiplier,
        curDV: 0, curP: 0, curV: 0
    }));

    workingDepo.sort((a, b) => b['Широта'] - a['Широта']);

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
            pt.assigned = depo['Вузол']; pt.color = depo.color;
            depo.curDV += pt.dv; depo.curP += pt.p; depo.curV += pt.v;
        }
    }

    let remainders = filteredPoints.filter(p => !p.assigned && !p.isExcluded);
    remainders.forEach(pt => {
        const ptGeo = turf.point([pt.lng, pt.lat]);
        let bestDepo = null, bestScore = Infinity;
        workingDepo.forEach(depo => {
            const dist = turf.distance(ptGeo, turf.point([parseFloat(depo['Довгота']), parseFloat(depo['Широта'])]));
            let maxOverload = 0;
            if (useDV && depo.capDV > 0) maxOverload = Math.max(maxOverload, depo.curDV / depo.capDV);
            if (useP && depo.capP > 0) maxOverload = Math.max(maxOverload, depo.curP / depo.capP);
            if (useV && depo.capV > 0) maxOverload = Math.max(maxOverload, depo.curV / depo.capV);
            const penalty = Math.pow(Math.max(0, maxOverload - 1), 2) * 50; 
            const score = dist * (1 + penalty);
            if (score < bestScore) { bestScore = score; bestDepo = depo; }
        });
        if (bestDepo) {
            pt.assigned = bestDepo['Вузол']; pt.color = bestDepo.color;
            bestDepo.curDV += pt.dv; bestDepo.curP += pt.p; bestDepo.curV += pt.v;
        }
    });
    drawMap(workingDepo);
}

window.exportToExcel = function() {
    if (!filteredPoints || filteredPoints.length === 0) return alert("Немає даних для вивантаження! Спочатку завантажте файл.");
    const exportData = filteredPoints.map(p => ({
        "Відділення (Вузол)": p.id,
        "Широта": p.lat, "Довгота": p.lng,
        "ДВ (сума)": p.dv, "П (сума)": p.p, "В (сума)": p.v,
        "Призначене Депо": p.assigned || "Не розподілено",
        "Колір": p.color || "",
        "Є Зірочкою": p.isStar ? "Так" : "Ні",
        "Затверджена Зірочка": p.isApprovedStar ? "Так" : "Ні",
        "Прив'язана Зірочка": p.assignedStar || "",
        "Колір Зірки": p.starColor || ""
    }));

    const rawDataExport = rawPointsData.map(r => ({
        "Відділення (Вузол)": r.id, "Широта": r.lat, "Довгота": r.lng,
        "Година": r.hour, "ДВ": r.dv, "П": r.p, "В": r.v
    }));

    const wsResult = XLSX.utils.json_to_sheet(exportData);
    const wsRaw = XLSX.utils.json_to_sheet(rawDataExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsResult, "Розподіл_Депо");
    XLSX.utils.book_append_sheet(wb, wsRaw, "Сирі_Дані");
    XLSX.writeFile(wb, "Депошкер_Результат.xlsx");
};

// --- РОЗУМНЕ ПЕРЕМІЩЕННЯ МІЖ ДЕПО ТА ВІДВ'ЯЗУВАННЯ ---
window.movePoint = function(pointId) {
    const input = document.getElementById('move-input-' + pointId);
    if (!input) return;
    const newDepoId = input.value;
    const pt = filteredPoints.find(p => p.id === pointId);
    if (!pt) return;
    const currentAssigned = pt.assigned || "";
    if (currentAssigned === newDepoId) return; 

    const oldDepo = workingDepo.find(d => d['Вузол'] === pt.assigned);
    const newDepo = workingDepo.find(d => d['Вузол'] === newDepoId);

    if (newDepoId !== "" && !newDepo) return;

    if (pt.isStar) {
        if (oldDepo) { oldDepo.curDV -= pt.dv; oldDepo.curP -= pt.p; oldDepo.curV -= pt.v; }
        if (newDepo) { newDepo.curDV += pt.dv; newDepo.curP += pt.p; newDepo.curV += pt.v; }
        
        pt.assigned = newDepoId === "" ? null : newDepoId;
        pt.color = newDepo ? newDepo.color : (pt.starColor || '#eab308'); 

        filteredPoints.forEach(child => {
            if (child.assignedStar === pt.id) {
                const childOldDepo = workingDepo.find(d => d['Вузол'] === child.assigned);
                if (childOldDepo) { childOldDepo.curDV -= child.dv; childOldDepo.curP -= child.p; childOldDepo.curV -= child.v; }
                if (newDepo) { newDepo.curDV += child.dv; newDepo.curP += child.p; newDepo.curV += child.v; }
                child.assigned = pt.assigned;
                child.color = pt.color;
            }
        });
    } else {
        if (pt.assignedStar) {
            const oldStar = filteredPoints.find(s => s.id === pt.assignedStar);
            if (oldStar) oldStar.starChildren--;
            pt.assignedStar = null; pt.starColor = null;
        }
        if (oldDepo) { oldDepo.curDV -= pt.dv; oldDepo.curP -= pt.p; oldDepo.curV -= pt.v; }
        if (newDepo) { newDepo.curDV += pt.dv; newDepo.curP += pt.p; newDepo.curV += pt.v; }

        pt.assigned = newDepoId === "" ? null : newDepoId;
        pt.color = newDepo ? newDepo.color : '#64748b'; 
    }
    
    map.closePopup();
    drawMap(workingDepo.length > 0 ? workingDepo : null); 
};

window.toggleExcludePoint = function(pointId) {
    const pt = filteredPoints.find(p => p.id === pointId);
    if (!pt) return;
    pt.isExcluded = !pt.isExcluded;
    if (pt.isExcluded) {
        if (pt.assigned) {
            const depo = workingDepo.find(d => d['Вузол'] === pt.assigned);
            if (depo) { depo.curDV -= pt.dv; depo.curP -= pt.p; depo.curV -= pt.v; }
            pt.assigned = null;
        }
        if (pt.isStar) {
            pt.isStar = false; pt.isApprovedStar = false; pt.starChildren = 0; pt.starColor = null;
            filteredPoints.forEach(child => { if (child.assignedStar === pt.id) { child.assignedStar = null; child.starColor = null; } });
        }
        if (pt.assignedStar) {
            const oldStar = filteredPoints.find(s => s.id === pt.assignedStar);
            if (oldStar) oldStar.starChildren--;
            pt.assignedStar = null; pt.starColor = null;
        }
    }
    map.closePopup(); drawMap(workingDepo);
};

window.movePointStar = function(pointId) {
    const sel = document.getElementById('move-star-sel-' + pointId);
    if (!sel) return;
    const newStarId = sel.value;
    const pt = filteredPoints.find(p => p.id === pointId);
    if (!pt || pt.isStar) return; 

    const oldStarId = pt.assignedStar;
    if (oldStarId !== newStarId) {
        if (oldStarId && oldStarId !== 'none') {
            const oldStar = filteredPoints.find(s => s.id === oldStarId);
            if (oldStar) oldStar.starChildren--;
        }
        if (newStarId !== 'none') {
            const newStar = filteredPoints.find(s => s.id === newStarId);
            if (newStar) { newStar.starChildren++; pt.starColor = newStar.starColor || newStar.color; }
            pt.assignedStar = newStarId;
        } else { pt.assignedStar = null; pt.starColor = null; }
        map.closePopup(); drawMap(workingDepo); 
    }
};

window.toggleStarApproval = function(pointId, isApproved) {
    const pt = filteredPoints.find(p => p.id === pointId);
    if (pt) {
        pt.isApprovedStar = isApproved;
    }
};

// --- РОЗУМНЕ СТВОРЕННЯ ЗІРОЧКИ (РЕЖИМИ В ДЕПО ТА БЕЗ) ---
window.makeManualStar = function(pointId, mode = 'depot') {
    const pt = filteredPoints.find(p => p.id === pointId);
    if (!pt || pt.isStar) return;

    if (pt.assignedStar) {
        const oldStar = filteredPoints.find(s => s.id === pt.assignedStar);
        if (oldStar) oldStar.starChildren--;
        pt.assignedStar = null;
    }

    pt.isStar = true;
    pt.isApprovedStar = true; // Створені руками автоматично затверджуються
    pt.starChildren = 0;
    pt.starColor = pt.color || '#eab308';
    
    const maxPointsPerStar = parseInt(document.getElementById('starMaxPoints').value) || 15;
    const starGeo = turf.point([pt.lng, pt.lat]);

    if (mode === 'depot') {
        const depotId = pt.assigned;
        if (!depotId) {
            alert("Точка не прив'язана до Депо! Застосовано пошук без прив'язки."); mode = 'global'; 
        } else {
            const depotInfo = depoData.find(d => d['Вузол'] === depotId);
            if (depotInfo) {
                const depotPt = turf.point([parseFloat(depotInfo['Довгота']), parseFloat(depotInfo['Широта'])]);
                const distDepotToStar = turf.distance(depotPt, starGeo);
                const bearingDepotToStar = turf.bearing(depotPt, starGeo);

                let regularPoints = filteredPoints.filter(p => p.assigned === depotId && !p.isStar && !p.assignedStar);
                regularPoints.forEach(child => {
                    if (pt.starChildren >= maxPointsPerStar) return; 
                    const childGeo = turf.point([child.lng, child.lat]);
                    const distDepotToChild = turf.distance(depotPt, childGeo);
                    if (distDepotToChild > distDepotToStar) {
                        const bearingDepotToChild = turf.bearing(depotPt, childGeo);
                        let diff = Math.abs(bearingDepotToChild - bearingDepotToStar);
                        if (diff > 180) diff = 360 - diff;
                        if (diff <= 35) {
                            child.assignedStar = pt.id; child.starColor = pt.starColor; pt.starChildren++;
                        }
                    }
                });
            }
        }
    }
    
    if (mode === 'global') {
        if (pt.assigned) {
            const oldDepo = workingDepo.find(d => d['Вузол'] === pt.assigned);
            if (oldDepo) { oldDepo.curDV -= pt.dv; oldDepo.curP -= pt.p; oldDepo.curV -= pt.v; }
            pt.assigned = null;
            if (pt.starColor === pt.color) { pt.starColor = '#eab308'; pt.color = '#eab308'; }
        }

        let candidates = filteredPoints.filter(p => !p.isStar && !p.assignedStar && p.id !== pt.id);
        let withDist = candidates.map(child => ({
            child, dist: turf.distance(starGeo, turf.point([child.lng, child.lat]))
        })).sort((a,b) => a.dist - b.dist);
        
        let taken = 0;
        for (let item of withDist) {
            if (taken >= maxPointsPerStar) break;
            if (item.child.assigned) {
                const childOldDepo = workingDepo.find(d => d['Вузол'] === item.child.assigned);
                if (childOldDepo) { childOldDepo.curDV -= item.child.dv; childOldDepo.curP -= item.child.p; childOldDepo.curV -= item.child.v; }
                item.child.assigned = null;
            }
            item.child.assignedStar = pt.id; item.child.starColor = pt.starColor; item.child.color = pt.starColor; 
            pt.starChildren++; taken++;
        }
    }

    map.closePopup(); drawMap(workingDepo);
};

window.unmakeManualStar = function(pointId) {
    const pt = filteredPoints.find(p => p.id === pointId);
    if (!pt || !pt.isStar) return;
    pt.isStar = false; pt.isApprovedStar = false; pt.starChildren = 0; pt.starColor = null;
    filteredPoints.forEach(child => {
        if (child.assignedStar === pt.id) { child.assignedStar = null; child.starColor = null; }
    });
    map.closePopup(); drawMap(workingDepo); 
};

// --- ГЛОБАЛЬНИЙ ПОШУК ЗІРОЧОК (БЕЗ ДЕПО) ---
window.globalDistributeStars = function() {
    const minArea = parseFloat(document.getElementById('starMinArea').value) || 150;
    const maxStars = parseInt(document.getElementById('starMaxCount').value) || 15;
    const maxPointsPerStar = parseInt(document.getElementById('starMaxPoints').value) || 15;
    const keepExisting = document.getElementById('keepExistingStars') ? document.getElementById('keepExistingStars').checked : false;

    if (keepExisting) {
        // Розформовуємо НЕЗАТВЕРДЖЕНІ зірки та їхніх дітей
        filteredPoints.forEach(p => {
            if (p.isStar && !p.isApprovedStar && !p.isExcluded) {
                p.isStar = false; p.starChildren = 0; p.starColor = null;
                // Повертаємо базовий колір точці
                p.color = p.assigned ? (workingDepo.find(x => x['Вузол'] === p.assigned)?.color || '#64748b') : '#64748b';
                
                filteredPoints.forEach(child => {
                    if (child.assignedStar === p.id) {
                        child.assignedStar = null; child.starColor = null;
                    }
                });
            }
        });
    } else {
        // Повністю скидаємо всі зірки (навіть затверджені)
        filteredPoints.forEach(p => { 
            if (!p.isExcluded && p.isStar) {
                p.isStar = false; p.isApprovedStar = false; p.starChildren = 0; p.starColor = null; 
                p.color = p.assigned ? (workingDepo.find(x => x['Вузол'] === p.assigned)?.color || '#64748b') : '#64748b';
                
                filteredPoints.forEach(child => {
                    if (child.assignedStar === p.id) {
                        child.assignedStar = null; child.starColor = null;
                    }
                });
            }
        });
    }

    // Беремо в роботу всі точки, які не виключені і ще не зайняті зірками
    let targetPoints = filteredPoints.filter(p => !p.isExcluded && !p.isStar && !p.assignedStar);
    
    if(targetPoints.length === 0) return alert("Немає доступних відділень для створення нових зірок.");

    let potentialStars = [];
    targetPoints.forEach(p => {
        const areaInfo = areaData.find(a => String(a['Вузол']) === String(p.id));
        const area = areaInfo ? parseFloat(areaInfo['Площа']) : 0; 
        if (area >= minArea) potentialStars.push({ pt: p, area: area });
    });

    potentialStars.sort((a, b) => b.area - a.area);
    let actualStars = potentialStars.slice(0, maxStars).map(s => s.pt);
    
    actualStars.forEach((s, i) => { 
        s.isStar = true; 
        s.isApprovedStar = false; // Автоматичні зірки не затверджуються самі по собі
        s.starChildren = 0; 
        s.starColor = distinctColors[i % distinctColors.length];
        s.color = s.starColor;
    });

    let regularPoints = targetPoints.filter(p => !p.isStar);
    
    // Враховуємо ВСІ актуальні зірки (навіть затверджені старі) для прилипання залишків
    let allAvailableStars = filteredPoints.filter(p => p.isStar && !p.isExcluded);

    regularPoints.forEach(pt => {
        const ptGeo = turf.point([pt.lng, pt.lat]);
        let bestStar = null; let minStarDist = Infinity;
        allAvailableStars.forEach(star => {
            if (star.starChildren >= maxPointsPerStar) return;
            const starGeo = turf.point([star.lng, star.lat]);
            const dist = turf.distance(starGeo, ptGeo);
            if (dist < minStarDist) { minStarDist = dist; bestStar = star; }
        });
        if (bestStar) {
            pt.assignedStar = bestStar.id; pt.starColor = bestStar.starColor; bestStar.starChildren++;
        }
    });

    map.closePopup(); drawMap(workingDepo.length > 0 ? workingDepo : null);
};

// --- ЛОКАЛЬНИЙ ПОШУК ЗІРОЧОК (В МЕЖАХ ДЕПО) ---
window.distributeStars = function(depotId) {
    const minArea = parseFloat(document.getElementById('starMinArea').value) || 150;
    const maxStars = parseInt(document.getElementById('starMaxCount').value) || 15;
    const maxPointsPerStar = parseInt(document.getElementById('starMaxPoints').value) || 15;
    const keepExisting = document.getElementById('keepExistingStars') ? document.getElementById('keepExistingStars').checked : false;

    const depotInfo = depoData.find(d => d['Вузол'] === depotId);
    if (!depotInfo) return;
    const depotPt = turf.point([parseFloat(depotInfo['Довгота']), parseFloat(depotInfo['Широта'])]);

    let depotPoints = filteredPoints.filter(p => p.assigned === depotId && !p.isExcluded);
    
    if (keepExisting) {
        depotPoints.forEach(p => {
            if (p.isStar && !p.isApprovedStar) {
                p.isStar = false; p.starChildren = 0; p.starColor = null;
                depotPoints.forEach(child => {
                    if (child.assignedStar === p.id) { child.assignedStar = null; child.starColor = null; }
                });
            }
        });
    } else {
        depotPoints.forEach(p => { 
            if (p.isStar) {
                p.isStar = false; p.isApprovedStar = false; p.starChildren = 0; p.starColor = null; 
                depotPoints.forEach(child => {
                    if (child.assignedStar === p.id) { child.assignedStar = null; child.starColor = null; }
                });
            }
        });
    }

    let targetPoints = depotPoints.filter(p => !p.isStar && !p.assignedStar);
    if(targetPoints.length === 0) return alert("Всі доступні відділення цього депо вже зайняті зірочками.");

    let potentialStars = [];
    targetPoints.forEach(p => {
        const areaInfo = areaData.find(a => String(a['Вузол']) === String(p.id));
        const area = areaInfo ? parseFloat(areaInfo['Площа']) : 0; 
        if (area >= minArea) potentialStars.push({ pt: p, area: area, dist: turf.distance(depotPt, turf.point([p.lng, p.lat])) });
    });

    potentialStars.sort((a, b) => b.area - a.area);
    let actualStars = potentialStars.slice(0, maxStars).map(s => s.pt);
    
    actualStars.forEach((s, idx) => { 
        s.isStar = true; s.isApprovedStar = false; s.starChildren = 0; s.starColor = distinctColors[idx % distinctColors.length]; 
    });

    let regularPoints = targetPoints.filter(p => !p.isStar);
    let allAvailableStars = filteredPoints.filter(p => p.isStar && p.assigned === depotId && !p.isExcluded);

    regularPoints.forEach(pt => {
        const ptGeo = turf.point([pt.lng, pt.lat]);
        const distDepotToPt = turf.distance(depotPt, ptGeo);
        const bearingDepotToPt = turf.bearing(depotPt, ptGeo); 

        let bestStar = null; let minStarDist = Infinity;
        allAvailableStars.forEach(star => {
            if (star.starChildren >= maxPointsPerStar) return; 
            const starGeo = turf.point([star.lng, star.lat]);
            const distDepotToStar = turf.distance(depotPt, starGeo);
            
            if (distDepotToPt > distDepotToStar) {
                const bearingDepotToStar = turf.bearing(depotPt, starGeo);
                let diff = Math.abs(bearingDepotToPt - bearingDepotToStar);
                if (diff > 180) diff = 360 - diff; 
                if (diff <= 35) { 
                    const distStarToPt = turf.distance(starGeo, ptGeo);
                    if (distStarToPt < minStarDist) { minStarDist = distStarToPt; bestStar = star; }
                }
            }
        });

        if (bestStar) {
            pt.assignedStar = bestStar.id; pt.starColor = bestStar.starColor; bestStar.starChildren++;
        }
    });

    map.closePopup(); drawMap(workingDepo);
};

window.changeDepoColor = function(depoId, newColor) {
    if (!workingDepo || workingDepo.length === 0) return;
    const depo = workingDepo.find(d => d['Вузол'] === depoId);
    if (depo) {
        depo.color = newColor; 
        filteredPoints.forEach(p => { if (p.assigned === depoId) p.color = newColor; });
        drawMap(workingDepo);
    }
};

window.changeStarColor = function(starId, newColor) {
    const star = filteredPoints.find(p => p.id === starId);
    if (star) {
        star.starColor = newColor; if(!star.color) star.color = newColor; 
        filteredPoints.forEach(p => { if (p.assignedStar === starId) p.starColor = newColor; });
        drawMap(workingDepo.length > 0 ? workingDepo : null);
    }
};

function drawMap(processedDepo = null) {
    layers.depo.clearLayers();
    layers.points.clearLayers();
    layers.polygons.clearLayers();
    layers.lines.clearLayers(); 
    
    // ОЧИЩАЄМО ОРЕОЛИ ПРИ КОЖНІЙ ПЕРЕМАЛЬОВЦІ КАРТИ
    if (layers.potentialStarHalos) {
        layers.potentialStarHalos.clearLayers();
    }
    
    oms.clearMarkers(); 

    const showDepotZones = document.getElementById('showDepotZones') ? document.getElementById('showDepotZones').checked : true;
    const showStarZones = document.getElementById('showStarZones') ? document.getElementById('showStarZones').checked : true;

    if (processedDepo && showDepotZones) {
        processedDepo.forEach(depo => {
            if (window.activeDepotFilters.size > 0 && !window.activeDepotFilters.has(depo['Вузол'])) return;
            const depoPoints = filteredPoints.filter(p => p.assigned === depo['Вузол'] && !p.isStar && !p.assignedStar);
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

    if (showStarZones) {
        const allStars = filteredPoints.filter(p => p.isStar);
        allStars.forEach(star => {
            if (window.activeDepotFilters.size > 0 && star.assigned && !window.activeDepotFilters.has(star.assigned)) return;

            const starChildren = filteredPoints.filter(p => p.assignedStar === star.id);
            if (starChildren.length >= 2) { 
                const pts = starChildren.map(p => turf.point([p.lng, p.lat]));
                pts.push(turf.point([star.lng, star.lat]));
                let hull = turf.convex(turf.featureCollection(pts));
                if (hull) {
                    hull = turf.buffer(hull, 0.2, { units: 'kilometers' });
                    hull = turf.polygonSmooth(hull, { iterations: 2 });
                    const polyColor = star.starColor || star.color || '#eab308';
                    L.geoJSON(hull, { style: { color: polyColor, weight: 3, opacity: 1, fillOpacity: 0.2, dashArray: '5,5' } }).addTo(layers.polygons);
                }
            }
        });
    }

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

        const dPts = filteredPoints.filter(p => p.assigned === d['Вузол']);
        const starsCount = dPts.filter(p => p.isStar).length;
        const orphansCount = dPts.filter(p => !p.isStar && !p.assignedStar).length;
        const depotRamps = d['Рампи'] || '0'; 

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

                <button onclick="distributeStars('${d['Вузол']}')" class="popup-btn btn-blue" style="color:#000; background:#eab308; border-color:#eab308;" ${isOff || !isDist ? 'disabled' : ''}>🌟 Шукати Звіздочки (Локально)</button>
                <div class="popup-row">
                    <button onclick="toggleDepo('${d['Вузол']}')" class="popup-btn btn-blue" style="flex:1; background:${isOff ? '#4ade80' : '#ef4444'}; color:${isOff ? '#000' : '#fff'};">${isOff ? 'ВКЛ' : 'ВИКЛ'}</button>
                    <button onclick="openDepotChart('${d['Вузол']}')" class="popup-btn btn-blue" style="flex:1;" ${isOff ? 'disabled' : ''}>📊 Графік</button>
                </div>
            </div>
        `;
        
        marker.customPopupHtml = popupHtml;
        marker.on('mouseover', function() { window.isHoverClick = true; marker.fire('click'); window.isHoverClick = false; });
        marker.depotData = d;
        layers.depo.addLayer(marker); oms.addMarker(marker); 
    });

    filteredPoints.forEach(p => {
        if (window.activeDepotFilters.size > 0 && p.assigned && !window.activeDepotFilters.has(p.assigned)) return;
        const isZero = (p.dv + p.p + p.v) === 0;
        const baseColor = p.starColor || p.color || '#64748b'; 
        const col = isZero ? '#334155' : baseColor; 
        
        let iconHtml, iconSize;
        if (p.isExcluded) {
            iconHtml = `<div style="background:#000; width:16px; height:16px; border: 2px solid #ef4444; border-radius: 50%; display:flex; align-items:center; justify-content:center; color:#ef4444; font-size:12px; font-weight:bold; box-shadow: 0 0 8px #ef4444;">X</div>`;
            iconSize = [20, 20];
        } else if (p.isStar) {
            iconHtml = `<div style="background:${col}; width:20px; height:20px; border: 2px solid #fff; border-radius: 50%; box-shadow: 0 0 10px ${col}; display:flex; align-items:center; justify-content:center; font-size:12px; text-shadow: 1px 1px 2px #000;">⭐</div>`;
            iconSize = [24, 24];
        } else {
            const starBorder = p.assignedStar ? baseColor : (isZero ? '#1e293b' : '#ffffff'); 
            iconHtml = `<div style="background:${col}; width:14px; height:14px; border: 2px solid ${starBorder}; border-radius: 50%;"></div>`;
            iconSize = [18, 18];
        }

        const icon = L.divIcon({ className: '', html: iconHtml, iconSize: iconSize });
        
        if (p.assignedStar) {
            const star = filteredPoints.find(s => s.id === p.assignedStar);
            if (star) {
                const lineColor = star.starColor || star.color || p.color;
                L.polyline([[star.lat, star.lng], [p.lat, p.lng]], { color: lineColor, weight: 2, dashArray: '5, 5', opacity: 0.8 }).addTo(layers.lines);
            }
        }
        if (p.isStar && p.assigned) {
            const dPt = depoData.find(d => d['Вузол'] === p.assigned);
            if (dPt) L.polyline([[dPt['Широта'], dPt['Довгота']], [p.lat, p.lng]], { color: col, weight: 3, opacity: 0.9 }).addTo(layers.lines);
        }

        const ptGeo = turf.point([p.lng, p.lat]);
        let depoSelectHtml = '<option value="">Нерозподілено</option>';
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
            filteredPoints.filter(s => s.isStar && (s.assigned === p.assigned || !p.assigned)).forEach(s => {
                starOptionsHtml += `<option value="${s.id}" ${s.id === p.assignedStar ? 'selected' : ''}>${s.id}</option>`;
            });
        }

        const areaInfo = areaData.find(a => String(a['Вузол']) === String(p.id)) || {};
        const ptArea = areaInfo['Площа'] || 0;
        const ptRamps = areaInfo['Рампи'] || '0';

        let starStatsHtml = '';
        if (p.isStar) {
            const currentStarColor = p.starColor || p.color || '#ffffff';
            starStatsHtml = `
                <div class="stats-grid">
                    <div class="stats-item"><span>Підлеглих</span><b>${p.starChildren}</b></div>
                    <div class="stats-item"><span>Площа Звіздочки</span><b>${ptArea}</b></div>
                </div>
                <div class="popup-row" style="margin-top:5px; align-items:center;">
                    <label style="color: #fff; cursor: pointer; display: flex; align-items: center; gap: 8px; font-size: 0.8rem; width: 100%;">
                        <input type="checkbox" ${p.isApprovedStar ? 'checked' : ''} onchange="toggleStarApproval('${p.id}', this.checked)"> 
                        ✅ Затверджена Зірочка
                    </label>
                </div>
                <div class="popup-row" style="font-size:0.75rem; margin-top:5px; align-items:center;">Колір Звіздочки: 
                    <input type="color" value="${currentStarColor}" onchange="changeStarColor('${p.id}', this.value)" style="border:none; background:transparent; cursor:pointer; width:30px; height:30px; padding:0; margin-left:10px;">
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
                <div class="popup-row" style="gap:5px; margin-bottom:10px;">
                    <button onclick="makeManualStar('${p.id}', 'depot')" class="popup-btn btn-yellow" style="flex:1; padding:6px 2px; font-size:0.75rem;" title="З вектором від Депо">⭐ В Депо</button>
                    <button onclick="makeManualStar('${p.id}', 'global')" class="popup-btn btn-yellow" style="flex:1; padding:6px 2px; font-size:0.75rem;" title="По радіусу (відв'яже від Депо)">🌟 Без Депо</button>
                </div>
                <div style="font-size:0.75rem; color:var(--text-muted); margin-bottom:4px;">Прив'язка до звіздочки:</div>
                <div class="popup-row">
                    <select id="move-star-sel-${p.id}" class="smart-input">${starOptionsHtml}</select>
                    <button onclick="movePointStar('${p.id}')" class="popup-btn btn-yellow">ОК</button>
                </div>` : `
                <button onclick="unmakeManualStar('${p.id}')" class="popup-btn btn-red" style="width:100%; margin-bottom:10px;">❌ Роззвіздувати</button>
                `}

                <div style="font-size:0.75rem; color:var(--text-muted); margin-bottom:4px;">Депо (< 200км):</div>
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
        m.on('mouseover', function() { window.isHoverClick = true; m.fire('click'); window.isHoverClick = false; });
        m.pointData = p;

        layers.points.addLayer(m); oms.addMarker(m); 
    });

    // ОНОВЛЮЄМО ПІДСВІТКУ В КІНЦІ ВІДМАЛЬОВКИ (ЯКЩО РЕЖИМ УВІМКНЕНО)
    if (window.isStarHighlightMode) {
        window.updateStarHighlights();
    }
}

let depotChartInstance = null;
let currentChartData = null; 
let currentChartTab = 'all';

window.openDepotChart = function(depotId) {
    if (!filteredPoints || filteredPoints.length === 0) return alert("Спочатку запустіть розподіл!");
    document.getElementById('chartDepotTitle').innerText = `Навантаження: ${depotId}`;
    const assignedBranchIds = filteredPoints.filter(p => p.assigned === depotId).map(p => String(p.id).trim());
    const depoInfo = depoData.find(d => String(d['Вузол']).trim() === String(depotId).trim());
    if (!depoInfo) return alert("Дані депо не знайдені. Спочатку оновіть Депо (Крок 1).");

    const caps = { dv: parseFloat(depoInfo['ДВ']) || 0, p: parseFloat(depoInfo['П']) || 0, v: parseFloat(depoInfo['В']) || 0 };
    const loads = { dv: new Array(24).fill(0), p: new Array(24).fill(0), v: new Array(24).fill(0) };
    
    rawPointsData.forEach(r => {
        const rId = String(r.id).trim(); const hr = parseInt(r.hour) || 0;
        if (assignedBranchIds.includes(rId) && hr >= 0 && hr <= 23) {
            loads.dv[hr] += parseFloat(r.dv) || 0; loads.p[hr] += parseFloat(r.p) || 0; loads.v[hr] += parseFloat(r.v) || 0;
        }
    });
    currentChartData = { loads, caps }; switchChartTab('all'); document.getElementById('chartModal').style.display = 'flex';
};

window.openStarChart = function(starId) {
    document.getElementById('chartDepotTitle').innerText = `Навантаження Звіздочки: ${starId}`;
    const childrenIds = filteredPoints.filter(p => p.assignedStar === starId || p.id === starId).map(p => String(p.id).trim());
    const loads = { dv: new Array(24).fill(0), p: new Array(24).fill(0), v: new Array(24).fill(0) };
    
    rawPointsData.forEach(r => {
        const rId = String(r.id).trim(); const hr = parseInt(r.hour) || 0;
        if (childrenIds.includes(rId) && hr >= 0 && hr <= 23) {
            loads.dv[hr] += parseFloat(r.dv) || 0; loads.p[hr] += parseFloat(r.p) || 0; loads.v[hr] += parseFloat(r.v) || 0;
        }
    });
    currentChartData = { loads, caps: {dv: 0, p: 0, v: 0} }; switchChartTab('all'); document.getElementById('chartModal').style.display = 'flex';
};

window.switchChartTab = function(tabId) {
    currentChartTab = tabId;
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.getAttribute('data-tab') === tabId) btn.classList.add('active');
    });
    renderChart();
};

window.closeChartModal = function() { document.getElementById('chartModal').style.display = 'none'; };

function renderChart() {
    const ctx = document.getElementById('depotChart').getContext('2d');
    if (depotChartInstance) depotChartInstance.destroy();

    const labels = Array.from({length: 24}, (_, i) => `${i}:00`);
    const { loads, caps } = currentChartData;
    let datasets = [];

    const pushDV = () => {
        datasets.push({ label: 'Фактично ДВ', data: loads.dv, borderColor: '#38bdf8', borderWidth: 3, tension: 0.3, fill: false });
        if (caps.dv > 0) datasets.push({ label: 'Ліміт ДВ', data: new Array(24).fill(caps.dv), borderColor: '#38bdf8', borderWidth: 2, borderDash: [5, 5], pointRadius: 0, fill: false });
    };
    const pushP = () => {
        datasets.push({ label: 'Фактично П', data: loads.p, borderColor: '#4ade80', borderWidth: 3, tension: 0.3, fill: false });
        if (caps.p > 0) datasets.push({ label: 'Ліміт П', data: new Array(24).fill(caps.p), borderColor: '#4ade80', borderWidth: 2, borderDash: [5, 5], pointRadius: 0, fill: false });
    };
    const pushV = () => {
        datasets.push({ label: 'Фактично В', data: loads.v, borderColor: '#facc15', borderWidth: 3, tension: 0.3, fill: false });
        if (caps.v > 0) datasets.push({ label: 'Ліміт В', data: new Array(24).fill(caps.v), borderColor: '#facc15', borderWidth: 2, borderDash: [5, 5], pointRadius: 0, fill: false });
    };

    if (currentChartTab === 'all' || currentChartTab === 'dv') pushDV();
    if (currentChartTab === 'all' || currentChartTab === 'p') pushP();
    if (currentChartTab === 'all' || currentChartTab === 'v') pushV();

    depotChartInstance = new Chart(ctx, {
        type: 'line', data: { labels: labels, datasets: datasets },
        options: {
            responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false },
            plugins: { legend: { labels: { color: '#f1f5f9', usePointStyle: true, boxWidth: 8 } } },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' }, beginAtZero: true }
            }
        }
    });
}

window.toggleDepo = function(depoId) {
    const d = depoData.find(x => x['Вузол'] === depoId);
    if (d) {
        d.isActive = (d.isActive === false) ? true : false;
        localStorage.setItem('deposhker_depo_cache', JSON.stringify(depoData));
        if (filteredPoints && filteredPoints.length > 0) {
            filteredPoints.forEach(p => { p.assigned = null; p.color = '#64748b'; }); runDistribution(); 
        } else { drawMap(); }
    }
};

window.populateMassMoveDatalist = function() {
    const dl = document.getElementById('mass-depo-list'); let html = '';
    workingDepo.forEach(d => { html += `<option value="${d['Вузол']}"></option>`; }); dl.innerHTML = html;
};

window.closeMassMoveModal = function() { document.getElementById('massMoveModal').style.display = 'none'; if(drawnItems) drawnItems.clearLayers(); };

window.applyMassMove = function() {
    const newDepoId = document.getElementById('massMoveDepoInput').value;
    const newDepo = workingDepo.find(d => d['Вузол'] === newDepoId);
    if (!newDepo) return alert("Оберіть існуюче депо зі списку!");

    window.selectedPointsForMassMove.forEach(pt => {
        if (pt.assigned === newDepoId) return;
        const oldDepo = workingDepo.find(d => d['Вузол'] === pt.assigned);
        if (pt.assignedStar) {
            const oldStar = filteredPoints.find(s => s.id === pt.assignedStar);
            if (oldStar) oldStar.starChildren--;
            pt.assignedStar = null;
        }
        if (pt.isStar) {
            pt.isStar = false; pt.isApprovedStar = false; pt.starChildren = 0;
            filteredPoints.forEach(child => { if (child.assignedStar === pt.id) child.assignedStar = null; });
        }
        if (oldDepo) { oldDepo.curDV -= pt.dv; oldDepo.curP -= pt.p; oldDepo.curV -= pt.v; }
        newDepo.curDV += pt.dv; newDepo.curP += pt.p; newDepo.curV += pt.v;
        pt.assigned = newDepoId; pt.color = newDepo.color;
    });
    closeMassMoveModal(); drawMap(workingDepo);
};

window.toggleFilterMenu = function() {
    const menu = document.getElementById('mapFilterMenu');
    if (menu.style.display === 'none' || menu.style.display === '') { buildFilterMenu(); menu.style.display = 'flex'; } else { menu.style.display = 'none'; }
};

window.buildFilterMenu = function() {
    const menu = document.getElementById('mapFilterMenu');
    const activeDepots = [...new Set(filteredPoints.filter(p => p.assigned).map(p => p.assigned))].sort();
    if (activeDepots.length === 0) { menu.innerHTML = '<div style="color:var(--text-muted); font-size:0.8rem;">Немає призначених депо</div>'; return; }

    const isAllChecked = window.activeDepotFilters.size === 0 || window.activeDepotFilters.size === activeDepots.length;
    let html = `<label style="font-weight: bold; border-bottom: 1px solid var(--border); padding-bottom: 8px; margin-bottom: 8px;"><input type="checkbox" id="selectAllDepots" onchange="toggleAllDepots(this)" ${isAllChecked ? 'checked' : ''}> Обрати всі</label>`;
    activeDepots.forEach(depoId => {
        const isChecked = window.activeDepotFilters.size === 0 || window.activeDepotFilters.has(depoId);
        html += `<label><input type="checkbox" class="depot-filter-cb" value="${depoId}" onchange="updateDepotFilter()" ${isChecked ? 'checked' : ''}> ${depoId}</label>`;
    });
    menu.innerHTML = html;
};

window.toggleAllDepots = function(selectAllCb) {
    const checkboxes = document.querySelectorAll('.depot-filter-cb'); checkboxes.forEach(cb => { cb.checked = selectAllCb.checked; }); updateDepotFilter();
};

window.updateDepotFilter = function() {
    const checkboxes = document.querySelectorAll('.depot-filter-cb'); const selectAllCb = document.getElementById('selectAllDepots');
    window.activeDepotFilters.clear(); let selectedNames = []; let allChecked = true;

    checkboxes.forEach(cb => { if (cb.checked) { window.activeDepotFilters.add(cb.value); selectedNames.push(cb.value); } else { allChecked = false; } });
    if (selectAllCb) { selectAllCb.checked = allChecked; }

    const input = document.getElementById('mapFilterInput');
    if (allChecked) { window.activeDepotFilters.clear(); input.value = ''; } 
    else if (selectedNames.length === 0) { window.activeDepotFilters.add('__NONE__'); input.value = 'Нічого не обрано'; } 
    else { input.value = selectedNames.join(', '); }
    drawMap(workingDepo.length > 0 ? workingDepo : null);
};

document.addEventListener('click', function(e) {
    const container = document.getElementById('mapFilterContainer');
    if (container && !container.contains(e.target)) { document.getElementById('mapFilterMenu').style.display = 'none'; }
});

let highlightLayer = null;
window.handleSearchInput = function(query) {
    const resultsContainer = document.getElementById('pointSearchResults');
    query = query.trim().toLowerCase();
    if (query.length < 3) { resultsContainer.style.display = 'none'; return; }

    const matches = filteredPoints.filter(p => String(p.id).toLowerCase().includes(query));
    if (matches.length === 0) {
        resultsContainer.innerHTML = '<div class="search-result-item" style="color:var(--text-muted); cursor:default;">Нічого не знайдено</div>';
        resultsContainer.style.display = 'flex';
    } else if (matches.length === 1) {
        resultsContainer.style.display = 'none'; document.getElementById('pointSearchInput').value = matches[0].id; focusAndHighlightPoint(matches[0].id);
    } else {
        resultsContainer.innerHTML = '';
        matches.forEach(match => {
            const div = document.createElement('div'); div.className = 'search-result-item'; div.innerText = match.id;
            div.onclick = () => { document.getElementById('pointSearchInput').value = match.id; resultsContainer.style.display = 'none'; focusAndHighlightPoint(match.id); };
            resultsContainer.appendChild(div);
        });
        resultsContainer.style.display = 'flex';
    }
};

window.focusAndHighlightPoint = function(pointId) {
    const pt = filteredPoints.find(p => p.id === pointId); if (!pt) return;
    map.setView([pt.lat, pt.lng], 14);
    if (highlightLayer) map.removeLayer(highlightLayer);
    const haloIcon = L.divIcon({ className: 'highlight-halo', iconSize: [40, 40], iconAnchor: [20, 20] });
    highlightLayer = L.marker([pt.lat, pt.lng], { icon: haloIcon, interactive: false, zIndexOffset: -100 }).addTo(map);

    layers.points.eachLayer(layer => {
        const latlng = layer.getLatLng(); if (latlng.lat === pt.lat && latlng.lng === pt.lng) layer.fire('click');
    });
    setTimeout(() => { if (highlightLayer) { map.removeLayer(highlightLayer); highlightLayer = null; } }, 4000);
};

document.addEventListener('click', function(e) {
    const searchContainer = document.querySelector('.map-search-container'); const filterContainer = document.getElementById('mapFilterContainer');
    if (searchContainer && !searchContainer.contains(e.target)) document.getElementById('pointSearchResults').style.display = 'none';
    if (filterContainer && !filterContainer.contains(e.target)) document.getElementById('mapFilterMenu').style.display = 'none';
});

window.toggleAllElements = function(turnOn) {
    let changed = false;
    if (depoData && depoData.length > 0) { depoData.forEach(d => { d.isActive = turnOn; }); localStorage.setItem('deposhker_depo_cache', JSON.stringify(depoData)); changed = true; }
    if (filteredPoints && filteredPoints.length > 0) {
        filteredPoints.forEach(pt => {
            pt.isExcluded = !turnOn;
            if (!turnOn) { pt.assigned = null; pt.isStar = false; pt.isApprovedStar = false; pt.starChildren = 0; pt.assignedStar = null; pt.starColor = null; }
        }); changed = true;
    }
    if (changed) {
        if (!turnOn && workingDepo && workingDepo.length > 0) { workingDepo.forEach(d => { d.curDV = 0; d.curP = 0; d.curV = 0; }); }
        map.closePopup(); drawMap(workingDepo && workingDepo.length > 0 ? workingDepo : null);
    } else { alert("Немає завантажених даних для зміни статусу."); }
};

window.isRoutingMode = false; window.routeNodes = []; window.osrmLegs = []; window.savedRoutes = []; let routeSortableInst = null;
window.toggleRoutingMode = function() {
    window.isRoutingMode = !window.isRoutingMode;
    const btn = document.getElementById('btnRoutingToggle'); const panel = document.getElementById('routingRightPanel'); const logo = document.querySelector('.map-floating-logo');
    if (window.isRoutingMode) {
        btn.style.background = '#ef4444'; btn.style.color = '#fff'; btn.style.borderColor = '#ef4444'; btn.innerHTML = '🛑'; panel.style.display = 'flex';
        if (logo) logo.style.right = '365px'; map.closePopup(); updateRouteHalos();
        if (!routeSortableInst) { routeSortableInst = new Sortable(document.getElementById('routeList'), { animation: 150, onEnd: function() { rebuildRouteArrayFromDOM(); calculateOSRMRoute(); } }); }
    } else {
        btn.style.background = ''; btn.style.color = ''; btn.style.borderColor = ''; btn.innerHTML = '🚗'; panel.style.display = 'none';
        if (logo) logo.style.right = '25px'; clearRoute();
    }
};

window.handleRouteClick = function(data, type) {
    if (routeNodes.length === 0 && type === 'point') {
        let parentNode = null;
        if (data.assignedStar) {
            const star = filteredPoints.find(p => p.id === data.assignedStar);
            if (star) parentNode = { id: star.id, lat: star.lat, lng: star.lng, isParent: true, stopTime: 0 };
        } else if (data.assigned) {
            const depo = depoData.find(d => d['Вузол'] === data.assigned);
            if (depo) parentNode = { id: depo['Вузол'], lat: parseFloat(depo['Широта']), lng: parseFloat(depo['Довгота']), isParent: true, stopTime: 0 };
        }
        if (parentNode) routeNodes.push(parentNode);
    }
    const id = type === 'point' ? data.id : data['Вузол']; const lat = type === 'point' ? data.lat : parseFloat(data['Широта']); const lng = type === 'point' ? data.lng : parseFloat(data['Довгота']);
    const defaultStop = parseInt(document.getElementById('routeDefaultStop').value) || 10;
    if (!routeNodes.find(n => n.id === id)) { routeNodes.push({ id, lat, lng, isParent: false, stopTime: routeNodes.length === 0 ? 0 : defaultStop }); calculateOSRMRoute(); }
};

window.calculateOSRMRoute = async function() {
    layers.route.clearLayers(); if (routeNodes.length < 2) return renderRouteList();
    const isRoundTrip = document.getElementById('routeRoundTrip').checked;
    let reqNodes = [...routeNodes]; if (isRoundTrip) reqNodes.push(routeNodes[0]);
    const coordsStr = reqNodes.map(n => `${n.lng},${n.lat}`).join(';');
    const url = `https://router.project-osrm.org/route/v1/driving/${coordsStr}?overview=full&geometries=geojson`;

    try {
        const response = await fetch(url); const data = await response.json();
        if (data.code === 'Ok') {
            const route = data.routes[0]; const distKm = (route.distance / 1000).toFixed(1);
            document.getElementById('routeDist').innerText = `${distKm} км`;
            window.osrmLegs = route.legs.map(leg => Math.round(leg.duration / 60));
            L.geoJSON(route.geometry, { style: { color: '#facc15', weight: 5, opacity: 0.8 }, className: 'route-anim' }).addTo(layers.route);
            recalculateSchedule(); 
        }
    } catch (e) { console.error("OSRM Error: ", e); }
};

window.recalculateSchedule = function() {
    if (routeNodes.length < 2 || window.osrmLegs.length === 0) return renderRouteList();
    const startTimeStr = document.getElementById('routeStartTime').value; const isRoundTrip = document.getElementById('routeRoundTrip').checked;
    let currentTime = parseTimeToMinutes(startTimeStr); let totalTravelMin = 0;
    routeNodes[0].arrTime = null; routeNodes[0].depTime = currentTime;

    for (let i = 0; i < window.osrmLegs.length; i++) {
        let legMin = window.osrmLegs[i]; totalTravelMin += legMin; currentTime += legMin; 
        if (isRoundTrip && i === window.osrmLegs.length - 1) { routeNodes[0].returnArrTime = currentTime; } 
        else {
            let targetNode = routeNodes[i + 1]; targetNode.arrTime = currentTime;
            let stopM = parseInt(targetNode.stopTime) || 0; targetNode.depTime = currentTime + stopM; currentTime = targetNode.depTime; totalTravelMin += stopM;
        }
    }
    document.getElementById('routeTotalTime').innerText = formatMinutesToTimeSpan(totalTravelMin); renderRouteList();
};

window.updateNodeStopTime = function(id, newVal) { const node = routeNodes.find(n => n.id === id); if (node) { node.stopTime = Math.max(0, parseInt(newVal) || 0); recalculateSchedule(); } };
window.applyDefaultStopTime = function() { const def = Math.max(0, parseInt(document.getElementById('routeDefaultStop').value) || 0); routeNodes.forEach((n, idx) => { if (idx !== 0) n.stopTime = def; }); recalculateSchedule(); };

window.updateRouteHalos = function() {
    layers.routeHalos.clearLayers(); if (!window.isRoutingMode) return;
    const haloIcon = L.divIcon({ className: '', html: '<div class="route-halo"></div>', iconSize: [30, 30], iconAnchor: [15, 15] });
    const savedHaloIcon = L.divIcon({ className: '', html: '<div class="route-halo saved-halo"></div>', iconSize: [30, 30], iconAnchor: [15, 15] });
    let handledIds = new Set();
    routeNodes.forEach(node => { handledIds.add(node.id); L.marker([node.lat, node.lng], { icon: haloIcon, interactive: false, zIndexOffset: -40 }).addTo(layers.routeHalos); });
    window.savedRoutes.forEach(route => { route.nodes.forEach(node => { if (!handledIds.has(node.id)) { handledIds.add(node.id); L.marker([node.lat, node.lng], { icon: savedHaloIcon, interactive: false, zIndexOffset: -50 }).addTo(layers.routeHalos); } }); });
};

window.renderRouteList = function() {
    const tbody = document.getElementById('routeList'); tbody.innerHTML = '';
    const isRoundTrip = document.getElementById('routeRoundTrip').checked;
    routeNodes.forEach((node, idx) => {
        const tr = document.createElement('tr'); tr.className = node.isParent ? 'parent' : ''; tr.dataset.id = node.id;
        const arrStr = node.arrTime !== null ? formatTimeOfDay(node.arrTime) : '--:--'; const depStr = node.depTime !== null ? formatTimeOfDay(node.depTime) : '--:--';
        const isStart = idx === 0;
        tr.innerHTML = `
            <td style="color:var(--text-muted);">${idx + 1}</td>
            <td style="text-align: left;">${node.id} ${node.isParent ? '<span style="font-size: 0.6rem;">⭐</span>' : ''}</td>
            <td style="color:#4ade80; font-weight:bold;">${arrStr}</td>
            <td>${isStart ? '-' : `<input type="number" class="stop-input" value="${node.stopTime}" onchange="updateNodeStopTime('${node.id}', this.value)">`}</td>
            <td style="color:#facc15; font-weight:bold;">${depStr}</td>
            <td><span class="remove-node" onclick="removeRouteNode('${node.id}')">✖</span></td>`;
        tbody.appendChild(tr);
    });
    const tfoot = document.getElementById('routeFooter');
    if (isRoundTrip && routeNodes.length > 1 && routeNodes[0].returnArrTime) {
        document.getElementById('returnNodeName').innerText = routeNodes[0].id; document.getElementById('returnArrTime').innerText = formatTimeOfDay(routeNodes[0].returnArrTime); tfoot.style.display = 'table-footer-group';
    } else { tfoot.style.display = 'none'; }
    updateRouteHalos();
};

window.removeRouteNode = function(id) { routeNodes = routeNodes.filter(n => n.id !== id); if (routeNodes.length === 1 && routeNodes[0].isParent) routeNodes = []; calculateOSRMRoute(); };
window.rebuildRouteArrayFromDOM = function() { const listItems = document.querySelectorAll('#routeList tr[data-id]'); const newArr = []; listItems.forEach(tr => { const node = routeNodes.find(n => n.id === tr.dataset.id); if (node) newArr.push(node); }); routeNodes = newArr; };
window.clearRoute = function() { routeNodes = []; window.osrmLegs = []; layers.route.clearLayers(); document.getElementById('routeTotalTime').innerText = '0 хв'; document.getElementById('routeDist').innerText = '0 км'; renderRouteList(); };

window.saveCurrentRoute = function() {
    if (routeNodes.length < 2) return alert("Маршрут занадто короткий!");
    const routeName = routeNodes.map(n => n.id).join('_'); const distKm = document.getElementById('routeDist').innerText;
    const startTimeStr = document.getElementById('routeStartTime').value; const isRoundTrip = document.getElementById('routeRoundTrip').checked;
    const routeCopy = JSON.parse(JSON.stringify(routeNodes));
    window.savedRoutes.push({ id: Date.now(), name: routeName, distance: distKm, startTime: startTimeStr, isRoundTrip: isRoundTrip, nodes: routeCopy });
    document.getElementById('savedRoutesCount').innerText = window.savedRoutes.length; renderSavedRoutesList(); clearRoute(); 
};

window.renderSavedRoutesList = function() {
    const container = document.getElementById('savedRoutesList'); if (!container) return; container.innerHTML = '';
    window.savedRoutes.forEach(route => {
        const div = document.createElement('div'); div.className = 'saved-route-item';
        div.innerHTML = `<span class="saved-route-name" title="${route.name}">${route.name}</span><span class="remove-node" onclick="deleteSavedRoute(${route.id}, event)" title="Видалити">✖</span>`;
        div.onclick = () => loadSavedRoute(route.id); container.appendChild(div);
    }); updateRouteHalos(); 
};

window.deleteSavedRoute = function(id, event) { event.stopPropagation(); window.savedRoutes = window.savedRoutes.filter(r => r.id !== id); document.getElementById('savedRoutesCount').innerText = window.savedRoutes.length; renderSavedRoutesList(); };
window.loadSavedRoute = function(id) {
    const route = window.savedRoutes.find(r => r.id === id); if (!route) return;
    document.getElementById('routeRoundTrip').checked = route.isRoundTrip; document.getElementById('routeStartTime').value = route.startTime;
    routeNodes = JSON.parse(JSON.stringify(route.nodes)); calculateOSRMRoute();
};

window.exportRoutesToExcel = function() {
    if (window.savedRoutes.length === 0) return alert('Немає збережених ланцюгів!');
    let exportData = [];
    window.savedRoutes.forEach((route, rIdx) => {
        route.nodes.forEach((node, nIdx) => {
            exportData.push({ "Ланцюг": route.name, "№ п/п": nIdx + 1, "Вузол": node.id, "Тип": node.isParent ? (nIdx === 0 ? "Старт" : "Зірочка/Депо") : "Відділення", "Приїзд": node.arrTime !== null ? formatTimeOfDay(node.arrTime) : '-', "Виїзд": node.depTime !== null ? formatTimeOfDay(node.depTime) : '-', "Кілометраж": nIdx === 0 ? route.distance : "-" });
        });
        if (route.isRoundTrip && route.nodes[0].returnArrTime) {
            exportData.push({ "Ланцюг": route.name, "№ п/п": "Фініш", "Вузол": route.nodes[0].id, "Тип": "Повернення на базу", "Приїзд": formatTimeOfDay(route.nodes[0].returnArrTime), "Виїзд": "-", "Кілометраж": "-" });
        }
    });
    const ws = XLSX.utils.json_to_sheet(exportData); const wb = XLSX.utils.book_new(); ws['!cols'] = [{wch: 35}, {wch: 8}, {wch: 20}, {wch: 15}, {wch: 10}, {wch: 10}, {wch: 15}];
    XLSX.utils.book_append_sheet(wb, ws, "Ланцюги Руху"); XLSX.writeFile(wb, "Маршрути_Звіздар.xlsx");
};

// --- ПІДСВІЧУВАННЯ КАНДИДАТІВ У ЗІРКИ ПРИ ЗМІНІ ПЛОЩІ ---
let starHighlightTimeout;
window.highlightPotentialStars = function() {
    if (!layers.potentialStarHalos) return;
    
    // Очищаємо попередні ореоли та таймер
    layers.potentialStarHalos.clearLayers();
    clearTimeout(starHighlightTimeout);
    
    const minArea = parseFloat(document.getElementById('starMinArea').value) || 0;
    if (minArea <= 0 || filteredPoints.length === 0) return;

    // Динамічно додаємо CSS для красивої пульсації (якщо його ще немає)
    if (!document.getElementById('potential-star-style')) {
        const style = document.createElement('style');
        style.id = 'potential-star-style';
        style.innerHTML = `
            .potential-star-halo {
                border-radius: 50%;
                background: rgba(234, 179, 8, 0.4); 
                box-shadow: 0 0 15px 5px rgba(234, 179, 8, 0.8);
                animation: pulse-potential 1s infinite alternate;
            }
            @keyframes pulse-potential {
                0% { transform: scale(0.7); opacity: 0.5; }
                100% { transform: scale(1.3); opacity: 1; }
            }
        `;
        document.head.appendChild(style);
    }

    // Перебираємо точки і шукаємо ті, що проходять по площі
    filteredPoints.forEach(p => {
        if (p.isExcluded) return; // Ігноруємо виключені (червоні хрестики)
        
        const areaInfo = areaData.find(a => String(a['Вузол']) === String(p.id));
        const area = areaInfo ? parseFloat(areaInfo['Площа']) : 0; 
        
        if (area >= minArea) {
            const haloIcon = L.divIcon({
                className: 'potential-star-halo',
                iconSize: [30, 30],
                iconAnchor: [15, 15]
            });
            // Додаємо пульсуючий маркер під саму точку
            L.marker([p.lat, p.lng], { 
                icon: haloIcon, 
                interactive: false, 
                zIndexOffset: -30 
            }).addTo(layers.potentialStarHalos);
        }
    });

    // Ховаємо підсвітку через 3 секунди після останнього введення цифри
    starHighlightTimeout = setTimeout(() => {
        layers.potentialStarHalos.clearLayers();
    }, 3000);
};
function parseTimeToMinutes(timeStr) { let [h, m] = timeStr.split(':').map(Number); return (h * 60) + m; }
function formatTimeOfDay(totalMinutes) { let h = Math.floor(totalMinutes / 60) % 24; let m = totalMinutes % 60; return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`; }
function formatMinutesToTimeSpan(totalMinutes) { let h = Math.floor(totalMinutes / 60); let m = totalMinutes % 60; return h > 0 ? `${h} год ${m} хв` : `${m} хв`; }

// --- КЕРУВАННЯ ПІДСВІТКОЮ КАНДИДАТІВ У ЗІРКИ ---
window.isStarHighlightMode = false;

window.toggleStarHighlightMode = function() {
    window.isStarHighlightMode = !window.isStarHighlightMode;
    const btn = document.getElementById('btnStarHighlightToggle');
    
    if (window.isStarHighlightMode) {
        btn.style.background = '#eab308'; // Активний жовтий колір
        btn.style.color = '#000';
        updateStarHighlights();
    } else {
        btn.style.background = '';
        btn.style.color = '';
        if (layers.potentialStarHalos) {
            layers.potentialStarHalos.clearLayers();
        }
    }
};

window.updateStarHighlights = function() {
    if (!layers.potentialStarHalos) return;
    
    // Завжди очищаємо попередні ореоли перед перерахунком
    layers.potentialStarHalos.clearLayers();
    
    // Якщо режим вимкнено - просто виходимо
    if (!window.isStarHighlightMode) return;
    
    const minArea = parseFloat(document.getElementById('starMinArea').value) || 0;
    if (minArea <= 0 || filteredPoints.length === 0) return;

    if (!document.getElementById('potential-star-style')) {
        const style = document.createElement('style');
        style.id = 'potential-star-style';
        style.innerHTML = `
            .potential-star-halo {
                width: 100%;
                height: 100%;
                border-radius: 50%;
                background: rgba(234, 179, 8, 0.4); 
                box-shadow: 0 0 15px 5px rgba(234, 179, 8, 0.8);
                animation: pulse-potential 1.5s infinite alternate;
            }
            @keyframes pulse-potential {
                0% { transform: scale(0.8); opacity: 0.6; }
                100% { transform: scale(1.4); opacity: 1; }
            }
        `;
        document.head.appendChild(style);
    }

    filteredPoints.forEach(p => {
        // 1. Ігноруємо виключені точки
        if (p.isExcluded) return; 
        
        // 2. Ігноруємо точки, які зараз приховані фільтром Депо
        if (window.activeDepotFilters.size > 0 && p.assigned && !window.activeDepotFilters.has(p.assigned)) return;
        
        const areaInfo = areaData.find(a => String(a['Вузол']) === String(p.id));
        const area = areaInfo ? parseFloat(areaInfo['Площа']) : 0; 
        
        if (area >= minArea) {
            // ФІКС: Пустий className, анімація застосовується до внутрішнього div через html
            const haloIcon = L.divIcon({
                className: '', 
                html: '<div class="potential-star-halo"></div>',
                iconSize: [30, 30],
                iconAnchor: [15, 15]
            });
            L.marker([p.lat, p.lng], { 
                icon: haloIcon, 
                interactive: false, 
                zIndexOffset: -30 
            }).addTo(layers.potentialStarHalos);
        }
    });
};