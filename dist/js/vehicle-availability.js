/**
 * Vehicle Availability Tracker - MyGeotab Add-In
 *
 * Categorizes vehicles as:
 *   - Available  (inside a zone with ZoneType "Office")
 *   - Dispatched (outside all office zones)
 *
 * Uses point-in-polygon to check if a vehicle's GPS coordinates
 * fall inside zones that have ZoneTypeOfficeId assigned.
 */
geotab.addin.vehicleAvailabilityTracker = () => {
    'use strict';

    const REFRESH_INTERVAL_MS = 60000;
    const OFFICE_ZONE_TYPE_ID = 'ZoneTypeOfficeId';

    let api;
    let refreshTimer = null;
    let map = null;
    let markersLayer = null;
    let zonesLayer = null;
    let vehicleData = [];
    let sortColumn = 'name';
    let sortAsc = true;

    // ── DOM helpers ──────────────────────────────────────────

    const $ = (sel) => document.querySelector(sel);
    const $$ = (sel) => document.querySelectorAll(sel);

    // ── Leaflet marker icons ─────────────────────────────────

    function createIcon(color) {
        return L.divIcon({
            className: 'va-marker',
            html: `<svg width="28" height="40" viewBox="0 0 28 40" xmlns="http://www.w3.org/2000/svg">
                <path d="M14 0C6.27 0 0 6.27 0 14c0 10.5 14 26 14 26s14-15.5 14-26C28 6.27 21.73 0 14 0z" fill="${color}"/>
                <circle cx="14" cy="14" r="6" fill="#fff"/>
            </svg>`,
            iconSize: [28, 40],
            iconAnchor: [14, 40],
            popupAnchor: [0, -36]
        });
    }

    // Defer icon creation until initialize() when Leaflet is guaranteed loaded
    let ICON_AVAILABLE = null;
    let ICON_DISPATCHED = null;

    // ── Point-in-polygon ───────────────────────────────────────

    function pointInPolygon(lat, lng, points) {
        var n = points.length;
        var inside = false;
        var j = n - 1;
        for (var i = 0; i < n; i++) {
            var yi = points[i].y, xi = points[i].x; // y=lat, x=lng
            var yj = points[j].y, xj = points[j].x;
            if (((yi > lat) !== (yj > lat)) &&
                (lng < (xj - xi) * (lat - yi) / (yj - yi) + xi)) {
                inside = !inside;
            }
            j = i;
        }
        return inside;
    }

    // ── Data loading ─────────────────────────────────────────

    function loadData() {
        const btn = $('#va-refresh-btn');
        const loading = $('#va-loading');
        if (btn) btn.disabled = true;
        if (loading) loading.classList.remove('hidden');

        api.multiCall([
            ['Get', { typeName: 'Device' }],
            ['Get', { typeName: 'DeviceStatusInfo' }],
            ['Get', { typeName: 'Zone' }]
        ], function (results) {
            const devices = results[0];
            const statuses = results[1];
            const zones = results[2];

            processData(devices, statuses, zones);

            if (btn) btn.disabled = false;
            if (loading) loading.classList.add('hidden');
            updateLastRefreshed();
        }, function (error) {
            console.error('Vehicle Availability - API error:', error);
            if (btn) btn.disabled = false;
            if (loading) loading.classList.add('hidden');
        });
    }

    function processData(devices, statuses, zones) {
        // Build lookup maps
        const statusByDeviceId = {};
        statuses.forEach(function (s) {
            if (s.device && s.device.id) {
                statusByDeviceId[s.device.id] = s;
            }
        });

        // Identify office zones by ZoneType (ZoneTypeOfficeId)
        const officeZones = zones.filter(function (z) {
            var ztypes = z.zoneTypes || [];
            return ztypes.some(function (zt) {
                var id = (typeof zt === 'string') ? zt : (zt.id || '');
                return id === OFFICE_ZONE_TYPE_ID;
            }) && z.points && z.points.length >= 3;
        });

        // Build vehicle data using point-in-polygon for zone containment
        vehicleData = devices.map(function (device) {
            const status = statusByDeviceId[device.id];
            const lat = status ? status.latitude : null;
            const lng = status ? status.longitude : null;
            const speed = status ? status.speed : 0;
            const dateTime = status ? status.dateTime : null;
            const communicating = status ? status.isDeviceCommunicating : false;

            // Check if vehicle is inside any office zone using point-in-polygon
            var matchedZoneNames = [];
            var inOfficeZone = false;
            if (lat && lng && (lat !== 0 || lng !== 0)) {
                officeZones.forEach(function (zone) {
                    if (pointInPolygon(lat, lng, zone.points)) {
                        inOfficeZone = true;
                        matchedZoneNames.push(zone.name);
                    }
                });
            }

            var location = (lat && lng && (lat !== 0 || lng !== 0))
                ? lat.toFixed(5) + ', ' + lng.toFixed(5)
                : '-';

            return {
                id: device.id,
                name: device.name || '(unnamed)',
                serialNumber: device.serialNumber || '',
                status: inOfficeZone ? 'available' : 'dispatched',
                statusLabel: inOfficeZone ? 'Available' : 'Dispatched',
                location: location,
                lat: lat,
                lng: lng,
                speed: speed,
                communicating: communicating,
                currentZones: matchedZoneNames.join(', ') || '-',
                lastUpdated: dateTime ? formatDateTime(dateTime) : '-'
            };
        });

        // Filter out devices with no GPS position (lat/lng = 0)
        const mappableVehicles = vehicleData.filter(function (v) {
            return v.lat && v.lng && (v.lat !== 0 || v.lng !== 0);
        });

        renderSummary();
        renderTable();
        renderMap(mappableVehicles, officeZones);
    }

    // ── Rendering ────────────────────────────────────────────

    function renderSummary() {
        const available = vehicleData.filter(function (v) { return v.status === 'available'; }).length;
        const dispatched = vehicleData.filter(function (v) { return v.status === 'dispatched'; }).length;
        const total = vehicleData.length;

        $('#va-count-available').textContent = available;
        $('#va-count-dispatched').textContent = dispatched;
        $('#va-count-total').textContent = total;
    }

    function renderTable() {
        const filter = ($('#va-filter-status') || {}).value || 'all';
        const search = (($('#va-search') || {}).value || '').toLowerCase();

        let filtered = vehicleData;

        if (filter !== 'all') {
            filtered = filtered.filter(function (v) { return v.status === filter; });
        }

        if (search) {
            filtered = filtered.filter(function (v) {
                return v.name.toLowerCase().indexOf(search) !== -1 ||
                       v.serialNumber.toLowerCase().indexOf(search) !== -1 ||
                       v.currentZones.toLowerCase().indexOf(search) !== -1;
            });
        }

        // Sort
        filtered.sort(function (a, b) {
            let valA = a[sortColumn] || '';
            let valB = b[sortColumn] || '';
            if (typeof valA === 'string') valA = valA.toLowerCase();
            if (typeof valB === 'string') valB = valB.toLowerCase();
            if (valA < valB) return sortAsc ? -1 : 1;
            if (valA > valB) return sortAsc ? 1 : -1;
            return 0;
        });

        const tbody = $('#va-table-body');
        if (!tbody) return;

        if (filtered.length === 0) {
            tbody.innerHTML = '<tr><td colspan="6" class="va-empty">No vehicles found</td></tr>';
            updateSortArrows();
            return;
        }

        tbody.innerHTML = filtered.map(function (v) {
            return '<tr data-id="' + v.id + '" data-lat="' + v.lat + '" data-lng="' + v.lng + '">' +
                '<td>' + escapeHtml(v.name) + '</td>' +
                '<td>' + escapeHtml(v.serialNumber) + '</td>' +
                '<td><span class="va-status ' + v.status + '">' + v.statusLabel + '</span></td>' +
                '<td>' + escapeHtml(v.location) + '</td>' +
                '<td>' + escapeHtml(v.currentZones) + '</td>' +
                '<td>' + escapeHtml(v.lastUpdated) + '</td>' +
                '</tr>';
        }).join('');

        // Row click → center map
        $$('#va-table-body tr').forEach(function (row) {
            row.addEventListener('click', function () {
                const lat = parseFloat(row.dataset.lat);
                const lng = parseFloat(row.dataset.lng);
                if (lat && lng && map) {
                    map.setView([lat, lng], 15);
                    // Highlight row
                    $$('#va-table-body tr').forEach(function (r) { r.classList.remove('highlighted'); });
                    row.classList.add('highlighted');
                }
            });
        });

        updateSortArrows();
    }

    function renderMap(vehicles, officeZones) {
        if (!map) return;

        // Clear layers
        if (markersLayer) markersLayer.clearLayers();
        if (zonesLayer) zonesLayer.clearLayers();

        // Draw office zone polygons
        officeZones.forEach(function (zone) {
            const latlngs = zone.points.map(function (p) {
                return [p.y, p.x]; // y = lat, x = lng
            });
            L.polygon(latlngs, {
                color: '#2196F3',
                weight: 2,
                fillColor: '#2196F3',
                fillOpacity: 0.1
            }).addTo(zonesLayer).bindPopup('<b>' + escapeHtml(zone.name) + '</b><br>Office Zone');
        });

        // Place vehicle markers
        vehicles.forEach(function (v) {
            if (!v.lat || !v.lng) return;

            const icon = v.status === 'available' ? ICON_AVAILABLE : ICON_DISPATCHED;
            const marker = L.marker([v.lat, v.lng], { icon: icon }).addTo(markersLayer);
            marker.bindPopup(
                '<b>' + escapeHtml(v.name) + '</b><br>' +
                '<span class="va-status ' + v.status + '" style="font-size:11px;">' + v.statusLabel + '</span><br>' +
                'Serial: ' + escapeHtml(v.serialNumber) + '<br>' +
                'Zone: ' + escapeHtml(v.currentZones) + '<br>' +
                'Speed: ' + v.speed + ' km/h<br>' +
                'Updated: ' + escapeHtml(v.lastUpdated)
            );
        });

        // Fit bounds to all markers
        if (vehicles.length > 0) {
            const bounds = markersLayer.getBounds();
            if (bounds.isValid()) {
                map.fitBounds(bounds, { padding: [40, 40] });
            }
        }
    }

    // ── Helpers ──────────────────────────────────────────────

    function formatDateTime(dt) {
        try {
            const d = new Date(dt);
            if (isNaN(d.getTime())) return String(dt);
            return d.toLocaleString();
        } catch (e) {
            return String(dt);
        }
    }

    function escapeHtml(str) {
        if (!str) return '';
        var div = document.createElement('div');
        div.appendChild(document.createTextNode(String(str)));
        return div.innerHTML;
    }

    function updateLastRefreshed() {
        const el = $('#va-last-updated');
        if (el) el.textContent = 'Last refreshed: ' + new Date().toLocaleTimeString();
    }

    function updateSortArrows() {
        $$('.va-table th[data-sort]').forEach(function (th) {
            const arrow = th.querySelector('.sort-arrow');
            if (!arrow) return;
            if (th.dataset.sort === sortColumn) {
                arrow.textContent = sortAsc ? '\u25B2' : '\u25BC';
                arrow.classList.add('active');
            } else {
                arrow.textContent = '\u25B2';
                arrow.classList.remove('active');
            }
        });
    }

    // ── Add-In lifecycle ─────────────────────────────────────

    return {
        initialize: function (freshApi, state, callback) {
            api = freshApi;

            // Guard: if Leaflet failed to load, show error and continue without map
            if (typeof L === 'undefined') {
                console.error('Vehicle Availability - Leaflet library failed to load');
                var mapEl = document.getElementById('va-map');
                if (mapEl) mapEl.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#999;">Map unavailable (library failed to load)</div>';
                callback();
                return;
            }

            // Create marker icons now that Leaflet is loaded
            ICON_AVAILABLE = createIcon('#4CAF50');
            ICON_DISPATCHED = createIcon('#FF9800');

            // Initialize Leaflet map
            map = L.map('va-map').setView([43.65, -79.38], 10);
            L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                attribution: '&copy; OpenStreetMap contributors',
                maxZoom: 19
            }).addTo(map);

            markersLayer = L.featureGroup().addTo(map);
            zonesLayer = L.featureGroup().addTo(map);

            // Bind refresh button
            var btn = $('#va-refresh-btn');
            if (btn) btn.addEventListener('click', loadData);

            // Bind filter controls
            var filterEl = $('#va-filter-status');
            if (filterEl) filterEl.addEventListener('change', renderTable);

            var searchEl = $('#va-search');
            if (searchEl) searchEl.addEventListener('input', renderTable);

            // Bind table sorting
            $$('.va-table th[data-sort]').forEach(function (th) {
                th.addEventListener('click', function () {
                    var col = th.dataset.sort;
                    if (sortColumn === col) {
                        sortAsc = !sortAsc;
                    } else {
                        sortColumn = col;
                        sortAsc = true;
                    }
                    renderTable();
                });
            });

            callback();
        },

        focus: function (freshApi, state) {
            api = freshApi;

            // Fix Leaflet rendering when container was hidden
            if (map) {
                setTimeout(function () { map.invalidateSize(); }, 200);
            }

            // Load data immediately and start auto-refresh
            loadData();
            refreshTimer = setInterval(loadData, REFRESH_INTERVAL_MS);
        },

        blur: function () {
            if (refreshTimer) {
                clearInterval(refreshTimer);
                refreshTimer = null;
            }
        }
    };
};
