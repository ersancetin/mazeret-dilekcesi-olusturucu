(function () {
    'use strict';

    var allHearings = [];
    var el = {};
    var MAX_TABLE_HEARINGS = 5;
    var grouped = {};
    var conflictKeys = [];
    var lawyerInfo = { name: '', baro: '', baroSicil: '', adres: '' };
    var DAY_NAMES = ['Paz', 'Pzt', 'Sal', 'Çar', 'Per', 'Cum', 'Cmt'];
    var openDays = {};

    function init() {
        el = {
            uploadSection: document.getElementById('upload-section'),
            infoSection: document.getElementById('info-section'),
            resultSection: document.getElementById('result-section'),
            dropZone: document.getElementById('drop-zone'),
            fileInput: document.getElementById('file-input'),
            lawyerName: document.getElementById('lawyer-name'),
            baroName: document.getElementById('baro-name'),
            baroSicil: document.getElementById('baro-sicil'),
            lawyerAddress: document.getElementById('lawyer-address'),
            continueBtn: document.getElementById('continue-btn'),
            backToUpload: document.getElementById('back-to-upload'),
            stats: document.getElementById('stats'),
            searchInput: document.getElementById('search-input'),
            searchClear: document.getElementById('search-clear'),
            downloadAllBtn: document.getElementById('download-all-btn'),
            resetBtn: document.getElementById('reset-btn'),
            daysList: document.getElementById('days-list'),
            progressBar: document.getElementById('progress-bar'),
            progressFill: document.getElementById('progress-fill'),
            progressText: document.getElementById('progress-text'),
            errorToast: document.getElementById('error-toast')
        };

        setupDragDrop();
        el.fileInput.addEventListener('change', function (e) {
            if (e.target.files[0]) handleFile(e.target.files[0]);
        });
        el.continueBtn.addEventListener('click', handleContinue);
        el.lawyerAddress.addEventListener('input', autoWrapTextarea);
        el.backToUpload.addEventListener('click', function () {
            el.infoSection.classList.add('hidden');
            el.uploadSection.classList.remove('hidden');
        });
        el.downloadAllBtn.addEventListener('click', handleDownloadAll);
        el.resetBtn.addEventListener('click', handleReset);

        el.searchInput.addEventListener('input', handleSearch);
        el.searchClear.addEventListener('click', function () {
            el.searchInput.value = '';
            el.searchClear.classList.add('hidden');
            renderDaysList('');
        });

        el.daysList.addEventListener('click', handleDaysListClick);
    }

    function handleSearch() {
        var q = el.searchInput.value.trim();
        if (q.length > 0) {
            el.searchClear.classList.remove('hidden');
        } else {
            el.searchClear.classList.add('hidden');
        }
        renderDaysList(q);
    }

    function handleDaysListClick(e) {
        var target = e.target;

        var dlBtn = target.closest('.day-download-btn');
        if (dlBtn) {
            e.stopPropagation();
            downloadDay(dlBtn.getAttribute('data-key'));
            return;
        }

        var singleBtn = target.closest('.single-download-btn');
        if (singleBtn) {
            var hid = +singleBtn.getAttribute('data-hid');
            var key = singleBtn.getAttribute('data-key');
            downloadSingleHearing(hid, key);
            return;
        }

        var header = target.closest('.day-header');
        if (header) {
            if (el.searchInput.value.trim()) return;
            var dayKey = header.getAttribute('data-key');
            openDays[dayKey] = !openDays[dayKey];
            var card = header.closest('.day-card');
            var hearings = card.querySelector('.day-hearings');
            var toggle = card.querySelector('.day-toggle');
            if (hearings) hearings.classList.toggle('open');
            if (toggle) toggle.classList.toggle('open');
        }
    }

    // ── Drag & Drop ──

    function setupDragDrop() {
        var dz = el.dropZone;
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(function (evt) {
            dz.addEventListener(evt, function (e) { e.preventDefault(); e.stopPropagation(); });
        });
        dz.addEventListener('dragenter', function () { dz.classList.add('drag-over'); });
        dz.addEventListener('dragover', function () { dz.classList.add('drag-over'); });
        dz.addEventListener('dragleave', function () { dz.classList.remove('drag-over'); });
        dz.addEventListener('drop', function (e) {
            dz.classList.remove('drag-over');
            if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
        });
        dz.addEventListener('click', function () { el.fileInput.click(); });
        dz.addEventListener('keydown', function (e) {
            if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); el.fileInput.click(); }
        });
    }

    // ── File Handling ──

    function handleFile(file) {
        if (!file.name.match(/\.xlsx?$/i)) {
            showError('Lütfen bir Excel (.xlsx veya .xls) dosyası yükleyin.');
            return;
        }
        var reader = new FileReader();
        reader.onload = function (e) {
            try {
                allHearings = parseExcel(e.target.result);
                if (allHearings.length === 0) {
                    showError('Dosyada duruşma verisi bulunamadı.');
                    return;
                }
                grouped = groupByDate(allHearings);
                var keys = Object.keys(grouped).sort();
                conflictKeys = keys.filter(function (k) { return grouped[k].length > 1; });

                if (conflictKeys.length === 0) {
                    showError('Aynı güne denk gelen duruşma bulunamadı.');
                    return;
                }

                showInfoForm();
            } catch (err) {
                showError('Dosya okunamadı: ' + err.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // ── Excel Parsing ──

    function parseExcel(buffer) {
        var wb = XLSX.read(buffer, { type: 'array', cellDates: true });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (raw.length < 2) return [];

        var startRow = 0;
        var first = raw[0];
        if (first && typeof first[0] === 'string') {
            var f = first[0].toLowerCase();
            if (f.indexOf('birim') >= 0 || f.indexOf('mahkeme') >= 0 || f.indexOf('merci') >= 0) {
                startRow = 1;
            }
        }

        var hearings = [];
        for (var i = startRow; i < raw.length; i++) {
            var row = raw[i];
            if (!row || (!row[0] && !row[1])) continue;
            var birim = String(row[0] || '').trim();
            var dosyaNo = String(row[1] || '').trim();
            var tarih = parseDateValue(row[3]);
            if (birim || dosyaNo) {
                hearings.push({ birim: birim, dosyaNo: dosyaNo, tarih: tarih, id: i });
            }
        }
        return hearings;
    }

    function parseDateValue(val) {
        if (!val) return null;
        if (val instanceof Date && !isNaN(val.getTime())) return val;
        var str = String(val).trim();
        var m = str.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})\s+(\d{1,2}):(\d{2})/);
        if (m) return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]);
        m = str.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/);
        if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
        m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
        var d = new Date(str);
        return isNaN(d.getTime()) ? null : d;
    }

    // ── Date Utilities ──

    function dateKey(d) {
        if (!d) return 'tarihsiz';
        return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    }
    function formatDate(d) {
        if (!d) return '-';
        return pad(d.getDate()) + '.' + pad(d.getMonth() + 1) + '.' + d.getFullYear();
    }
    function formatDateTime(d) {
        if (!d) return '-';
        var base = formatDate(d);
        var h = d.getHours(), m = d.getMinutes();
        if (h === 0 && m === 0) return base;
        return base + ' ' + pad(h) + ':' + pad(m);
    }
    function formatTime(d) {
        if (!d) return '-';
        var h = d.getHours(), m = d.getMinutes();
        if (h === 0 && m === 0) return '-';
        return pad(h) + ':' + pad(m);
    }
    function pad(n) { return n < 10 ? '0' + n : '' + n; }

    // ── Grouping ──

    function groupByDate(hearings) {
        var groups = {};
        hearings.forEach(function (h) {
            var key = dateKey(h.tarih);
            if (!groups[key]) groups[key] = [];
            groups[key].push(h);
        });
        return groups;
    }

    // ── Closest Hearings ──

    function getClosestHearings(hearing, group, count) {
        var others = group.filter(function (h) { return h.id !== hearing.id; });
        if (!hearing.tarih) return others.slice(0, count);
        var t = hearing.tarih.getTime();
        others.sort(function (a, b) {
            var da = Math.abs((a.tarih ? a.tarih.getTime() : 0) - t);
            var db = Math.abs((b.tarih ? b.tarih.getTime() : 0) - t);
            return da - db;
        });
        return others.slice(0, count);
    }

    // ── UI: Info Form ──

    function showInfoForm() {
        el.uploadSection.classList.add('hidden');
        el.infoSection.classList.remove('hidden');
        el.lawyerName.focus();
    }

    function handleContinue() {
        lawyerInfo.name = el.lawyerName.value.trim();
        lawyerInfo.baro = el.baroName.value.trim();
        lawyerInfo.baroSicil = el.baroSicil.value.trim();
        lawyerInfo.adres = el.lawyerAddress.value.trim();
        showResults();
    }

    // ── UI: Results ──

    function showResults() {
        el.infoSection.classList.add('hidden');
        el.resultSection.classList.remove('hidden');

        var totalPetitions = 0;
        conflictKeys.forEach(function (k) { totalPetitions += grouped[k].length; });

        el.stats.innerHTML =
            '<div class="stat-item"><div class="stat-value">' + allHearings.length + '</div><div class="stat-label">Toplam Duruşma</div></div>' +
            '<div class="stat-item accent"><div class="stat-value accent">' + conflictKeys.length + '</div><div class="stat-label">Duruşma Günü</div></div>' +
            '<div class="stat-item"><div class="stat-value">' + totalPetitions + '</div><div class="stat-label">Toplam Dilekçe</div></div>';

        openDays = {};
        conflictKeys.forEach(function (k) { openDays[k] = false; });
        renderDaysList('');
    }

    function renderDaysList(filter) {
        var q = (filter || '').toLocaleLowerCase('tr');
        var isSearching = q.length > 0;
        var html = '';
        var visibleCount = 0;

        conflictKeys.forEach(function (key) {
            var group = grouped[key];
            var filtered;

            if (isSearching) {
                filtered = group.filter(function (h) {
                    return h.birim.toLocaleLowerCase('tr').indexOf(q) >= 0 ||
                           h.dosyaNo.toLocaleLowerCase('tr').indexOf(q) >= 0;
                });
            } else {
                filtered = group;
            }

            if (filtered.length === 0) return;
            visibleCount++;

            var d = group[0].tarih;
            var dayName = d ? DAY_NAMES[d.getDay()] : '';
            var dateDisplay = d ? formatDate(d) : key;
            var isOpen = isSearching ? true : !!openDays[key];

            html += '<div class="day-card" data-key="' + key + '">';
            html += '<div class="day-header" data-key="' + key + '">';
            html += '<div class="day-info">';
            html += '<span class="day-name">' + esc(dayName) + '</span>';
            html += '<span class="day-date">' + esc(dateDisplay) + '</span>';
            html += '<span class="day-count">' + group.length + ' duruşma';
            if (isSearching && filtered.length !== group.length) {
                html += ' (' + filtered.length + ' eşleşme)';
            }
            html += '</span>';
            html += '</div>';
            html += '<div class="day-actions">';
            html += '<button class="btn btn-sm btn-primary day-download-btn" data-key="' + key + '">';
            html += '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>';
            html += 'İndir</button>';
            html += '<span class="day-toggle' + (isOpen ? ' open' : '') + '">&#9660;</span>';
            html += '</div>';
            html += '</div>';

            html += '<div class="day-hearings' + (isOpen ? ' open' : '') + '">';
            filtered.forEach(function (h) {
                var time = formatTime(h.tarih);
                html += '<div class="hearing-row">';
                html += '<div class="hearing-info">';
                if (isSearching) {
                    html += '<span class="hearing-birim">' + highlightMatch(h.birim, q) + '</span>';
                    html += '<span class="hearing-meta">' + highlightMatch(h.dosyaNo, q);
                } else {
                    html += '<span class="hearing-birim">' + esc(h.birim) + '</span>';
                    html += '<span class="hearing-meta">' + esc(h.dosyaNo);
                }
                if (time !== '-') html += ' &middot; ' + time;
                html += '</span>';
                html += '</div>';
                html += '<button class="btn-icon single-download-btn" data-hid="' + h.id + '" data-key="' + key + '" title="UDF İndir">';
                html += '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>';
                html += '</button>';
                html += '</div>';
            });
            html += '</div>';
            html += '</div>';
        });

        if (visibleCount === 0) {
            html = '<div class="empty-state">Sonuç bulunamadı.</div>';
        }

        el.daysList.innerHTML = html;
    }

    function highlightMatch(text, query) {
        var safe = esc(text);
        if (!query) return safe;
        var lower = text.toLocaleLowerCase('tr');
        var idx = lower.indexOf(query);
        if (idx === -1) return safe;
        var before = esc(text.substring(0, idx));
        var match = esc(text.substring(idx, idx + query.length));
        var after = esc(text.substring(idx + query.length));
        return before + '<span class="highlight">' + match + '</span>' + after;
    }

    // ── Address Auto Wrap ──

    var WRAP_LEN = 30;
    var wrapping = false;

    function autoWrapTextarea() {
        if (wrapping) return;
        wrapping = true;
        var ta = el.lawyerAddress;
        var pos = ta.selectionStart;
        var val = ta.value;
        var lines = val.split('\n');
        var newLines = [];
        var charsBefore = 0;
        var newPos = pos;

        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            if (line.length <= WRAP_LEN) {
                newLines.push(line);
                charsBefore += line.length + 1;
                continue;
            }
            var wrapped = wrapLine(line, WRAP_LEN);
            var oldLen = line.length;
            var lineStart = charsBefore;
            if (pos > lineStart && pos <= lineStart + oldLen) {
                var cursorInLine = pos - lineStart;
                var acc = 0;
                for (var w = 0; w < wrapped.length; w++) {
                    var wLen = wrapped[w].length;
                    if (cursorInLine <= acc + wLen) {
                        newPos = lineStart + acc + (cursorInLine - acc) + w;
                        break;
                    }
                    acc += wLen;
                    if (w < wrapped.length - 1) acc++;
                }
            }
            for (var w = 0; w < wrapped.length; w++) {
                newLines.push(wrapped[w]);
            }
            charsBefore += oldLen + 1;
        }

        var result = newLines.join('\n');
        if (result !== val) {
            ta.value = result;
            if (newPos > result.length) newPos = result.length;
            ta.selectionStart = ta.selectionEnd = newPos;
        }
        wrapping = false;
    }

    function wrapLine(text, maxLen) {
        text = text.trim();
        if (text.length <= maxLen) return [text];
        var lines = [];
        while (text.length > maxLen) {
            var cut = -1;
            for (var i = maxLen; i >= Math.floor(maxLen / 2); i--) {
                if (text.charAt(i) === ',') { cut = i + 1; break; }
            }
            if (cut === -1) {
                for (var i = maxLen; i >= Math.floor(maxLen / 2); i--) {
                    if (text.charAt(i) === ' ') { cut = i; break; }
                }
            }
            if (cut === -1) cut = maxLen;
            lines.push(text.substring(0, cut).trim());
            text = text.substring(cut).trim();
        }
        if (text) lines.push(text);
        return lines;
    }

    // ── UDF Builder ──

    var ZWSP = '​';

    function UDFBuilder() {
        this.pool = '';
        this.elems = [];
    }

    UDFBuilder.prototype.p = function (runs, opts) {
        if (typeof runs === 'string') runs = [{ text: runs }];
        if (typeof opts === 'number') opts = { align: opts };
        opts = opts || {};

        var parts = [];
        for (var i = 0; i < runs.length; i++) {
            var r = runs[i];
            var family = r.family || 'Times New Roman';
            var size = r.size || 12;
            var segments = r.text.split('\t');
            for (var j = 0; j < segments.length; j++) {
                if (j > 0) {
                    var tabOff = this.pool.length;
                    this.pool += '\t';
                    parts.push('<tab startOffset="' + tabOff + '" length="1" family="' + xmlAttr(family) + '" size="' + size + '" />');
                }
                var seg = segments[j];
                if (seg.length > 0) {
                    var off = this.pool.length;
                    this.pool += seg;
                    var a = 'startOffset="' + off + '" length="' + seg.length + '"';
                    a += ' family="' + xmlAttr(family) + '"';
                    a += ' size="' + size + '"';
                    if (r.bold) a += ' bold="true"';
                    if (r.italic) a += ' italic="true"';
                    if (r.underline) a += ' underline="true"';
                    parts.push('<content ' + a + ' />');
                }
            }
        }
        if (parts.length === 0) {
            var off2 = this.pool.length;
            this.pool += ZWSP;
            parts.push('<content startOffset="' + off2 + '" length="1" family="Times New Roman" size="' + (opts.size || 12) + '" />');
        }

        var pa = 'Alignment="' + (opts.align || 0) + '"';
        pa += ' LeftIndent="' + (opts.left || 0) + '"';
        pa += ' RightIndent="' + (opts.right || 0) + '"';
        if (opts.first) pa += ' FirstLineIndent="' + opts.first + '"';
        if (opts.line) pa += ' LineSpacing="' + opts.line + '"';
        if (opts.above) pa += ' SpaceAbove="' + opts.above + '"';
        if (opts.below) pa += ' SpaceBelow="' + opts.below + '"';

        this.elems.push('<paragraph ' + pa + '>' + parts.join('') + '</paragraph>');
    };

    UDFBuilder.prototype.table = function (headers, rows, colSpans, colAligns) {
        var colCount = headers.length;
        var spans = colSpans || '30,220,100,120';
        var allRows = [headers].concat(rows);
        var rowsXml = '';

        for (var r = 0; r < allRows.length; r++) {
            var isHdr = r === 0;
            var cellsXml = '';
            for (var c = 0; c < colCount; c++) {
                var text = String(allRows[r][c] || '') || ' ';
                var off = this.pool.length;
                this.pool += text;
                var a = 'startOffset="' + off + '" length="' + text.length + '"';
                a += ' family="Times New Roman" size="12"';
                if (isHdr) a += ' bold="true"';
                var align = colAligns && colAligns[c] !== undefined ? colAligns[c] : 1;
                var cellPa = 'Alignment="' + align + '" SpaceAbove="3" SpaceBelow="3"';
                if (align === 0) cellPa += ' LeftIndent="4"';
                cellsXml += '<cell><paragraph ' + cellPa + '><content ' + a + ' /></paragraph></cell>';
            }
            rowsXml += '<row rowName="row' + (r + 1) + '" rowType="dataRow">' + cellsXml + '</row>';
        }

        this.elems.push(
            '<table tableName="Sabit" columnCount="' + colCount + '"' +
            ' columnSpans="' + spans + '" border="borderCell">' +
            rowsXml + '</table>'
        );
    };

    UDFBuilder.prototype.fields = function (rows) {
        var rowsXml = '';
        for (var r = 0; r < rows.length; r++) {
            var f = rows[r];
            var cellsXml = '';
            var lblOff = this.pool.length;
            this.pool += f.label;
            cellsXml += '<cell><paragraph Alignment="0"><content startOffset="' + lblOff +
                '" length="' + f.label.length + '" family="Times New Roman" size="12" bold="true" /></paragraph></cell>';
            var colOff = this.pool.length;
            this.pool += ': ';
            var valOff = this.pool.length;
            this.pool += f.value;
            var pa = 'Alignment="0"';
            if (f.below) pa += ' SpaceBelow="' + f.below + '"';
            if (f.above) pa += ' SpaceAbove="' + f.above + '"';
            cellsXml += '<cell><paragraph ' + pa + '>' +
                '<content startOffset="' + colOff + '" length="2" family="Times New Roman" size="12" bold="true" />' +
                '<content startOffset="' + valOff + '" length="' + f.value.length + '" family="Times New Roman" size="12" />' +
                '</paragraph></cell>';
            rowsXml += '<row rowName="row' + (r + 1) + '" rowType="dataRow">' + cellsXml + '</row>';
        }
        this.elems.push('<table tableName="Sabit" columnCount="2" columnSpans="80,400" border="borderNone">' + rowsXml + '</table>');
    };

    UDFBuilder.prototype.build = function () {
        return '<?xml version="1.0" encoding="UTF-8" ?>\n' +
            '<template format_id="1.8">\n' +
            '<content><![CDATA[' + this.pool + ']]></content>\n' +
            '<properties><pageFormat mediaSizeName="1"' +
            ' leftMargin="56.692913385826778" rightMargin="56.692913385826778"' +
            ' topMargin="56.692913385826778" bottomMargin="42.51968503937008"' +
            ' paperOrientation="1" headerFOffset="20.0" footerFOffset="20.0" /></properties>\n' +
            '<elements resolver="hvl-default">\n' +
            this.elems.join('\n') + '\n' +
            '</elements>\n' +
            '<styles><style name="default" description="Geçerli" family="Dialog" size="12"' +
            ' bold="false" italic="false" foreground="-13421773"' +
            ' FONT_ATTRIBUTE_KEY="javax.swing.plaf.FontUIResource[family=Dialog,name=Dialog,style=plain,size=12]" />' +
            '<style name="hvl-default" family="Times New Roman" size="12" description="Gövde" /></styles>\n' +
            '</template>';
    };

    // ── Dilekçe metni ──

    function birimToAddress(text) {
        var upper = text.toLocaleUpperCase('tr');
        if (upper.charAt(upper.length - 1) === 'I') return upper + 'NA';
        return upper + 'NE';
    }

    function petitionText(h, ds, ts, info, others, totalOnDay) {
        var lawyer = info.name ? 'Av. ' + info.name : 'Av. _______________';
        var tarihExpr = ds + ' tarihli' + (ts !== '-' ? ' saat ' + ts : '') + ' duruşmasına';
        var baroLine = '';
        if (info.baro && info.baroSicil) {
            baroLine = info.baro + ' Barosu, Sicil No: ' + info.baroSicil;
        } else if (info.baro) {
            baroLine = info.baro + ' Barosu';
        } else if (info.baroSicil) {
            baroLine = 'Sicil No: ' + info.baroSicil;
        }
        var adresLines = [];
        if (info.adres) {
            info.adres.split('\n').filter(function (l) { return l.trim(); }).forEach(function (line) {
                wrapLine(line, WRAP_LEN).forEach(function (wl) { adresLines.push(wl); });
            });
        }
        var allOthersShown = others.length >= totalOnDay - 1;
        var ekListeIfade = allOthersShown
            ? 'Aynı güne ait diğer duruşmalarımın listesi dilekçemiz ekinde arz olunmaktadır.'
            : 'Aynı güne ait duruşmalarımın bir kısmına ilişkin liste dilekçemiz ekinde arz olunmaktadır.';
        var ekBaslik = allOthersShown
            ? 'EK: Aynı Güne Ait Duruşma Listesi'
            : 'EK: Aynı Güne Ait Duruşmalarımın Bir Kısmına İlişkin Liste';
        return {
            birim: birimToAddress(h.birim),
            dosyaNo: h.dosyaNo,
            konu: 'Mesleki mazeretimizin sunulmasından ibarettir.',
            ekBaslik: ekBaslik,
            aciklama: 'Mahkemenizin yukarıda esas numarası yazılı dosyasının ' +
                tarihExpr + ', aynı tarihte muhtelif mahkemelerde toplam ' +
                totalOnDay + ' adet duruşmam bulunması sebebiyle katılma imkânım ' +
                'bulunmamaktadır. ' + ekListeIfade,
            sonuc: 'Yukarıda arz ve izah olunan nedenlerle; mesleki mazeretimin kabulüne, ' +
                'duruşmanın başka bir güne ertelenmesine ve yeni duruşma gün ve saatinin ' +
                'tarafıma tebliğine karar verilmesini saygılarımla arz ve talep ederim. ',
            tarih: ds,
            lawyer: lawyer,
            baroLine: baroLine,
            adresLines: adresLines,
            others: others
        };
    }

    // ── UDF Generation ──

    function createUDF(hearing, group) {
        var ds = formatDate(hearing.tarih);
        var ts = formatTime(hearing.tarih);
        var closest = getClosestHearings(hearing, group, MAX_TABLE_HEARINGS);
        var t = petitionText(hearing, ds, ts, lawyerInfo, closest, group.length);

        var b = new UDFBuilder();

        b.p('', { above: 40 });
        b.p([{ text: t.birim, bold: true, size: 13 }], { align: 1, below: 20 });

        b.fields([
            { label: 'DOSYA NO', value: t.dosyaNo, below: 4 },
            { label: 'KONU', value: t.konu, below: 36 }
        ]);

        b.p([{ text: 'AÇIKLAMALAR', bold: true, underline: true }], { below: 10 });
        b.p([{ text: t.aciklama }], { align: 3, line: 0.15, first: 35.0, below: 10 });

        b.p([{ text: 'NETİCE VE TALEP', bold: true, underline: true }], { above: 20, below: 10 });
        b.p([{ text: t.sonuc }, { text: t.tarih, bold: true }], { align: 3, line: 0.15, first: 35.0, below: 28 });

        b.p([{ text: t.lawyer }], { align: 2 });
        if (t.baroLine) {
            b.p([{ text: t.baroLine }], { align: 2 });
        }
        if (t.adresLines.length > 0) {
            t.adresLines.forEach(function (line, idx) {
                var opts = { align: 2 };
                if (idx === t.adresLines.length - 1) opts.below = 28;
                b.p([{ text: line }], opts);
            });
        } else {
            b.p('', { below: 28 });
        }

        b.p([{ text: t.ekBaslik, bold: true }], { below: 8 });

        var headers = ['No', 'Birim', 'Dosya No', 'Duruşma Tarihi'];
        var rows = t.others.map(function (oh, i) {
            return [String(i + 1), oh.birim, oh.dosyaNo, formatDateTime(oh.tarih)];
        });
        b.table(headers, rows, '35,260,115,130', [1, 0, 1, 1]);

        var xml = b.build();
        var zip = new JSZip();
        zip.file('content.xml', xml);
        return zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
    }

    // ── Downloads ──

    function downloadSingleHearing(hid, key) {
        var group = grouped[key];
        var hearing = null;
        for (var i = 0; i < group.length; i++) {
            if (group[i].id === hid) { hearing = group[i]; break; }
        }
        if (!hearing) return;

        createUDF(hearing, group).then(function (blob) {
            saveAs(blob, makeFilename(hearing));
        }).catch(function (err) {
            showError('UDF oluşturulamadı: ' + err.message);
        });
    }

    function downloadDay(key) {
        var group = grouped[key];
        if (!group || group.length === 0) return;

        var outerZip = new JSZip();
        var done = 0;
        var total = group.length;

        el.progressBar.classList.remove('hidden');
        updateProgress(done, total);

        function processNext(i) {
            if (i >= total) {
                el.progressText.textContent = 'ZIP hazırlanıyor...';
                outerZip.generateAsync({ type: 'blob' }).then(function (blob) {
                    var dateStr = formatDate(group[0].tarih).replace(/\./g, '-');
                    saveAs(blob, 'Mazeret_' + dateStr + '.zip');
                    el.progressBar.classList.add('hidden');
                });
                return;
            }
            createUDF(group[i], group).then(function (blob) {
                outerZip.file(makeFilename(group[i]), blob);
                done++;
                updateProgress(done, total);
                processNext(i + 1);
            }).catch(function () {
                done++;
                updateProgress(done, total);
                processNext(i + 1);
            });
        }

        processNext(0);
    }

    function handleDownloadAll() {
        if (conflictKeys.length === 0) return;

        var allPetitions = [];
        conflictKeys.forEach(function (key) {
            grouped[key].forEach(function (h) {
                allPetitions.push({ hearing: h, group: grouped[key] });
            });
        });

        var outerZip = new JSZip();
        var done = 0;
        var total = allPetitions.length;

        el.progressBar.classList.remove('hidden');
        el.downloadAllBtn.disabled = true;
        updateProgress(done, total);

        function processNext(i) {
            if (i >= total) {
                el.progressText.textContent = 'ZIP hazırlanıyor...';
                outerZip.generateAsync({ type: 'blob' }).then(function (blob) {
                    saveAs(blob, 'mazeret-dilekceleri.zip');
                    el.progressBar.classList.add('hidden');
                    el.downloadAllBtn.disabled = false;
                });
                return;
            }
            var p = allPetitions[i];
            createUDF(p.hearing, p.group).then(function (blob) {
                outerZip.file(makeFilename(p.hearing), blob);
                done++;
                updateProgress(done, total);
                processNext(i + 1);
            }).catch(function () {
                done++;
                updateProgress(done, total);
                processNext(i + 1);
            });
        }

        processNext(0);
    }

    function updateProgress(done, total) {
        var pct = Math.round((done / total) * 100);
        el.progressFill.style.width = pct + '%';
        el.progressText.textContent = done + ' / ' + total + ' dilekçe';
    }

    function makeFilename(hearing) {
        var ds = formatDate(hearing.tarih).replace(/\./g, '-');
        var birim = hearing.birim.substring(0, 40).replace(/[\/\\?%*:|"<>]/g, '_').replace(/\s+/g, '_');
        var dosya = hearing.dosyaNo.replace(/[\/\\?%*:|"<>]/g, '-');
        return 'Mazeret_' + birim + '_' + dosya + '_' + ds + '.udf';
    }

    // ── Reset ──

    function handleReset() {
        allHearings = [];
        grouped = {};
        conflictKeys = [];
        openDays = {};
        el.resultSection.classList.add('hidden');
        el.infoSection.classList.add('hidden');
        el.uploadSection.classList.remove('hidden');
        el.fileInput.value = '';
        el.searchInput.value = '';
        el.searchClear.classList.add('hidden');
        el.daysList.innerHTML = '';
        el.progressBar.classList.add('hidden');
    }

    // ── Utils ──

    function esc(text) {
        var div = document.createElement('div');
        div.textContent = text || '';
        return div.innerHTML;
    }

    function xmlAttr(text) {
        return String(text).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    function showError(msg) {
        el.errorToast.textContent = msg;
        el.errorToast.classList.remove('hidden');
        setTimeout(function () { el.errorToast.classList.add('hidden'); }, 4000);
    }

    document.addEventListener('DOMContentLoaded', init);
})();
