/*
 *  Таблица запчастей — данные берутся из CSV-файла, а не из HTML.
 *  Чтобы добавить/изменить позицию, правьте data/parts.csv (см. README.md).
 */
(function () {
	'use strict';

	/* ------------------------------------------------------------------
	 *  НАСТРОЙКА: откуда брать данные.
	 *
	 *  Основной источник — Google Таблица (правится прямо в браузере,
	 *  сайт подхватывает изменения сразу, без коммита в GitHub):
	 *  https://docs.google.com/spreadsheets/d/1kIhsasNJc_ByQAFeM96MynhXWAjgqAsnrTY4PU_g9JM/edit
	 *
	 *  Таблица должна быть открыта на чтение: Доступ → «Все, у кого есть
	 *  ссылка» → Читатель. Чтобы переехать на другую таблицу, достаточно
	 *  заменить SHEET_ID — это часть ссылки между /d/ и /edit.
	 *
	 *  Если Google недоступен, список берётся из data/parts.csv в репозитории.
	 * ------------------------------------------------------------------ */
	var SHEET_ID = '1kIhsasNJc_ByQAFeM96MynhXWAjgqAsnrTY4PU_g9JM';
	var SHEET_URL = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/gviz/tq?tqx=out:csv';
	var BACKUP_URL = 'data/parts.csv';

	var COLUMNS = ['№', 'Наименование', 'Код/чертеж', 'Количество шт.', 'Фото'];
	var PHOTO_SEPARATOR = /\s*[|;]\s*/;   // как разделены пути к фото внутри ячейки

	/* ---------------------------- CSV ---------------------------------- */

	// Разбор CSV с поддержкой кавычек и переносов строк внутри ячеек.
	function parseCSV(text, delimiter) {
		var rows = [], row = [], field = '', inQuotes = false, i = 0;

		if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);   // BOM из Excel

		while (i < text.length) {
			var ch = text[i];

			if (inQuotes) {
				if (ch === '"') {
					if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
					inQuotes = false; i++; continue;
				}
				field += ch; i++; continue;
			}

			if (ch === '"') { inQuotes = true; i++; continue; }
			if (ch === delimiter) { row.push(field); field = ''; i++; continue; }
			if (ch === '\r') { i++; continue; }
			if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }

			field += ch; i++;
		}
		row.push(field);
		rows.push(row);

		return rows.filter(function (r) {
			return r.some(function (c) { return c.trim() !== ''; });
		});
	}

	// Excel в русской локали сохраняет с ";", остальные — с ",".
	function detectDelimiter(text) {
		var head = text.split('\n')[0];
		return (head.split(';').length > head.split(',').length) ? ';' : ',';
	}

	function toObjects(rows) {
		var header = rows[0].map(function (h) { return h.trim(); });
		return rows.slice(1).map(function (r) {
			var o = {};
			header.forEach(function (h, i) { o[h] = (r[i] || '').trim(); });
			return o;
		});
	}

	/* --------------------------- вспомогательное ------------------------ */

	function esc(s) {
		return String(s).replace(/[&<>"']/g, function (c) {
			return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
		});
	}

	function el(tag, attrs, html) {
		var n = document.createElement(tag);
		if (attrs) Object.keys(attrs).forEach(function (k) { n.setAttribute(k, attrs[k]); });
		if (html != null) n.innerHTML = html;
		return n;
	}

	// Группировка с сохранением порядка появления в файле.
	function groupBy(items, key) {
		var order = [], map = {};
		items.forEach(function (it) {
			var k = it[key] || '';
			if (!(k in map)) { map[k] = []; order.push(k); }
			map[k].push(it);
		});
		return order.map(function (k) { return { name: k, items: map[k] }; });
	}

	/* ------------------------------ рендер ------------------------------ */

	function buildTable(rowsOfCategory) {
		var table = el('table');
		var head = el('tr');
		COLUMNS.forEach(function (c) {
			head.appendChild(el('th', { style: 'text-align: center;', 'class': 'first' }, esc(c)));
		});
		table.appendChild(head);

		groupBy(rowsOfCategory, 'Раздел').forEach(function (section) {
			if (section.name) {
				var tr = el('tr', { 'data-section': '1' });
				tr.appendChild(el('th', {
					colspan: COLUMNS.length,
					style: 'text-align: center;',
					'class': 'first'
				}, esc(section.name) + ':'));
				table.appendChild(tr);
			}

			section.items.forEach(function (item, idx) {
				var photos = (item['Фото'] || '').split(PHOTO_SEPARATOR).filter(Boolean);
				var tr = el('tr');
				tr.dataset.search = [item['Наименование'], item['Код/чертеж'], section.name]
					.join(' ').toLowerCase();

				tr.appendChild(el('td', null, String(idx + 1)));
				tr.appendChild(el('td', null, esc(item['Наименование'] || '')));
				tr.appendChild(el('td', null, esc(item['Код/чертеж'] || '')));
				tr.appendChild(el('td', null, esc(item['Количество'] || '')));

				var photoCell = el('td');
				if (photos.length) {
					var btn = el('button', {
						type: 'button',
						'class': 'btn btn-primary',
						'data-photos': photos.join('|'),
						'data-title': item['Наименование'] || 'Фото'
					}, 'Фото');
					photoCell.appendChild(btn);
				}
				tr.appendChild(photoCell);
				table.appendChild(tr);
			});
		});

		return table;
	}

	function buildTabs(container, categories) {
		var tabs = el('div', { 'class': 'tabs' });
		var panels = [];

		// Сначала все переключатели с подписями — иначе подписи неоткрытых вкладок
		// оказываются под таблицей открытой вкладки.
		categories.forEach(function (cat, i) {
			var id = 'tab-btn-' + (i + 1);
			var input = el('input', { type: 'radio', name: 'tab-btn', id: id, value: '' });
			if (i === 0) input.checked = true;

			var label = el('label', { 'for': id });
			label.appendChild(el('span', { 'class': 'tab-name' }, esc(cat.name)));
			label.appendChild(el('span', { 'class': 'tab-count' }, ''));

			input.addEventListener('change', function () {
				panels.forEach(function (d) { d.classList.remove('is-active'); });
				panels[i].classList.add('is-active');
			});

			tabs.appendChild(input);
			tabs.appendChild(label);
		});

		// Затем сами таблицы — все ниже строки с вкладками.
		categories.forEach(function (cat, i) {
			var panel = el('div', { id: 'content-' + (i + 1) });
			if (i === 0) panel.classList.add('is-active');
			panel.appendChild(buildTable(cat.items));
			panel.appendChild(el('p', { 'class': 'parts-empty', hidden: 'hidden' }, 'Ничего не найдено.'));

			panels.push(panel);
			tabs.appendChild(panel);
		});

		container.innerHTML = '';
		container.appendChild(buildSearchBox(tabs));
		container.appendChild(tabs);
	}

	/* ------------------------------ поиск ------------------------------- */

	function buildSearchBox(tabs) {
		var wrap = el('div', { 'class': 'parts-search' });
		var input = el('input', {
			type: 'search',
			placeholder: 'Поиск по наименованию или коду…',
			'aria-label': 'Поиск по списку запчастей'
		});
		wrap.appendChild(input);

		input.addEventListener('input', function () {
			var q = input.value.trim().toLowerCase();

			tabs.querySelectorAll(':scope > div').forEach(function (panel, i) {
				var found = 0;

				panel.querySelectorAll('tr[data-search]').forEach(function (tr) {
					var hit = !q || tr.dataset.search.indexOf(q) !== -1;
					tr.hidden = !hit;
					if (hit) found++;
				});

				// подзаголовки разделов прячем, если в разделе ничего не осталось
				var visibleInSection = 0, lastSection = null;
				panel.querySelectorAll('tr').forEach(function (tr) {
					if (tr.dataset.section) {
						if (lastSection) lastSection.hidden = visibleInSection === 0;
						lastSection = tr; visibleInSection = 0;
					} else if (!tr.hidden && tr.dataset.search) {
						visibleInSection++;
					}
				});
				if (lastSection) lastSection.hidden = visibleInSection === 0;

				panel.querySelector('table').hidden = (found === 0);
				panel.querySelector('.parts-empty').hidden = (found !== 0);

				var counter = tabs.querySelectorAll('label')[i].querySelector('.tab-count');
				counter.textContent = q ? ' (' + found + ')' : '';
			});
		});

		return wrap;
	}

	/* ------------------------- модальное окно фото ---------------------- */

	function setupPhotoModal(container) {
		var modal = el('div', {
			'class': 'modal fade',
			id: 'partsPhotoModal',
			tabindex: '-1',
			'aria-hidden': 'true'
		},
			'<div class="modal-dialog modal-dialog-centered modal-dialog-scrollable">' +
				'<div class="modal-content">' +
					'<div class="modal-header">' +
						'<h5 class="modal-title"></h5>' +
						'<button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Закрыть"></button>' +
					'</div>' +
					'<div class="modal-body"></div>' +
					'<div class="modal-footer">' +
						'<button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Закрыть</button>' +
					'</div>' +
				'</div>' +
			'</div>');
		document.body.appendChild(modal);

		var bsModal = new bootstrap.Modal(modal, { backdrop: 'static', keyboard: false });

		container.addEventListener('click', function (e) {
			var btn = e.target.closest('button[data-photos]');
			if (!btn) return;

			modal.querySelector('.modal-title').textContent = btn.dataset.title;
			modal.querySelector('.modal-body').innerHTML = btn.dataset.photos
				.split('|')
				.map(function (src) {
					return '<img src="' + esc(src) + '" width="100%" height="100%" loading="lazy" alt="">';
				})
				.join('<br>');

			bsModal.show();
		});
	}

	/* ------------------------------ запуск ------------------------------ */

	function showError(container, message) {
		container.innerHTML = '';
		container.appendChild(el('p', { 'class': 'parts-error' }, message));
	}

	function showNotice(container, message) {
		container.insertBefore(el('p', { 'class': 'parts-notice' }, message), container.firstChild);
	}

	// Свежая копия в обход кэша браузера; у ссылки Google уже есть «?».
	function noCache(url) {
		return url + (url.indexOf('?') === -1 ? '?' : '&') + '_=' + Date.now();
	}

	function load(url) {
		return fetch(noCache(url))
			.then(function (r) {
				if (!r.ok) throw new Error('HTTP ' + r.status);
				return r.text();
			})
			.then(function (text) {
				var items = toObjects(parseCSV(text, detectDelimiter(text)));
				if (!items.length) throw new Error('таблица пуста');
				return items;
			});
	}

	function render(container, items) {
		buildTabs(container, groupBy(items, 'Категория'));
		setupPhotoModal(container);
	}

	document.addEventListener('DOMContentLoaded', function () {
		var container = document.getElementById('parts');
		if (!container) return;

		load(SHEET_URL)
			.then(function (items) {
				render(container, items);
			})
			.catch(function (sheetErr) {
				// Google не ответил — показываем копию из репозитория.
				return load(BACKUP_URL)
					.then(function (items) {
						render(container, items);
						showNotice(container,
							'Google Таблица сейчас недоступна (' + esc(sheetErr.message) + '), ' +
							'показан сохранённый список — возможно, без последних изменений.');
					})
					.catch(function (backupErr) {
						showError(container,
							'Не удалось загрузить список запчастей ' +
							'(Google: ' + esc(sheetErr.message) + ', ' +
							'резервная копия: ' + esc(backupErr.message) + '). ' +
							'Если вы открыли страницу двойным щелчком по файлу — так браузер ' +
							'не даёт читать данные; откройте сайт по адресу ' +
							'<a href="https://imnaa.github.io/zipdetals/">imnaa.github.io/zipdetals</a>.');
					});
			});
	});
})();
